# -*- coding: utf-8 -*-
"""
Modelos de deteccion de anomalias - version final (v12) de la monografia.

Lee la base anonimizada Datos_PowerBI.xlsx (hojas Ordenes y Diferencias) y genera:
  - Tabla 4-3: comparacion de 14 enfoques (reglas, 10 no supervisados, 3 supervisados)
  - Tabla 4-4: ordenes observadas capturadas segun la carga de revision
  - Tabla 4-5: perfil de los grupos de K-medias
  - Tablas de sensibilidad (etiqueta e hiperparametros) y de diagnostico
  - Figuras 4-4 a 4-8
Uso:  python modelos_v12.py Datos_PowerBI.xlsx carpeta_salida/

Decisiones de diseno (apartado 3.6 de la monografia):
  * El tecnico y la zona NO se usan como variables (el tecnico, codificado como entero,
    producia un efecto artificial; la zona es urbana en 469 de 471 ordenes).
  * El tipo de servicio entra en los modelos supervisados como variables indicadoras (one-hot).
  * Los no supervisados se ajustan sobre las 471 ordenes SIN la etiqueta y se evaluan con ella.
  * Los supervisados se evaluan con validacion cruzada estratificada 5 x 10 (probabilidades
    fuera de muestra promediadas por orden).
"""
import sys, os, json, warnings
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy import stats
from scipy.cluster.hierarchy import linkage, fcluster, dendrogram
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import make_pipeline
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier, IsolationForest
from sklearn.neighbors import LocalOutlierFactor, NearestNeighbors
from sklearn.svm import OneClassSVM
from sklearn.cluster import KMeans, DBSCAN
from sklearn.mixture import GaussianMixture
from sklearn.covariance import MinCovDet
from sklearn.decomposition import PCA
from sklearn.neural_network import MLPRegressor
from sklearn.metrics import roc_auc_score, average_precision_score, roc_curve, precision_recall_curve, silhouette_score
from sklearn.model_selection import StratifiedKFold
from sklearn.inspection import permutation_importance
from xgboost import XGBClassifier

warnings.filterwarnings("ignore")
RUTA, OUT = sys.argv[1], sys.argv[2]
os.makedirs(OUT + "/figuras", exist_ok=True)
RES = {}

# ---------------------------------------------------------------- datos
o = pd.read_excel(RUTA, sheet_name="Ordenes")
dif = pd.read_excel(RUTA, sheet_name="Diferencias")
MAT = list(o.columns[o.columns.get_loc("ANILLAS_GUIA"): o.columns.get_loc("ADAPTADOR_ACDC") + 1])
NUM = MAT + ["postes", "cantidad_iptv"]
y = o["observada_operador"].astype(int).values
n, pos = len(y), int(y.sum())
base = pos / n
regla = o["puntaje_riesgo"].fillna(0).values
servicio = o["tipo_servicio"].astype(str)

# z por tipo de servicio (21 variables); materiales sin uso -> z = 0
Z = o[NUM].astype(float).copy()
for _, g in o.groupby("tipo_servicio"):
    mu, sd = o.loc[g.index, NUM].mean(), o.loc[g.index, NUM].std().replace(0, np.nan)
    Z.loc[g.index] = ((o.loc[g.index, NUM] - mu) / sd).fillna(0)
var_ok = [c for c in NUM if Z[c].std() > 0]          # quita columnas constantes (4 materiales sin uso)
Ze = Z[var_ok].values
RES["variables_no_supervisados"] = len(var_ok)

# variables supervisadas: one-hot del tipo de servicio + 21 numericas (sin tecnico ni zona)
Xs = pd.concat([pd.get_dummies(servicio, prefix="serv", drop_first=True).astype(float),
                o[NUM].astype(float)], axis=1)
RES["variables_supervisados"] = Xs.shape[1]

# ---------------------------------------------------------------- utilidades
rng0 = np.random.default_rng(12)
I1, I0 = np.where(y == 1)[0], np.where(y == 0)[0]
BOOT = [np.r_[rng0.choice(I1, len(I1)), rng0.choice(I0, len(I0))] for _ in range(2000)]

def ic_auc(s, yy=y, boot=BOOT):
    v = [roc_auc_score(yy[b], s[b]) for b in boot]
    return np.percentile(v, [2.5, 97.5])

def p_perm_fijo(s, yy=y, B=5000, seed=3):
    """Permutacion de la etiqueta con puntajes fijos (valido para metodos que no usan la etiqueta)."""
    r = np.random.default_rng(seed); a = roc_auc_score(yy, s)
    nulo = np.array([roc_auc_score(r.permutation(yy), s) for _ in range(B)])
    return (1 + (nulo >= a).sum()) / (B + 1)

def capturadas(s, k, yy=y, seed=0):
    r = np.random.default_rng(seed)
    orden = np.lexsort((r.random(len(s)), -s))     # desempate aleatorio
    return int(yy[orden[:k]].sum())

# ---------------------------------------------------------------- 1. diagnostico
diag = {}
# asimetria por material y servicio (supuesto del puntaje z)
asim = []
for sv, g in o.groupby("tipo_servicio"):
    for c in NUM:
        x = g[c].astype(float)
        if x.std() > 0 and len(x) >= 8:
            asim.append({"servicio": sv, "variable": c, "asimetria": stats.skew(x), "pct_ceros": (x == 0).mean()})
asim = pd.DataFrame(asim)
diag["combinaciones_evaluadas"] = len(asim)
diag["asimetria_mayor_1"] = int((asim["asimetria"].abs() > 1).sum())
diag["asimetria_mediana"] = float(asim["asimetria"].abs().median())
sw = [stats.shapiro(g["CABLE_DROP"].astype(float))[1] for _, g in o.groupby("tipo_servicio") if g["CABLE_DROP"].std() > 0]
diag["shapiro_drop_p_max"] = float(max(sw))
sh = [stats.shapiro(g[c].astype(float))[1] for c in ["CABLE_DROP", "CABLE_UTP", "TENSORES", "postes"]
      for _, g in o.groupby("tipo_servicio") if g[c].std() > 0 and len(g) >= 8]
diag["shapiro_rechazos"] = int(sum(p < 0.05 for p in sh)); diag["shapiro_pruebas"] = len(sh)
asim.round(3).to_csv(OUT + "/Tabla_A-4_asimetria.csv", index=False)

# observadas por tipo de servicio y tipo de trabajo, con IC de Wilson
def wilson(k, m, z=1.96):
    if m == 0: return (np.nan, np.nan)
    p = k / m; d = 1 + z * z / m; c = (p + z * z / (2 * m)) / d
    h = z * np.sqrt(p * (1 - p) / m + z * z / (4 * m * m)) / d
    return (c - h, c + h)
filas = []
for col in ["tipo_servicio", "tipo_trabajo"]:
    for v, g in o.groupby(col):
        k, m = int(g["observada_operador"].sum()), len(g); lo, hi = wilson(k, m)
        filas.append({"agrupacion": col, "categoria": v, "ordenes": m, "observadas": k,
                      "tasa": round(k / m, 3), "ic95_inf": round(lo, 3), "ic95_sup": round(hi, 3)})
obs_grupo = pd.DataFrame(filas); obs_grupo.to_csv(OUT + "/Tabla_4-1b_observadas_por_grupo.csv", index=False)
m2 = servicio == "ONT + 2 STB"
diag["fisher_ont2stb_p"] = float(stats.fisher_exact([[y[m2].sum(), (1 - y[m2]).sum()], [y[~m2].sum(), (1 - y[~m2]).sum()]])[1])
# chi-cuadrado global servicio x observada por permutacion (celdas pequenas)
def chi(lbl):
    t = pd.crosstab(servicio, lbl); return stats.chi2_contingency(t)[0]
c0 = chi(y); r = np.random.default_rng(5)
diag["servicio_chi2_perm_p"] = float((1 + sum(chi(r.permutation(y)) >= c0 for _ in range(5000))) / 5001)
# tecnico x observada (prueba por permutacion)
tec = o["tecnico1_cod"].astype(str)
def chi_t(lbl): return stats.chi2_contingency(pd.crosstab(tec, lbl))[0]
c0 = chi_t(y)
diag["tecnico_chi2_perm_p"] = float((1 + sum(chi_t(r.permutation(y)) >= c0 for _ in range(5000))) / 5001)
RES["diagnostico"] = diag

# ---------------------------------------------------------------- 2. seleccion del numero de grupos
ks = list(range(2, 11))
inercia, silu = [], []
for k in ks:
    km = KMeans(k, n_init=20, random_state=0).fit(Ze)
    inercia.append(km.inertia_); silu.append(silhouette_score(Ze, km.labels_))
K = ks[int(np.argmax(silu))]
bic = [GaussianMixture(c, covariance_type="diag", reg_covar=1e-3, n_init=5, random_state=0).fit(Ze).bic(Ze) for c in range(1, 11)]
G = int(np.argmin(bic)) + 1
Lw = linkage(Ze, method="ward")
sil_ward = {k: silhouette_score(Ze, fcluster(Lw, k, "maxclust")) for k in ks}
MS = 10
kd = np.sort(NearestNeighbors(n_neighbors=MS).fit(Ze).kneighbors(Ze)[0][:, -1])
# codo de la curva k-distancia: punto mas alejado de la recta entre extremos
xk = np.arange(len(kd)); p1, p2 = np.array([0, kd[0]]), np.array([len(kd) - 1, kd[-1]])
v = p2 - p1; w = np.c_[xk, kd] - p1
dist = np.abs(v[0] * w[:, 1] - v[1] * w[:, 0]) / np.linalg.norm(v)
EPS = float(kd[int(np.argmax(dist))])
RES["seleccion"] = {"k_kmeans": K, "silueta_k": dict(zip(ks, np.round(silu, 3))), "componentes_gmm": G,
                    "bic": [round(b, 1) for b in bic], "silueta_ward": {k: round(v, 3) for k, v in sil_ward.items()},
                    "dbscan_min_samples": MS, "dbscan_eps": round(EPS, 2)}

# ---------------------------------------------------------------- 3. no supervisados (sin etiqueta)
S = {}
S["Puntaje z maximo"] = Z.abs().max(axis=1).values
try:
    mcd = MinCovDet(random_state=0, support_fraction=0.9).fit(Ze); S["Mahalanobis robusta (MCD)"] = mcd.mahalanobis(Ze)
except Exception:
    cov = np.cov(Ze, rowvar=False); ip = np.linalg.pinv(cov); d = Ze - Ze.mean(0)
    S["Mahalanobis robusta (MCD)"] = np.einsum("ij,jk,ik->i", d, ip, d)
lof = LocalOutlierFactor(n_neighbors=20).fit(Ze); S["LOF"] = -lof.negative_outlier_factor_
S["Isolation Forest"] = -IsolationForest(n_estimators=300, random_state=0).fit(Ze).score_samples(Ze)
S["One-Class SVM"] = -OneClassSVM(kernel="rbf", gamma="scale", nu=0.05).fit(Ze).decision_function(Ze)
km = KMeans(K, n_init=20, random_state=0).fit(Ze)
S["K-medias (distancia al centro)"] = np.linalg.norm(Ze - km.cluster_centers_[km.labels_], axis=1)
lab_w = fcluster(Lw, K, "maxclust")
cent_w = {c: Ze[lab_w == c].mean(0) for c in np.unique(lab_w)}
S["Jerarquico de Ward (distancia al centro)"] = np.array([np.linalg.norm(Ze[i] - cent_w[lab_w[i]]) for i in range(n)])
db = DBSCAN(eps=EPS, min_samples=MS).fit(Ze)
S["DBSCAN (distancia k-vecino)"] = NearestNeighbors(n_neighbors=MS).fit(Ze).kneighbors(Ze)[0][:, -1]
gmm = GaussianMixture(G, covariance_type="diag", reg_covar=1e-3, n_init=5, random_state=0).fit(Ze)
S["Mezcla gaussiana (GMM)"] = -gmm.score_samples(Ze)
ae_err = np.zeros(n); Zs = StandardScaler().fit_transform(Ze)
for s in range(5):
    ae = MLPRegressor(hidden_layer_sizes=(8, 3, 8), activation="tanh", max_iter=3000, alpha=1e-3,
                      random_state=s).fit(Zs, Zs)
    ae_err += ((ae.predict(Zs) - Zs) ** 2).mean(1) / 5
S["Autocodificador"] = ae_err
noise = db.labels_ == -1
RES["dbscan"] = {"grupos": int(len(set(db.labels_)) - (1 if noise.any() else 0)), "ruido": int(noise.sum()),
                 "observadas_en_ruido": int(y[noise].sum())}

# ---------------------------------------------------------------- 4. supervisados (CV 5 x 10)
peso = (n - pos) / pos
SUP = {
    "Regresion logistica": lambda: make_pipeline(StandardScaler(), LogisticRegression(class_weight="balanced", C=0.1, max_iter=5000)),
    "Random Forest": lambda: RandomForestClassifier(n_estimators=300, min_samples_leaf=3, class_weight="balanced_subsample", random_state=0, n_jobs=-1),
    "XGBoost": lambda: XGBClassifier(n_estimators=200, max_depth=2, learning_rate=0.05, subsample=0.8, colsample_bytree=0.8,
                                     scale_pos_weight=peso, eval_metric="logloss", random_state=0, n_jobs=4),
}
X = Xs.values
def oof(fab, yy, seed):
    p = np.zeros(n)
    for tr, te in StratifiedKFold(5, shuffle=True, random_state=seed).split(X, yy):
        p[te] = fab().fit(X[tr], yy[tr]).predict_proba(X[te])[:, 1]
    return p
auc_rep = {}
for nom, fab in SUP.items():
    P = np.array([oof(fab, y, s) for s in range(10)])
    S[nom] = P.mean(0)
    auc_rep[nom] = [roc_auc_score(y, p) for p in P]
S["Indice de riesgo (reglas)"] = regla

# ---------------------------------------------------------------- 5. evaluacion
FAM = {"Indice de riesgo (reglas)": "Reglas", "Puntaje z maximo": "Estadistico", "Mahalanobis robusta (MCD)": "Estadistico",
       "LOF": "Densidad", "Isolation Forest": "Aislamiento", "One-Class SVM": "Frontera",
       "K-medias (distancia al centro)": "Agrupamiento", "Jerarquico de Ward (distancia al centro)": "Agrupamiento",
       "DBSCAN (distancia k-vecino)": "Agrupamiento", "Mezcla gaussiana (GMM)": "Agrupamiento",
       "Autocodificador": "Reconstruccion", "Regresion logistica": "Supervisado", "Random Forest": "Supervisado", "XGBoost": "Supervisado"}
PERM_SUP_B = int(os.environ.get("PERM_SUP_B", "200"))   # 0 = modo rapido (sin permutaciones de supervisados)
filas = []
for nom, s in S.items():
    a = roc_auc_score(y, s); lo, hi = ic_auc(s); ap = average_precision_score(y, s)
    if nom in SUP:
        if PERM_SUP_B == 0:
            p = np.nan; extra = f"{np.mean(auc_rep[nom]):.2f} +- {np.std(auc_rep[nom]):.2f}"
            filas.append({"enfoque": nom, "familia": FAM[nom], "AUC": a, "IC95_inf": lo, "IC95_sup": hi, "AP": ap,
                          "AP_tasa_base": ap / base, "p_permutacion": p, "AUC_por_repeticion": extra,
                          "capt_60": capturadas(s, 60), "capt_119": capturadas(s, 119), "capt_150": capturadas(s, 150)})
            continue
        a0 = roc_auc_score(y, oof(SUP[nom], y, 0)); r = np.random.default_rng(7); nulo = []
        for b in range(PERM_SUP_B):
            yp = r.permutation(y); nulo.append(roc_auc_score(yp, oof(SUP[nom], yp, b + 1)))
        p = (1 + (np.array(nulo) >= a0).sum()) / (PERM_SUP_B + 1)
        extra = f"{np.mean(auc_rep[nom]):.2f} +- {np.std(auc_rep[nom]):.2f}"
    else:
        p = p_perm_fijo(s); extra = ""
    filas.append({"enfoque": nom, "familia": FAM[nom], "AUC": a, "IC95_inf": lo, "IC95_sup": hi, "AP": ap,
                  "AP_tasa_base": ap / base, "p_permutacion": p, "AUC_por_repeticion": extra,
                  "capt_60": capturadas(s, 60), "capt_119": capturadas(s, 119), "capt_150": capturadas(s, 150)})
T = pd.DataFrame(filas).sort_values("AUC", ascending=False).reset_index(drop=True)
# Holm
pv = T["p_permutacion"].fillna(1.0).values; orden = np.argsort(pv); m = len(pv); adj = np.empty(m); run = 0
for i, j in enumerate(orden):
    run = max(run, min(1, (m - i) * pv[j])); adj[j] = run
T["p_Holm"] = adj
T.round(4).to_csv(OUT + "/Tabla_4-3_comparacion_enfoques.csv", index=False)
RES["tasa_base"] = base

# carga de revision: promedio sobre 50 desempates/semillas para los metodos principales
carga = []
for nom in T["enfoque"]:
    for k in (60, 119, 150):
        c = [capturadas(S[nom], k, seed=s) for s in range(50)]
        carga.append({"enfoque": nom, "k": k, "pct": round(100 * k / n, 1), "capturadas_media": np.mean(c),
                      "min": min(c), "max": max(c), "exhaustividad": np.mean(c) / pos, "precision": np.mean(c) / k,
                      "esperado_azar": k * base})
carga = pd.DataFrame(carga); carga.round(3).to_csv(OUT + "/Tabla_4-4_carga_de_revision.csv", index=False)

# matriz de confusion con k = 119 para los tres mejores no supervisados, LR y reglas
conf = []
for nom in list(T["enfoque"][:3]) + ["Regresion logistica", "Indice de riesgo (reglas)"]:
    rr = np.random.default_rng(0); idx = np.lexsort((rr.random(n), -S[nom]))[:119]; mk = np.zeros(n, bool); mk[idx] = True
    conf.append({"enfoque": nom, "VP": int((mk & (y == 1)).sum()), "FP": int((mk & (y == 0)).sum()),
                 "FN": int((~mk & (y == 1)).sum()), "VN": int((~mk & (y == 0)).sum())})
pd.DataFrame(conf).drop_duplicates("enfoque").to_csv(OUT + "/Tabla_4-6_matriz_confusion_k119.csv", index=False)

# diferencia de capturadas (k = 119) mejor metodo vs reglas, bootstrap
mejor = T["enfoque"][0]
def top(s, k=119):
    rr = np.random.default_rng(0); mk = np.zeros(n, bool); mk[np.lexsort((rr.random(n), -s))[:k]] = True; return mk
mb, mr = top(S[mejor]), regla >= 1
rb = np.random.default_rng(9); dd = []
for _ in range(5000):
    i = rb.choice(n, n); dd.append(y[i][mb[i]].sum() - y[i][mr[i]].sum())
RES["dif_mejor_vs_reglas"] = {"mejor": mejor, "dif": int(y[mb].sum() - y[mr].sum()), "ic95": list(np.percentile(dd, [2.5, 97.5]))}

# ---------------------------------------------------------------- 6. perfiles de los grupos de K-medias
prof = []
for c in range(K):
    g = km.labels_ == c
    fila = {"grupo": c + 1, "ordenes": int(g.sum()), "observadas": int(y[g].sum()), "tasa_obs": y[g].mean(),
            "servicio_dominante": servicio[g].value_counts().index[0],
            "pct_servicio_dominante": servicio[g].value_counts(normalize=True).iloc[0]}
    for v in ["CABLE_DROP", "CABLE_DUCTO", "TENSORES", "postes", "ANILLAS_GUIA", "CABLE_UTP", "CONECTOR_RJ45", "TERMINAL_SC"]:
        fila["media_" + v] = o.loc[g, v].mean()
    zc = pd.Series(km.cluster_centers_[c], index=var_ok)
    fila["variables_que_lo_definen"] = ", ".join(f"{k}({v:+.1f})" for k, v in zc.reindex(zc.abs().sort_values(ascending=False).index)[:3].items())
    prof.append(fila)
prof = pd.DataFrame(prof); prof.round(2).to_csv(OUT + "/Tabla_4-5_perfil_grupos_kmedias.csv", index=False)
# segundo nivel: sin el grupo pequeno de ordenes extremas, se busca una segmentacion operativa
g_ext = np.argmin(np.bincount(km.labels_)); ext = km.labels_ == g_ext
Rz = Ze[~ext]; sil2 = {}
for k in range(2, 9):
    lab = KMeans(k, n_init=20, random_state=0).fit(Rz).labels_
    if np.bincount(lab).min() >= 0.05 * len(Rz):
        sil2[k] = silhouette_score(Rz, lab)
K2 = max(sil2, key=sil2.get)
km2 = KMeans(K2, n_init=20, random_state=0).fit(Rz)
RES["segundo_nivel"] = {"extremas": int(ext.sum()), "k": K2, "silueta": round(sil2[K2], 3),
                        "silueta_por_k_validos": {k: round(v, 3) for k, v in sil2.items()}}
seg = np.full(n, -1); seg[np.where(~ext)[0]] = km2.labels_
nombres_seg = {-1: "A (extremas)"} | {c: f"B{c+1}" for c in range(K2)}
prof2 = []
for c in [-1] + list(range(K2)):
    g = seg == c
    fila = {"segmento": nombres_seg[c], "ordenes": int(g.sum()), "observadas": int(y[g].sum()), "tasa_obs": y[g].mean(),
            "servicio_dominante": servicio[g].value_counts().index[0], "pct_servicio_dominante": servicio[g].value_counts(normalize=True).iloc[0]}
    for v in ["CABLE_DROP", "TENSORES", "postes", "CABLE_DUCTO", "ANILLAS_GUIA", "CINTA_ACERO", "CABLE_UTP", "TERMINAL_SC"]:
        fila["media_" + v] = o.loc[g, v].mean()
    zc = Z.loc[g, var_ok].mean()
    fila["z_medio_rasgos"] = ", ".join(f"{k}({v:+.1f})" for k, v in zc.reindex(zc.abs().sort_values(ascending=False).index)[:4].items())
    prof2.append(fila)
prof2 = pd.DataFrame(prof2); prof2.round(2).to_csv(OUT + "/Tabla_4-5_perfil_segmentos.csv", index=False)
tt = pd.crosstab(seg[seg >= 0], y[seg >= 0]); RES["segmentos_vs_observada_fisher_p"] = float(stats.fisher_exact(tt.values)[1]) if tt.shape == (2, 2) else float(stats.chi2_contingency(tt)[1])
t = pd.crosstab(km.labels_, y)
RES["kmeans_grupo_vs_observada_p"] = float(stats.chi2_contingency(t)[1]) if t.shape[1] > 1 else None
# estabilidad de los grupos: indice de Rand ajustado entre semillas y con Ward
from sklearn.metrics import adjusted_rand_score
ari = [adjusted_rand_score(km.labels_, KMeans(K, n_init=20, random_state=s).fit(Ze).labels_) for s in range(1, 11)]
RES["kmeans_ari_semillas_min"] = float(min(ari)); RES["kmeans_ari_vs_ward"] = float(adjusted_rand_score(km.labels_, lab_w))

# ---------------------------------------------------------------- 7. importancia de variables (supervisados)
imp = {}
lr = SUP["Regresion logistica"]().fit(X, y)
imp["coef_LR"] = pd.Series(lr[-1].coef_[0], index=Xs.columns).sort_values(key=abs, ascending=False)
pi = permutation_importance(SUP["Random Forest"]().fit(X, y), X, y, scoring="roc_auc", n_repeats=30, random_state=0)
imp["perm_RF"] = pd.Series(pi.importances_mean, index=Xs.columns).sort_values(ascending=False)
pd.DataFrame({"coef_regresion_logistica": imp["coef_LR"].round(3)}).to_csv(OUT + "/coeficientes_regresion_logistica.csv")
pd.DataFrame({"importancia_permutacion_RF": imp["perm_RF"].round(4)}).to_csv(OUT + "/importancia_permutacion_RF.csv")
# contribucion por variable en LOF: z medio de las observadas
RES["z_medio_observadas"] = Z[y == 1].abs().mean().sort_values(ascending=False).round(2).head(6).to_dict()

# ---------------------------------------------------------------- 8. sensibilidad de la etiqueta
sust = set(dif.loc[dif["tipo"] != "REDONDEO", "solicitud"])
y13 = y.copy(); y13[(o["aviso_iptv"] == 1).values & (o["con_diferencia_sistema"] == 0).values & (y == 1)] = 0
yerr = ((y == 1) | o["solicitud"].isin(sust).values).astype(int)
pdf = (o["con_pdf"] == 1).values
ypdf = (o["solicitud"].isin(sust).values).astype(int)
sens = []
for nom in T["enfoque"]:
    s = S[nom]
    sens.append({"enfoque": nom, "AUC_14_observadas": roc_auc_score(y, s),
                 "AUC_13_sin_lectura_parcial": roc_auc_score(y13, s),
                 "AUC_error_real_20": roc_auc_score(yerr, s),
                 "AUC_dentro_72_PDF": roc_auc_score(ypdf[pdf], s[pdf])})
sens = pd.DataFrame(sens); sens.round(3).to_csv(OUT + "/Tabla_A-5_sensibilidad_etiqueta.csv", index=False)
RES["etiquetas"] = {"pos_13": int(y13.sum()), "pos_error_real": int(yerr.sum()), "pdf_n": int(pdf.sum()), "pdf_pos": int(ypdf[pdf].sum()),
                    "no_obs_con_dif": int(((o["con_pdf"] == 1) & (o["con_diferencia_sistema"] == 1) & (o["observada_operador"] == 0)).sum()),
                    "no_obs_con_pdf": int(((o["con_pdf"] == 1) & (o["observada_operador"] == 0)).sum()),
                    "no_obs_dif_sustantiva": int(len(sust - set(o.loc[y == 1, "solicitud"])))}

# ---------------------------------------------------------------- 9. sensibilidad de hiperparametros
hp = []
def add(mod, par, val, s):
    hp.append({"modelo": mod, "hiperparametro": par, "valor": val, "AUC": roc_auc_score(y, s), "AP": average_precision_score(y, s)})
for k in (10, 20, 30, 50): add("LOF", "vecinos k", k, -LocalOutlierFactor(n_neighbors=k).fit(Ze).negative_outlier_factor_)
for nu in (0.03, 0.05, 0.10): add("One-Class SVM", "nu", nu, -OneClassSVM(nu=nu, gamma="scale").fit(Ze).decision_function(Ze))
for t in (100, 300, 500): add("Isolation Forest", "arboles", t, -IsolationForest(n_estimators=t, random_state=0).fit(Ze).score_samples(Ze))
for k in sorted({max(2, K - 1), K, K + 1, K + 2}):
    mk = KMeans(k, n_init=20, random_state=0).fit(Ze); add("K-medias", "grupos k", k, np.linalg.norm(Ze - mk.cluster_centers_[mk.labels_], axis=1))
for g in sorted({max(1, G - 1), G, G + 1}):
    add("GMM", "componentes", g, -GaussianMixture(g, covariance_type="diag", reg_covar=1e-3, n_init=5, random_state=0).fit(Ze).score_samples(Ze))
for ms in (5, 10, 20): add("DBSCAN", "min_samples", ms, NearestNeighbors(n_neighbors=ms).fit(Ze).kneighbors(Ze)[0][:, -1])
for c in (0.01, 0.1, 1.0):
    f = lambda c=c: make_pipeline(StandardScaler(), LogisticRegression(class_weight="balanced", C=c, max_iter=5000))
    add("Regresion logistica", "C", c, np.mean([oof(f, y, s) for s in range(10)], 0))
pd.DataFrame(hp).round(3).to_csv(OUT + "/Tabla_A-3_sensibilidad_hiperparametros.csv", index=False)

# version con tecnico codificado como entero (la de la v11), para documentar el artefacto
from sklearn.preprocessing import LabelEncoder
Xv = pd.concat([pd.DataFrame({c: LabelEncoder().fit_transform(o[c].astype(str)) for c in ["tipo_servicio", "zona", "tecnico1_cod"]}),
                o[NUM].astype(float)], axis=1).values
def oof_v(yy, seed):
    p = np.zeros(n)
    for tr, te in StratifiedKFold(5, shuffle=True, random_state=seed).split(Xv, yy):
        p[te] = SUP["Regresion logistica"]().fit(Xv[tr], yy[tr]).predict_proba(Xv[te])[:, 1]
    return p
pv11 = np.mean([oof_v(y, s) for s in range(10)], 0)
RES["v11_tecnico_entero"] = {"AUC": roc_auc_score(y, pv11), "capt_119": capturadas(pv11, 119)}
Xoh = pd.concat([Xs, pd.get_dummies(tec, prefix="tec", drop_first=True).astype(float)], axis=1).values
def oof_oh(seed):
    p = np.zeros(n)
    for tr, te in StratifiedKFold(5, shuffle=True, random_state=seed).split(Xoh, y):
        p[te] = SUP["Regresion logistica"]().fit(Xoh[tr], y[tr]).predict_proba(Xoh[te])[:, 1]
    return p
poh = np.mean([oof_oh(s) for s in range(10)], 0)
RES["tecnico_onehot"] = {"AUC": roc_auc_score(y, poh), "capt_119": capturadas(poh, 119)}

# ---------------------------------------------------------------- 10. prioridad para el tablero y descarga de PDF
mejor_ns = T[T["familia"] != "Supervisado"]["enfoque"].iloc[0] if T["familia"].iloc[0] != "Reglas" else T["enfoque"].iloc[1]
pri = pd.Series(S[mejor_ns]).rank(ascending=False, method="first").astype(int)
pd.DataFrame({"solicitud": o["solicitud"], "puntaje_" + mejor_ns: np.round(S[mejor_ns], 4), "prioridad": pri}).to_csv(OUT + "/prioridad_ordenes.csv", index=False)
RES["metodo_prioridad"] = mejor_ns

# ---------------------------------------------------------------- 11. figuras
AZ, NA, AQ, GR, TX, MU = "#2a78d6", "#eb6834", "#1baf7a", "#9aa3ad", "#0b0b0b", "#52514e"
COLF = {"Reglas": NA, "Supervisado": AZ, "Estadistico": AQ, "Densidad": AQ, "Aislamiento": AQ, "Frontera": AQ,
        "Agrupamiento": AQ, "Reconstruccion": AQ}
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 9, "axes.spines.top": False, "axes.spines.right": False,
                     "axes.edgecolor": "#b5b4ae", "axes.labelcolor": TX, "xtick.color": MU, "ytick.color": MU})
corto = {"Indice de riesgo (reglas)": "Índice de reglas", "Puntaje z maximo": "Puntaje z máximo",
         "Mahalanobis robusta (MCD)": "Mahalanobis robusta", "K-medias (distancia al centro)": "K-medias",
         "Jerarquico de Ward (distancia al centro)": "Jerárquico (Ward)", "DBSCAN (distancia k-vecino)": "DBSCAN",
         "Mezcla gaussiana (GMM)": "Mezcla gaussiana", "Regresion logistica": "Regresión logística"}
nm = lambda s: corto.get(s, s)

# Figura 4-4: AUC con IC 95 % y AP
Tp = T.sort_values("AUC")
fig, ax = plt.subplots(1, 2, figsize=(10, 5.6), dpi=220, gridspec_kw={"wspace": 0.55})
yy_ = np.arange(len(Tp)); cols = [COLF[f] for f in Tp["familia"]]
ax[0].hlines(yy_, Tp["IC95_inf"], Tp["IC95_sup"], color=GR, lw=2)
ax[0].scatter(Tp["AUC"], yy_, c=cols, s=46, zorder=3, edgecolor="white", lw=1.2)
ax[0].axvline(0.5, ls="--", color=MU, lw=1); ax[0].set_yticks(yy_); ax[0].set_yticklabels([nm(e) for e in Tp["enfoque"]])
ax[0].set_xlabel("AUC-ROC (punto) e IC 95 % bootstrap (línea)"); ax[0].set_xlim(0.25, 1.0)
ax[0].text(0.505, len(Tp) - 0.4, "azar", color=MU, fontsize=8)
ax[1].barh(yy_, Tp["AP"], color=cols, height=0.6)
ax[1].axvline(base, ls="--", color=MU, lw=1); ax[1].text(base + 0.003, len(Tp) - 0.4, f"azar = {base:.3f}", color=MU, fontsize=8)
ax[1].set_yticks(yy_); ax[1].set_yticklabels([nm(e) for e in Tp["enfoque"]]); ax[1].set_xlabel("Precisión promedio (AP)")
for i, v in enumerate(Tp["AP"]): ax[1].text(v + 0.003, i, f"{v:.3f}", va="center", fontsize=7.5, color=MU)
from matplotlib.lines import Line2D
ax[0].legend(handles=[Line2D([], [], marker="o", ls="", color=c, label=l) for c, l in [(NA, "Reglas"), (AQ, "No supervisado"), (AZ, "Supervisado")]],
             frameon=False, loc="lower right", fontsize=8)
fig.savefig(OUT + "/figuras/Figura_4-4_comparacion_enfoques.png", bbox_inches="tight"); plt.close(fig)

# Figura 4-5: curvas ROC y PR de los 4 mejores + reglas
sel = list(T["enfoque"][:4]); sel += ["Indice de riesgo (reglas)"] if "Indice de riesgo (reglas)" not in sel else []
COLE = {"Indice de riesgo (reglas)": NA, "LOF": AQ, "Random Forest": AZ, "Regresion logistica": "#008300",
        "Mahalanobis robusta (MCD)": "#eda100", "Puntaje z maximo": "#6250d6", "One-Class SVM": "#e87ba4", "XGBoost": "#e34948"}
pal = [COLE.get(e, "#555555") for e in sel]
fig, ax = plt.subplots(1, 2, figsize=(10, 4.4), dpi=220)
for c, nom in zip(pal, sel):
    f, tp_, _ = roc_curve(y, S[nom]); ax[0].plot(f, tp_, color=c, lw=2, label=f"{nm(nom)} ({roc_auc_score(y, S[nom]):.2f})")
    pr, rc, _ = precision_recall_curve(y, S[nom]); ax[1].plot(rc, pr, color=c, lw=2, label=nm(nom))
ax[0].plot([0, 1], [0, 1], "--", color=MU, lw=1); ax[0].set_xlabel("Tasa de falsos positivos"); ax[0].set_ylabel("Tasa de verdaderos positivos (recall)")
ax[0].legend(frameon=False, fontsize=7.5, loc="lower right", title="AUC", title_fontsize=8)
ax[1].axhline(base, ls="--", color=MU, lw=1); ax[1].set_xlabel("Recall"); ax[1].set_ylabel("Precisión"); ax[1].set_ylim(0, 0.6)
ax[1].legend(frameon=False, fontsize=7.5, loc="upper right")
fig.tight_layout(); fig.savefig(OUT + "/figuras/Figura_4-5_curvas_ROC_PR.png"); plt.close(fig)

# Figura 4-6: seleccion del numero de grupos
fig, ax = plt.subplots(1, 4, figsize=(13, 3.3), dpi=220)
ax[0].plot(ks, inercia, "-o", color=AZ, lw=2, ms=5); ax[0].set_xlabel("k (grupos)"); ax[0].set_ylabel("Inercia (SSE)"); ax[0].set_title("a) Método del codo", loc="left", fontsize=9)
ax[1].plot(ks, silu, "-o", color=AZ, lw=2, ms=5, label="K-medias"); ax[1].plot(ks, [sil_ward[k] for k in ks], "-s", color=NA, lw=2, ms=5, label="Ward")
ax[1].axvline(K, ls="--", color=MU, lw=1); ax[1].set_xlabel("k (grupos)"); ax[1].set_ylabel("Silueta media"); ax[1].legend(frameon=False, fontsize=7.5); ax[1].set_title("b) Silueta", loc="left", fontsize=9)
ax[2].plot(range(1, 11), bic, "-o", color=AQ, lw=2, ms=5); ax[2].axvline(G, ls="--", color=MU, lw=1); ax[2].set_xlabel("Componentes"); ax[2].set_ylabel("BIC"); ax[2].set_title("c) BIC de la mezcla gaussiana", loc="left", fontsize=9)
ax[3].plot(kd, color=AZ, lw=2); ax[3].axhline(EPS, ls="--", color=MU, lw=1); ax[3].text(5, EPS * 1.05, f"ε = {EPS:.1f}", color=MU, fontsize=8)
ax[3].set_xlabel("Órdenes ordenadas"); ax[3].set_ylabel(f"Distancia al {MS}.º vecino"); ax[3].set_title("d) Curva k-distancia (DBSCAN)", loc="left", fontsize=9)
fig.tight_layout(); fig.savefig(OUT + "/figuras/Figura_4-6_seleccion_grupos.png"); plt.close(fig)

# Figura 4-7: perfil de los segmentos (z medio dentro del tipo de servicio)
segs = [-1] + list(range(K2))
M = pd.DataFrame([Z.loc[seg == c, var_ok].mean() for c in segs], index=[nombres_seg[c] for c in segs])
top_v = M.abs().max().sort_values(ascending=False).index[:10]; M = M[top_v]
fig, ax = plt.subplots(figsize=(9.2, 0.6 * len(segs) + 2.0), dpi=220)
vmax = float(np.abs(M.values).max())
im = ax.imshow(M.values, cmap="RdBu_r", vmin=-vmax, vmax=vmax, aspect="auto")
ax.set_xticks(range(len(top_v))); ax.set_xticklabels([v.replace("_", " ").capitalize() for v in top_v], rotation=35, ha="right", fontsize=8)
ax.set_yticks(range(len(segs))); ax.set_yticklabels([f"Segmento {nombres_seg[c]} (n={int((seg==c).sum())}; obs.={int(y[seg==c].sum())})" for c in segs], fontsize=8)
for i in range(len(segs)):
    for j in range(len(top_v)):
        ax.text(j, i, f"{M.values[i, j]:.1f}", ha="center", va="center", fontsize=7, color="white" if abs(M.values[i, j]) > vmax * 0.55 else TX)
cb = fig.colorbar(im, ax=ax, fraction=0.03); cb.set_label("z medio", fontsize=8)
for sp in ax.spines.values(): sp.set_visible(False)
fig.savefig(OUT + "/figuras/Figura_4-7_perfil_grupos.png", bbox_inches="tight"); plt.close(fig)

# Figura 4-8: curva de captura
fig, ax = plt.subplots(figsize=(7.4, 4.4), dpi=220)
kk = np.arange(0, n + 1)
for c, nom in zip(pal, sel):
    rr = np.random.default_rng(0); orden = np.lexsort((rr.random(n), -S[nom])); cap = np.r_[0, np.cumsum(y[orden])]
    ax.plot(kk, cap, color=c, lw=2, label=nm(nom))
ax.plot(kk, kk * base, "--", color=MU, lw=1, label="Azar")
ax.axvline(119, color=GR, lw=1); ax.text(121, 0.5, "119 órdenes (25,3 %)", color=MU, fontsize=8)
ax.set_xlabel("Órdenes revisadas (en el orden que propone cada método)"); ax.set_ylabel("Observadas incluidas (de 14)")
ax.legend(frameon=False, fontsize=8, loc="lower right"); ax.set_ylim(0, 14.5)
fig.tight_layout(); fig.savefig(OUT + "/figuras/Figura_4-8_curva_captura.png"); plt.close(fig)

# Figura A-2: puntaje z maximo en ordenes observadas vs no observadas
zmax = Z.abs().max(axis=1).values
a_, b_ = zmax[y == 0], zmax[y == 1]
u_mw, p_mw = stats.mannwhitneyu(b_, a_, alternative="greater")
RES["z_maximo_obs_vs_no"] = {"mediana_no_observadas": round(float(np.median(a_)), 2),
                             "mediana_observadas": round(float(np.median(b_)), 2),
                             "U_Mann_Whitney": float(u_mw), "p_una_cola": round(float(p_mw), 3)}
fig, ax = plt.subplots(figsize=(6.2, 5), dpi=220)
bp = ax.boxplot([a_, b_], widths=0.5, patch_artist=True, showfliers=False, medianprops=dict(color=TX, lw=1.6))
for patch, c in zip(bp["boxes"], [GR, NA]):
    patch.set_facecolor(c); patch.set_alpha(0.55)
rr = np.random.default_rng(0)
ax.scatter(1 + rr.normal(0, 0.06, len(a_)), a_, s=6, c="#555555", alpha=0.35, lw=0)
ax.scatter(2 + rr.normal(0, 0.05, len(b_)), b_, s=30, c=NA, edgecolor="white", lw=0.6, zorder=3)
ax.axhline(2, ls="--", color=MU, lw=1); ax.text(2.45, 2.08, "|z| = 2", fontsize=8, color=MU, ha="right")
ax.set_xticks([1, 2]); ax.set_xticklabels([f"No observadas\n(n={len(a_)})", f"Observadas\n(n={len(b_)})"])
ax.set_ylabel("Puntaje z máximo (por tipo de servicio)")
fig.tight_layout(); fig.savefig(OUT + "/figuras/Figura_A-2_boxplot_z_maximo.png"); plt.close(fig)

# Figura A-1: PCA coloreada por grupo, observadas resaltadas
P2 = PCA(2, random_state=0).fit(Ze); C2 = P2.transform(Ze); RES["pca_var"] = list(np.round(P2.explained_variance_ratio_, 3))
fig, ax = plt.subplots(figsize=(7.2, 5.2), dpi=220)
pal8 = [AZ, NA, AQ, "#eda100", "#e87ba4", "#008300", "#6250d6", "#e34948"]
for c in range(K):
    g = km.labels_ == c; ax.scatter(C2[g, 0], C2[g, 1], s=12, color=pal8[c % 8], alpha=0.55, lw=0, label=f"Grupo {c+1}")
ax.scatter(C2[y == 1, 0], C2[y == 1, 1], s=60, facecolor="none", edgecolor=TX, lw=1.4, label="Observada por el operador")
ax.set_xlabel(f"Componente 1 ({RES['pca_var'][0]*100:.1f} %)"); ax.set_ylabel(f"Componente 2 ({RES['pca_var'][1]*100:.1f} %)")
ax.legend(frameon=False, fontsize=7.5, loc="best"); fig.tight_layout(); fig.savefig(OUT + "/figuras/Figura_A-1_proyeccion_PCA.png"); plt.close(fig)

json.dump(RES, open(OUT + "/resumen_resultados.json", "w"), indent=1, default=lambda x: x.item() if hasattr(x, "item") else str(x), ensure_ascii=False)
pd.set_option("display.width", 250)
print(T.round(3).to_string())
print(json.dumps(RES, indent=1, default=lambda x: x.item() if hasattr(x, "item") else str(x), ensure_ascii=False))
print(prof.round(2).to_string()); print(prof2.round(2).to_string())
print(carga[carga.k == 119].round(2).to_string())
print(sens.round(3).to_string())
print(pd.read_csv(OUT + "/Tabla_A-3_sensibilidad_hiperparametros.csv").to_string())
print(imp["coef_LR"].round(2).head(8).to_string()); print(imp["perm_RF"].round(4).head(8).to_string())
print(pd.read_csv(OUT + "/Tabla_4-6_matriz_confusion_k119.csv").to_string())
print(obs_grupo.to_string())
