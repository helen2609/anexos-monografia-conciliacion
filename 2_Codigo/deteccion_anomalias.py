"""
Proyecto Final Diplomado - Deteccion de anomalias en planillas de conciliacion
Empresa contratista / operador - Ordenes de trabajo FTTH/IPTV

Este script:
1. Extrae la tabla "Maestro Ejecutados" de la planilla de conciliacion (.xlsx)
2. Construye un perfil esperado de materiales por tipo de servicio ejecutado
   (aprendido de los propios datos, no reglas fijas a mano)
3. Detecta anomalias: ordenes cuyos materiales se desvian del perfil esperado
   de su categoria, usando z-score por grupo
4. Exporta una tabla de ordenes sospechosas, ordenadas por severidad

Uso:
    python deteccion_anomalias.py ruta_a_planilla.xlsx
"""

import sys
import re
import difflib
import openpyxl
import pandas as pd
import numpy as np

# --- 1. Configuracion de columnas (posicion 0-index en "Maestro Ejecutados") ---

COL_SOLICITUD = 1
COL_TIPO_TRABAJO = 7
COL_CANTIDAD_IPTV = 10
COL_TIPO_SERVICIO = 24
COL_TECNICO1 = 26
COL_TECNICO2 = 27
COL_AREA_COBERTURA = 3  # URBANO / RURAL
COL_POSTES = 81         # "Apoyos Instalados en Postes" -> cantidad de postes usados

# Bloque "MATERIAL INSTALADO": columnas 52 a 70 (19 materiales)
MATERIALES_COL_START = 52
MATERIALES = [
    "ANILLAS_GUIA", "TENSORES", "SOPORTE_MEDIO_TRAMO", "CABLE_DUCTO",
    "CINTA_ACERO", "CABLE_ACERO_CRUCE", "ACCESORIOS_POSTE", "CABLE_DROP",
    "ROSETA_OPTICA", "TERMINAL_SC", "CABLE_UTP", "CONECTOR_RJ45",
    "CAPUCHON_RJ45", "PATCH_CORD_SCSC", "ADAPTADOR_NAP", "CABLE_HDMI",
    "CABLE_RCA", "PATCH_CORD_UTP", "ADAPTADOR_ACDC",
]

# TENSORES se valida con una regla exacta contra la cantidad de postes,
# no con z-score, por eso se excluye del chequeo estadistico general.
MATERIALES_ESTADISTICOS = [m for m in MATERIALES if m != "TENSORES"]

# Tipos de servicio validos segun el tipo de trabajo declarado
SERVICIOS_DE_INSTALACION = {"SOLO ONT", "ONT + 1 STB", "ONT + 2 STB", "ONT + 3 STB"}
SERVICIOS_DE_ADICION = {"ADICION DE SERVICIO"}

# Materiales medidos en metros: el sistema los guarda como entero,
# un decimal en la conciliacion es señal de error de conversion/tipeo.
MATERIALES_METROS_ENTEROS = ["CABLE_DUCTO", "CABLE_DROP", "CABLE_UTP", "CINTA_ACERO", "CABLE_ACERO_CRUCE"]

# Posible FTTB (fibra hasta el edificio): sin postes ni materiales de tendido aereo.
# En esas ordenes es normal que falten o se desvien estos materiales, asi que la
# alerta baja a riesgo BAJO y se sugiere "posible FTTB" en lugar de marcar error.
MATERIALES_TENDIDO_AEREO = ["ANILLAS_GUIA", "TENSORES", "SOPORTE_MEDIO_TRAMO",
                            "CINTA_ACERO", "CABLE_ACERO_CRUCE", "ACCESORIOS_POSTE"]
# CABLE_DROP no se incluye: el drop siempre existe, un drop en 0 es error real aunque sea FTTB.
MATERIALES_AFECTADOS_POR_FTTB = MATERIALES_TENDIDO_AEREO + ["CABLE_DUCTO"]

# Formato esperado del numero de solicitud: solo se valida el primer digito
# y la cantidad de digitos (ajustar si cambia en otros meses)
SOLICITUD_PRIMER_DIGITO = "3"
SOLICITUD_LARGO = 9

# Umbral de similitud para detectar posible error de tipeo en nombres de tecnico
# (distancia de edicion maxima entre un nombre raro y uno frecuente)
TECNICO_DISTANCIA_MAXIMA = 2
TECNICO_FRECUENCIA_MINIMA_CONFIABLE = 3  # veces que debe aparecer un nombre para considerarlo "correcto"

# Materiales clave que NO deberian faltar si el perfil del grupo los usa siempre
# (media alta y desviacion baja -> son "obligatorios" para ese tipo de servicio)
UMBRAL_OBLIGATORIO_MEDIA = 0.9   # si >=90% de las ordenes del grupo lo usan
UMBRAL_ANOMALIA_ZSCORE = 2.0     # desviaciones estandar para marcar anomalia

# --- Pesos del indice de riesgo ---
# Alto (3): inconsistencias que suelen implicar que la orden esta mal cerrada
#           o que puede no coincidir con lo que el sistema del operador tiene registrado.
# Medio (2): inconsistencias fisicas/logicas detectables pero de menor gravedad.
# Bajo (1): errores de formato/redondeo, se corrigen facil y no cuestionan
#           si el trabajo se hizo bien.
PESOS = {
    "tipo_trabajo_vs_servicio": 3,
    "solicitud_formato": 3,
    "tensores_vs_postes": 2,
    "tecnico_tipeo": 2,
}

# Peso especifico por material para las alertas "material_faltante",
# "material_inusual" y "decimal_metros": depende de QUIEN asume el costo.
# - Materiales que paga/factura la empresa contratista (contratista): un error aqui
#   afecta directamente su costo/cobro -> peso alto.
# - Materiales que entrega la empresa de telecomunicaciones (operador): el error sigue
#   siendo una inconsistencia a corregir, pero no representa el mismo riesgo de
#   costo para el contratista -> peso bajo.
# Los marcados como "asumido" no fueron confirmados explicitamente por el usuario,
# se agruparon con el resto de herrajes de instalacion por similitud.
# Criterio: la lista de precios unitarios de la propia planilla del operador (hoja "Resumen").
# - Material con precio mayor a 0: la planilla del operador lo valora y lo descuenta si hay
#   diferencia -> afecta el costo/cobro de la contratista -> peso alto.
# - Material con precio 0 en esa planilla: el error sigue siendo una inconsistencia a corregir,
#   pero no tiene valor en Bs -> peso bajo.
MATERIALES_OPERADOR = {
    "CABLE_HDMI", "CABLE_RCA", "PATCH_CORD_UTP", "ADAPTADOR_ACDC",   # precio 0 en la planilla del operador
}
MATERIALES_CONTRATISTA = set(MATERIALES) - MATERIALES_OPERADOR

# Dos ejes: (1) quien paga el material y (2) que tan fuerte es la evidencia.
#   - "faltante": evidencia fuerte (un material casi obligatorio no esta)
#   - "decimal" / "inusual": evidencia mas debil (puede ser un caso legitimo,
#     por ejemplo una casa mas lejos que el promedio)
PESO_MATERIAL = {
    "contratista": {"faltante": 3, "decimal": 2, "inusual": 2},
    "operador":       {"faltante": 2, "decimal": 1, "inusual": 1},
}


def peso_material(nombre_material: str, tipo_alerta: str) -> int:
    """Peso segun quien paga el material y el tipo de alerta
    (tipo_alerta: 'faltante', 'decimal' o 'inusual')."""
    if nombre_material in MATERIALES_OPERADOR:
        quien = "operador"
    else:
        quien = "contratista"  # incluye los materiales asumidos y cualquier no clasificado
    return PESO_MATERIAL[quien][tipo_alerta]


NIVELES_RIESGO = ["BAJO", "MEDIO", "ALTO"]


def extraer_tabla(path_excel: str) -> pd.DataFrame:
    """Lee la hoja 'Maestro Ejecutados' y arma un DataFrame limpio, una fila por orden."""
    wb = openpyxl.load_workbook(path_excel, data_only=True)
    ws = wb["Maestro Ejecutados"]
    filas = list(ws.iter_rows(min_row=3, values_only=True))

    registros = []
    for r in filas:
        if r[COL_SOLICITUD] is None:
            continue  # fila vacia / de formato
        registro = {
            "solicitud": r[COL_SOLICITUD],
            "tipo_trabajo": r[COL_TIPO_TRABAJO],
            "tipo_servicio": r[COL_TIPO_SERVICIO],
            "cantidad_iptv": r[COL_CANTIDAD_IPTV] or 0,
            "area_cobertura": r[COL_AREA_COBERTURA],
            "tecnico1": r[COL_TECNICO1],
            "tecnico2": r[COL_TECNICO2],
            "postes": r[COL_POSTES] if isinstance(r[COL_POSTES], (int, float)) else 0,
        }
        materiales = r[MATERIALES_COL_START: MATERIALES_COL_START + len(MATERIALES)]
        for nombre, valor in zip(MATERIALES, materiales):
            registro[nombre] = valor if isinstance(valor, (int, float)) else 0
        registros.append(registro)

    df = pd.DataFrame(registros)
    return df


def construir_perfil_esperado(df: pd.DataFrame) -> pd.DataFrame:
    """Calcula media, desviacion estandar y % de uso de cada material, por tipo_servicio."""
    stats = df.groupby("tipo_servicio")[MATERIALES_ESTADISTICOS].agg(["mean", "std"])
    uso = df.groupby("tipo_servicio")[MATERIALES_ESTADISTICOS].apply(lambda g: (g > 0).mean())
    return stats, uso


def validar_solicitud(solicitud):
    """Regla exacta: el numero de solicitud debe empezar con el digito esperado
    y tener el largo esperado. Detecta el caso de tipear el primer digito mal
    (ej. empezar con 2 en vez de 3). Devuelve (categoria, mensaje) o None."""
    texto = str(solicitud)
    if len(texto) != SOLICITUD_LARGO or not texto.startswith(SOLICITUD_PRIMER_DIGITO):
        msg = f"SOLICITUD='{texto}' no tiene el formato esperado (debe empezar con {SOLICITUD_PRIMER_DIGITO} y tener {SOLICITUD_LARGO} digitos)"
        return ("solicitud_formato", msg)
    return None


def validar_decimales(fila) -> list:
    """Regla exacta: los materiales medidos en metros se guardan como entero en el
    sistema; un valor con decimales en la conciliacion senala un error de tipeo/redondeo.
    Devuelve una lista de (categoria, mensaje)."""
    alertas = []
    for mat in MATERIALES_METROS_ENTEROS:
        valor = fila[mat]
        if valor and valor != int(valor):
            msg = f"{mat}={valor:g} tiene decimales (el sistema lo guarda como entero)"
            alertas.append(("decimal_metros", msg))
    return alertas


def construir_roster_tecnicos(df: pd.DataFrame) -> dict:
    """Cuenta cuantas veces aparece cada nombre de tecnico (tecnico1 + tecnico2)."""
    nombres = pd.concat([df["tecnico1"], df["tecnico2"]]).dropna().astype(str).str.strip()
    return nombres.value_counts().to_dict()


def validar_nombre_tecnico(nombre, roster: dict):
    """Si un nombre aparece pocas veces y es muy parecido a un nombre frecuente,
    es candidato a error de tipeo (ej. 'RUBEN GARCIA' en vez de 'RAUL GARCIA')."""
    if not nombre:
        return None
    nombre = str(nombre).strip()
    frecuencia = roster.get(nombre, 0)
    if frecuencia >= TECNICO_FRECUENCIA_MINIMA_CONFIABLE:
        return None  # nombre ya confiable, se repite bastante

    nombres_confiables = [n for n, f in roster.items() if f >= TECNICO_FRECUENCIA_MINIMA_CONFIABLE and n != nombre]
    parecidos = difflib.get_close_matches(nombre, nombres_confiables, n=1, cutoff=0.75)
    if parecidos:
        msg = f"TECNICO='{nombre}' (aparece {frecuencia}x) es muy similar a '{parecidos[0]}' — posible error de tipeo"
        return ("tecnico_tipeo", msg)
    return None


def validar_tensores_vs_postes(postes: float, tensores: float):
    """Regla exacta: por cada poste se usan entre 1 y 2 tensores (relacion fisica real)."""
    minimo_esperado = postes
    maximo_esperado = postes * 2
    if not (minimo_esperado <= tensores <= maximo_esperado):
        msg = (
            f"TENSORES={tensores:g} no coincide con POSTES={postes:g} "
            f"(esperado entre {minimo_esperado:g} y {maximo_esperado:g})"
        )
        return ("tensores_vs_postes", msg)
    return None


def es_posible_fttb(fila) -> bool:
    """Heuristica: instalacion nueva sin postes y sin ningun material de tendido aereo.
    (La planilla no distingue FTTH de FTTB, por eso es una sugerencia y no una certeza.)"""
    if fila["tipo_trabajo"] != "INSTALACION" or fila["postes"] != 0:
        return False
    return all(fila[m] == 0 for m in MATERIALES_TENDIDO_AEREO)


def validar_tipo_trabajo_vs_servicio(tipo_trabajo: str, tipo_servicio: str):
    """Regla exacta: el tipo de trabajo declarado debe ser consistente con el
    tipo de servicio ejecutado seleccionado (evita el caso de tipear 'ADICION'
    pero dejar seleccionado un tipo de servicio de instalacion nueva, o viceversa)."""
    if tipo_trabajo == "ADICION DE SERVICIO" and tipo_servicio not in SERVICIOS_DE_ADICION:
        msg = f"TIPO DE TRABAJO='ADICION DE SERVICIO' pero TIPO DE SERVICIO EJECUTADO='{tipo_servicio}'"
        return ("tipo_trabajo_vs_servicio", msg)
    if tipo_trabajo == "INSTALACION" and tipo_servicio not in SERVICIOS_DE_INSTALACION:
        msg = f"TIPO DE TRABAJO='INSTALACION' pero TIPO DE SERVICIO EJECUTADO='{tipo_servicio}'"
        return ("tipo_trabajo_vs_servicio", msg)
    return None


def detectar_anomalias(df: pd.DataFrame) -> pd.DataFrame:
    """Marca, por orden, que materiales o campos se desvian de lo esperado,
    y calcula un indice de riesgo ponderado (no todas las alertas pesan igual)."""
    stats, uso = construir_perfil_esperado(df)
    roster = construir_roster_tecnicos(df)

    resultados = []
    for _, fila in df.iterrows():
        grupo = fila["tipo_servicio"]
        alertas = []            # lista de (categoria, mensaje) -- peso fijo por categoria
        alertas_con_peso = []   # lista de (peso, mensaje) -- peso especifico por material
        notas_fttb = []         # alertas de material que bajan a riesgo BAJO por posible FTTB
        posible_fttb = es_posible_fttb(fila)

        alerta_solicitud = validar_solicitud(fila["solicitud"])
        if alerta_solicitud:
            alertas.append(alerta_solicitud)

        alerta_tipo = validar_tipo_trabajo_vs_servicio(fila["tipo_trabajo"], grupo)
        if alerta_tipo:
            alertas.append(alerta_tipo)

        alerta_tensores = validar_tensores_vs_postes(fila["postes"], fila["TENSORES"])
        if alerta_tensores:
            alertas.append(alerta_tensores)

        # --- Materiales con decimales: peso segun quien paga el material ---
        for mat in MATERIALES_METROS_ENTEROS:
            valor = fila[mat]
            if valor and valor != int(valor):
                msg = f"{mat}={valor:g} tiene decimales (el sistema lo guarda como entero)"
                alertas_con_peso.append((peso_material(mat, "decimal"), msg))

        for campo in ("tecnico1", "tecnico2"):
            alerta_nombre = validar_nombre_tecnico(fila[campo], roster)
            if alerta_nombre:
                alertas.append(alerta_nombre)

        for mat in MATERIALES_ESTADISTICOS:
            media = stats.loc[grupo, (mat, "mean")]
            std = stats.loc[grupo, (mat, "std")] or 0
            pct_uso = uso.loc[grupo, mat]
            valor = fila[mat]

            if pct_uso >= UMBRAL_OBLIGATORIO_MEDIA and valor == 0:
                msg = f"Falta {mat} (se usa en {pct_uso:.0%} de '{grupo}')"
                if posible_fttb and mat in MATERIALES_AFECTADOS_POR_FTTB:
                    notas_fttb.append(msg)
                else:
                    alertas_con_peso.append((peso_material(mat, "faltante"), msg))
            elif std > 0:
                z = abs(valor - media) / std
                if z >= UMBRAL_ANOMALIA_ZSCORE:
                    msg = f"{mat}={valor:g} inusual (esperado ~{media:.1f}, grupo '{grupo}')"
                    if posible_fttb and mat in MATERIALES_AFECTADOS_POR_FTTB:
                        notas_fttb.append(msg)
                    else:
                        alertas_con_peso.append((peso_material(mat, "inusual"), msg))

        puntaje = sum(PESOS[categoria] for categoria, _ in alertas) + sum(p for p, _ in alertas_con_peso)
        todas_las_alertas_texto = [m for _, m in alertas] + [m for _, m in alertas_con_peso]

        resultados.append({
            "solicitud": fila["solicitud"],
            "tipo_trabajo": fila["tipo_trabajo"],
            "tipo_servicio": grupo,
            "area_cobertura": fila["area_cobertura"],
            "cantidad_iptv": fila["cantidad_iptv"],
            "tecnico1": fila["tecnico1"],
            "posible_fttb": posible_fttb,
            "nota_fttb": ("POSIBLE FTTB (sin postes ni materiales de tendido aereo), riesgo reducido: "
                          + " | ".join(notas_fttb)) if notas_fttb else "",
            "n_alertas": len(todas_las_alertas_texto),
            "puntaje_riesgo": puntaje,
            "detalle_alertas": " | ".join(todas_las_alertas_texto) if todas_las_alertas_texto else "",
        })

    resultado_df = pd.DataFrame(resultados)
    resultado_df["nivel_riesgo"] = asignar_nivel_riesgo(resultado_df["puntaje_riesgo"])
    return resultado_df.sort_values("puntaje_riesgo", ascending=False)


def asignar_nivel_riesgo(puntajes: pd.Series) -> pd.Series:
    """Convierte el puntaje ponderado en BAJO / MEDIO / ALTO.

    Los cortes se derivan directamente del significado de los pesos (no son
    arbitrarios): una sola alerta de peso bajo/medio (1 o 2) deja la orden en
    MEDIO; alcanzar o superar 3 -- el peso de una alerta grave por si sola
    (material faltante, inconsistencia de tipo de trabajo, solicitud mal
    formada), o varias alertas menores combinadas -- la sube a ALTO.
    """
    def clasificar(p):
        if p == 0:
            return "BAJO"
        elif p <= 2:
            return "MEDIO"
        else:
            return "ALTO"

    return puntajes.apply(clasificar)


def resumen_por_tecnico(df: pd.DataFrame) -> pd.DataFrame:
    """Cuadro simple: nombre de tecnico, cuantas instalaciones y cuantas adiciones hizo
    (cuenta al tecnico tanto si aparece en TECNICO1 como en TECNICO2 de la orden)."""
    filas = []
    for _, fila in df.iterrows():
        es_adicion = fila["tipo_trabajo"] == "ADICION DE SERVICIO"
        for campo in ("tecnico1", "tecnico2"):
            nombre = fila[campo]
            if not nombre:
                continue
            filas.append({"tecnico": str(nombre).strip(), "es_adicion": es_adicion})

    detalle = pd.DataFrame(filas)
    resumen = detalle.groupby("tecnico")["es_adicion"].agg(
        instalaciones=lambda s: (~s).sum(),
        adiciones=lambda s: s.sum(),
    ).reset_index()
    resumen["total"] = resumen["instalaciones"] + resumen["adiciones"]
    return resumen.sort_values("total", ascending=False)


def entrenar_clasificador_riesgo(anomalias: pd.DataFrame):
    """Entrena un Random Forest que predice el nivel de riesgo (BAJO/MEDIO/ALTO)
    a partir de caracteristicas del CONTEXTO de la orden (tipo de trabajo, tipo
    de servicio, zona, cantidad de IPTV, tecnico) -- es decir, variables que se
    conocen antes/independientemente de revisar material por material.

    Esto es complementario al motor de reglas: el motor de reglas te dice
    exactamente QUE esta mal en una orden ya cerrada; este modelo aprende que
    patrones (ej. cierto tecnico, cierta zona, cierto tipo de servicio) se
    asocian historicamente a mayor riesgo, util para priorizar revision incluso
    antes de que la orden este completamente cargada.
    """
    from sklearn.ensemble import RandomForestClassifier
    from sklearn.model_selection import StratifiedKFold, cross_val_predict
    from sklearn.preprocessing import LabelEncoder
    from sklearn.metrics import classification_report

    datos = anomalias.copy()
    columnas_categoricas = ["tipo_trabajo", "tipo_servicio", "area_cobertura", "tecnico1"]

    X = pd.DataFrame(index=datos.index)
    codificadores = {}
    for col in columnas_categoricas:
        le = LabelEncoder()
        X[col] = le.fit_transform(datos[col].astype(str))
        codificadores[col] = le
    X["cantidad_iptv"] = datos["cantidad_iptv"]

    y = datos["nivel_riesgo"]

    modelo = RandomForestClassifier(n_estimators=200, random_state=42, class_weight="balanced")

    # Dataset chico -> validacion cruzada (k=5) en vez de un solo train/test split
    n_clases_presentes = y.value_counts().min()
    k = min(5, n_clases_presentes) if n_clases_presentes >= 2 else 2
    cv = StratifiedKFold(n_splits=max(k, 2), shuffle=True, random_state=42)

    predicciones_cv = cross_val_predict(modelo, X, y, cv=cv)
    print("--- Evaluacion por validacion cruzada (k=%d) ---" % cv.get_n_splits())
    print(classification_report(y, predicciones_cv, zero_division=0))

    # Entrenamiento final con todos los datos, para usar en ordenes nuevas
    modelo.fit(X, y)

    importancias = pd.Series(modelo.feature_importances_, index=X.columns).sort_values(ascending=False)
    print("Importancia de variables:")
    print(importancias.to_string())

    return modelo, codificadores, importancias


# === EVALUACION ===
# Precios unitarios (Bs) de la tabla "Evaluacion de materiales" de la conciliacion.
# Incluye todos los materiales de la lista de precios del operador (los de precio 0 no tienen valor en Bs).
PRECIOS_UNITARIOS = {
    "ANILLAS_GUIA": 4.70, "TENSORES": 4.70, "SOPORTE_MEDIO_TRAMO": 12.00,
    "CABLE_DUCTO": 9.30, "CINTA_ACERO": 15.00, "CABLE_ACERO_CRUCE": 8.80,
    "ACCESORIOS_POSTE": 205.00, "CABLE_DROP": 0.60, "ROSETA_OPTICA": 8.10,
    "TERMINAL_SC": 6.72, "CABLE_UTP": 6.00, "CONECTOR_RJ45": 4.80, "CAPUCHON_RJ45": 2.50,
    "PATCH_CORD_SCSC": 45.00, "ADAPTADOR_NAP": 18.36,
    # Precio 0 en la planilla del operador (sin valor en Bs)
    "CABLE_HDMI": 0.0, "CABLE_RCA": 0.0, "PATCH_CORD_UTP": 0.0, "ADAPTADOR_ACDC": 0.0,
}


def preparar_observaciones(observaciones: list) -> pd.DataFrame:
    """Convierte la lista de observaciones del operador en tabla, con la diferencia
    (conciliacion - sistema) y su valor en Bs."""
    obs = pd.DataFrame(
        observaciones,
        columns=["solicitud", "material", "valor_conciliacion", "valor_sistema"],
    )
    obs["solicitud"] = obs["solicitud"].astype(int)
    obs["diferencia"] = obs["valor_conciliacion"] - obs["valor_sistema"]
    obs["diferencia_bs"] = obs.apply(
        lambda r: r["diferencia"] * PRECIOS_UNITARIOS.get(r["material"], 0.0), axis=1
    )
    return obs


def evaluar_detector(anomalias: pd.DataFrame, obs: pd.DataFrame):
    """Compara el nivel de riesgo calculado contra las ordenes que el operador observo.
    Devuelve (resumen, detalle). Ordenes no observadas se toman como sin error."""
    from sklearn.metrics import roc_auc_score, average_precision_score

    a = anomalias.copy()
    a["solicitud"] = a["solicitud"].astype(int)
    a["con_error"] = a["solicitud"].isin(set(obs["solicitud"])).astype(int)
    total = int(a["con_error"].sum())

    filas = []
    for nombre, cond in [("Cualquier alerta (MEDIO o ALTO)", a["n_alertas"] > 0),
                         ("Solo ALTO", a["nivel_riesgo"] == "ALTO")]:
        sel = a[cond]
        tp = int(sel["con_error"].sum())
        filas.append({
            "criterio": nombre,
            "ordenes_marcadas": len(sel),
            "errores_atrapados": f"{tp}/{total}",
            "recall": round(tp / total, 2) if total else 0,
            "precision": round(tp / len(sel), 3) if len(sel) else 0,
        })
    resumen = pd.DataFrame(filas)
    print(f"Tasa base de error: {total}/{len(a)} = {total / len(a):.1%}")
    print(f"AUC del puntaje de riesgo: {roc_auc_score(a['con_error'], a['puntaje_riesgo']):.2f} "
          f"(0.50 = azar) | precision promedio: "
          f"{average_precision_score(a['con_error'], a['puntaje_riesgo']):.3f}")

    detalle = []
    for sol, grupo in obs.groupby("solicitud"):
        fila = a[a["solicitud"] == sol]
        if fila.empty:
            detalle.append({"solicitud": sol, "nivel_riesgo": "NO ESTA EN LA PLANILLA"})
            continue
        fila = fila.iloc[0]
        mats = list(grupo["material"])
        detalle.append({
            "solicitud": sol,
            "materiales_con_diferencia": ", ".join(mats),
            "puntaje_riesgo": fila["puntaje_riesgo"],
            "nivel_riesgo": fila["nivel_riesgo"],
            "alerta_sobre_ese_material": any(m in fila["detalle_alertas"] for m in mats),
        })
    return resumen, pd.DataFrame(detalle)


def entrenar_clasificador_con_etiquetas_reales(df: pd.DataFrame, obs: pd.DataFrame, repeticiones: int = 10):
    """Random Forest entrenado con etiquetas REALES (orden observada por el operador = 1).
    Usa validacion cruzada estratificada repetida porque hay muy pocos positivos."""
    import numpy as np
    from sklearn.ensemble import RandomForestClassifier
    from sklearn.model_selection import RepeatedStratifiedKFold
    from sklearn.metrics import roc_auc_score, average_precision_score
    from sklearn.preprocessing import LabelEncoder

    d = df.copy()
    d["solicitud"] = d["solicitud"].astype(int)
    y = d["solicitud"].isin(set(obs["solicitud"])).astype(int).values
    positivos = int(y.sum())
    if positivos < 2:
        print("Muy pocos positivos para entrenar.")
        return None

    X = pd.DataFrame(index=d.index)
    for col in ["tipo_servicio", "area_cobertura", "tecnico1"]:
        X[col] = LabelEncoder().fit_transform(d[col].astype(str))
    for col in MATERIALES + ["postes", "cantidad_iptv"]:
        X[col] = d[col].astype(float)

    splits = min(5, positivos)
    cv = RepeatedStratifiedKFold(n_splits=splits, n_repeats=repeticiones, random_state=42)
    aucs, aps = [], []
    for tr, te in cv.split(X, y):
        rf = RandomForestClassifier(n_estimators=300, class_weight="balanced_subsample",
                                    min_samples_leaf=3, random_state=0).fit(X.iloc[tr], y[tr])
        prob = rf.predict_proba(X.iloc[te])[:, 1]
        aucs.append(roc_auc_score(y[te], prob))
        aps.append(average_precision_score(y[te], prob))

    print(f"Positivos: {positivos} de {len(y)} (tasa base {y.mean():.1%})")
    print(f"AUC: {np.mean(aucs):.2f} (+-{np.std(aucs):.2f}) | precision promedio: "
          f"{np.mean(aps):.3f} (tasa base {y.mean():.3f})")
    modelo = RandomForestClassifier(n_estimators=300, class_weight="balanced_subsample",
                                    min_samples_leaf=3, random_state=0).fit(X, y)
    importancias = pd.Series(modelo.feature_importances_, index=X.columns).sort_values(ascending=False)
    return modelo, importancias


def main():
    if len(sys.argv) < 2:
        print("Uso: python deteccion_anomalias.py ruta_a_planilla.xlsx")
        sys.exit(1)

    path_excel = sys.argv[1]
    print(f"Leyendo: {path_excel}")
    df = extraer_tabla(path_excel)
    print(f"Ordenes cargadas: {len(df)}")

    anomalias = detectar_anomalias(df)
    n_con_alerta = (anomalias["n_alertas"] > 0).sum()
    print(f"Ordenes con al menos una alerta: {n_con_alerta} / {len(df)}")
    print("\nDistribucion de nivel de riesgo:")
    print(anomalias["nivel_riesgo"].value_counts().to_string())

    print("\nTop 10 ordenes de mayor riesgo:")
    cols = ["solicitud", "tipo_servicio", "puntaje_riesgo", "nivel_riesgo", "detalle_alertas"]
    print(anomalias[cols].head(10).to_string(index=False))

    print("\nResumen por tecnico (instalaciones / adiciones):")
    print(resumen_por_tecnico(df).to_string(index=False))

    print("\n--- Clasificador de riesgo (Random Forest) ---")
    entrenar_clasificador_riesgo(anomalias)


if __name__ == "__main__":
    main()
