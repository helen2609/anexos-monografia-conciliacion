# Analisis estadistico de las reglas (apartados 4.2 y 4.6.3 de la monografia)
# Lee la base anonimizada Datos_PowerBI.xlsx (hoja Ordenes) y calcula intervalos de Wilson, la prueba exacta
# de Fisher y la U de Mann-Whitney para el indice de reglas. Las pruebas de permutacion y los intervalos
# bootstrap de los catorce enfoques estan en modelos_v12.py.
# Uso: python analisis_estadistico.py ruta/Datos_PowerBI.xlsx
import sys
import numpy as np, pandas as pd
from scipy import stats

o = pd.read_excel(sys.argv[1] if len(sys.argv) > 1 else "Datos_PowerBI.xlsx", sheet_name="Ordenes")
y = o["observada_operador"].astype(int).values
alerta = (o["nivel_riesgo"] != "BAJO").values
puntaje = o["puntaje_riesgo"].fillna(0).values

def wilson(k, n, z=1.96):
    p = k / n; den = 1 + z**2 / n
    c = (p + z**2 / (2 * n)) / den; h = z * np.sqrt(p * (1 - p) / n + z**2 / (4 * n**2)) / den
    return float(round(c - h, 3)), float(round(c + h, 3))

k, n = int(y.sum()), len(y)
tp, m = int(y[alerta].sum()), int(alerta.sum())
print(f"Tasa de observacion: {k}/{n} = {k/n:.3f}  IC95% {wilson(k, n)}")
print(f"Reglas: exhaustividad {tp}/{k} IC95% {wilson(tp, k)}; precision {tp}/{m} IC95% {wilson(tp, m)}")
tabla = [[int((alerta & (y == 1)).sum()), int((alerta & (y == 0)).sum())],
         [int((~alerta & (y == 1)).sum()), int((~alerta & (y == 0)).sum())]]
print("Fisher bilateral (alerta vs observada): p =", round(stats.fisher_exact(tabla)[1], 3))
print("Mann-Whitney unilateral (puntaje): p =", round(stats.mannwhitneyu(puntaje[y == 1], puntaje[y == 0], alternative="greater")[1], 3))
print(f"Esperado al azar con {m} ordenes: {k * m / n:.1f}")
