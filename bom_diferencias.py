"""Diferencias BOM: costo unitario de la hoja Recetas (export local) vs motor_recetas. NO corrige nada: clasifica.
Fuente: Copias_Gestion/originales_export/Recetas.xlsx (EXPORT LOCAL, no Drive en vivo). Salida: Reportes_Ahorro/bom_diferencias.csv"""
import csv
import statistics
from pathlib import Path

from motor_recetas import MotorRecetas

BASE = Path(__file__).resolve().parent
m = MotorRecetas(BASE / "Copias_Gestion" / "originales_export" / "Recetas.xlsx")
filas = []
for cod, rc in m.recetas.items():
    ref = rc["c_un_hoja"]
    r = m.explotar(cod, rc["rend"] or 1)
    if not r or not ref:
        continue
    motor = r.costo_unitario if hasattr(r, "costo_unitario") else None
    if motor is None:
        continue
    ratio = ref / motor if motor else None
    if abs(ref - motor) / max(abs(ref), 1e-9) <= 0.02 and r.completo:
        continue
    if not r.completo:
        clase = "INCOMPLETA (insumo sin precio/factor)"
    elif ratio and abs(ratio - round(ratio)) < 0.02 * max(ratio, 1) and ratio > 1.5:
        clase = f"hoja = motor × {ratio:.1f}: probable REND/unidad no dividido en la hoja"
    elif ratio and 1 / ratio > 1.5 and abs(1 / ratio - round(1 / ratio)) < 0.02 * (1 / ratio):
        clase = f"motor = hoja × {1 / ratio:.1f}: probable factor de unidad (g↔kg) en el motor"
    else:
        clase = "diferencia sin patrón: revisar receta"
    filas.append({"cod": cod, "nombre": rc.get("nombre", ""), "rend_hoja": rc["rend"], "c_un_hoja": round(ref, 4),
                  "c_un_motor": round(motor, 4), "ratio": round(ratio, 3) if ratio else "", "clase": clase})
out = BASE / "Reportes_Ahorro" / "bom_diferencias.csv"
with out.open("w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=list(filas[0]))
    w.writeheader()
    w.writerows(filas)
from collections import Counter
print(len(filas), "recetas con diferencia")
for k, v in Counter(x["clase"].split(":")[0].split(" ×")[0] for x in filas).most_common():
    print(v, k)
