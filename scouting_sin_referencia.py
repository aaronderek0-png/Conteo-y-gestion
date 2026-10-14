"""Universo de comparativa: productos comprados (Dossier 'Cadena Octubre') sin referencia comparable.

Cruza el Dossier con Reportes_Ahorro/Reporte_Ahorro_Resumen.xlsx (salida de comparador_precios.py) y deja:
  Reportes_Ahorro/Productos_Sin_Referencia.xlsx  -> SIN REFERENCIA (gasto desc.), NUEVOS (no estaban en Septiembre)
No inventa precios ni contacta proveedores: solo ordena lo que falta investigar, por impacto en el gasto.
Uso: python scouting_sin_referencia.py
"""
from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE = Path(__file__).resolve().parent
DOSSIER = BASE / "Dossier_Mensual_Mishiguene.xlsx"
RESUMEN = BASE / "Reportes_Ahorro" / "Reporte_Ahorro_Resumen.xlsx"
SALIDA = BASE / "Reportes_Ahorro" / "Productos_Sin_Referencia.xlsx"


def norm(s) -> str:
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"\s+", " ", re.sub(r"\(opci[oó]n \d+\)", "", s)).strip()


def main() -> dict:
    oct_ = pd.read_excel(DOSSIER, sheet_name="Cadena Octubre")
    sep = pd.read_excel(DOSSIER, sheet_name="Cadena Septiembre")
    comp = pd.read_excel(RESUMEN, sheet_name="Comparativa general")
    con_ref = {(str(r["Código SKU"]), norm(r["Producto (Dossier)"])) for _, r in comp.iterrows()}
    ya_sep = {(str(r["Código SKU"]), norm(r["Productos"])) for _, r in sep.iterrows()}
    filas = []
    for _, r in oct_.iterrows():
        clave = (str(r["Código SKU"]), norm(r["Productos"]))
        if clave in con_ref:
            continue
        filas.append({"SKU": r["Código SKU"], "Producto": r["Productos"], "Categoría": r["Categoría 2"],
                      "Proveedor actual": r["Proveedor"], "Cantidad inventario": r["Cantidad Inventario"],
                      "Monto total": r["Monto Total"], "Precio ponderado": r["Precio Ponderado"],
                      "Nuevo vs septiembre": "SÍ" if clave not in ya_sep else "no",
                      "Estado": "SIN REFERENCIA", "Acción": "buscar equivalente verificable (especie/calidad/presentación/neto)"})
    df = pd.DataFrame(filas).sort_values("Monto total", ascending=False, na_position="last")
    cubierto = float(oct_["Monto Total"].sum())
    sin_ref = float(df["Monto total"].sum())
    resumen = pd.DataFrame([{"Generado": datetime.now().isoformat(timespec="seconds"), "Productos en Dossier": len(oct_),
                             "Con referencia": len(oct_) - len(df), "SIN REFERENCIA": len(df),
                             "Gasto total Dossier": cubierto, "Gasto sin referencia": sin_ref,
                             "% del gasto sin referencia": round(100 * sin_ref / cubierto, 1) if cubierto else None,
                             "Nuevos vs septiembre": int((df["Nuevo vs septiembre"] == "SÍ").sum())}])
    with pd.ExcelWriter(SALIDA) as w:
        resumen.to_excel(w, sheet_name="Resumen", index=False)
        df.to_excel(w, sheet_name="SIN REFERENCIA", index=False)
        df[df["Nuevo vs septiembre"] == "SÍ"].to_excel(w, sheet_name="NUEVOS", index=False)
    print(resumen.T.to_string(header=False))
    print(df.head(15)[["SKU", "Producto", "Categoría", "Proveedor actual", "Monto total"]].to_string(index=False))
    return resumen.iloc[0].to_dict()


if __name__ == "__main__":
    main()
