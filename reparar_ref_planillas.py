"""Repara el derrame `#REF!` de las pestañas `Facturas` (K:M, T:U, V:W) sin depender del separador de matrices del idioma.

Causa (verificada en el export de Lafinur): el comodín de las fórmulas `{"";"";""}` es una COLUMNA de 3 filas; cada celda
intenta derramar hacia abajo, choca con la fórmula de la fila siguiente y Sheets devuelve #REF!. V además usa
`HSTACK("","",)` (3 columnas) y pisa la columna X (SKU). Reparación: `HSTACK("","","")` en K:M, `HSTACK("","")` en T:U
y en V:W.

Uso:
  python reparar_ref_planillas.py --local Lafinur                 # dry-run (por defecto): solo lectura, informa
  MISHI_ENTORNO=PRUEBA python reparar_ref_planillas.py --local Lafinur --apply
`--apply` exige MISHI_ENTORNO=PRUEBA (entornos.py valida que el ID sea una copia registrada). En PRODUCCIÓN se rechaza.
También informa filas con N° de factura y PRODUCTOS vacío (p. ej. 111 y 854 en Lafinur): el dato hay que completarlo a mano.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

# {"";"";""} / {"";""} con espacios opcionales; HSTACK("","",) con la coma final
RX_V3 = re.compile(r'\{\s*""\s*;\s*""\s*;\s*""\s*\}')
RX_V2 = re.compile(r'\{\s*""\s*;\s*""\s*\}')
RX_V_TRAIL = re.compile(r'HSTACK\(\s*""\s*[,;]\s*""\s*[,;]\s*\)')


def reparar_formula(f: str) -> str:
    """Respeta el separador de argumentos que ya usa la fórmula en la planilla (es-AR: ';'; en-US: ',')."""
    sep = ";" if re.search(r'=\s*""\s*;', f) or re.search(r'\)\s*;', f) else ","
    fix3, fix2 = f'HSTACK("","",""' .replace('","', f'"{sep}"') + ")", f'HSTACK("{sep}"' + ')'
    fix3 = f'HSTACK(""{sep}""{sep}"")'
    fix2 = f'HSTACK(""{sep}"")'
    g = RX_V3.sub(fix3, f)
    g = RX_V2.sub(fix2, g)
    g = RX_V_TRAIL.sub(fix2, g)
    return g


def analizar(valores_formula: list[list[str]]) -> dict:
    """valores_formula: filas 1..n (índice 0 = encabezado) de A:X con value_render_option=FORMULA."""
    cambios, a_sin_e = [], []
    for i, fila in enumerate(valores_formula[1:], 2):
        fila = list(fila) + [""] * (24 - len(fila))
        if fila[0] not in ("", None) and fila[4] in ("", None):
            a_sin_e.append(i)
        for col in (10, 19, 21):                       # K, T, V
            f = fila[col]
            if isinstance(f, str) and f.startswith("="):
                g = reparar_formula(f)
                if g != f:
                    cambios.append({"fila": i, "col": "KTV"[(10, 19, 21).index(col)], "antes": f[:90], "despues": g[:90]})
    return {"celdas_a_reparar": len(cambios), "muestra": cambios[:4], "filas_con_N_sin_producto": a_sin_e}


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--local", required=True)
    ap.add_argument("--apply", action="store_true")
    a = ap.parse_args()
    import entornos
    import imputacion_facturas as imp
    if a.apply and entornos.entorno() != "PRUEBA":
        print("RECHAZADO: --apply solo se permite con MISHI_ENTORNO=PRUEBA (copia registrada en entornos.json).")
        return 2
    gc, _ = imp._gspread()
    sh = gc.open_by_key(imp.IDS_PLANILLAS[a.local])
    ws = imp._hoja_por_clave(sh, "factura")
    vals = ws.get(f"A1:X{ws.row_count}", value_render_option="FORMULA")
    res = analizar(vals)
    res.update({"local": a.local, "planilla": sh.title, "entorno": entornos.entorno(), "modo": "APPLY" if a.apply else "DRY-RUN"})
    print(json.dumps(res, ensure_ascii=False, indent=1))
    if not a.apply:
        return 0
    por_col: dict[str, list] = {"K": [], "T": [], "V": []}
    for i, fila in enumerate(vals[1:], 2):
        fila = list(fila) + [""] * (24 - len(fila))
        for col, L in ((10, "K"), (19, "T"), (21, "V")):
            f = fila[col]
            if isinstance(f, str) and f.startswith("=") and reparar_formula(f) != f:
                por_col[L].append((i, reparar_formula(f)))
    data = []
    for L, items in por_col.items():
        for i, f in items:
            data.append({"range": f"'{ws.title}'!{L}{i}", "values": [[f]]})
    for k in range(0, len(data), 2000):
        sh.values_batch_update({"valueInputOption": "USER_ENTERED", "data": data[k:k + 2000]})
    print(f"reparadas {len(data)} celdas en {sh.title}/{ws.title} (entorno PRUEBA)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
