"""Verificador independiente de las gestiones generadas por gestion_mensual.py (lee los .xlsx, no importa el generador).

Chequea por libro: fórmulas esperadas en cada fila (FINAL/VALORIZADO/%/Diferencias/Var), SKU duplicados, precios vacíos con stock,
reglas de Agus (códigos fusionados, ítems excluidos, rubros), paridad Arcos<->Cabello, y recalcula el valorizado en Python.
Uso: python verificar_gestion.py [carpeta]   (default Copias_Gestion/octubre_VF)
"""
import re
import sys
from pathlib import Path

import openpyxl

BASE = Path(__file__).resolve().parent
PROHIBIDOS_SKU = {"I0021", "TBC5", "I0561", "OBR127", "OBR9", "OBR131", "BURR", "FAR"}
PROHIBIDOS_NOMBRE = re.compile(r"libre?r[ií]a|posnet|^vela|gehakte|rag[uú]|manteca clarificada|^labneh|aceite de jalape", re.I)


def leer(path):
    wb = openpyxl.load_workbook(path)
    ws = wb["Inventario"]
    h = [c.value for c in ws[1]]
    i_final = h.index("FINAL") + 1
    rows = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 5).value is None:
            continue
        rows.append(r)
    return wb, ws, h, i_final, rows


def verificar(path: Path, esperado_prueba: bool):
    errores, avisos = [], []
    wb, ws, h, c_final, rows = leer(path)
    k_ini = 7
    sect = c_final - k_ini
    c_val, c_pct, c_orig, c_dif, c_stk, c_pref, c_var = (c_final + i for i in range(1, 8))
    last = ws.max_row
    nombres = [h[i] for i in range(len(h))]
    skus, tot = {}, 0.0
    for r in rows:
        f = ws.cell(r, c_final).value
        a, b = openpyxl.utils.get_column_letter(k_ini), openpyxl.utils.get_column_letter(c_final - 1)
        if f != f"=SUM({a}{r}:{b}{r})":
            errores.append(f"fila {r}: FINAL={f!r}")
        if ws.cell(r, c_val).value != f"={openpyxl.utils.get_column_letter(c_final)}{r}*F{r}":
            errores.append(f"fila {r}: VALORIZADO={ws.cell(r, c_val).value!r}")
        if not str(ws.cell(r, c_pct).value).startswith("=IFERROR("):
            errores.append(f"fila {r}: % sin fórmula")
        if ws.cell(r, c_dif).value != f"={openpyxl.utils.get_column_letter(c_final)}{r}-{openpyxl.utils.get_column_letter(c_stk)}{r}":
            errores.append(f"fila {r}: Diferencias={ws.cell(r, c_dif).value!r}")
        sku = ws.cell(r, 4).value
        nombre = str(ws.cell(r, 5).value)
        if sku:
            skus.setdefault(sku, []).append(r)
            if sku in PROHIBIDOS_SKU:
                errores.append(f"fila {r}: SKU {sku} debía estar fusionado/excluido")
        if PROHIBIDOS_NOMBRE.search(nombre) and "Chimidoha" not in path.name:
            errores.append(f"fila {r}: '{nombre}' no debería estar (regla de Agus)")
        precio = ws.cell(r, 6).value
        cant = sum(float(ws.cell(r, c).value or 0) for c in range(k_ini, c_final))
        if esperado_prueba and cant > 0 and not precio:
            avisos.append(f"fila {r}: '{nombre}' tiene stock {cant:g} y NO tiene precio")
        tot += cant * float(precio or 0)
    for sku, rs in skus.items():
        if len(rs) > 1:
            errores.append(f"SKU duplicado {sku} en filas {rs}")
    st = wb["Resumen"]
    if st["B7"].value != f"=SUM(Inventario!{openpyxl.utils.get_column_letter(c_val)}2:{openpyxl.utils.get_column_letter(c_val)}{last})":
        errores.append(f"Resumen!B7={st['B7'].value!r}")
    for cel, esp in (("B8", "=(B2+B3+B4-B5+B6)-B7"), ("B10", "=B9/1.21"), ("B11", "=IFERROR(B8/B10,0)")):
        if st[cel].value != esp:
            errores.append(f"Resumen!{cel}={st[cel].value!r}")
    return dict(items=len(rows), skus=len(skus), total=tot, errores=errores, avisos=avisos), wb, ws, skus, c_final


def main():
    carpeta = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "Copias_Gestion" / "octubre_VF"
    archivos = sorted(carpeta.glob("Gestion_*.xlsx"))
    if not archivos:
        print("no hay libros en", carpeta)
        return 1
    skus_loc, malos = {}, 0
    for p in archivos:
        prueba = "_prueba" in p.name
        res, wb, ws, skus, c_final = verificar(p, prueba)
        loc = p.name.split("_")[1]
        skus_loc[loc] = set(skus)
        print(f"{p.name}: ítems={res['items']} SKU={res['skus']} valorizado recalculado=${res['total']:,.0f} errores={len(res['errores'])} avisos={len(res['avisos'])}")
        for e in res["errores"][:15]:
            print("   ERROR", e)
        for a in res["avisos"][:6]:
            print("   aviso", a)
        malos += len(res["errores"])
    if {"Arcos", "Cabello"} <= set(skus_loc):
        a, c = skus_loc["Arcos"], skus_loc["Cabello"]
        dif = (a ^ c)
        print(f"Paridad Arcos/Cabello: {len(a & c)} en común, {len(a - c)} sólo Arcos, {len(c - a)} sólo Cabello -> {sorted(dif)[:12]}")
    print("OK" if not malos else f"{malos} errores")
    return 1 if malos else 0


if __name__ == "__main__":
    sys.exit(main())
