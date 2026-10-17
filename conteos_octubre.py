"""Conteos de octubre estandarizados al formato de Arcos (copias locales; no toca Drive).

Para Arcos / Cabello / Obrador(Congreso) / Chimidoha:
  * hojas de sector: conteos en blanco (se borran SOLO las celdas de cantidad), SKU unificado según las reglas de Agus
    (códigos fusionados, SKU faltantes resueltos por nombre), filas excluidas ocultas (no se borran),
  * hoja de totales: `Total2` (Arcos/Cabello/Chimidoha) o `Total` (Obrador: su GAS lee ese nombre) reconstruida con el layout y estilo de
    `Total2` de Arcos (RUBRO, UN., MEDIDA, CÓDIGO, PRODUCTO, <sectores>, FINAL), ítems = los de la gestión de octubre,
  * Cabello: se eliminan `Total` y `Total22` (duplicados con datos de septiembre; el GAS usa `Total2`).
  * Chimidoha (sin GAS): los totales por sector son fórmulas SUMIF sobre las hojas de sector, por SKU.
Salida: Copias_Gestion/octubre_VF/Conteo_<Local>_octubre.xlsx
"""
from __future__ import annotations

import re
from collections import defaultdict
from copy import copy

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

import gestion_mensual as gm

THIN = Side(style="thin")
BORDE = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
AZUL = PatternFill("solid", fgColor="9CC2E5")
ROSA = PatternFill("solid", fgColor="F4CCCC")

CFG = {
    "Arcos": dict(src="Conteo_Arcos.xlsx", total="Total2", primera=4, hoja_hdr=3, conteo={"2 DESPACHO": 6, "7 DESCARTABLES DEPO": 6}, default_col=5,
                  sectores=["1 SALON + BARRA", "2 DESPACHO", "3 DEPOSITO-ALMACEN", "4 ROTISERÍA", "5 DESCARTABLES ROTI", "6 DESCARTABLES BARRA", "7 DESCARTABLES DEPO", "8 CAJA", "9 BACHA", "10 CAMARA"],
                  restyle=False, borrar=[]),
    "Cabello": dict(src="Conteo_Cabello.xlsx", total="Total2", primera=2, hoja_hdr=1, conteo={}, default_col=5,
                    sectores=gm.SECTORES["Cabello"], restyle=True, borrar=["Total", "Total22"]),
    "Congreso": dict(src="Conteo_Obrador.xlsx", total="Total", primera=6, hoja_hdr=5, conteo={}, default_col=5,
                     sectores=gm.SECTORES["Congreso"], restyle=True, borrar=[]),
    "Chimidoha": dict(src="Conteo_Chimi_SKU.xlsx", total="Total2", primera=4, hoja_hdr=3, conteo={}, default_col=5,
                      sectores=["DESPACHO 1", "DEPOSITO - ALMACEN 2", "PLAZA FRIA-PASTELERIA 3", "CAMARA 4", "BARRA 5", "DESCARTABLES 6", "BACHA 7"], restyle=True, borrar=[]),
}
# columnas de sector del Total <- hoja(s) de sector (Arcos agrupa 8 CAJA/9 BACHA según el GAS: CAJA sólo I0021/BKP4 -> DESC. BARRA; BACHA se ignora)
MAP_TOTAL = {
    "Arcos": [["1 SALON + BARRA"], ["2 DESPACHO"], ["3 DEPOSITO-ALMACEN"], ["4 ROTISERÍA"], ["5 DESCARTABLES ROTI"], ["6 DESCARTABLES BARRA"], ["7 DESCARTABLES DEPO"], ["10 CAMARA"]],
    "Cabello": [[s] for s in gm.SECTORES["Cabello"]],
    "Congreso": [[s] for s in gm.SECTORES["Congreso"]],
    "Chimidoha": [["DESPACHO 1"], ["DEPOSITO - ALMACEN 2"], ["PLAZA FRIA-PASTELERIA 3"], ["CAMARA 4"], ["BARRA 5"], ["DESCARTABLES 6"], ["BACHA 7"]],
}


def reglas_para_hoja(local, F, df_local, log):
    """Devuelve (sku_nuevo(sku, nombre) -> sku, excluido(sku, nombre) -> motivo|None)."""
    por_nombre = {gm.norm(r.producto): r.sku for _, r in df_local.iterrows() if r.sku}
    pool = F.pool_nombres()

    def sku_nuevo(sku, nombre):
        sku = (sku or "").strip().upper()
        if sku in gm.REGLAS["fusionar_sku"]:
            return gm.REGLAS["fusionar_sku"][sku][0]
        if not sku:
            for pat, s, _ in gm.REGLAS["sku_por_nombre"]:
                if re.search(pat, nombre or "", re.I):
                    return s
            hit = por_nombre.get(gm.norm(nombre))
            if hit:
                return hit
            h = F.buscar_sku_por_nombre(nombre or "", pool)
            return h[0] if h else ""
        return sku

    def excluido(sku, nombre):
        m = gm.REGLAS["excluir_sku"].get((local if local != "Congreso" else "Congreso", (sku or "").upper()))
        if m:
            return m
        loc = local if local != "Congreso" else "Congreso"
        for locs, pat, why in gm.REGLAS["excluir_nombre"]:
            if ("*" in locs or loc in locs) and re.search(pat, nombre or "", re.I):
                return why
        return None

    return sku_nuevo, excluido


def procesar_hojas_sector(wb, local, F, df_local):
    cfg = CFG[local]
    sku_nuevo, excluido = reglas_para_hoja(local, F, df_local, None)
    cambios = []
    for nombre in cfg["sectores"]:
        if nombre not in wb.sheetnames:
            continue
        ws = wb[nombre]
        ccol = cfg["conteo"].get(nombre, cfg["default_col"])
        if local == "Arcos" and nombre == "1 SALON + BARRA":
            ws["A2"].value = None  # 'Completar fecha'
        # cantidades: se borran TODAS (también filas sin nombre = productos no listados que se cargaron durante el mes)
        for r in range(cfg["primera"], ws.max_row + 1):
            cell = ws.cell(r, ccol)
            v = cell.value
            if v is not None and (isinstance(v, (int, float)) or (isinstance(v, str) and v.startswith("="))):
                try:
                    cell.value = None
                except AttributeError:
                    pass
        for r in range(cfg["primera"], ws.max_row + 1):
            sku, prod = ws.cell(r, 2).value, ws.cell(r, 4).value
            un = ws.cell(r, 1).value
            if prod is None or str(prod).strip() == "" or (un is None and not sku):
                continue
            if str(un).strip().upper().startswith("UN") and str(ws.cell(r, 3).value or "").strip().upper() == "MEDIDA":
                continue
            # borrar SÓLO la cantidad
            cell = ws.cell(r, ccol)
            try:
                if cell.value is not None:
                    cell.value = None
            except AttributeError:   # celda combinada
                pass
            ns = sku_nuevo(str(sku or ""), str(prod))
            if ns != (str(sku).strip().upper() if sku else ""):
                ws.cell(r, 2).value = ns or None
                cambios.append((nombre, r, str(sku or ""), ns, str(prod)))
            mot = excluido(ns or str(sku or ""), str(prod))
            if mot:
                ws.row_dimensions[r].hidden = True
                cambios.append((nombre, r, str(sku or ""), "OCULTA", f"{prod}: {mot}"))
        if cfg["restyle"]:
            restyle_sector(ws, cfg, ccol)
    return cambios


def restyle_sector(ws, cfg, ccol):
    """Mismo estilo que las hojas de sector de Arcos (Arial 14, encabezado celeste, bordes finos)."""
    h = cfg["hoja_hdr"]
    last = h
    for r in range(h + 1, ws.max_row + 1):
        if ws.cell(r, 4).value not in (None, "") or ws.cell(r, 3).value not in (None, ""):
            last = r
    for c in range(1, 6):
        x = ws.cell(h, c)
        x.font, x.fill, x.border = Font(name="Arial", size=14, bold=True), AZUL, BORDE
        x.alignment = Alignment(horizontal="center" if c == 1 else None, wrap_text=True)
    for r in range(h + 1, last + 1):
        for c in range(1, 6):
            x = ws.cell(r, c)
            if type(x).__name__ == "MergedCell":
                continue
            x.font, x.border = Font(name="Arial", size=14, bold=True), BORDE
            x.alignment = Alignment(horizontal="center" if c in (1, ccol) else None, wrap_text=True)
    for col, w in (("A", 10.1), ("B", 13.9), ("C", 12.0), ("D", 59.9), ("E", 14.0)):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = f"A{h + 1}"


def construir_total(wb, local, df_local, F):
    cfg = CFG[local]
    for n in cfg["borrar"]:
        if n in wb.sheetnames:
            del wb[n]
    nombre = cfg["total"]
    idx = wb.sheetnames.index(nombre) if nombre in wb.sheetnames else len(wb.sheetnames)
    if nombre in wb.sheetnames:
        del wb[nombre]
    ws = wb.create_sheet(nombre, idx)
    sect = gm.SECTORES[local]
    heads = ["RUBRO", "UN.", "MEDIDA", "CÓDIGO", "PRODUCTO"] + sect + ["FINAL"]
    for j, h in enumerate(heads, start=1):
        c = ws.cell(1, j, h)
        c.font, c.fill = Font(name="Arial", bold=True), ROSA
    df = df_local.assign(_r=df_local.rubro.str.upper(), _p=df_local.producto.str.upper()).sort_values(["_r", "_p"])
    k = len(sect)
    mapa = MAP_TOTAL[local]
    for i, (_, r) in enumerate(df.iterrows(), start=2):
        for j, v in enumerate([r.rubro, r.un, r.medida, r.sku or None, r.producto], start=1):
            ws.cell(i, j, v).font = Font(name="Arial")
        for s in range(k):
            c = ws.cell(i, 6 + s)
            c.font = Font(name="Arial")
            if local == "Chimidoha":
                if r.sku:
                    partes = [f"SUMIF('{h}'!$B:$B,$D{i},'{h}'!$E:$E)" for h in mapa[s]]
                else:   # ítem sin SKU: se suma por nombre exacto (TRIM por los espacios finales de los conteos)
                    partes = [f"SUMPRODUCT((TRIM('{h}'!$D$4:$D$200)=TRIM($E{i}))*1,'{h}'!$E$4:$E$200)" for h in mapa[s]]
                c.value = "=" + "+".join(partes)
        f = ws.cell(i, 6 + k, f"=sum({L(6)}{i}:{L(5 + k)}{i})")
        f.font = Font(name="Arial")
    ws.freeze_panes = "A2"
    ws.column_dimensions["B"].width = 7.2
    ws.column_dimensions["E"].width = 36.2
    return ws


def main():
    F = gm.Fuentes(10, 2026)
    med = gm._medida_por_sku(F)
    from collections import defaultdict as dd
    log = dd(list)
    dfs = {l: gm.aplicar_reglas(F, l, gm.items_base(F, l, med), log[l]) for l in gm.SECTORES}
    gm.espejar_arcos_cabello(dfs, log)
    gm.OUT.mkdir(parents=True, exist_ok=True)
    for local, cfg in CFG.items():
        wb = openpyxl.load_workbook(gm.WORK / cfg["src"])
        cambios = procesar_hojas_sector(wb, local, F, dfs[local])
        construir_total(wb, local, dfs[local], F)
        # hoja de cambios aplicados (trazabilidad)
        if "Cambios octubre" in wb.sheetnames:
            del wb["Cambios octubre"]
        wc = wb.create_sheet("Cambios octubre")
        for j, h in enumerate(["Hoja", "Fila", "SKU anterior", "SKU nuevo / acción", "Producto / motivo"], start=1):
            c = wc.cell(1, j, h)
            c.font, c.fill = Font(name="Arial", bold=True), ROSA
        for i, row in enumerate(cambios, start=2):
            for j, v in enumerate(row, start=1):
                wc.cell(i, j, v)
        wc.column_dimensions["E"].width = 90
        out = gm.OUT / f"Conteo_{ {'Congreso': 'Obrador'}.get(local, local)}_octubre.xlsx"
        wb.save(out)
        print(f"{local:10s} -> {out.name}: {len(cambios)} cambios en hojas de sector; hojas={wb.sheetnames[:12]}")


if __name__ == "__main__":
    main()
