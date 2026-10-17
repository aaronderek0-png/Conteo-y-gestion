"""Gestión mensual Mishiguene — generador local (sin tocar Drive de producción).

Arma, para Arcos / Cabello / Lafinur / Congreso(Obrador) / Chimidoha, el libro `Gestion_<Local>_<mes>.xlsx`
con el formato de Arcos (hoja Inventario), a partir de:
  * cierre del mes anterior (gestión de septiembre, copias locales en Copias_Gestion/octubre_trabajo/),
  * Maestro de Productos (SKU, nombre, rubro, unidad),
  * lista de precios Obrador (PRECIOS Auto / Precios_Actuales + Recetas!Obrador COSTO),
  * Facturas de cada local (cache en octubre_trabajo/cache/facturas_<Local>.pkl),
  * correcciones de Agus (reglas en REGLAS, abajo).

Orden de precio por ítem (el GAS de precios pisaba G con F cuando no había factura; acá no):
  1. Ítem de transferencia del Obrador (con markup)  -> precio Obrador del local (Lafinur=LAFINUR; cafés/Chimi=CHIMIDOHA).
  2. Factura del local en el mes en curso.
  3. Precio de la lista Obrador (insumo pasamano, coef 1.0).
  4. Factura del local en los 3 meses previos.
  5. Maestro de Productos (P.U).
  6. Precio del cierre anterior (se mantiene, queda marcado).
Cualquier variación > ±50 % contra el cierre anterior NO se inyecta: queda el precio anterior y la fila va a `Controles`.
Congreso valoriza su producción propia al COSTO del Obrador (Recetas!Obrador) — sin margen de transferencia.

Uso:  python gestion_mensual.py [--mes 10 --anio 2026] [--prueba]   (--prueba rellena conteos con los actuales para validar)
Solo escribe en Copias_Gestion/octubre_VF/.  Nada se sube a Drive.
"""
from __future__ import annotations

import argparse
import re
import unicodedata
from collections import OrderedDict, defaultdict
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from rapidfuzz import fuzz

BASE = Path(__file__).resolve().parent
WORK = BASE / "Copias_Gestion" / "octubre_trabajo"
CACHE = WORK / "cache"
OUT = BASE / "Copias_Gestion" / "octubre_VF"

MESES = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

SECTORES = {
    "Arcos": ["BARRA / SALON", "COCINA", "DEPO", "ROTISERIA", "DESCARTABLES ROTI", "DESCARTABLES BARRA", "DESCARTABLES DEPO", "CAMARA"],
    "Cabello": ["DESPACHO CALIENTE", "DELY", "PASTELERIA", "PRODUCCION - CAMARA - ALMACEN", "BARRA - DEPO", "BARRA", "SALON"],
    "Lafinur": ["CONTEO"],
    "Congreso": ["PRODUCCIÓN", "PASTELERÍA", "DEPÓSITO"],
    "Chimidoha": ["DESPACHO", "DEPOSITO - ALMACEN", "PLAZA FRIA-PASTELERIA", "CAMARA", "BARRA", "DESCARTABLES", "BACHA"],
}
# columna de la lista de precios Obrador que corresponde a cada local
COL_OBRADOR = {"Lafinur": "PRECIO_LAFINUR", "Arcos": "PRECIO_CHIMIDOHA", "Cabello": "PRECIO_CHIMIDOHA", "Chimidoha": "PRECIO_CHIMIDOHA", "Congreso": "COSTO"}

# ───────────────────────────── Reglas de Agus (feedback 14/10 sobre los mensuales de septiembre) ─────────────────────────────
REGLAS = {
    # códigos distintos para lo mismo -> se unifican (stock se suma). clave = código que se descarta
    "fusionar_sku": {
        "I0021": ("I0025", "Bolsa Kraft N5 (Laminados) = Bolsa Kraft 5: mismo producto con dos códigos"),
        "TBC5": ("I0560", "Tapa bowl compostable 500 cc = Tapa bowl biodegradable 500 cc"),
        "I0561": ("I0560", "Tapa bowl ovalado 500 cc = Tapa bowl biodegradable 500 cc"),
        "OBR127": ("BAMAPOLA", "Bagel amapola: Cabello usaba otro código que Arcos"),
        "OBR9": ("BCEBOLLA", "Bagel cebolla: Cabello usaba otro código que Arcos"),
        "OBR131": ("BSESAMO", "Bagel sésamo: Cabello usaba otro código que Arcos"),
        "BURR": ("PHBURR", "Philas de burrecas = Phila de ricota y queso (código de Arcos)"),
        "FAR": ("FER", "Farfalaj: dos códigos en Arcos para el mismo producto (caja de 400 g)"),
        "OBR20": ("OBR3031", "Guefilte: el código vigente en Recetas/lista Obrador es OBR3031 (memoria 15/10)"),
        "POGUEFI": ("OBR3031", "Guefilte: código vigente OBR3031"),
    },
    # (local|*, regex sobre nombre, motivo): no deben figurar en el conteo del local
    "excluir_nombre": [
        (("Arcos", "Cabello"), r"jalape[nñ]o.*aceite|aceite.*jalape[nñ]o|^labneh|gehakte|rag[uú]|manteca clarificada", "Producción propia: no va en el conteo (Agus)"),
        (("*",), r"libre?r[ií]a|rollo posnet|posnet|^vela", "No corresponde al inventario (librería/posnet/velas) (Agus)"),
    ],
    # (local), sku -> motivo: Arcos/Cabello no compran Harina 0000 (la compran al Obrador como 000)
    "excluir_sku": {("Arcos", "H0000"): "No compran Harina 0000; compran Harina 000 al Obrador (Agus)",
                    ("Cabello", "H0000"): "No compran Harina 0000; compran Harina 000 al Obrador (Agus)"},
    # rubro canónico forzado por SKU (Agus: figuraban como Almacén)
    "rubro": {"M24": "OBRADOR", "M254": "OBRADOR", "M246": "OBRADOR", "MO2015": "OBRADOR", "VDVB": "Almacén"},
    # SKU forzado por nombre (ítems sin código o con código repetido para cosas distintas)
    "sku_por_nombre": [
        (r"vino bikurim", "VBIK", "Vino Bikurim sin código en Cabello = VBIK (el mismo que Arcos)"),
    ],
    # ítems que Arcos tiene y Cabello no (y viceversa) por carta — no se espejan
    "no_espejar": {"Arcos": {"TSAL"}, "Cabello": set()},
    # notas de Agus para que la persona que cierra las verifique (SKU -> texto) — van a Controles y a Obs
    "verificar": {
        "OBR3031": "Guefiltefish: ¿por qué aumentó tanto? (Agus)",
        "OBR20": "Guefiltefish código viejo: usar OBR3031 (Recetas)",
        "M254": "Shakshuka verde: de 3 mil a 8 mil — validar costo en Obrador (Agus)",
        "SING": "Salsa inglesa: el precio venía mal desde agosto (6 mil en vez de 5 mil y pico) (Agus)",
        "FER": "Farfalaj: la caja de 400 g sale $7.000 — el conteo debe ser en cajas o en kg reales (en el conteo figuraban 360 g) (Agus)",
        "TTRIG": "Tomate perita italiano: última compra Alyser, lata 2,5 kg = $9.816,80 (Agus)",
        "MD200": "Manga descartable Alyser: el paquete trae 100 mangas, precio por unidad = paquete/100 (Agus)",
        "I0177": "Papel parafinado: el precio no cambió para Arcos; confirmar contra Obrador (Agus)",
        "ALC": "Alcaparras: se cobran desde Obrador, no usar precio anterior (Agus)",
        "I0538": "Nicassia Cabernet Malbec de $7.893 (agosto) a $48.843: chequear (Agus)",
        "I0283": "Serie A (bonarda): precio sospechoso, chequear facturas de junio/septiembre (Agus)",
        "I0284": "Serie A (chardonay): precio sospechoso, chequear facturas (Agus)",
        "I0286": "Serie A (torrontés): precio sospechoso, chequear facturas (Agus)",
        "RDCL": "Roll de canela: precio de Cabello no aparece en facturas Congreso (Agus)",
        "QFE": "Queso feta 20.420: no aparece en Obrador, ¿a qué corresponde? (Agus)",
        "QSE": "Queso sardo 13.730: no aparece en Obrador; Agus lo usará igual en Arcos (Agus)",
        "QCTI": "Queso crema típico: quedó desactualizado en Cabello (Agus)",
        "VBIK": "Vino Bikurim/Pesaj: diferencia de precio Arcos vs Cabello arrastrada varios meses (Agus)",
        "I0502": "Bowl compostable cholent 960 cc (más caro): chequear stock — decía 67 unidades (Agus)",
        "PLETZ": "Pletzalej: Cabello quedó con precio desactualizado (Agus)",
        "LAFFA": "Laffa: Cabello quedó con precio desactualizado (Agus)",
        "PITA": "Pita: Cabello quedó con precio desactualizado (Agus)",
        "PHES": "Phila espinaca: Cabello quedó con precio desactualizado (Agus)",
    },
}

# ───────────────────────────── utilidades ─────────────────────────────

def norm(s) -> str:
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().upper()
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"[^A-Z0-9 ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def norm_obr(s) -> str:
    s = unicodedata.normalize("NFD", str(s or "")).encode("ascii", "ignore").decode().upper().strip()
    s = re.sub(r"^(GR|UN|UND|PORC|KG|LT|ML)\.?\s+", "", s)
    return norm(s)


def num(x):
    try:
        v = float(x)
        return None if v != v else v
    except (TypeError, ValueError):
        return None


MASA = {"kg", "kilo", "kilos", "lt", "lts", "litro", "litros", "l", "lt."}
UNID = {"un", "und", "unid", "unidad", "unidades", "botella", "frasco", "caja", "pote", "rollo", "porc", "paquete", "paq", "pack", "atado", "bandeja", "balde", "lata", "docena"}


def kind_medida(m) -> str:
    m = str(m or "").strip().lower().rstrip(".")
    if m in MASA:
        return "masa"
    if m in UNID:
        return "unid"
    return "?"


# ───────────────────────────── carga de fuentes ─────────────────────────────

class Fuentes:
    def __init__(self, mes: int, anio: int):
        self.mes, self.anio = mes, anio
        ro = lambda f, sh: pd.read_excel(WORK / f, sheet_name=sh)
        # Maestro
        m = ro("Maestro_RO.xlsx", "Productos")
        m = m.rename(columns={m.columns[1]: "producto", m.columns[3]: "rubro", m.columns[5]: "um", m.columns[7]: "pu"})
        m["sku"] = m["SKU"].astype(str).str.strip().str.upper()
        self.maestro = m.drop_duplicates("sku").set_index("sku")[["producto", "rubro", "um", "pu"]]
        # lista de precios Obrador (vigente) + historial
        pa = ro("PRECIOS_Auto.xlsx", "Precios_Actuales")
        cc = ro("PRECIOS_Auto.xlsx", "Control_Cambios")[["CODIGO_MAXIREST", "CATEGORIA", "COEF_CHIMI"]]
        pa = pa.merge(cc, on="CODIGO_MAXIREST", how="left")
        pa = pa[pa.CODIGO_OBRADOR.notna()].copy()
        pa["sku"] = pa.CODIGO_OBRADOR.astype(str).str.strip().str.upper()
        pa["kind"] = pa.DESCRIPCION.astype(str).str.extract(r"^\s*([A-Za-z]+)\.")[0].str.lower().fillna("")
        pa["nname"] = pa.DESCRIPCION.map(norm_obr)
        self.obr = pa.drop_duplicates("sku", keep="last").set_index("sku")
        self.obr_by_name = defaultdict(list)
        for sku, r in self.obr.iterrows():
            self.obr_by_name[r.nname].append(sku)
        # costo Obrador (Recetas!Obrador)
        o = ro("Recetas_RO.xlsx", "Obrador")
        o["sku"] = o["SKU"].astype(str).str.strip().str.upper()
        o["nname"] = o["NOMBRE"].map(norm_obr)
        o["COSTO"] = pd.to_numeric(o["COSTO"], errors="coerce")
        o["kind"] = o["NOMBRE"].astype(str).str.upper().map(lambda s: "gr" if "GRAMO" in s else ("un" if ("UNIDAD" in s or "UN)" in s) else ("kg" if "1 KG" in s else "")))
        self.obr_tab = o.drop_duplicates("sku").set_index("sku")
        self.obr_tab_by_name = defaultdict(list)
        for sku, r in self.obr_tab.iterrows():
            self.obr_tab_by_name[r.nname].append(sku)
        # facturas
        self.fact = {}
        for loc in SECTORES:
            p = CACHE / f"facturas_{loc}.pkl"
            if not p.exists():
                continue
            f = pd.read_pickle(p)
            f["fecha"] = pd.to_datetime(f.f_ingreso.where(f.f_ingreso.notna(), f.f_factura), errors="coerce")
            f["p"] = pd.to_numeric(f.precio_un_inv, errors="coerce")
            f["sku"] = f.sku.astype(str).str.strip().str.upper()
            f = f[(f.p > 0) & f.fecha.notna() & ~f.sku.isin(["", "NAN", "NONE", "SIN SKU"]) & ~f.sku.str.contains("#N/A|ERROR", na=False)]
            self.fact[loc] = f.sort_values(["fecha"]).drop_duplicates("sku", keep="last").set_index("sku")  # la más reciente por SKU
            self.fact[loc + "_todas"] = f
        self.sept = pd.read_pickle(CACHE / "sept_items.pkl")
        # Chimidoha: compra centralizada en Congreso -> precio de transferencia = factura de Congreso x recargo por rubro (CLAUDE.md, matriz de recargos)
        if "Congreso_todas" in self.fact:
            cg = self.fact["Congreso_todas"].copy()
            def coef_rubro(sku):
                rub = str(self.maestro.loc[sku, "rubro"]).lower() if sku in self.maestro.index else ""
                if any(k in rub for k in ("descartable", "paniﬁc", "panific", "harinera")):
                    return 1.2
                if any(k in rub for k in ("carne", "pescado")):
                    return 1.1
                return 1.0
            cg["p"] = cg["p"] * cg["sku"].map(coef_rubro)
            propio = self.fact.get("Chimidoha_todas")
            both = pd.concat([cg, propio]) if propio is not None else cg
            self.fact["Chimidoha_todas"] = both.sort_values("fecha")

    # ── SKU por nombre
    def pool_nombres(self):
        pool = defaultdict(set)
        for sku, r in self.maestro.iterrows():
            pool[norm(r.producto)].add(sku)
        for sku, r in self.obr.iterrows():
            pool[r.nname].add(sku)
        for sku, r in self.obr_tab.iterrows():
            pool[r.nname].add(sku)
        return {k: v for k, v in pool.items() if k}

    def buscar_sku_por_nombre(self, nombre, pool):
        n = norm_obr(nombre)
        if n in pool and len(pool[n]) == 1:
            return next(iter(pool[n])), "exacto", n
        best, bs = None, 0
        for k in pool:
            sc = fuzz.token_sort_ratio(n, k)
            if sc > bs:
                best, bs = k, sc
        if bs >= 93 and len(pool[best]) == 1:
            ta, tb = set(n.split()), set(best.split())
            if all(any(fuzz.ratio(x, y) >= 80 for y in tb) for x in ta if len(x) > 2 or x.isdigit()):
                return next(iter(pool[best])), f"fuzzy {bs:.0f}", best
        return None

    # ── bridges
    def obrador_entry(self, sku: str, nombre: str):
        """Devuelve (sku_obrador, fila lista de precios) o (None, None)."""
        if sku in self.obr.index:
            return sku, self.obr.loc[sku]
        n = norm_obr(nombre)
        c = self.obr_by_name.get(n, [])
        if len(c) >= 1:
            return c[0], self.obr.loc[c[0]]
        best, bs = None, 0
        for k in self.obr_by_name:
            s = fuzz.token_sort_ratio(n, k)
            if s > bs:
                best, bs = k, s
        if bs >= 93:
            c = self.obr_by_name[best]
            return c[0], self.obr.loc[c[0]]
        return None, None

    def costo_tab(self, sku: str, nombre: str):
        if sku in self.obr_tab.index and pd.notna(self.obr_tab.loc[sku, "COSTO"]):
            return sku, self.obr_tab.loc[sku]
        c = self.obr_tab_by_name.get(norm_obr(nombre), [])
        if c and pd.notna(self.obr_tab.loc[c[0], "COSTO"]):
            return c[0], self.obr_tab.loc[c[0]]
        return None, None


# ───────────────────────────── armado de ítems por local ─────────────────────────────

def _medida_por_sku(F: Fuentes) -> dict:
    """Unidad de gestión por SKU tomada de los locales que sí la tienen (Cabello no trae columna MEDIDA)."""
    d = {}
    for _, r in F.sept.iterrows():
        m = str(r.medida or "").strip()
        if m and m.upper() != "NONE" and r.sku and r.sku not in d:
            d[r.sku] = m
    return d


def _canon_rubro(r, sku, F: Fuentes) -> str:
    if sku in REGLAS["rubro"]:
        return REGLAS["rubro"][sku]
    r = str(r or "").strip()
    if r.upper() == "OBRADOR":
        return "OBRADOR"
    if sku in F.maestro.index and isinstance(F.maestro.loc[sku, "rubro"], str):
        return F.maestro.loc[sku, "rubro"].strip()
    return r[:1].upper() + r[1:] if r else ""


def items_base(F: Fuentes, local: str, medida_sku: dict) -> pd.DataFrame:
    if local == "Chimidoha":
        mp = pd.read_csv(WORK / "mapeo_sku_chimi_v3.csv", encoding="utf-8-sig").fillna("")
        mp = mp[~mp.medida.str.upper().isin(["CANTIDAD", "MEDIDA"])]
        rows = []
        for _, r in mp.iterrows():
            rows.append(dict(sku=r.sku.upper(), producto=r.producto.strip(), rubro="", un=1, medida=r.medida.strip(), final_sept=0.0, p_sept=None,
                             sect=[0.0] * len(SECTORES[local]), hoja=r.hoja, fuente="conteo Chimidoha"))
        return pd.DataFrame(rows)
    d = F.sept[F.sept.local == local]
    rows = []
    for _, r in d.iterrows():
        sect = list(r.sect) + [0.0] * (len(SECTORES[local]) - len(r.sect))
        if not sum(sect) and (r.final or 0) > 0:      # Cabello cargó el cierre sólo en FINAL (sin abrir por sector)
            sect[0] = float(r.final)
        med = str(r.medida or "").strip()
        if not med or med.upper() == "NONE":
            med = medida_sku.get(r.sku, "")
        rows.append(dict(sku=r.sku, producto=r.producto, rubro=r.rubro, un=r.un if r.un is not None else 1, medida=med, final_sept=r.final or 0.0,
                         p_sept=r.p_act if r.p_act else None, sect=sect, hoja="", fuente="gestión sept"))
    df = pd.DataFrame(rows)
    if local == "Congreso":  # producción propia del Obrador: sale del conteo Obrador (Total) y hoy no se valoriza
        ws = openpyxl.load_workbook(WORK / "Conteo_Obrador.xlsx", data_only=True)["Total"]
        have = set(df.sku)
        add = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            sku = str(r[3] or "").strip().upper()
            if not sku or sku in have or sku == "NO ENCONTRADO" or not r[4]:
                continue
            have.add(sku)
            add.append(dict(sku=sku, producto=str(r[4]).strip(), rubro=r[0] or "OBRADOR", un=num(r[1]) or 1, medida=str(r[2] or "").strip(), final_sept=num(r[8]) or 0.0,
                            p_sept=None, sect=[num(x) or 0.0 for x in r[5:8]], hoja="", fuente="conteo Obrador (producción)"))
        df = pd.concat([df, pd.DataFrame(add)], ignore_index=True)
    return df


def aplicar_reglas(F: Fuentes, local: str, df: pd.DataFrame, log: list) -> pd.DataFrame:
    df = df.copy()
    # 1) SKU por nombre cuando falta
    for i, r in df.iterrows():
        if not r.sku:
            for pat, sku, why in REGLAS["sku_por_nombre"]:
                if re.search(pat, r.producto, re.I):
                    df.at[i, "sku"] = sku
                    log.append(("SKU asignado", sku, r.producto, why))
    # 1b) SKU por nombre contra Maestro / lista Obrador / costos Obrador (match exacto o fuzzy >= 93 con verificación de tokens)
    pool = F.pool_nombres()
    for i, r in df.iterrows():
        if df.at[i, "sku"]:
            continue
        hit = F.buscar_sku_por_nombre(r.producto, pool)
        if hit:
            df.at[i, "sku"] = hit[0]
            log.append(("SKU asignado (por nombre)", hit[0], r.producto, f"match {hit[1]} con '{hit[2]}'"))
    # 2) fusionar códigos duplicados
    for i, r in df.iterrows():
        if r.sku in REGLAS["fusionar_sku"]:
            nuevo, why = REGLAS["fusionar_sku"][r.sku]
            log.append(("SKU unificado", r.sku + " → " + nuevo, r.producto, why))
            df.at[i, "sku"] = nuevo
    # 3) exclusiones
    keep = []
    for i, r in df.iterrows():
        motivo = REGLAS["excluir_sku"].get((local, r.sku))
        if not motivo:
            for locs, pat, why in REGLAS["excluir_nombre"]:
                if ("*" in locs or local in locs) and re.search(pat, r.producto, re.I):
                    motivo = why
                    break
        if motivo:
            extra = f" (stock cierre sept {r.final_sept:g})" if r.final_sept else ""
            log.append(("Ítem excluido", r.sku, r.producto + extra, motivo))
        keep.append(not motivo)
    df = df[keep].copy()
    # 4) agrupar por SKU (suma de stock, primera ocurrencia manda); sin SKU se agrupa por nombre
    df["_k"] = df.apply(lambda r: r.sku or "~" + norm(r.producto), axis=1)
    out = []
    for k, g in df.groupby("_k", sort=False):
        r0 = g.iloc[0].copy()
        if len(g) > 1:
            r0["final_sept"] = float(g.final_sept.sum())
            r0["sect"] = [sum(x) for x in zip(*g.sect)]
            ps = [p for p in g.p_sept if p]
            r0["p_sept"] = ps[0] if ps else None
            if r0.sku:
                log.append(("Ítems agrupados", r0.sku, " + ".join(g.producto), f"mismo SKU repetido {len(g)} veces: se suma el stock"))
        out.append(r0)
    df = pd.DataFrame(out).drop(columns="_k").reset_index(drop=True)
    df["rubro"] = [_canon_rubro(r.rubro, r.sku, F) for _, r in df.iterrows()]
    for _, r in df[df.sku.isin(REGLAS["rubro"])].iterrows():
        log.append(("Rubro corregido", r.sku, r.producto, f"rubro = {r.rubro}"))
    return df


def espejar_arcos_cabello(dfs: dict, log: list):
    """Arcos y Cabello deben tener casi los mismos ítems (stock 0 en el que no lo tenía)."""
    for a, b in (("Arcos", "Cabello"), ("Cabello", "Arcos")):
        sb = set(dfs[b].sku)
        add = []
        for _, r in dfs[a].iterrows():
            if not r.sku or r.sku in sb or r.sku in REGLAS["no_espejar"].get(a, set()):
                continue
            n = len(SECTORES[b])
            add.append(dict(sku=r.sku, producto=r.producto, rubro=r.rubro, un=r.un, medida=r.medida, final_sept=0.0, p_sept=r.p_sept, sect=[0.0] * n, hoja="", fuente=f"espejo de {a}"))
            log[b].append(("Ítem agregado (espejo)", r.sku, r.producto, f"{a} lo tiene y {b} no: se agrega con stock 0 (Agus: casi los mismos ítems)"))
        if add:
            dfs[b] = pd.concat([dfs[b], pd.DataFrame(add)], ignore_index=True)


# ───────────────────────────── precios ─────────────────────────────

def conv_obrador(price, kind, medida, p_ref):
    """Convierte precio de la lista Obrador (por gramo / unidad / porción) a la unidad de la gestión. -> (precio, nota) o (None, motivo)."""
    mk = kind_medida(medida)
    k = (kind or "").lower()
    if k == "gr":
        if mk == "masa":
            return price * 1000.0, "g→kg ×1000"
        return None, f"lista en $/g pero la gestión usa '{medida}'"
    if k in ("un", "und", "porc"):
        if mk in ("unid", "?"):
            return price, ""
        return None, f"lista en $/{k} pero la gestión usa '{medida}'"
    # sin prefijo: ya viene en unidad de insumo; salvo que difiera ~1000 veces del precio de referencia
    if p_ref and p_ref > 0 and mk == "masa":
        r = price / p_ref
        if 700 < r < 1400:
            return price / 1000.0, "auto ÷1000 (lista en unidad base)"
        if 0.0007 < r < 0.0014:
            return price * 1000.0, "auto ×1000 (lista en unidad base)"
    if mk == "masa" and price < 300 and (not p_ref or price / p_ref < 0.05):
        return price * 1000.0, "auto ×1000 ($/g sin prefijo)"
    return price, ""


def conv_costo_tab(costo, kind, medida):
    mk = kind_medida(medida)
    if kind == "gr":
        return (costo * 1000.0, "costo $/g→kg ×1000") if mk == "masa" else (None, f"costo en $/g pero la gestión usa '{medida}'")
    if kind == "un":
        return (costo, "") if mk in ("unid", "?") else (None, f"costo por unidad pero la gestión usa '{medida}'")
    if kind == "kg":
        return (costo, "") if mk == "masa" else (None, f"costo por kg pero la gestión usa '{medida}'")
    return (costo, "") if mk == "unid" else (None, "costo sin unidad reconocida")


def maestro_a_unidad(F: Fuentes, sku, medida):
    if sku not in F.maestro.index:
        return None
    pu, um = num(F.maestro.loc[sku, "pu"]), str(F.maestro.loc[sku, "um"] or "").strip().lower()
    if not pu or pu <= 0:
        return None
    mk = kind_medida(medida)
    if um in ("g", "gr", "ml") and mk == "masa":
        return pu * 1000.0
    if um in ("un", "und", "unid") and mk in ("unid", "?"):
        return pu
    return None


def resolver_precio(F: Fuentes, local: str, r, mes: int, anio: int) -> dict:
    sku, nombre, medida, p_ref = r.sku, r.producto, r.medida, num(r.p_sept)
    p_ref = p_ref if p_ref and p_ref > 0 else None
    res = dict(precio=None, origen="", estado="OK", obr_sku="", nota="", candidato=None)
    ini_mes = pd.Timestamp(anio, mes, 1)
    todas = F.fact.get(local + "_todas")
    inv_cur = inv_prev = None
    if todas is not None and sku:
        s = todas[todas.sku == sku].sort_values("fecha")
        cur = s[s.fecha >= ini_mes]
        prev = s[(s.fecha < ini_mes) & (s.fecha >= ini_mes - pd.DateOffset(months=3))]
        if len(cur):
            inv_cur = cur.iloc[-1]
        elif len(prev):
            inv_prev = prev.iloc[-1]

    def fmt_f(f, cur):
        d = f.fecha
        return f"Factura: {d:%d/%m/%Y}" if cur else f"Histórico: {MESES[d.month]} ({d:%d/%m})"

    cands = []
    entry_sku, entry = F.obrador_entry(sku, nombre)
    transfer, pas, coef = False, None, None
    if entry is not None:
        res["obr_sku"] = entry_sku
        coef = num(entry.get("COEF_CHIMI"))
        transfer = bool(coef and coef > 1.0001)
        col = COL_OBRADOR[local] if COL_OBRADOR[local] in entry.index else "PRECIO_LAFINUR"
        p0 = num(entry[col])
        if p0:
            pc, nota = conv_obrador(p0, entry["kind"], medida, p_ref)
            if pc is None:
                res["nota"] = "Obrador: " + nota
            else:
                pas = (pc, f"Obrador {col.replace('PRECIO_', '')} ({entry_sku}){' ' + nota if nota else ''}")
    if local == "Congreso":
        ts, te = F.costo_tab(sku, nombre)
        if te is not None:
            pc, nota = conv_costo_tab(float(te["COSTO"]), te["kind"], medida)
            if pc is not None:
                cands.append((pc, f"Costo Obrador {ts} (sin margen){' ' + nota if nota else ''}"))
            else:
                res["nota"] = (res["nota"] + " | " if res["nota"] else "") + "Costo Obrador: " + nota
        elif pas and transfer:
            cands.append((pas[0], pas[1] + " (producción valorizada al costo)"))
    if not cands:
        if transfer and pas:
            cands.append(pas)
        elif inv_cur is not None:
            cands.append((float(inv_cur.p), fmt_f(inv_cur, True)))
        elif pas:
            cands.append((pas[0], pas[1].replace("Obrador", "Lista Obrador (pasamano)", 1)))
        elif inv_prev is not None:
            cands.append((float(inv_prev.p), fmt_f(inv_prev, False)))
        else:
            pm = maestro_a_unidad(F, sku, medida)
            if pm:
                cands.append((pm, "Maestro de Productos (P.U)"))
    if cands:
        p, o = cands[0]
        res["candidato"] = p
        oficial = o.startswith(("Obrador", "Costo Obrador", "Lista Obrador"))
        # lista Obrador = fuente oficial curada (los precios de septiembre estaban desactualizados, Agus): se acepta hasta +300 % / -70 %;
        # facturas / Maestro: regla del proyecto, ±50 % no se inyecta solo.
        lim_sup, lim_inf = (4.0, 0.3) if oficial else (1.5, 0.5)
        if p_ref and p_ref > 0 and (p / p_ref > lim_sup or p / p_ref < lim_inf):
            res.update(precio=p_ref, origen=f"REVISAR: propuesto {p:,.2f} ({p / p_ref - 1:+.0%}) vs cierre sept - {o}", estado="REVISAR")
        elif oficial and p_ref and p_ref > 0 and (p / p_ref > 1.5 or p / p_ref < 0.5):
            res.update(precio=p, origen=o, estado="APLICADO_REV", nota=(res["nota"] + " | " if res["nota"] else "") + f"Variación {p / p_ref - 1:+.0%} vs cierre sept (lista Obrador oficial, aplicada)")
        else:
            res.update(precio=p, origen=o)
    else:
        if p_ref:
            res.update(precio=p_ref, origen="Sin factura ni lista Obrador - se mantuvo precio de cierre sept", estado="SIN_FACTURA")
        else:
            res.update(precio=None, origen="SIN PRECIO", estado="SIN_PRECIO")
    if sku in REGLAS["verificar"]:
        res["nota"] = (res["nota"] + " | " if res["nota"] else "") + REGLAS["verificar"][sku]
    return res


# ───────────────────────────── escritura (formato Arcos) ─────────────────────────────

NARANJA = PatternFill("solid", fgColor="F4B083")
GRIS = PatternFill("solid", fgColor="D0CECE")
AMARILLO = PatternFill("solid", fgColor="FFF2CC")
ROJO = PatternFill("solid", fgColor="F4CCCC")
AVISO = PatternFill("solid", fgColor="FFF8E1")
THIN = Side(style="thin")
BORDE = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_HDR = Font(name="Arial", size=12, bold=True)
F_BODY = Font(name="Arial", size=10)
F_BOLD = Font(name="Arial", size=10, bold=True)
NF_PESOS = "[$ $]#,##0.00"
NF_VAL = '_-"$"\\ * #,##0.00_-;\\-"$"\\ * #,##0.00_-;_-"$"\\ * "-"??_-;_-@'
ANCHOS = {"A": 10.1, "B": 7.4, "C": 8.7, "D": 13.0, "E": 31.3, "F": 16.4}


def _hdr(ws, row, col, text, fill=NARANJA, wrap=True):
    c = ws.cell(row, col, text)
    c.font, c.fill, c.border = F_HDR, fill, BORDE
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=wrap)
    return c


def escribir_gestion(local: str, df: pd.DataFrame, res: list, log: list, F: Fuentes, mes: int, anio: int, path: Path, prueba: bool) -> dict:
    sect = SECTORES[local]
    k = len(sect)
    c_sect0 = 7
    c_final = c_sect0 + k
    c_val, c_pct, c_orig, c_dif, c_stk, c_pref, c_var, c_obr, c_obs = (c_final + i for i in range(1, 10))
    n = len(df)
    wb = openpyxl.Workbook()
    wsR = wb.active
    wsR.title = "Resumen"
    ws = wb.create_sheet("Inventario")
    heads = ["RUBRO", "UN.", "MEDIDA", "CODIGO", "PRODUCTO", "$ unitario actual"] + sect + ["FINAL", "VALORIZADO", "% valorizado", " Fecha ultima factura / origen del precio",
             "Diferencias vs cierre " + MESES[mes - 1][:3].lower(), "Stock cierre " + MESES[mes - 1][:3].lower(), "$ cierre " + MESES[mes - 1][:3].lower() + " (ref.)", "Var. % precio", "SKU Obrador / Recetas", "Observaciones"]
    for j, h in enumerate(heads, start=1):
        fill = GRIS if c_sect0 <= j < c_final else NARANJA
        _hdr(ws, 1, j, h, fill)
    ws.row_dimensions[1].height = 49.5
    ws.freeze_panes = "A2"
    order = df.assign(_r=df.rubro.str.upper(), _p=df.producto.str.upper()).sort_values(["_r", "_p"]).index
    last = n + 1
    estados = []
    for i, ix in enumerate(order, start=2):
        r, x = df.loc[ix], res[ix]
        vals = {1: r.rubro, 2: r.un, 3: r.medida, 4: r.sku, 5: r.producto}
        for c, v in vals.items():
            ws.cell(i, c, v if v not in ("", None) else None).font = F_BODY
        p = ws.cell(i, 6, x["precio"])
        p.font, p.number_format, p.alignment = F_BODY, NF_PESOS, Alignment(horizontal="right")
        for s in range(k):
            v = r.sect[s] if prueba and s < len(r.sect) and r.sect[s] else None
            c = ws.cell(i, c_sect0 + s, v)
            c.font, c.alignment = F_BODY, Alignment(horizontal="right")
        a, b = L(c_sect0), L(c_sect0 + k - 1)
        cf = ws.cell(i, c_final, f"=SUM({a}{i}:{b}{i})")
        cf.font, cf.fill, cf.border, cf.alignment, cf.number_format = F_BOLD, AMARILLO, BORDE, Alignment(horizontal="right"), "0.00"
        cv = ws.cell(i, c_val, f"={L(c_final)}{i}*F{i}")
        cv.font, cv.border, cv.number_format = F_BODY, BORDE, NF_PESOS
        cp = ws.cell(i, c_pct, f"=IFERROR({L(c_val)}{i}/SUM({L(c_val)}$2:{L(c_val)}${last}),0)")
        cp.font, cp.number_format = F_BODY, "0.00%"
        co = ws.cell(i, c_orig, x["origen"])
        co.font, co.border = F_BODY, BORDE
        if x["estado"] == "REVISAR" or x["estado"] == "SIN_PRECIO":
            co.fill = ROJO
        elif x["estado"] == "APLICADO_REV":
            co.fill = AVISO
        cd = ws.cell(i, c_dif, f"={L(c_final)}{i}-{L(c_stk)}{i}")
        cd.font, cd.number_format = F_BODY, "0.00"
        cs = ws.cell(i, c_stk, round(float(r.final_sept), 6) if r.final_sept else 0)
        cs.font, cs.number_format = F_BODY, "0.00"
        cr = ws.cell(i, c_pref, r.p_sept if r.p_sept else None)
        cr.font, cr.number_format = F_BODY, NF_PESOS
        cvv = ws.cell(i, c_var, f'=IF(N({L(c_pref)}{i})>0,F{i}/{L(c_pref)}{i}-1,"")')
        cvv.font, cvv.number_format = F_BODY, "0.0%"
        ws.cell(i, c_obr, x["obr_sku"] or None).font = F_BODY
        ws.cell(i, c_obs, x["nota"] or None).font = F_BODY
        estados.append((r, x))
    ws.auto_filter.ref = f"A1:{L(c_obs)}{last}"
    for col, w in ANCHOS.items():
        ws.column_dimensions[col].width = w
    for j in range(c_sect0, c_final):
        ws.column_dimensions[L(j)].width = 14
    for j, w in ((c_final, 12.4), (c_val, 19.6), (c_pct, 12), (c_orig, 46), (c_dif, 14), (c_stk, 14), (c_pref, 16), (c_var, 11), (c_obr, 16), (c_obs, 60)):
        ws.column_dimensions[L(j)].width = w
    # ── Resumen (bloque del Resumen de Arcos, columna del mes)
    mesn = MESES[mes]
    wsR["A1"].value = None
    _hdr(wsR, 1, 2, mesn, NARANJA)
    stock_sept = float(F.sept[F.sept.local == local].eval("final * p_act").fillna(0).sum())
    filas = [("STOCK INICIAL", round(stock_sept, 2)), ("COMPRAS ", None), ("TRASPASOS ( INGRESOS )", None), ("TRASPASOS ( EGRESOS )", None), ("OBRADOR", None),
             ("STOCK FINAL", f"=SUM(Inventario!{L(c_val)}2:{L(c_val)}{last})"), ("CONSUMO $$$", "=(B2+B3+B4-B5+B6)-B7"), ("VENTA CON IVA ", None), ("VENTA TOTAL SIN IVA", "=B9/1.21"),
             ("FOOD COST", "=IFERROR(B8/B10,0)")]
    for i, (lab, v) in enumerate(filas, start=2):
        a = wsR.cell(i, 1, lab)
        a.font = F_BOLD
        b = wsR.cell(i, 2, v)
        b.font = F_BODY
        b.number_format = "0.00%" if lab == "FOOD COST" else NF_PESOS
        if v is None:
            b.fill = AVISO
    wsR["A13"] = "Celdas amarillas = a completar al cierre (compras, traspasos, ventas). STOCK INICIAL = cierre " + MESES[mes - 1] + " valorizado a sus precios."
    wsR["A13"].font = Font(name="Arial", size=9, italic=True)
    wsR.column_dimensions["A"].width = 30
    wsR.column_dimensions["B"].width = 20
    # ── Controles
    wc = wb.create_sheet("Controles")
    for j, h in enumerate(["Tipo", "SKU", "Producto", "$ cierre ant.", "$ propuesto", "Var. %", "Origen / detalle", "Nota"], start=1):
        _hdr(wc, 1, j, h, NARANJA)
    fila = 2
    prior = {"REVISAR": 0, "SIN_PRECIO": 1, "APLICADO_REV": 2, "SIN_FACTURA": 3}
    for r, x in sorted(estados, key=lambda t: (prior.get(t[1]["estado"], 9), t[0].rubro, t[0].producto)):
        tipo = {"REVISAR": "REVISAR (no inyectado)", "SIN_PRECIO": "SIN PRECIO", "APLICADO_REV": "Aplicado (lista Obrador)", "SIN_FACTURA": "Sin factura (se mantuvo)"}.get(x["estado"])
        if not tipo and not x["nota"]:
            continue
        if not tipo:
            tipo = "Verificar (Agus)"
        var = (x["candidato"] / r.p_sept - 1) if (x["candidato"] and r.p_sept) else None
        vals = [tipo, r.sku, r.producto, r.p_sept, x["candidato"], var, x["origen"], x["nota"]]
        for j, v in enumerate(vals, start=1):
            c = wc.cell(fila, j, v)
            c.font = F_BODY
            if j in (4, 5):
                c.number_format = NF_PESOS
            if j == 6:
                c.number_format = "0.0%"
        if x["estado"] in ("REVISAR", "SIN_PRECIO"):
            wc.cell(fila, 1).fill = ROJO
        fila += 1
    for col, w in zip("ABCDEFGH", (24, 12, 40, 14, 14, 9, 80, 80)):
        wc.column_dimensions[col].width = w
    wc.freeze_panes = "A2"
    wc.auto_filter.ref = f"A1:H{max(fila - 1, 1)}"
    # ── Comentarios
    wm = wb.create_sheet("Comentarios")
    for j, h in enumerate(["Cambio", "SKU", "Producto", "Motivo"], start=1):
        _hdr(wm, 1, j, h, NARANJA)
    for i, row in enumerate(log, start=2):
        for j, v in enumerate(row, start=1):
            wm.cell(i, j, v).font = F_BODY
    sin = [(r.sku, r.producto) for r, x in estados if not r.sku]
    i = len(log) + 2
    for sku, nombre in sin:
        for j, v in enumerate(["Sin SKU", "", nombre, "No se encontró SKU seguro en Maestro/Insumos/Recetas: asignar a mano o dar de alta en el Maestro"], start=1):
            wm.cell(i, j, v).font = F_BODY
        i += 1
    for col, w in zip("ABCD", (26, 22, 60, 110)):
        wm.column_dimensions[col].width = w
    wm.freeze_panes = "A2"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    est = pd.Series([x["estado"] for _, x in estados]).value_counts().to_dict()
    total = sum((sum(r.sect) if r.sect else 0) * (num(x["precio"]) or 0) for r, x in estados) if prueba else None
    return dict(local=local, items=n, estados=est, sin_sku=len(sin), total_prueba=total, stock_sept=stock_sept, path=str(path))


def fallback_cafes(df: pd.DataFrame, res: list, dfs: dict, resueltos: dict) -> list:
    """Chimidoha no tiene cierre propio: sin factura/lista propia usa el precio del mismo ítem en Arcos/Cabello (mismo precio de transferencia)."""
    por_sku, por_nombre = {}, {}
    for l in ("Arcos", "Cabello"):
        if l not in resueltos:
            continue
        for (_, r), x in zip(dfs[l].iterrows(), resueltos[l]):
            if x["estado"] in ("OK", "APLICADO_REV") and num(x["precio"]):
                if r.sku:
                    por_sku.setdefault(r.sku, (x["precio"], x["origen"], l))
                por_nombre.setdefault(norm(r.producto), (x["precio"], x["origen"], l, r.sku))
    for i, ((_, r), x) in enumerate(zip(df.iterrows(), res)):
        if x["estado"] != "SIN_PRECIO":
            continue
        hit = por_sku.get(r.sku) if r.sku else None
        if not hit:
            h = por_nombre.get(norm(r.producto))
            if h:
                hit = h[:3]
        if hit:
            x.update(precio=hit[0], estado="OK", origen=f"Mismo ítem en {hit[2]}: {hit[1]}")
    return res


# ───────────────────────────── main ─────────────────────────────

def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--mes", type=int, default=10)
    ap.add_argument("--anio", type=int, default=2026)
    ap.add_argument("--prueba", action="store_true", help="rellena los conteos con los del cierre anterior para validar fórmulas/precios (salida en _prueba/)")
    ap.add_argument("--locales", nargs="*", default=list(SECTORES))
    a = ap.parse_args(argv)
    F = Fuentes(a.mes, a.anio)
    med = _medida_por_sku(F)
    log = defaultdict(list)
    dfs = {l: aplicar_reglas(F, l, items_base(F, l, med), log[l]) for l in SECTORES}
    espejar_arcos_cabello(dfs, log)
    salida = []
    resueltos = {}
    orden = [l for l in SECTORES if l in a.locales or l in ("Arcos", "Cabello")]
    orden = sorted(set(orden), key=list(SECTORES).index)
    for l in orden:
        res = [resolver_precio(F, l, r, a.mes, a.anio) for _, r in dfs[l].iterrows()]
        if l == "Chimidoha":
            res = fallback_cafes(dfs[l], res, dfs, resueltos)
        resueltos[l] = res
        if l not in a.locales:
            continue
        sub = OUT / ("_prueba" if a.prueba else "")
        p = sub / f"Gestion_{l}_{MESES[a.mes].lower()}{'_prueba' if a.prueba else ''}.xlsx"
        info = escribir_gestion(l, dfs[l], res, log[l], F, a.mes, a.anio, p, a.prueba)
        salida.append(info)
        print(f"{l:10s} ítems={info['items']:4d} estados={info['estados']} sin_sku={info['sin_sku']}"
              + (f" | valorizado prueba ${info['total_prueba']:,.0f} vs cierre ant. ${info['stock_sept']:,.0f}" if a.prueba else ""))
    return salida


if __name__ == "__main__":
    main()
