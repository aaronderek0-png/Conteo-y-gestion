"""Imputación automática: comprobante (JSON de invoice_parser.py) -> filas de la pestaña `Facturas` de un local.

Flujo por comprobante
  1. Proveedor detectado -> lista oficial (`Listas!B4:B`, la que alimenta el desplegable de la columna D).
  2. Cada renglón -> presentación del catálogo (`Productos!C`, desplegable de la columna E), en este orden:
       a) aprendizaje_proveedores.json (cómo se cargó ese texto de ese proveedor antes)  -> confianza 0.99
       b) nombre idéntico a una presentación                                            -> 1.00
       c) fuzzy (rapidfuzz + los filtros duros de comparador_precios: marcas, especie, variedad, medidas)
  3. Reglas de negocio: Verdulería (bulto/cajón solo para Limón, Pepino, Berenjena, Tomate Perita y Cebolla; kilo
     para el resto) y conversión de unidad (kg/lt/gr/bulto) a la unidad de la presentación.
  4. Decisión autónoma:
       confianza >= 0.90 y precio neto dentro de ±25 % del costo vigente  -> INYECTAR (cols A:I + P_UN y SKU)
       si no                                                            -> fila 'REVISAR / NO CATALOGADO' (en rojo)
       y, si el producto/proveedor no existe, un registro para SOLICITUDES_PRODUCTOS.

Seguridad: nunca toca la planilla original. `--apply` escribe SOLO en
`Copias_Gestion/<Local>_FACTURAS_COPIA.xlsx` (copia de Arcos_data/<local>.xlsx, que es el export de la
planilla de Facturas) y en `Copias_Gestion/Maestro_COPIA.xlsx` (hoja Solicitudes), y verifica cada celda
escrita leyendo el archivo de nuevo. Sin `--apply` es un dry-run que no escribe nada.

Uso:
    python imputacion_facturas.py factura.json|foto.jpg|carpeta/ [--local Arcos] [--apply]
    python imputacion_facturas.py --demo-historico 6 [--local Arcos]       # prueba con facturas reales reconstruidas
    python imputacion_facturas.py --bootstrap-aprendizaje                   # arma aprendizaje_proveedores.json
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import shutil
import sys
import tempfile
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
from rapidfuzz import fuzz, process

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "Arcos_data"
CARPETA_ENTRADA = BASE_DIR / "Facturas Mishiguene"          # carpeta diaria de comprobantes (sin guiones)
CARPETA_ENTRADA_VIEJA = BASE_DIR / "Facturas_WPP"           # lectura retrocompatible
CARPETA_REVISION = BASE_DIR / "Carpeta Revision"            # espejo de lo inyectado (OK / a revisar)
COLOR_OK = "FFF8E1"                                         # amarillo suave: inyectada, falta aprobar
COLOR_REVISAR = "FFEBEE"                                    # rojo suave: requiere revisión humana
COPIAS_DIR = BASE_DIR / "Copias_Gestion"
APRENDIZAJE_JSON = BASE_DIR / "aprendizaje_proveedores.json"
MAESTRO_XLSX = DATA_DIR / "maestro.xlsx"

# Planilla de Facturas de cada local (export local de la planilla de producción)
PLANILLAS = {"Arcos": "arcos.xlsx", "Cabello": "cabello.xlsx", "Lafinur": "lafinur.xlsx",
             "Congreso": "congreso_facturas.xlsx",       # 'Obrador' == 'Congreso'
             "Chimidoha": "congreso_facturas.xlsx"}      # sin planilla propia: se crea desde la de Congreso
SINONIMOS_LOCAL = {"obrador": "Congreso", "congreso": "Congreso", "arcos": "Arcos", "cabello": "Cabello",
                   "lafinur": "Lafinur", "chimidoha": "Chimidoha", "chimi": "Chimidoha"}

CONF_MINIMA = 0.90
TOLERANCIA_PRECIO = 0.25
REVISAR_TXT = "REVISAR / NO CATALOGADO"
# Verdulería: solo estos van por bulto/cajón; el resto, por kilo (se compara por nombre de presentación)
VERDULERIA_BULTO = {"limon", "pepino", "berenjena", "tomate perita", "cebolla"}
BULTO_UNIDADES = {"BULTO", "CAJA", "CAJON", "CAJÓN", "BAL", "BALDE"}
# kg que trae un bulto/cajón cuando el comprobante no lo dice ('cajón 18 kg'). Sin dato -> REVISAR.
BULTO_KG_DEFAULT: dict[str, float] = {"tomate perita": 16.0, "limon": 16.0, "berenjena": 9.0, "cebolla": 17.0,
                                       "pepino": 13.0}

def _fecha_futura_invalida(f) -> bool:
    from defensivo import fecha_futura_invalida
    return fecha_futura_invalida(f)


_EQUIV_CACHE: dict = {}


def bulto_kg_equivalencias() -> dict[str, float]:
    """kg por bulto según la pestaña `Equivalencias` del Maestro (proveedor Raíces/Open Fruits, unidad KG, factor > 1,5).
    Fuente primaria; `BULTO_KG_DEFAULT` queda solo de respaldo. Cacheado por mtime del export."""
    p = BASE_DIR / "Arcos_data" / "maestro.xlsx"
    try:
        mt = p.stat().st_mtime
    except OSError:
        return {}
    if _EQUIV_CACHE.get("mt") == mt:
        return _EQUIV_CACHE["d"]
    d: dict[str, float] = {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        if "Equivalencias" in wb.sheetnames:
            for r in wb["Equivalencias"].iter_rows(min_row=2, values_only=True):
                if len(r) > 5 and str(r[4] or "").upper() == "KG" and isinstance(r[5], (int, float)) and r[5] > 1.5:
                    d[norm(r[1] or r[3])] = float(r[5])
    except Exception as e:                                   # noqa: BLE001
        LOG.warning("no pude leer Equivalencias del Maestro: %s", e)
    _EQUIV_CACHE.update({"mt": mt, "d": d})
    return d


def bulto_kg(pres_norm: str) -> float | None:
    return bulto_kg_equivalencias().get(pres_norm) or BULTO_KG_DEFAULT.get(pres_norm)


LOG = logging.getLogger("imputacion_facturas")


def _sin_acentos(t: Any) -> str:
    return unicodedata.normalize("NFKD", str(t if t is not None and t == t else "")).encode("ascii", "ignore").decode()


def norm(t: Any) -> str:
    """Clave de comparación exacta: sin tildes, minúsculas, sin signos, espacios simples."""
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", _sin_acentos(t).lower())).strip()


def _num(v: Any) -> float | None:
    try:
        f = float(v)
        return None if f != f else f
    except (TypeError, ValueError):
        return None


# --------------------------------------------------------------------------- #
# Datos del local: catálogo, proveedores oficiales, historial de precios
# --------------------------------------------------------------------------- #
class Local:
    def __init__(self, nombre: str):
        self.nombre = nombre
        if nombre not in PLANILLAS:
            raise ValueError(f"'{nombre}' no tiene planilla de Facturas (disponibles: {', '.join(PLANILLAS)})")
        self.original = DATA_DIR / PLANILLAS[nombre]
        self.copia = COPIAS_DIR / f"{nombre}_FACTURAS_COPIA.xlsx"
        # Catálogo, proveedores e historial salen SIEMPRE del export original (trae los valores calculados por Google);
        # las copias guardadas con openpyxl no conservan los valores cacheados de las fórmulas. De la copia solo se leen
        # las filas que este agente agregó (valores literales A:I debajo de la última fila del original).
        LOG.info("%s: leyendo %s", nombre, self.original.name)
        xl = pd.ExcelFile(self.original)
        prod = pd.read_excel(xl, "Productos", header=0).iloc[:, :8]
        prod.columns = ["SKU", "PRODUCTO_BASE", "PRESENTACION", "CAT1", "CAT2", "UN_PRES", "UN_MEDIDA", "CANT_PRES"]
        prod = prod[prod["PRESENTACION"].notna()].copy()
        prod["_norm"] = prod["PRESENTACION"].map(norm)
        self.catalogo = prod.drop_duplicates("_norm").reset_index(drop=True)
        listas = pd.read_excel(xl, "Listas", header=None).iloc[3:, 1].dropna().astype(str).str.strip()
        self.proveedores = sorted(set(listas[listas != ""]))
        cols = ["N", "F_FACTURA", "F_INGRESO", "PROVEEDOR", "PRODUCTO", "NOTA", "CANT", "PRECIO", "DESC"]
        crudo = pd.read_excel(xl, "Facturas", header=0, usecols="A:I")
        crudo.columns = cols
        ultimo_original = int(crudo.index[crudo["N"].notna()].max()) if crudo["N"].notna().any() else -1

        def preparar(d: pd.DataFrame) -> pd.DataFrame:
            d = d[d["N"].notna() & d["PRODUCTO"].notna()].copy()
            d["F_FACTURA"] = pd.to_datetime(d["F_FACTURA"], errors="coerce")
            fut = d["F_FACTURA"].map(_fecha_futura_invalida).astype(bool)
            if fut.any():                                    # FECHA_FUTURA_INVALIDA: no entra al costo vigente
                LOG.warning("%s: %d factura(s) con fecha futura descartadas del costo vigente (N° %s)", self.nombre if hasattr(self, "nombre") else "?",
                            int(fut.sum()), sorted(set(d.loc[fut, "N"].astype(str)))[:5])
                d = d[~fut].copy()
            d["PRECIO"] = pd.to_numeric(d["PRECIO"], errors="coerce")
            d["CANT"] = pd.to_numeric(d["CANT"], errors="coerce")
            d["DESC"] = pd.to_numeric(d["DESC"], errors="coerce").fillna(0.0)
            d["P_NETO"] = d["PRECIO"] * (1 - d["DESC"])
            return d
        original = preparar(crudo)
        agregadas = original.iloc[0:0]
        if self.copia.is_file():
            c = pd.read_excel(self.copia, "Facturas", header=0, usecols="A:I")
            c.columns = cols
            agregadas = preparar(c.iloc[ultimo_original + 1:] if nombre != "Chimidoha" else c)
        self.facturas_precios = pd.concat([original, agregadas])        # historial para costo vigente
        # Chimidoha: sus propias facturas son solo las de su copia; los precios salen de la planilla de Congreso
        fac = agregadas if nombre == "Chimidoha" else self.facturas_precios
        self.facturas = fac
        self.ultima_fila = None            # se completa al escribir
        self._idx_prov: dict[str, set[str]] = {}
        for p, g in fac.groupby(fac["PROVEEDOR"].map(norm)):
            self._idx_prov[p] = set(g["PRODUCTO"].map(norm))

    def costo_vigente(self, presentacion: str, antes_de: date | None = None) -> tuple[float | None, str | None]:
        """Último precio neto cargado para la presentación (anterior a `antes_de` si se indica)."""
        h = self.facturas_precios
        f = h[h["PRODUCTO"].map(norm) == norm(presentacion)]
        f = f[f["P_NETO"].notna() & (f["P_NETO"] > 0)]
        f = f[~f["F_FACTURA"].map(_fecha_futura_invalida).astype(bool)]      # FECHA_FUTURA_INVALIDA
        if antes_de is not None:
            f = f[f["F_FACTURA"] < pd.Timestamp(antes_de)]
        if f.empty:
            return None, None
        r = f.sort_values("F_FACTURA").iloc[-1]
        return float(r["P_NETO"]), str(pd.Timestamp(r["F_FACTURA"]).date())

    def productos_del_proveedor(self, proveedor: str) -> set[str]:
        return self._idx_prov.get(norm(proveedor), set())


# --------------------------------------------------------------------------- #
# Aprendizaje (texto del proveedor -> presentación)
# --------------------------------------------------------------------------- #
def cargar_aprendizaje() -> dict:
    try:
        return json.loads(APRENDIZAJE_JSON.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"version": 1, "mapeos": {}}


def guardar_aprendizaje(ap: dict) -> None:
    """Escritura atómica (tmp + replace): un corte no deja el aprendizaje validado vacío o truncado."""
    tmp = APRENDIZAJE_JSON.with_suffix(".tmp")
    tmp.write_text(json.dumps(ap, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")
    tmp.replace(APRENDIZAJE_JSON)


def aprender(ap: dict, proveedor: str, descripcion: str, presentacion: str, local: str, n: int = 1) -> None:
    k = f"{norm(proveedor)}|{norm(descripcion)}"
    e = ap["mapeos"].setdefault(k, {"presentacion": presentacion, "n": 0, "locales": []})
    if e["presentacion"] != presentacion:       # la carga más reciente manda
        e["presentacion"], e["n"] = presentacion, 0
    e["n"] += n
    if local not in e["locales"]:
        e["locales"].append(local)
    e["ultima"] = date.today().isoformat()


def bootstrap_aprendizaje() -> dict:
    """Siembra aprendizaje_proveedores.json con (a) las solicitudes APROBADAS del Maestro (texto del proveedor ->
    producto aprobado) y (b) el historial de Facturas de los 4 locales (proveedor, presentación) -> sí mismo."""
    ap = cargar_aprendizaje()
    sol = pd.read_excel(MAESTRO_XLSX, sheet_name="Solicitudes")
    ok = sol[sol["Estado"].astype(str).str.upper().eq("APROBADO") & sol["¿"].astype(str).str.upper().eq("PRODUCTO")]
    for _, r in ok.iterrows():
        if pd.notna(r["Producto Base"]) and pd.notna(r["Producto Propuesto"]) and pd.notna(r["Proveedor"]):
            aprender(ap, r["Proveedor"], r["Producto Base"], str(r["Producto Propuesto"]).strip(),
                     str(r["Local Solicitante"]).title())
    n_sol = len(ap["mapeos"])
    for nombre in (n for n in PLANILLAS if n != "Chimidoha"):
        loc = Local(nombre)
        g = loc.facturas.groupby([loc.facturas["PROVEEDOR"].map(_sin_acentos), loc.facturas["PRODUCTO"]]).size()
        pres_ok = set(loc.catalogo["PRESENTACION"])
        for (prov, pres), n in g.items():
            if pres in pres_ok and prov.strip():
                k = f"{norm(prov)}|{norm(pres)}"
                if k not in ap["mapeos"]:
                    aprender(ap, prov, pres, pres, nombre, int(n))
    guardar_aprendizaje(ap)
    LOG.info("aprendizaje: %d mapeos (%d de solicitudes aprobadas) -> %s", len(ap["mapeos"]), n_sol,
             APRENDIZAJE_JSON.name)
    return ap


# --------------------------------------------------------------------------- #
# Matching
# --------------------------------------------------------------------------- #
_CMP = None


def _comparador():
    """Reutiliza la lógica de matching del comparador (marcas, especie, variedad, medidas)."""
    global _CMP
    if _CMP is None:
        sys.path.insert(0, str(BASE_DIR))
        import comparador_precios as cmp
        _CMP = cmp
    return _CMP


def mapear_proveedor(detectado: str | None, texto_encabezado: str, oficiales: list[str]) -> tuple[str | None, float]:
    if not oficiales:
        return None, 0.0
    candidatos = [c for c in (detectado, ) if c]
    mejor, mejor_s = None, 0.0
    n_of = {norm(o): o for o in oficiales}
    for c in candidatos:
        k = norm(c)
        if k in n_of:
            return n_of[k], 1.0
        r = process.extractOne(k, list(n_of), scorer=fuzz.token_set_ratio)
        if r and r[1] > mejor_s:
            mejor, mejor_s = n_of[r[0]], r[1] / 100
    # el nombre oficial puede figurar en cualquier parte del encabezado (razón social)
    enc = norm(texto_encabezado)
    for k, o in n_of.items():
        if len(k) >= 5 and f" {k} " in f" {enc} " and mejor_s < 0.95:
            mejor, mejor_s = o, 0.95
    return (mejor, mejor_s) if mejor_s >= 0.85 else (None, mejor_s)


def _medida_envase(cat_row: pd.Series) -> tuple[float, str] | None:
    u = str(cat_row["UN_MEDIDA"]).strip().upper()
    f, h = _num(cat_row["UN_PRES"]), _num(cat_row["CANT_PRES"])
    if u in {"KG", "LT", "G", "GR", "ML"} and f and h:
        v = f * h
        return (v / 1000, "KG" if u in {"G", "GR"} else "LT") if u in {"G", "GR", "ML"} else (v, u)
    return None


def mapear_producto(desc: str, proveedor: str | None, loc: Local, ap: dict, alias: str | None = None) -> dict:
    """Devuelve {'presentacion', 'confianza', 'metodo', 'alternativas'}; presentacion None si no hay candidato.
    `alias`: nombre con que el proveedor figura en la planilla (si difiere del oficial de Listas!B)."""
    k = norm(desc)
    provs = [p for p in dict.fromkeys((proveedor, alias)) if p]
    for p_ in provs:
        e = ap["mapeos"].get(f"{norm(p_)}|{k}")
        if e and e["presentacion"] in set(loc.catalogo["PRESENTACION"]):
            return {"presentacion": e["presentacion"], "confianza": 0.99, "metodo": f"aprendizaje (n={e['n']})",
                    "alternativas": []}
    proveedor = provs[0] if provs else None
    exacto = loc.catalogo[loc.catalogo["_norm"] == k]
    if len(exacto):
        return {"presentacion": exacto.iloc[0]["PRESENTACION"], "confianza": 1.0, "metodo": "nombre idéntico",
                "alternativas": []}
    if provs:           # texto casi igual a uno ya aprendido de ESE proveedor ('Caja 3 pouch hojas de Parra x 3 kg')
        prefs = tuple(norm(p_) + "|" for p_ in provs)
        validas = set(loc.catalogo["PRESENTACION"])
        mejor_s, mejor_p = 0, None
        for clave, e in ap["mapeos"].items():
            if not clave.startswith(prefs) or e["presentacion"] not in validas:
                continue
            pref = next(p_ for p_ in prefs if clave.startswith(p_))
            s_ = fuzz.token_set_ratio(k, clave[len(pref):])
            if s_ > mejor_s:
                mejor_s, mejor_p = s_, e["presentacion"]
        if mejor_s >= 88:
            return {"presentacion": mejor_p, "confianza": 0.92, "metodo": f"sinónimo aprendido del proveedor ({mejor_s:.0f})",
                    "alternativas": []}

    cmp = _comparador()
    q = cmp._consulta(desc)
    qk = cmp.match_key(q)
    qtokens = qk.split()
    if not qtokens:
        return {"presentacion": None, "confianza": 0.0, "metodo": "sin texto útil", "alternativas": []}
    habituales = loc.productos_del_proveedor(proveedor) if proveedor else set()
    tam_q = cmp.parse_presentacion(_sin_acentos(desc))
    hay_tam = bool(cmp._QTY_RE.search(_sin_acentos(desc)))
    cands: list[tuple[float, str, list[str]]] = []
    for _, c in loc.catalogo.iterrows():
        ck = cmp.match_key(c["PRESENTACION"])
        ctok = ck.split()
        if not ctok or not cmp.candidato_valido(qtokens, ctok):
            continue
        sc = cmp.score_match(qk, ck)
        if c["_norm"] in habituales:
            sc += 6                                        # ya se le compra ese producto a ese proveedor
        if hay_tam:
            env = _medida_envase(c)
            if env and env[1] == tam_q[1] and not 0.75 <= env[0] / tam_q[0] <= 1.34:
                sc -= 15                                   # otro tamaño de envase
        cands.append((sc, c["PRESENTACION"], ctok))
    if not cands:
        # sin candidato que pase los filtros duros (¿error de tipeo?): se sugiere el más parecido, tope 0.80
        nombres = {cmp.match_key(n): n for n in loc.catalogo["PRESENTACION"]}
        r = process.extractOne(qk, list(nombres), scorer=fuzz.token_set_ratio)
        if r and r[1] >= 80:
            return {"presentacion": nombres[r[0]], "confianza": round(min(r[1] / 100 * 0.85, 0.80), 3),
                    "metodo": f"sugerencia sin filtros {r[1]:.0f}", "alternativas": []}
        return {"presentacion": None, "confianza": 0.0, "metodo": "sin candidato", "alternativas": []}
    cands.sort(key=lambda x: -x[0])
    top = cands[0]
    conf = min(top[0], 100) / 100 * 0.97                   # el fuzzy nunca llega a 1.0: eso es del match exacto
    if len(cands) > 1 and top[0] - cands[1][0] < 4:
        conf -= 0.10                                       # dos candidatos casi iguales: ambiguo
    return {"presentacion": top[1], "confianza": round(max(conf, 0.0), 3), "metodo": f"fuzzy {top[0]:.0f}",
            "alternativas": [f"{p} ({s:.0f})" for s, p, _ in cands[1:4]]}


# --------------------------------------------------------------------------- #
# Reglas de negocio: unidad / bulto / verdulería
# --------------------------------------------------------------------------- #
_BULTO_KG = re.compile(r"(?:caj[oó]n|cajon|bulto|caja|bolsa|malla|jaula)?\s*(?:x|de)?\s*(\d+(?:[.,]\d+)?)\s*kgs?\b", re.I)


def convertir_unidad(item: dict, cat: pd.Series) -> tuple[float | None, float | None, list[str], str | None]:
    """Cantidad/precio del comprobante -> (G cantidad, H precio) en la unidad de la presentación.
    Devuelve (G, H, avisos, error)."""
    cant, precio = float(item["cantidad"]), float(item["precio_unitario"])
    u = (item.get("unidad_detectada") or "").upper()
    avisos: list[str] = []
    pres_norm = norm(cat["PRESENTACION"])
    es_verdura = norm(cat["CAT2"]) == "frutas y verduras"
    envase = _medida_envase(cat)                    # (cantidad por unidad de presentación, 'KG'/'LT')
    es_bulto = u in BULTO_UNIDADES

    if es_verdura:
        en_lista = pres_norm in VERDULERIA_BULTO
        if es_bulto:
            m = _BULTO_KG.search(item["descripcion_original"])
            kg = float(m.group(1).replace(",", ".")) if m else bulto_kg(pres_norm)
            if not kg:
                return None, None, avisos, ("bulto/cajón sin peso en el comprobante: no se puede llevar a kilos"
                                            if en_lista else
                                            "regla Verdulería: se espera kilo y vino por bulto/cajón sin peso")
            if not en_lista:
                avisos.append("regla Verdulería: este producto va por kilo y vino por bulto; convertido por el peso indicado")
            cant, precio, u = cant * kg, precio / kg, "KG"
        elif en_lista and u in {"KG", ""}:
            avisos.append("Verdulería: el comprobante ya está en kilos (el bulto se convirtió aguas arriba)")

    if u in {"KG", "LT", "GR", "G"} and envase:
        if u in {"GR", "G"}:
            cant, precio, u = cant / 1000, precio * 1000, "KG"
        base = envase[1]
        if u != base:
            return None, None, avisos, f"unidad del comprobante ({u}) no coincide con la presentación ({base})"
        return cant / envase[0], precio * envase[0], avisos, None
    if u in {"KG", "LT"} and not envase and str(cat["UN_MEDIDA"]).upper() not in {u}:
        return None, None, avisos, f"comprobante en {u} y la presentación es {cat['UN_MEDIDA']}"
    return cant, precio, avisos, None                  # UN / sin unidad: se carga tal cual


# --------------------------------------------------------------------------- #
# Decisión por renglón
# --------------------------------------------------------------------------- #
def decidir_renglon(item: dict, comp: dict, proveedor: str | None, loc: Local, ap: dict) -> dict:
    desc = item["descripcion_original"]
    m = mapear_producto(desc, proveedor, loc, ap, alias=comp.get("proveedor_planilla"))
    out = {"local": loc.nombre, "descripcion_original": desc, "presentacion": m["presentacion"],
           "confianza": m["confianza"], "metodo": m["metodo"], "alternativas": m["alternativas"],
           "cant_orig": item.get("cantidad"), "precio_orig": item.get("precio_unitario"),
           "motivos": [], "avisos": [], "solicitud": False, "accion": "REVISAR", "sku": None,
           "G": None, "H": None, "I": None, "J": None, "P_UN": None, "vigente": None, "var_pct": None}
    fecha_fact = comp.get("fecha_factura")
    out["I"] = round((item.get("descuento_pct") or 0) / 100, 4)
    if item.get("linea_consistente") is False:
        out["motivos"].append("el renglón no cierra: cantidad × precio ≠ total (¿cantidad mal leída?)")
    if m["presentacion"] is None:
        out["motivos"].append("producto no catalogado")
        out["solicitud"] = True
        return out
    cat = loc.catalogo[loc.catalogo["PRESENTACION"] == m["presentacion"]].iloc[0]
    out["sku"] = cat["SKU"]
    G, H, avisos, err = convertir_unidad(item, cat)
    out["avisos"] += avisos
    if err:
        out["motivos"].append(err)
    else:
        out["G"], out["H"] = round(G, 4), round(H, 4)
        out["J"] = round(G * H * (1 - out["I"]), 2)
        env = _medida_envase(cat)
        cant_inv = G * (_num(cat["UN_PRES"]) or 1) * (_num(cat["CANT_PRES"]) or 1)
        out["P_UN"] = round(out["J"] / cant_inv, 4) if cant_inv else None
        antes = datetime.strptime(fecha_fact, "%Y-%m-%d").date() if fecha_fact else None
        vig, fvig = loc.costo_vigente(m["presentacion"], antes)
        if vig is None:                                   # sin historial en este local: se prueba con los otros
            for otro in (n for n in PLANILLAS if n != "Chimidoha"):
                if otro != loc.nombre:
                    try:
                        vig, fvig = _vigente_otro_local(otro, m["presentacion"], antes)
                    except Exception:                    # noqa: BLE001
                        vig = None
                    if vig:
                        out["avisos"].append(f"costo vigente tomado de {otro}")
                        break
        out["vigente"] = vig
        if vig:
            neto = H * (1 - out["I"])
            out["var_pct"] = round((neto / vig - 1) * 100, 1)
            if abs(neto / vig - 1) > TOLERANCIA_PRECIO:
                out["motivos"].append(f"precio {out['var_pct']:+.0f} % vs vigente ${vig:,.2f} ({fvig})"
                                      + (" — > ±50 %: nunca se inyecta sola" if abs(neto / vig - 1) > 0.5 else ""))
        else:
            out["motivos"].append("sin costo vigente para validar el precio")
    if m["confianza"] < CONF_MINIMA:
        out["motivos"].append(f"confianza {m['confianza']:.2f} < {CONF_MINIMA:.2f}")
        out["solicitud"] = True
    if not proveedor:
        out["motivos"].append("proveedor fuera de la lista oficial")
        out["solicitud"] = True
    if not out["motivos"]:
        out["accion"] = "INYECTAR"
    return out


_VIG_CACHE: dict[str, Local] = {}


def _vigente_otro_local(nombre: str, pres: str, antes) -> tuple[float | None, str | None]:
    if nombre not in _VIG_CACHE:
        _VIG_CACHE[nombre] = Local(nombre)
    return _VIG_CACHE[nombre].costo_vigente(pres, antes)


def numero_a_celda(numero: str | None) -> float | int | str | None:
    """Solo los dígitos significativos, sin punto de venta ni ceros a la izquierda (así se carga en las planillas):
    '00003-00004462' -> 4462 · '5548-00008853' -> 8853 · '0001-00063554' -> 63554 · 'R-0001/00009194' -> 9194."""
    n = ultimo_entero(numero)
    return n if n is not None else (numero or None)


def imputar_comprobante(comp: dict, loc: Local, ap: dict, omitir_duplicados: bool = False) -> dict:
    proveedor, conf_p = mapear_proveedor(comp.get("proveedor_detectado"), comp.get("texto_encabezado", ""),
                                         loc.proveedores)
    filas = [decidir_renglon(i, comp, proveedor, loc, ap) for i in comp.get("items", [])]
    fecha_f, fecha_i = comp.get("fecha_factura"), comp.get("fecha_ingreso") or comp.get("fecha_factura")
    for f in filas:
        f.update({"A": numero_a_celda(comp.get("numero_factura")), "B": fecha_f, "C": fecha_i,
                  "D": proveedor or (comp.get("proveedor_detectado") or "").upper(),
                  "proveedor_oficial": proveedor})
        if not omitir_duplicados and f["presentacion"] and f["G"]:
            dup = loc.facturas[(loc.facturas["N"].map(_num) == _num(f["A"])) &
                               (loc.facturas["PRODUCTO"].map(norm) == norm(f["presentacion"])) &
                               (loc.facturas["CANT"].map(_num).round(3) == (round(f["G"], 3) if f["G"] else None))]
            if len(dup):
                f["accion"], f["motivos"] = "DUPLICADA", ["ya cargada en la planilla (mismo N°, producto y cantidad)"]
                f["solicitud"] = False
    v = comp.get("validacion", {})
    if not v.get("ok"):
        for f in filas:
            if f["accion"] == "INYECTAR":
                f["accion"] = "REVISAR"
            f["motivos"].append(f"comprobante no cierra: {v.get('motivo')}")
    return {"comprobante": {k: comp.get(k) for k in ("archivo", "local_destino", "numero_factura", "fecha_factura",
                                                    "proveedor_detectado", "total_comprobante")},
            "proveedor_oficial": proveedor, "confianza_proveedor": conf_p, "validacion": v, "filas": filas}


# --------------------------------------------------------------------------- #
# Escritura (solo copias) + Solicitudes
# --------------------------------------------------------------------------- #
def _fecha_celda(iso: str | None):
    return datetime.strptime(iso, "%Y-%m-%d") if iso else None


def aplicar_a_copia(resultados: list[dict], loc: Local, remitos: list[dict] | None = None,
                    marcar: list[tuple[int, str]] | None = None, remitos_obrador: list[dict] | None = None) -> dict:
    """Escribe las filas en la copia local de Facturas. Devuelve {'escritas': n, 'verificadas': n, 'filas': [...]}.
    `remitos`: registros para la pestaña 'Remitos' (o 'Remitos_Pendientes' si ya existe una ajena). `marcar`: [(fila, nota)] filas ya existentes que se
    señalan en la col F (p. ej. 'REASIGNAR A CHIMIDOHA') sin tocar sus importes."""
    import openpyxl
    from openpyxl.styles import PatternFill
    if not loc.copia.is_file():
        COPIAS_DIR.mkdir(exist_ok=True)
        if loc.nombre == "Chimidoha":
            crear_copia_chimidoha(loc.copia)
        else:
            shutil.copy2(loc.original, loc.copia)
        LOG.info("copia creada: %s", loc.copia.name)
    assert loc.copia.resolve().parent == COPIAS_DIR.resolve() and loc.copia.name.endswith("_FACTURAS_COPIA.xlsx")
    wb = openpyxl.load_workbook(loc.copia)
    ws = wb["Facturas"]
    ultima = max((r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value not in (None, "")), default=1)
    rojo = PatternFill("solid", fgColor="FFEBEE")         # requiere revisión (rojo suave)
    ok_fill = PatternFill("solid", fgColor="FFF8E1")      # inyectada OK, pendiente de aprobación (amarillo suave)
    escritas = []
    for res in resultados:
        for f in res["filas"]:
            if f["accion"] == "DUPLICADA":
                continue
            ultima += 1
            r = ultima
            reviso = f["accion"] == "REVISAR"
            valores = {1: f["A"], 2: _fecha_celda(f["B"]), 3: _fecha_celda(f["C"]), 4: f["D"],
                       5: REVISAR_TXT if reviso else f["presentacion"],
                       7: f["G"] if f["G"] is not None else f["cant_orig"],
                       8: f["H"] if f["H"] is not None else f["precio_orig"],
                       9: f["I"]}
            if reviso:   # nota en la columna F (libre) con lo original, la sugerencia y por qué
                sug = f" | sugerido: {f['presentacion']} (conf {f['confianza']:.2f})" if f["presentacion"] else ""
                valores[6] = f"{f['descripcion_original']}{sug} | {'; '.join(f['motivos'])}"[:250]
            for c, v in valores.items():
                ws.cell(r, c).value = v
            for c in range(1, 10):                      # toda fila nueva va pintada hasta que alguien la apruebe
                ws.cell(r, c).fill = rojo if reviso else ok_fill
            escritas.append((r, valores, f))
    amarillo = PatternFill("solid", fgColor="FFF2CC")
    for fila, nota in (marcar or []):
        ws.cell(fila, 6).value = nota[:250]
        ws.cell(fila, 6).fill = amarillo
    if remitos is not None:
        from copy import copy
        nombre_hoja = "Remitos_Pendientes"
        if nombre_hoja in wb.sheetnames:
            del wb[nombre_hoja]
        wr = wb.create_sheet(nombre_hoja)
        wr.sheet_view.showGridLines = False                       # sin líneas de cuadrícula
        for c in range(1, 25):                                    # encabezado, fuente, color y ancho clonados de Facturas
            src, dst = ws.cell(1, c), wr.cell(1, c)
            dst.value = src.value
            if src.has_style:
                dst.font, dst.fill, dst.border = copy(src.font), copy(src.fill), copy(src.border)
                dst.alignment, dst.number_format = copy(src.alignment), src.number_format
            letra = openpyxl.utils.get_column_letter(c)
            if ws.column_dimensions[letra].width:
                wr.column_dimensions[letra].width = ws.column_dimensions[letra].width
        for j, h in enumerate(["ESTADO REMITO IA", "PESO REAL KG", "COSTO REAL $/KG", "ANULADO ('NO')", "FACTURA DEFINITIVA"]):
            c = wr.cell(1, 25 + j)
            c.value = h
            src = ws.cell(1, 10)
            if src.has_style:
                c.font, c.fill, c.alignment = copy(src.font), copy(src.fill), copy(src.alignment)
        wr.freeze_panes = "A2"
        for rem in remitos:
            for r in rem["renglones"]:
                cant = 0 if r["anulado"] else (r.get("peso_real_recibido_kg") if (r.get("unidad") or "").upper() == "KG" and r.get("peso_real_recibido_kg") is not None else r["cantidad_real"])
                precio = r.get("precio_final") if r.get("precio_final") is not None else r["precio_unitario"]
                fila = wr.max_row + 1
                valores = {1: numero_a_celda(rem["remito"]), 2: _fecha_celda(rem["fecha"]), 3: _fecha_celda(rem["fecha"]),
                           4: rem.get("proveedor_oficial") or rem["proveedor"], 5: r.get("presentacion") or r["descripcion"],
                           6: r.get("nota") or None, 7: cant, 8: precio, 9: 0,
                           10: round(cant * (precio or 0), 2), 25: rem["estado"], 26: r.get("peso_real_recibido_kg"),
                           27: r.get("costo_real_por_kg"), 28: "SI" if r["anulado"] else None,
                           29: rem.get("factura_definitiva")}
                for c, v in valores.items():
                    wr.cell(fila, c).value = v
                for c in range(1, 10):
                    wr.cell(fila, c).fill = PatternFill("solid", fgColor="FFF8E1")
    if remitos_obrador:
        agregar_remitos_obrador_a_hoja(wb, remitos_obrador)
    wb.save(loc.copia)
    # read-back: cada celda escrita tiene que leerse igual
    chk = openpyxl.load_workbook(loc.copia, read_only=True)["Facturas"]
    ok = 0
    for r, valores, _ in escritas:
        fila = next(chk.iter_rows(min_row=r, max_row=r, max_col=9, values_only=True))
        if all((fila[c - 1] == v) or (v is None and fila[c - 1] in (None, "")) for c, v in valores.items()):
            ok += 1
    return {"escritas": len(escritas), "verificadas": ok, "filas": [(r, f) for r, _, f in escritas]}


def registrar_solicitudes(resultados: list[dict], locales: list[str], apply: bool) -> list[dict]:
    """Registros para SOLICITUDES_PRODUCTOS (hoja 'Solicitudes' del Maestro). Con apply=True se agregan a
    Copias_Gestion/Maestro_COPIA.xlsx; nunca al Maestro de producción."""
    regs = []
    ya = set()
    for res, loc in zip(resultados, locales):
        for f in res["filas"]:
            if not f["solicitud"] or f["accion"] == "DUPLICADA":
                continue
            clave = (norm(f["descripcion_original"]), norm(f["D"]))
            if clave in ya:
                continue
            ya.add(clave)
            if not f.get("proveedor_oficial"):
                tipo, prod_base, prod_prop = "PROVEEDOR", "", ""
            else:
                tipo, prod_base, prod_prop = "PRODUCTO", f["descripcion_original"], f["descripcion_original"]
            regs.append({"Fecha Solicitud": datetime.now().strftime("%d/%m/%Y %H:%M"),
                         "Local Solicitante": loc.upper(), "¿": tipo, "Tipo": "", "Rubro Propuesto": "",
                         "Producto Base": prod_base, "Producto Propuesto": prod_prop, "Contenido": "",
                         "UM Contenido": "", "Factor Conversión": "", "Proveedor": f["D"], "Estado": "PENDIENTE"})
    if apply and regs:
        import openpyxl
        destino = COPIAS_DIR / "Maestro_COPIA.xlsx"
        if not destino.is_file():
            shutil.copy2(MAESTRO_XLSX, destino)
        wb = openpyxl.load_workbook(destino)
        ws = wb["Solicitudes"]
        cols = [c.value for c in ws[1]]
        for reg in regs:
            ws.append([reg.get(c, "") for c in cols])
        wb.save(destino)
    return regs


def escribir_log(resultados: list[dict], locales: list[str]) -> list[Path]:
    COPIAS_DIR.mkdir(exist_ok=True)
    paths = []
    for nombre in sorted(set(locales)):
        p = COPIAS_DIR / f"log_Facturas_{nombre}.csv"
        nuevo = not p.is_file()
        with open(p, "a", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            if nuevo:
                w.writerow(["timestamp", "archivo", "N", "proveedor", "descripcion_original", "presentacion",
                            "sku", "confianza", "metodo", "accion", "G_cantidad", "H_precio", "P_UN", "vigente",
                            "var_pct", "motivos"])
            for res, loc in zip(resultados, locales):
                if loc != nombre:
                    continue
                for f in res["filas"]:
                    w.writerow([datetime.now().isoformat(timespec="seconds"), res["comprobante"]["archivo"], f["A"],
                                f["D"], f["descripcion_original"], f["presentacion"], f["sku"], f["confianza"],
                                f["metodo"], f["accion"], f["G"], f["H"], f["P_UN"], f["vigente"], f["var_pct"],
                                "; ".join(f["motivos"] + f["avisos"])])
        paths.append(p)
    return paths


# --------------------------------------------------------------------------- #
# Prueba con facturas reales reconstruidas
# --------------------------------------------------------------------------- #
def _pdf_texto(path: Path, lineas: list[str]) -> None:
    """PDF mínimo (Courier, una página) con texto extraíble: simula la factura digital de un proveedor."""
    def esc(s: str) -> str:
        return s.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")
    contenido = ["BT", "/F1 9 Tf", "11 TL", "30 810 Td"] + [f"({esc(l.encode('latin-1', 'replace').decode('latin-1'))}) '"
                                                           for l in lineas] + ["ET"]
    stream = "\n".join(contenido).encode("latin-1")
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
            b"<< /Length %d >>\nstream\n" % len(stream) + stream + b"\nendstream",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Courier /Encoding /WinAnsiEncoding >>"]
    out, offs = b"%PDF-1.4\n", []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1) + b"".join(b"%010d 00000 n \n" % o for o in offs)
    out += b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, xref)
    path.write_bytes(out)


def _ar(n: float) -> str:
    return f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def reconstruir_comprobantes(loc: Local, cuantos: int, carpeta: Path) -> list[dict]:
    """Toma las últimas `cuantos` facturas reales del local (mismo N° + proveedor), las vuelve a escribir como
    un PDF de proveedor con descripciones 'sucias' (mayúsculas, sin tildes, tamaño, errores de tipeo) y devuelve
    la verdad de terreno (presentación realmente cargada)."""
    cat = set(loc.catalogo["PRESENTACION"])
    f = loc.facturas[loc.facturas["PRODUCTO"].isin(cat) & loc.facturas["CANT"].notna()
                     & loc.facturas["PRECIO"].gt(0) & loc.facturas["F_FACTURA"].notna()]
    f = f[f["PROVEEDOR"].map(lambda p: norm(p) in {norm(x) for x in loc.proveedores})]
    grupos = [g for _, g in f.groupby([f["N"].map(_num), f["PROVEEDOR"]]) if 2 <= len(g) <= 25]
    grupos.sort(key=lambda g: g["F_FACTURA"].max())
    elegidos = grupos[-cuantos:]
    carpeta.mkdir(parents=True, exist_ok=True)
    reales = []
    for gi, g in enumerate(elegidos):
        prov, num, fecha = g["PROVEEDOR"].iloc[0], int(_num(g["N"].iloc[0])), g["F_FACTURA"].iloc[0]
        lineas = [f"{str(prov).upper()}  C.U.I.T. 30-70000000-1", f"REMITO X  N° 0001-{num:08d}",
                  f"Fecha: {fecha:%d/%m/%Y}   Señor/es: MISHIGUENE {loc.nombre.upper()}",
                  "CANT UN DESCRIPCION                              PRECIO   BONIF       TOTAL"]
        verdad, total = [], 0.0
        for k, (_, r) in enumerate(g.iterrows()):
            c = loc.catalogo[loc.catalogo["PRESENTACION"] == r["PRODUCTO"]].iloc[0]
            d = _sin_acentos(r["PRODUCTO"]).upper()
            env = _medida_envase(c)
            if k % 3 == 1 and env:
                d += f" X {env[0]:g} {env[1]}"                 # el proveedor le agrega el tamaño
            if k % 5 == 4 and len(d) > 6:
                d = d[:3] + d[4] + d[3] + d[5:]                # error de tipeo (letras 4 y 5 cambiadas)
            cant, precio = float(r["CANT"]), float(r["PRECIO"])
            bon = round(float(r["DESC"]) * 100, 2)
            linea_total = round(cant * precio * (1 - bon / 100), 2)
            total += linea_total
            un = "UN"
            cant_txt = (f"{cant:.3f}".rstrip("0").rstrip(".")).replace(".", ",")
            lineas.append(f"{cant_txt} {un} {d[:40]:<40} {_ar(precio):>12} {(_ar(bon) + '%') if bon else '':>7} "
                          f"{_ar(linea_total):>13}".replace("  ", " "))
            verdad.append({"descripcion": d, "presentacion": r["PRODUCTO"], "precio": precio, "cant": cant})
        lineas += ["", f"TOTAL {_ar(total)}"]
        pdf = carpeta / f"{loc.nombre}_{gi:02d}_{num}.pdf"
        _pdf_texto(pdf, lineas)
        reales.append({"pdf": pdf, "verdad": verdad, "proveedor": prov, "num": num, "fecha": fecha})
    return reales


# --------------------------------------------------------------------------- #
# Lote WhatsApp: retro-entrenamiento, carga agrupada, remitos con birome, NC, prorrateo, Chimidoha
# --------------------------------------------------------------------------- #
REMITOS_JSON = BASE_DIR / "remitos_pendientes.json"
GENERICOS_AGRUPADO = {"verduleria", "varios", "limpieza", "lavadero"}
CARGOS_RE = re.compile(r"servicio\s+log[ií]stico|\bflete\b|\benv[ií]o\b|acarreo|impuestos?\s+internos?", re.I)
PRESENTACION_NC_MP = "Nota de credito Materia Prima"       # SKU LV0218, categoría NC
NC_B_RE = re.compile(r"bonific|descuento|diferencia\s+de\s+precio|dif\.?\s+precio|ajuste|rebaja|error\s+de\s+precio", re.I)
NC_A_RE = re.compile(r"devol|rechaz|mal\s+estado|faltante|error\s+de\s+carga|no\s+fue\s+lo\s+solicitado|reversi|roto|vencid",
                     re.I)


def ultimo_entero(numero: Any) -> int | None:
    p = re.findall(r"\d+", str(numero or ""))
    return int(p[-1]) if p else None


def crear_copia_chimidoha(destino: Path) -> None:
    """Chimidoha (Café Mishiguene Núñez, PB del 1602) no tiene planilla propia: se crea desde la de Congreso con la
    misma estructura (Facturas, Productos, Listas, fórmulas) y la pestaña Facturas vacía (A:I)."""
    import openpyxl
    shutil.copy2(DATA_DIR / PLANILLAS["Congreso"], destino)
    wb = openpyxl.load_workbook(destino)
    ws = wb["Facturas"]
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value not in (None, ""):
            for c in range(1, 10):
                ws.cell(r, c).value = None if c != 9 else 0.0
    wb.save(destino)


# ---- historial -------------------------------------------------------------- #
_LOCALES_CACHE: dict[str, "Local"] = {}


def locales_todos() -> dict[str, "Local"]:
    """Los 4 locales con planilla propia (Chimidoha comparte el catálogo de Congreso)."""
    nombres = ["Arcos", "Cabello", "Lafinur", "Congreso"]
    if (COPIAS_DIR / "Chimidoha_FACTURAS_COPIA.xlsx").is_file():
        nombres.append("Chimidoha")
    for n in nombres:
        if n not in _LOCALES_CACHE:
            _LOCALES_CACHE[n] = Local(n)
    return {n: _LOCALES_CACHE[n] for n in nombres}


def neto_efectivo(comp: dict) -> float:
    s = sum(i["total_linea"] for i in comp.get("items", []))
    return round(s + float((comp.get("extra") or {}).get("imp_internos_total") or 0), 2)


def _subconjunto(valores: list[float], objetivo: float, tol: float = 0.01) -> list[int] | None:
    """Índices de un subconjunto de `valores` cuya suma ≈ objetivo (fuerza bruta hasta 14 filas)."""
    from itertools import combinations
    n = len(valores)
    if n == 0:
        return None
    for k in range(1, min(n, 14) + 1):
        for comb in combinations(range(n), k):
            if abs(sum(valores[i] for i in comb) - objetivo) <= max(1.0, tol * abs(objetivo)):
                return list(comb)
    return None


def buscar_en_historial(comp: dict, locales: dict[str, "Local"]) -> list[dict]:
    """Busca el comprobante por N° (columna A) en las 4 planillas. Exige proveedor parecido o (fecha ±4 días y monto)."""
    n = ultimo_entero(comp.get("numero_factura"))
    if n is None or str(comp.get("numero_factura", "")).upper().startswith("R-"):
        return []
    fecha = pd.Timestamp(comp["fecha_factura"]) if comp.get("fecha_factura") else None
    objetivo = neto_efectivo(comp)
    if comp.get("tipo") == "NC":
        objetivo = -objetivo                      # en la planilla las NC van en negativo
    prov = norm(comp.get("proveedor_detectado"))
    out = []
    for nombre, loc in locales.items():
        f = loc.facturas
        m = f[pd.to_numeric(f["N"], errors="coerce") == n]
        if m.empty:
            continue
        sim = max((fuzz.token_set_ratio(prov, norm(p)) for p in m["PROVEEDOR"].dropna().unique()), default=0) if prov else 0
        fecha_ok = fecha is not None and bool(((m["F_FACTURA"] - fecha).abs() <= pd.Timedelta(days=4)).any())
        totales = (m["CANT"].astype(float) * m["PRECIO"].astype(float) * (1 - m["DESC"].astype(float))).round(2).tolist()
        idx = _subconjunto(totales, objetivo)
        monto_ok = idx is not None
        if not ((sim >= 60 and (fecha_ok or monto_ok)) or (fecha_ok and monto_ok)):
            continue
        n_items = len([i for i in comp.get("items", [])
                       if not CARGOS_RE.search(i["descripcion_original"]) and abs(i["total_linea"]) > 0.005])
        if idx is None:
            modo, filas_idx = "MONTO_DISTINTO", list(range(len(m)))
        elif len(idx) == 1 and n_items > 1:
            modo, filas_idx = "AGRUPADO_TOTAL", idx
        elif len(idx) >= max(1, n_items - 1):
            modo, filas_idx = "FILAS", idx
        else:
            modo, filas_idx = "FILAS_PARCIAL", idx
        sel = m.iloc[filas_idx]
        out.append({"local": nombre, "modo": modo, "filas": sel, "todas": m, "proveedor_planilla": str(sel["PROVEEDOR"].iloc[0]),
                    "suma_filas": round(sum(totales[i] for i in filas_idx), 2), "objetivo": objetivo,
                    "similitud_proveedor": sim, "fecha_planilla": str(sel["F_FACTURA"].iloc[0].date())
                    if pd.notna(sel["F_FACTURA"].iloc[0]) else None})
    return out


def prorratear_cargos(comp: dict) -> list[dict]:
    """Prorrateo en memoria de impuestos internos y cargos (envío / servicio logístico) dentro del precio del producto.

    Modo exacto de la planilla de Cabello (factura Quilmes 5548-00008853): el neto de cada producto (ya con su
    descuento) más su parte de cargos e impuestos internos se expresa en la col H *bruta* y se conserva el % de
    descuento en la col I:  H = (neto + cargos prorrateados) / cantidad / (1 − desc).  Los cargos quedan como
    renglones de $0 ('envio', 'IMPUESTOS INTERNOS') solo para trazabilidad — no inflan 'materia prima perdida'."""
    items = comp.get("items", [])
    prods = [i for i in items if not CARGOS_RE.search(i["descripcion_original"])]
    cargos = [i for i in items if CARGOS_RE.search(i["descripcion_original"])]
    imp_int = float((comp.get("extra") or {}).get("imp_internos_total") or 0)
    pool = sum(c["total_linea"] for c in cargos) + imp_int
    if not prods or pool == 0:
        return [dict(i, _cargo=bool(CARGOS_RE.search(i["descripcion_original"]))) for i in items]
    base = sum(p["total_linea"] for p in prods)
    out = []
    for p in prods:
        parte = pool * p["total_linea"] / base if base else 0
        neto = p["total_linea"] + parte
        d = p.get("descuento_pct", 0.0) / 100
        h = neto / p["cantidad"] / (1 - d) if d < 1 else p["precio_unitario"]
        out.append(dict(p, precio_unitario=round(h, 2), total_linea=round(neto, 2), _prorrateo=round(parte, 2),
                        _nota=f"cargos e impuestos internos de ${pool:,.2f} prorrateados en col H"))
    out += [dict(c, precio_unitario=0.0, total_linea=0.0, cantidad=1.0, descuento_pct=0.0, _cargo=True,
                 _nota="prorrateado en el precio del producto") for c in cargos]
    if imp_int and not any(re.search(r"impuestos?\s+internos?", c["descripcion_original"], re.I) for c in cargos):
        out.append({"descripcion_original": "IMPUESTOS INTERNOS", "cantidad": 1.0, "unidad_detectada": "UN",
                    "precio_unitario": 0.0, "descuento_pct": 0.0, "total_linea": 0.0, "_cargo": True,
                    "_nota": f"${imp_int:,.2f} prorrateados en la col H del producto"})
    return out


def aprender_de_historial(ap: dict, comp: dict, match: dict) -> dict:
    """Retro-entrenamiento: aprende descripción de proveedor → presentación (+ SKU) o el modo de carga del proveedor."""
    items = prorratear_cargos(comp) if (comp.get("extra") or {}).get("imp_internos_total") else comp.get("items", [])
    filas = match["todas"] if match["modo"].startswith("FILAS") else match["filas"]
    prov = match["proveedor_planilla"]
    res = {"modo": match["modo"], "aprendidos": [], "regla": None}
    if match["modo"] == "AGRUPADO_TOTAL":
        r = filas.iloc[0]
        modos = ap.setdefault("modos_proveedor", {})
        e = modos.setdefault(norm(prov), {"modo_carga": "AGRUPADO_TOTAL", "proveedor_oficial": prov, "conceptos": {},
                                          "facturas_lote": []})
        e["modo_carga"] = "AGRUPADO_TOTAL"
        e["conceptos"][str(r["PRODUCTO"])] = e["conceptos"].get(str(r["PRODUCTO"]), 0) + 1
        e["concepto"] = str(r["PRODUCTO"])
        e["facturas_lote"] = sorted(set(e["facturas_lote"]) | {str(comp.get("numero_factura"))})
        e["ultima"] = date.today().isoformat()
        res["regla"] = f"{prov}: AGRUPADO_TOTAL → concepto '{r['PRODUCTO']}' (1 fila por factura con el neto total)"
        return res
    if match["modo"] not in ("FILAS", "FILAS_PARCIAL"):
        return res
    libres = set(range(len(filas)))
    pares = []
    for ii, it in enumerate(items):
        if it.get("_cargo"):
            continue
        for fi in range(len(filas)):
            r = filas.iloc[fi]
            tot = abs(float(r["CANT"]) * float(r["PRECIO"]) * (1 - float(r["DESC"])))
            ref = max(abs(it["total_linea"]), 1.0)
            pares.append((abs(tot - abs(it["total_linea"])) / ref, ii, fi))
    usados_i: set[int] = set()
    for dif, ii, fi in sorted(pares):
        if dif > 0.02 or ii in usados_i or fi not in libres:
            continue
        usados_i.add(ii)
        libres.discard(fi)
        pres = str(filas.iloc[fi]["PRODUCTO"])
        aprender(ap, prov, items[ii]["descripcion_original"], pres, match["local"])
        res["aprendidos"].append((items[ii]["descripcion_original"], pres))
    return res


def perfil_historico_proveedores(locales: dict[str, "Local"], minimo: int = 5, umbral: float = 0.8) -> dict:
    """Proveedores cuya historia es mayoritariamente 'una fila por factura con concepto genérico' (Verdulería, Varios…)."""
    todo = pd.concat([l.facturas.assign(LOC=n) for n, l in locales.items()])
    todo["NN"] = pd.to_numeric(todo["N"], errors="coerce")
    g = todo.groupby([todo["PROVEEDOR"].map(norm), "LOC", "NN"]).agg(filas=("PRODUCTO", "size"),
                                                                    prod=("PRODUCTO", "first")).reset_index()
    out = {}
    for prov, s in g.groupby("PROVEEDOR"):
        una = s[s["filas"] == 1]
        if len(s) < minimo or len(una) / len(s) < umbral:
            continue
        top = una["prod"].map(norm).value_counts()
        if top.index[0] in GENERICOS_AGRUPADO and top.iloc[0] / len(una) >= umbral:
            out[prov] = {"facturas": int(len(s)), "una_fila": int(len(una)),
                         "concepto": str(una["prod"][una["prod"].map(norm) == top.index[0]].iloc[0])}
    return out


# ---- remitos con birome ------------------------------------------------------ #
_PESO_RE = re.compile(r"^\s*(?P<p>[A-ZÁÉÍÓÚÑa-záéíóúñ ]+?)\s+(?P<kg>\d+(?:[.,]\d{1,3})?)\s*k\s*g\s*$", re.I)
_NO_RE = re.compile(r"^(?P<p>[^:]+?)\s*:\s*['\"]?\s*NO\b", re.I)


def parse_notas_birome(notas: list[str]) -> tuple[set[str], dict[str, float]]:
    """Anotaciones manuscritas del remito → (productos anulados con 'NO' / tachados, pesos reales en kg)."""
    anulados, pesos = set(), {}
    for n in notas or []:
        m = _NO_RE.match(n)
        if m or re.search(r"tachad|anulad", n, re.I):
            clave = norm((m.group("p") if m else n)).split(" ")[0] if (m or n) else ""
            if clave:
                anulados.add(clave)
            continue
        m = _PESO_RE.match(n)
        if m:
            pesos[norm(m.group("p")).split(" ")[0]] = float(m.group("kg").replace(",", "."))
    return anulados, pesos


def procesar_remito(comp: dict) -> dict:
    """Remito de verdulería → registro de Remitos_Pendientes con renglones anulados ('NO'), kilos reales de balanza
    anotados a mano y costo real efectivo por kilo ingresado."""
    anulados, pesos = parse_notas_birome(comp.get("anotaciones_manuscritas") or [])
    renglones, total_impreso, total_efectivo = [], 0.0, 0.0
    for it in comp.get("items", []):
        clave = norm(it["descripcion_original"]).split(" ")[0]
        impreso = round((it.get("cantidad") or 0) * (it.get("precio_unitario") or 0), 2)
        anulado = clave in anulados
        real_kg = pesos.get(clave)
        total_impreso += impreso
        cant_real = 0.0 if anulado else it["cantidad"]
        efectivo = 0.0 if anulado else impreso
        total_efectivo += efectivo
        r = {"descripcion": it["descripcion_original"], "unidad": it.get("unidad_detectada"),
             "cantidad_facturada_remito": it["cantidad"], "precio_unitario": it["precio_unitario"],
             "total_remito": impreso, "anulado": anulado, "cantidad_real": cant_real,
             "peso_real_recibido_kg": real_kg, "costo_real_por_kg": None, "reclamo_por_peso": 0.0, "nota": ""}
        if anulado:
            r["nota"] = "renglón anulado a mano con 'NO' (cantidad_real = 0)"
        elif real_kg:
            r["costo_real_por_kg"] = round(impreso / real_kg, 2)
            if (it.get("unidad_detectada") or "").upper() == "KG":
                dif = round(it["cantidad"] - real_kg, 3)
                r["reclamo_por_peso"] = round(max(dif, 0) * it["precio_unitario"], 2)
                r["nota"] = (f"remito factura {it['cantidad']:g} kg y balanza marcó {real_kg:g} kg"
                             + (f": diferencia {dif:g} kg = ${r['reclamo_por_peso']:,.2f}" if dif > 0 else " (a favor)"))
            else:
                r["nota"] = f"{it['cantidad']:g} {it.get('unidad_detectada') or 'u'} pesaron {real_kg:g} kg → ${r['costo_real_por_kg']:,.2f}/kg"
        renglones.append(r)
    return {"remito": comp.get("numero_factura"), "fecha": comp.get("fecha_factura"),
            "proveedor": comp.get("proveedor_detectado"), "archivo": comp.get("archivo"),
            "estado": "PENDIENTE", "factura_definitiva": None, "renglones": renglones,
            "total_impreso": round(total_impreso, 2), "total_efectivo": round(total_efectivo, 2),
            "anulados": sorted(anulados), "firmas": comp.get("firmas"), "alertas": []}


def cargar_remitos() -> dict:
    try:
        return json.loads(REMITOS_JSON.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def guardar_remitos(d: dict) -> None:
    REMITOS_JSON.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


def _clave_remito(r: str | None) -> str:
    p = re.findall(r"\d+", str(r or ""))
    return "-".join(p[-2:]) if len(p) >= 2 else (p[-1] if p else "")


def conciliar_remito_factura(remito: dict, factura: dict) -> dict:
    """Cruza un remito pendiente con la factura definitiva: renglones cobrados que estaban anulados ('NO'), kilos
    facturados de más contra la balanza y diferencia total. Devuelve {'alertas', 'diferencia_total', 'detalle'}."""
    alertas, detalle = [], []
    fact = {norm(i["descripcion_original"]).split(" ")[0]: i for i in factura.get("items", [])}
    for r in remito["renglones"]:
        clave = norm(r["descripcion"]).split(" ")[0]
        f = fact.get(clave)
        if r["anulado"]:
            if f:
                alertas.append(f"RECLAMO: '{r['descripcion']}' estaba anulado con 'NO' en el remito y figura cobrado "
                               f"${f['total_linea']:,.2f}")
            else:
                detalle.append(f"{r['descripcion']}: anulado con 'NO' y NO facturado ✓")
            continue
        if not f:
            alertas.append(f"'{r['descripcion']}' ${r['total_remito']:,.2f} del remito no figura en la factura")
            continue
        kg_real = r["peso_real_recibido_kg"]
        if kg_real and (f.get("unidad_detectada") or "").upper() == "KG" and f["cantidad"] - kg_real > 0.05:
            alertas.append(f"RECLAMO: {r['descripcion']} facturado {f['cantidad']:g} kg vs {kg_real:g} kg recibidos "
                           f"(${(f['cantidad'] - kg_real) * f['precio_unitario']:,.2f})")
        elif kg_real:
            detalle.append(f"{r['descripcion']}: facturado {f['cantidad']:g} {f.get('unidad_detectada') or ''} / recibido "
                           f"{kg_real:g} kg ✓")
        if (f.get("unidad_detectada") or "").upper() == "KG" and kg_real:
            r["reclamo_por_peso"] = round(max(f["cantidad"] - kg_real, 0) * f["precio_unitario"], 2)
            r["cantidad_facturada_definitiva"] = f["cantidad"]
        if r["precio_unitario"] and abs(f["precio_unitario"] - r["precio_unitario"]) > 0.01:
            alertas.append(f"{r['descripcion']}: precio remito ${r['precio_unitario']:,.2f} ≠ factura ${f['precio_unitario']:,.2f}")
    neto_f = neto_efectivo(factura)
    return {"alertas": alertas, "detalle": detalle, "total_remito_efectivo": remito["total_efectivo"],
            "total_factura": neto_f, "diferencia_total": round(remito["total_efectivo"] - neto_f, 2)}


# ---- Notas de crédito -------------------------------------------------------- #
def clasificar_nc(comp: dict, vigente_ratio: float | None = None) -> tuple[str, str]:
    """'A' devolución física / faltante (producto específico, cantidad negativa, resta stock) o 'B' diferencia de
    precio / bonificación comercial (sin movimiento físico → 'Nota de credito Materia Prima')."""
    txt = " ".join([str(comp.get("referencia") or ""), " ".join(i["descripcion_original"] for i in comp.get("items", []))])
    if NC_B_RE.search(txt) and not NC_A_RE.search(txt):
        return "B", "bonificación / diferencia de precio"
    if NC_A_RE.search(txt):
        return "A", "devolución física / faltante / error de carga"
    if vigente_ratio is not None and vigente_ratio < 0.5:
        return "B", "importe parcial frente al precio vigente (diferencia de precio)"
    return "A", "sin motivo escrito: se trata como devolución del producto"


def filas_nc(comp: dict, loc: "Local", ap: dict, proveedor: str | None) -> list[dict]:
    items = comp.get("items", [])
    ratio = None
    if len(items) == 1:
        m = mapear_producto(items[0]["descripcion_original"], proveedor, loc, ap)
        if m["presentacion"]:
            vig, _ = loc.costo_vigente(m["presentacion"])
            cat = loc.catalogo[loc.catalogo["PRESENTACION"] == m["presentacion"]].iloc[0]
            env = (_num(cat["UN_PRES"]) or 1) * (_num(cat["CANT_PRES"]) or 1)
            if vig and env:
                ratio = items[0]["precio_unitario"] / (vig * env)
    caso, motivo = clasificar_nc(comp, ratio)
    base = {"A": comp.get("numero_factura"), "B": comp.get("fecha_factura"), "C": comp.get("fecha_factura"),
            "D": proveedor or (comp.get("proveedor_detectado") or "").upper(), "proveedor_oficial": proveedor}
    filas = []
    if caso == "B":
        total = round(sum(i["total_linea"] for i in items), 2)
        cat = loc.catalogo[loc.catalogo["PRESENTACION"] == PRESENTACION_NC_MP]
        f = dict(base, descripcion_original=f"NC {comp.get('numero_factura')} ({motivo})", presentacion=PRESENTACION_NC_MP,
                 sku=cat.iloc[0]["SKU"] if len(cat) else "LV0218", confianza=1.0, metodo="NC caso B (dif. precio/bonificación)",
                 G=-1.0, H=total, I=0.0, J=-total, P_UN=-total, vigente=None, var_pct=None, accion="INYECTAR",
                 motivos=[], avisos=[f"NC caso B: {motivo}; no resta kilos del inventario físico"], solicitud=False,
                 cant_orig=-1.0, precio_orig=total, alternativas=[], tipo_nc="B")
        filas.append(f)
        return filas
    for it in items:
        d = decidir_renglon(dict(it, cantidad=it["cantidad"]), comp, proveedor, loc, ap)
        d.update(base)
        if d["G"] is not None:
            d["G"], d["J"] = -abs(d["G"]), -abs(d["J"] or 0)
            d["P_UN"] = d["P_UN"]
        d["cant_orig"] = -abs(it["cantidad"])
        d["avisos"].append(f"NC caso A: {motivo}; cantidad negativa resta stock y gasto")
        d["tipo_nc"] = "A"
        # el precio vigente se compara contra el precio del renglón (positivo), el signo va en la cantidad
        filas.append(d)
    return filas


# ---- carga agrupada, prorrateo y orquestación por comprobante ---------------- #
def fila_simple(**k) -> dict:
    base = {"descripcion_original": "", "presentacion": None, "sku": None, "confianza": 0.0, "metodo": "",
            "G": None, "H": None, "I": 0.0, "J": None, "P_UN": None, "vigente": None, "var_pct": None,
            "accion": "REVISAR", "motivos": [], "avisos": [], "solicitud": False, "cant_orig": None,
            "precio_orig": None, "alternativas": [], "proveedor_oficial": None}
    base.update(k)
    return base


def construir_filas(comp: dict, loc: "Local", ap: dict) -> dict:
    """Decide cómo se imputa un comprobante. Devuelve {'proveedor_oficial', 'modo', 'filas', 'notas'}."""
    proveedor, conf_p = mapear_proveedor(comp.get("proveedor_detectado"), comp.get("texto_encabezado", ""), loc.proveedores)
    notas: list[str] = []
    fecha_f = comp.get("fecha_factura")
    meta = {"A": numero_a_celda(comp.get("numero_factura")), "B": fecha_f, "C": comp.get("fecha_ingreso") or fecha_f,
            "D": proveedor or (comp.get("proveedor_detectado") or "").upper(), "proveedor_oficial": proveedor}
    modos = ap.get("modos_proveedor", {})
    modo_prov = modos.get(norm(proveedor)) or modos.get(norm(comp.get("proveedor_detectado"))) if proveedor else None
    if comp.get("tipo") == "NC":
        return {"proveedor_oficial": proveedor, "modo": "NC", "filas": filas_nc(comp, loc, ap, proveedor), "notas": notas}
    if modo_prov and len(comp.get("items", [])) == 1 and modo_prov.get("modo_carga") == "AGRUPADO_TOTAL":
        # un único renglón que ya se cargó por su nombre (p. ej. 'Huevo cajon' de Open Fruit) va por su nombre
        mp = mapear_producto(comp["items"][0]["descripcion_original"], proveedor, loc, ap)
        if mp["presentacion"] and mp["confianza"] >= CONF_MINIMA and mp["metodo"].startswith("aprendizaje")                 and norm(mp["presentacion"]) not in GENERICOS_AGRUPADO:
            modo_prov = None
    if modo_prov and modo_prov.get("modo_carga") == "AGRUPADO_TOTAL" and len(comp.get("items", [])) >= 1:
        total = neto_efectivo(comp)
        cat = loc.catalogo[loc.catalogo["_norm"] == norm(modo_prov["concepto"])]
        f = fila_simple(**meta, descripcion_original=f"{len(comp['items'])} renglones agrupados",
                        presentacion=modo_prov["concepto"], sku=cat.iloc[0]["SKU"] if len(cat) else None,
                        confianza=0.99, metodo="AGRUPADO_TOTAL (regla aprendida)", G=1.0, H=total, I=0.0, J=total,
                        P_UN=total, accion="INYECTAR" if len(cat) else "REVISAR", cant_orig=1.0, precio_orig=total,
                        avisos=[f"proveedor cargado como 1 fila '{modo_prov['concepto']}' por el neto total"])
        v = comp.get("validacion") or {}
        if v.get("ok") is False:
            f["accion"], f["motivos"] = "REVISAR", [f"comprobante no cierra: {v.get('motivo')}"]
        return {"proveedor_oficial": proveedor, "modo": "AGRUPADO_TOTAL", "filas": [f], "notas": notas}
    con_imp_internos = bool((comp.get("extra") or {}).get("imp_internos_total"))
    items = prorratear_cargos(comp) if con_imp_internos else comp.get("items", [])
    filas = []
    for it in items:
        if not con_imp_internos and CARGOS_RE.search(it["descripcion_original"]):
            # envío / flete / servicio logístico sin impuestos internos: renglón propio 'Envio' (convención de Congreso)
            cat = loc.catalogo[loc.catalogo["_norm"] == "envio"]
            filas.append(fila_simple(**meta, descripcion_original=it["descripcion_original"], presentacion="Envio",
                                     sku=cat.iloc[0]["SKU"] if len(cat) else None, confianza=1.0, metodo="cargo de envío",
                                     G=float(it["cantidad"]), H=float(it["precio_unitario"]),
                                     I=round((it.get("descuento_pct") or 0) / 100, 4), J=it["total_linea"],
                                     P_UN=it["total_linea"], accion="INYECTAR", cant_orig=it["cantidad"],
                                     precio_orig=it["precio_unitario"]))
            continue
        if it.get("_cargo"):
            nombre = "IMPUESTOS INTERNOS" if re.search(r"impuestos?\s+internos?", it["descripcion_original"], re.I) else "Envio"
            cat = loc.catalogo[loc.catalogo["_norm"] == norm(nombre)]
            filas.append(fila_simple(**meta, descripcion_original=it["descripcion_original"], presentacion=nombre,
                                     sku=cat.iloc[0]["SKU"] if len(cat) else None, confianza=1.0, metodo="cargo prorrateado",
                                     G=1.0, H=0.0, J=0.0, P_UN=0.0, accion="INYECTAR", cant_orig=1.0, precio_orig=0.0,
                                     avisos=[it.get("_nota", "")]))
            continue
        d = decidir_renglon(it, comp, proveedor, loc, ap)
        d.update(meta)
        if it.get("_prorrateo"):
            d["avisos"].append(it["_nota"])
        filas.append(d)
    v = comp.get("validacion") or {}
    if v.get("ok") is False:
        for f in filas:
            if f["accion"] == "INYECTAR":
                f["accion"] = "REVISAR"
            f["motivos"].append(f"comprobante no cierra: {v.get('motivo')}")
    _anotar_transferencia(filas, loc)
    return {"proveedor_oficial": proveedor, "modo": "PRORRATEO" if items is not comp.get("items") else "FILAS",
            "filas": filas, "notas": notas}


def _anotar_transferencia(filas: list[dict], loc: "Local") -> None:
    """Congreso (Obrador) compra y luego 'vende' a Chimidoha y a los locales: cada fila INYECTAR lleva el precio de
    transferencia por la matriz (panificados/descartables x1.20, carnes/pescados x1.10, almacén al costo)."""
    if loc.nombre != "Congreso":
        return
    import transferencias_obrador as tr
    for f in filas:
        if f.get("accion") != "INYECTAR" or not f.get("presentacion") or not f.get("P_UN"):
            continue
        cat = loc.catalogo[loc.catalogo["PRESENTACION"] == f["presentacion"]]
        if cat.empty:
            continue
        cat2 = cat.iloc[0]["CAT2"]
        f["transferencia"] = {"tipo": tr.tipo_transferencia(cat2), "factor": tr.factor_transferencia(cat2),
                              "p_un_transferencia": tr.precio_transferencia(f["P_UN"], cat2)}


# ---- publicación en la planilla real (solo con confirmación) ----------------- #
IDS_PLANILLAS = {"Congreso": "1udLEQpQ29avK4c6fNnQig-w02xIV82IgLe85yVJXaOY",
                 "Arcos": "1frh93qB8Od5mrtAfMuObKwRW93p7kev6DNgmERqz6ho",
                 "Lafinur": "1tv0uI4KsSEPRUecYfnwcDqpGbUDelGqWwe3Hp-fikOc",
                 "Chimidoha": "1gSHbDSgICeTy6lZpyIPRskFbEs0JB5mxw1e71oOkIzU",     # planilla "CHIMIDOHA Congreso"
                 "Cabello": "1zAxYyNB98Sb1-D0HtfokESXc6FrDS96EbFYyMmsseIE"}


_IDS_PRODUCCION = dict(IDS_PLANILLAS)
import entornos as _entornos                                # noqa: E402
if _entornos.entorno() == "PRUEBA":                          # rutas explícitas por ambiente
    IDS_PLANILLAS = _entornos.ids(_IDS_PRODUCCION)


def _hoja_por_clave(sh, clave: str):
    """Busca una pestaña por nombre ignorando emojis, candados y espacios ('🔒 Facturas', 'Remitos'...).
    Prefiere el nombre exacto (singular o plural) y descarta '... pendientes'."""
    letras = lambda s: re.sub(r"[^a-z]", "", unicodedata.normalize("NFKD", s.lower()).encode("ascii", "ignore").decode())
    clave = clave.lower()
    hojas = [w for w in sh.worksheets() if "pendiente" not in w.title.lower()]
    for w in hojas:
        if letras(w.title) in (clave, clave + "s"):
            return w
    return next((w for w in hojas if clave in letras(w.title)), None)


def _columnas_de_carga(ws) -> dict:
    """Mira la fila 2 de la pestaña real: las columnas con fórmula (=MAP/ARRAYFORMULA...) NO se pisan nunca.
    Devuelve las columnas A:I de carga manual y, si el código/SKU (col W/X) no tiene fórmula, dónde inyectarlo."""
    datos = ws.get("A1:AA8", value_render_option="FORMULA")
    head = datos[0] if datos else []
    # Una columna es "con fórmula" solo si la mayoría de las filas 2-8 trae una fórmula DERIVADA (función o referencia a
    # celdas). Un '=2*12' o '=10,5+10,6' tipeado como cantidad es un dato manual, no una columna calculada.
    derivada = re.compile(r"[A-Za-z]{1,3}\$?\d+|[A-Za-z_]+\(")
    filas = [r for r in datos[1:8]]
    con_formula = set()
    for i in range(max((len(r) for r in filas), default=0)):
        n = sum(1 for r in filas if i < len(r) and isinstance(r[i], str) and r[i].lstrip().startswith("=")
                and derivada.search(r[i]))
        if filas and n * 2 > len(filas) - 0:
            con_formula.add(i)
    fila2 = []
    manuales = [i for i in range(9) if i not in con_formula]
    sku_col = next((i for i, h in enumerate(head) if i >= 9 and re.search(r"codigo|sku", norm(h))), None)
    return {"manuales": manuales, "con_formula": sorted(con_formula),
            "sku_col": sku_col if (sku_col is not None and sku_col not in con_formula) else None,
            "sku_formula": sku_col is not None and sku_col in con_formula, "encabezado": head}


def _col_letra(i: int) -> str:
    s, i = "", i + 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _gspread():
    import gspread
    from google.oauth2.service_account import Credentials
    c = Credentials.from_service_account_file(
        str(BASE_DIR / "credentials.json"),
        scopes=["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.readonly"])
    gc = gspread.authorize(c)
    _abrir = gc.open_by_key

    def open_by_key(key, *a, **k):                           # control por ID: en PRUEBA rechaza producción
        import entornos
        entornos.validar_destino(key, _IDS_PRODUCCION)
        return _abrir(key, *a, **k)
    gc.open_by_key = open_by_key
    return gc, c


def verificar_acceso_planillas() -> dict:
    """Qué planillas reales ve la Service Account y con qué permiso (solo lectura de metadatos; no escribe)."""
    from google.auth.transport.requests import AuthorizedSession
    gc, creds = _gspread()
    ses = AuthorizedSession(creds)
    out = {}
    for local, sid in IDS_PLANILLAS.items():
        try:
            r = ses.get(f"https://www.googleapis.com/drive/v3/files/{sid}",
                        params={"fields": "name,capabilities(canEdit)", "supportsAllDrives": "true"})
            r.raise_for_status()
            d = r.json()
            sh = gc.open_by_key(sid)
            titulos = [w.title for w in sh.worksheets()]
            out[local] = {"ok": True, "nombre": d.get("name"), "editor": d.get("capabilities", {}).get("canEdit"),
                          "tiene_facturas": "Facturas" in titulos, "tiene_remitos": "Remitos" in titulos,
                          "pestañas": len(titulos)}
        except Exception as e:
            out[local] = {"ok": False, "error": f"{type(e).__name__}: {str(e)[:120]}"}
    return out


def _celda_a_texto_fecha(iso: str | None):
    if iso and re.fullmatch(r"\d{4}-\d{2}-\d{2}", str(iso)):
        y, m, d = str(iso).split("-")
        return f"{d}/{m}/{y}"
    return iso


def _fila_push(f: dict, demo: bool = False) -> list:
    reviso = f.get("accion") == "REVISAR"
    nota = ""
    if demo:
        nota = "DEMO IA - borrar con /deshacer"
    elif reviso:
        sug = f" | sugerido: {f['presentacion']} (conf {f['confianza']:.2f})" if f.get("presentacion") else ""
        nota = f"{f['descripcion_original']}{sug} | {'; '.join(f.get('motivos') or [])}"[:250]
    g = f["G"] if f.get("G") is not None else f.get("cant_orig")
    h = f["H"] if f.get("H") is not None else f.get("precio_orig")
    return [f["A"], _celda_a_texto_fecha(f["B"]), _celda_a_texto_fecha(f["C"]), f["D"],
            REVISAR_TXT if reviso else f["presentacion"], nota, g, h, f["I"]]


ENCABEZADO_REMITOS = ["FECHA", "PRODUCTO", "CANT. U/MEDIDA", "PRECIO U.", "PRECIO TOTAL"]


def filas_remito_obrador(rem: dict) -> list[list]:
    """Renglones de un remito del Obrador en las 5 columnas reales de la pestaña 'Remitos' del local:
    A FECHA (DD/MM) · B PRODUCTO · C CANT. U/MEDIDA · D PRECIO U. · E PRECIO TOTAL (=C*D, se completa con la fila)."""
    return [[r["fecha"], r["producto"], r["cantidad"], r["precio_unitario"]] for r in rem["renglones"]]


def agregar_remitos_obrador_a_hoja(wb, remitos: list[dict]) -> int:
    """Copia local: agrega los remitos del Obrador a la pestaña 'Remitos' (se crea si no existe: sin cuadrícula, formatos
    dd/MM, 0.##, $#,##0.00). Sin fila de total general. No duplica (fecha + producto + cantidad)."""
    from openpyxl.styles import Font, PatternFill
    if "Remitos" in wb.sheetnames:
        ws = wb["Remitos"]
    else:
        ws = wb.create_sheet("Remitos")
        ws.append(ENCABEZADO_REMITOS)
        for c in ws[1]:
            c.font = Font(bold=True)
        for col, w in zip("ABCDE", (12, 34, 16, 14, 16)):
            ws.column_dimensions[col].width = w
    ws.sheet_view.showGridLines = False
    ultima = max((r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value not in (None, "")), default=1)
    ya = {(str(ws.cell(r, 1).value)[:10], str(ws.cell(r, 2).value).strip().lower(), _num(ws.cell(r, 3).value))
          for r in range(2, ultima + 1)}
    n = 0
    for rem in remitos:
        for r in rem["renglones"]:
            f = _fecha_celda(r["fecha"])
            clave = (str(f)[:10] if f else None, r["producto"].strip().lower(), float(r["cantidad"]))
            if clave in ya:
                continue
            ultima += 1
            ws.cell(ultima, 1).value = f
            ws.cell(ultima, 1).number_format = "dd/MM"
            ws.cell(ultima, 2).value = r["producto"]
            ws.cell(ultima, 3).value = r["cantidad"]
            ws.cell(ultima, 3).number_format = "0.##"
            ws.cell(ultima, 4).value = r["precio_unitario"]
            ws.cell(ultima, 4).number_format = "$#,##0.00"
            ws.cell(ultima, 5).value = f"=C{ultima}*D{ultima}"
            ws.cell(ultima, 5).number_format = "$#,##0.00"
            for c in range(1, 6):
                ws.cell(ultima, c).fill = PatternFill("solid", fgColor="FFF8E1")
            ya.add(clave)
            n += 1
    return n


def procesar_remito_obrador(comp: dict, loc_receptor: "Local", ap: dict) -> dict:
    """Remito emitido por el Obrador (Mishiguene, Av. Cerviño 4417) hacia un local. Si un renglón viene sin precio se
    calcula con la matriz de transferencia (transferencias_obrador.py) sobre el último costo vigente en Congreso."""
    import transferencias_obrador as tr
    congreso = _LOCALES_CACHE.get("Congreso") or Local("Congreso")
    reng = []
    for it in comp.get("items", []):
        precio, origen = it.get("precio_unitario"), "remito"
        if not precio:
            m = mapear_producto(it["descripcion_original"], None, congreso, ap)
            vig, cat2 = None, None
            if m["presentacion"] and m["confianza"] >= 0.7:
                vig, _ = congreso.costo_vigente(m["presentacion"])
                cat2 = congreso.catalogo[congreso.catalogo["PRESENTACION"] == m["presentacion"]].iloc[0]["CAT2"]
            if vig:
                precio, origen = tr.precio_transferencia(vig, cat2), f"transferencia x{tr.factor_transferencia(cat2)} s/costo vigente"
            else:
                precio, origen = None, "sin precio: REVISAR"
        reng.append({"fecha": comp.get("fecha_factura"), "producto": it["descripcion_original"],
                     "cantidad": it["cantidad"], "precio_unitario": precio, "origen_precio": origen,
                     "total": round(it["cantidad"] * precio, 2) if precio else None})
    return {"remito": comp.get("numero_factura"), "tipo": "OBRADOR", "fecha": comp.get("fecha_factura"),
            "emisor": comp.get("proveedor_detectado"), "receptor": loc_receptor.nombre, "renglones": reng,
            "total": round(sum(r["total"] or 0 for r in reng), 2), "estado": "ENTREGADO_A_LOCAL"}


def asegurar_hoja_remitos_real(sh) -> str:
    """Planilla real: crea la pestaña 'Remitos' (5 columnas, sin cuadrícula) solo si no existe. Arcos, Cabello y Lafinur
    ya la tienen; Chimidoha no."""
    if _hoja_por_clave(sh, "remito") is not None:
        return "ya existía"
    rem = sh.add_worksheet(title="Remitos", rows=1000, cols=5)
    rem.update("A1:E1", [ENCABEZADO_REMITOS])
    sh.batch_update({"requests": [
        {"updateSheetProperties": {"properties": {"sheetId": rem.id, "gridProperties": {
            "hideGridlines": True, "frozenRowCount": 1}}, "fields": "gridProperties.hideGridlines,gridProperties.frozenRowCount"}},
        {"repeatCell": {"range": {"sheetId": rem.id, "startRowIndex": 0, "endRowIndex": 1},
                        "cell": {"userEnteredFormat": {"textFormat": {"bold": True}}}, "fields": "userEnteredFormat.textFormat.bold"}},
        {"repeatCell": {"range": {"sheetId": rem.id, "startRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": 1},
                        "cell": {"userEnteredFormat": {"numberFormat": {"type": "DATE", "pattern": "dd/MM"}}},
                        "fields": "userEnteredFormat.numberFormat"}},
        {"repeatCell": {"range": {"sheetId": rem.id, "startRowIndex": 1, "startColumnIndex": 3, "endColumnIndex": 5},
                        "cell": {"userEnteredFormat": {"numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}}},
                        "fields": "userEnteredFormat.numberFormat"}},
    ]})
    return "creada (5 columnas, sin cuadrícula)"


ENCABEZADO_PENDIENTES = ["ID REMITO", "N° REMITO", "FECHA", "PROVEEDOR", "PRODUCTO", "UNIDAD",
    "CANTIDAD IMPRESA", "CANTIDAD RECIBIDA", "PRECIO INFORMADO", "VALOR ESTIMADO", "ORIGEN PRECIO",
    "ANULADO", "ESTADO", "FACTURA ASOCIADA", "ARCHIVO", "ALERTAS"]


def filas_remitos_pendientes(remitos: list[dict]) -> list[list]:
    import hashlib
    filas = []
    for reg in remitos:
        ident = "|".join([str(reg.get("remito") or reg.get("archivo") or ""),
                          norm(reg.get("proveedor")), str(reg.get("local") or "")])
        rid = hashlib.sha256(ident.encode()).hexdigest()[:24]
        for i, r in enumerate(reg["renglones"]):
            recibida = 0 if r["anulado"] else r.get("cantidad_real")
            if not r["anulado"] and (r.get("unidad") or "").upper() == "KG" and r.get("peso_real_recibido_kg") is not None:
                recibida = r["peso_real_recibido_kg"]
            filas.append([f"{rid}:{i}", reg.get("remito") or "SIN NUMERO: REVISAR", reg.get("fecha") or "",
                          reg.get("proveedor_oficial") or reg.get("proveedor") or "", r["descripcion"], r.get("unidad") or "",
                          r.get("cantidad_facturada_remito"), recibida, r.get("precio_unitario"), r.get("total_estimado"),
                          r.get("precio_origen") or "sin estimación", "SI" if r["anulado"] else "NO", reg["estado"],
                          str(reg.get("factura_definitiva") or ""), reg.get("archivo") or "",
                          " | ".join(reg.get("alertas") or []) + (" | " + r["nota"] if r.get("nota") else "")])
    return filas


def push_remitos_proveedor_drive(local: str, remitos: list[dict], confirmado: bool = False) -> dict:
    """Registro separado de Facturas y de transferencias: no vuelve a imputar gasto/stock."""
    filas = filas_remitos_pendientes(remitos)
    if not confirmado or local not in IDS_PLANILLAS or not filas:
        return {"escrito": False, "local": local, "motivo": "sin confirmación, destino o renglones", "plan": filas}
    gc, _ = _gspread()
    sh = gc.open_by_key(IDS_PLANILLAS[local])
    ws = next((w for w in sh.worksheets() if w.title == "Remitos_Pendientes"), None)
    if ws is None:
        ws = sh.add_worksheet(title="Remitos_Pendientes", rows=1000, cols=16)
        ws.update(range_name="A1:P1", values=[ENCABEZADO_PENDIENTES], value_input_option="RAW")
        ws.freeze(rows=1)
    actuales = ws.get_all_values()
    if actuales and actuales[0][:16] != ENCABEZADO_PENDIENTES:
        raise ValueError("Remitos_Pendientes tiene otra estructura; no se modifica")
    idx = {fila[0]: n for n, fila in enumerate(actuales[1:], 2) if fila and fila[0]}
    proxima = len(actuales) + 1
    updates = []
    for fila in filas:
        n = idx.get(fila[0])
        if n is None:
            n = proxima
            proxima += 1
            idx[fila[0]] = n
        updates.append({"range": f"A{n}:P{n}", "values": [["" if x is None else x for x in fila]]})
    ws.batch_update([dict(u) for u in updates], value_input_option="RAW")   # gspread antepone el título al dict recibido
    for u in updates:
        leido = ws.get(u["range"], value_render_option="UNFORMATTED_VALUE")
        expected = u["values"][0]
        actual = (list(leido[0]) if leido else [])
        actual += [""] * (len(expected) - len(actual))
        if ["" if x is None else x for x in actual] != expected:
            raise ValueError("No se verificó la escritura de remitos; se conserva pendiente para reintentar")
    return {"escrito": True, "local": local, "tipo": "remitos_proveedor", "planilla": sh.title,
            "sheet_id": IDS_PLANILLAS[local], "hoja": ws.title, "filas": len(filas),
            "desde_fila": min(idx[f[0]] for f in filas), "hasta_fila": max(idx[f[0]] for f in filas), "detalle": []}


def _asegurar_filas(ws, ultima: int) -> int:
    """La API de Sheets rechaza (400 'exceeds grid limits') escribir fuera de la grilla: agrega las filas que falten.
    Devuelve cuántas agregó. Es la causa confirmada del APIError al cargar el remito 3727 en Cabello (grilla de 608)."""
    falta = ultima - ws.row_count
    if falta > 0:
        ws.add_rows(falta + 20)
    return max(falta, 0)


def _bloqueo_remitos(sh, local: str, remitos: list[dict], ws, primera: int, ultima: int, error: str) -> dict:
    """La pestaña Remitos está protegida o sin grilla: no se pierde el remito. Se registra el bloqueo en
    estado/bloqueos_permisos.json (cola de reintento limpio) y se vuelca un respaldo trazable en `Remitos_Pendientes`
    con ESTADO=BLOQUEADO_PERMISOS (no suma gasto ni stock). Devuelve el diagnóstico exacto."""
    import hashlib
    ahora = datetime.now().isoformat(timespec="seconds") if "datetime" in globals() else ""
    diag = {"escrito": False, "bloqueado": True, "local": local, "planilla": sh.title, "hoja": ws.title,
            "grilla_filas": ws.row_count, "rango_pretendido": f"A{primera}:E{ultima}",
            "error": re.sub(r"\s+", " ", error)[:240],
            "pasos": ["En Google Sheets: Datos ▸ Proteger hojas y rangos ▸ pestaña Remitos ▸ quitar la protección o agregar a "
                      "la Service Account agente-mishi@black-obelisk-460914-d9.iam.gserviceaccount.com como editora.",
                      f"Agregar al menos {max(ultima - ws.row_count, 0) + 20} filas a la pestaña (la grilla tiene {ws.row_count}).",
                      "Reintentar: python recuperar_lote.py apply (la cola estado/bloqueos_permisos.json conserva el remito)."]}
    f = BASE_DIR / "estado" / "bloqueos_permisos.json"
    f.parent.mkdir(exist_ok=True)
    try:
        cola = json.loads(f.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        cola = []
    for rem in remitos:
        cola.append({"ts": ahora, "tipo": "remito_obrador", "remito": rem.get("remito"), "local": local,
                     "renglones": rem.get("renglones"), "total": rem.get("total"), "diagnostico": diag["error"],
                     "estado": "BLOQUEADO_PERMISOS"})
    tmp = f.with_suffix(".tmp")
    tmp.write_text(json.dumps(cola, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    tmp.replace(f)
    try:                                                      # respaldo en la nube
        pend = next((w for w in sh.worksheets() if w.title == "Remitos_Pendientes"), None)
        if pend is None:
            pend = sh.add_worksheet(title="Remitos_Pendientes", rows=1000, cols=16)
            pend.update(range_name="A1:P1", values=[ENCABEZADO_PENDIENTES], value_input_option="RAW")
        vivos = {x[0] for x in pend.get_all_values()[1:]}
        nuevas = []
        for rem in remitos:
            base = hashlib.sha256(f"{rem.get('remito')}|obrador|{local}".encode()).hexdigest()[:24]
            for i, r in enumerate(rem["renglones"]):
                rid = f"{base}:{i}"
                if rid in vivos:
                    continue
                nuevas.append([rid, rem.get("remito"), str(r.get("fecha")), "OBRADOR (transferencia a " + local + ")",
                               r["producto"], "UN", r["cantidad"], r["cantidad"], r["precio_unitario"], r.get("total"),
                               "remito Obrador", "NO", "BLOQUEADO_PERMISOS", "", rem.get("archivo") or "",
                               "Pestaña Remitos protegida/llena: pendiente de cargar en Remitos"])
        if nuevas:
            pend.append_rows(nuevas, value_input_option="RAW")
        diag["respaldo_remitos_pendientes"] = len(nuevas)
    except Exception as e:                                    # noqa: BLE001
        diag["respaldo_remitos_pendientes"] = f"no se pudo: {type(e).__name__}"
    diag["motivo"] = "pestaña Remitos protegida o sin filas: ver pasos"
    return diag


def push_remitos_drive(local: str, remitos: list[dict], confirmado: bool = False) -> dict:
    """Remitos del Obrador hacia la pestaña 'Remitos' de la planilla REAL del local receptor, en sus 5 columnas:
    A FECHA (DD/MM) · B PRODUCTO · C CANT. U/MEDIDA · D PRECIO U. · E PRECIO TOTAL (=C*D). Sin fila de total general.
    No duplica (fecha + producto + cantidad). Pinta A:E de #FFF8E1."""
    if local not in IDS_PLANILLAS or not remitos:
        return {"escrito": False, "motivo": f"{local}: sin planilla real o sin remitos"}
    plan = [[r["fecha"], r["producto"], r["cantidad"], r["precio_unitario"]] for rem in remitos for r in rem["renglones"]]
    if not confirmado:
        return {"escrito": False, "motivo": "falta confirmación", "plan": plan}
    if any(p[3] is None for p in plan):
        return {"escrito": False, "motivo": "hay renglones sin precio: se revisan antes de escribir", "plan": plan}
    gc, _ = _gspread()
    sh = gc.open_by_key(IDS_PLANILLAS[local])
    estado_hoja = asegurar_hoja_remitos_real(sh)
    ws = _hoja_por_clave(sh, "remito")
    col_a, col_b, col_c = ws.col_values(1), ws.col_values(2), ws.col_values(3)
    n_filas = len(col_a)
    ya = {(col_a[i][:5], (col_b[i] if i < len(col_b) else "").strip().lower(), (col_c[i] if i < len(col_c) else "").replace(",", "."))
          for i in range(1, n_filas)}
    filas, detalle = [], []
    primera = n_filas + 1
    for rem in remitos:
        for r in rem["renglones"]:
            ddmm = _celda_a_texto_fecha(r["fecha"])
            clave = (ddmm[:5], r["producto"].strip().lower(), f"{r['cantidad']:g}")
            if clave in ya:
                continue
            fila_n = primera + len(filas)
            filas.append([ddmm, r["producto"], r["cantidad"], r["precio_unitario"], f"=C{fila_n}*D{fila_n}"])
            detalle.append({"fila": fila_n, "FECHA": ddmm[:5], "PRODUCTO": r["producto"], "CANT": r["cantidad"],
                            "PRECIO U": r["precio_unitario"], "TOTAL": r["total"]})
    if not filas:
        return {"escrito": False, "motivo": "el remito ya estaba cargado (no se duplica)", "hoja": estado_hoja}
    ultima = primera + len(filas) - 1
    try:
        _asegurar_filas(ws, ultima)
        ws.update(f"A{primera}:E{ultima}", filas, value_input_option="USER_ENTERED")
    except Exception as e:                                    # noqa: BLE001
        txt = str(e)
        if "protected" in txt or "exceeds grid limits" in txt:
            return _bloqueo_remitos(sh, local, remitos, ws, primera, ultima, txt)
        raise
    ws.batch_format([{"range": f"A{primera}:E{ultima}",
                      "format": {"backgroundColor": {"red": 1.0, "green": 0.973, "blue": 0.882},
                                 "numberFormat": {"type": "NUMBER", "pattern": "0.##"}}},
                     {"range": f"A{primera}:A{ultima}", "format": {"numberFormat": {"type": "DATE", "pattern": "dd/MM"}}},
                     {"range": f"D{primera}:E{ultima}", "format": {"numberFormat": {"type": "CURRENCY", "pattern": "$#,##0.00"}}}])
    return {"escrito": True, "tipo": "remitos", "local": local, "planilla": sh.title, "sheet_id": IDS_PLANILLAS[local],
            "hoja": ws.title, "hoja_estado": estado_hoja, "desde_fila": primera, "hasta_fila": ultima,
            "filas": len(filas), "detalle": detalle}


def push_drive(local: str, filas: list[dict], confirmado: bool = False, demo: bool = False) -> dict:
    """Agrega filas a la pestaña Facturas (se busca ignorando candados/emojis) de la planilla REAL del local, pintadas
    #FFF8E1 (OK) o #FFEBEE (REVISAR). Escribe SOLO en las columnas de carga manual (A:I) que NO tengan fórmula en la fila 2
    (=MAP/ARRAYFORMULA): jamás pisa una columna calculada. Si la columna de código/SKU (W/X) no tiene fórmula, inyecta
    ahí el SKU oficial. No duplica un N° + proveedor ya cargado (salvo demo=True).
    Solo escribe con confirmado=True. Devuelve el detalle exacto para poder deshacerlo."""
    if local not in IDS_PLANILLAS:
        return {"escrito": False, "local": local,
                "motivo": f"{local} no tiene planilla real de Facturas; queda en la copia local"}
    cand = [f for f in filas if f.get("accion") in ("INYECTAR", "REVISAR") and f.get("A") is not None]
    plan = [_fila_push(f, demo) for f in cand]
    if not confirmado:
        return {"escrito": False, "local": local, "motivo": "falta --confirmo-escribir-produccion", "plan": plan,
                "planilla": IDS_PLANILLAS[local]}
    gc, _ = _gspread()
    sh = gc.open_by_key(IDS_PLANILLAS[local])
    ws = _hoja_por_clave(sh, "factura")
    if ws is None:
        return {"escrito": False, "local": local, "motivo": "no encontré la pestaña Facturas"}
    cols = _columnas_de_carga(ws)
    col_a, col_d = ws.col_values(1), ws.col_values(4)
    existentes = {(ultimo_entero(a), norm(d)) for a, d in zip(col_a[1:], (col_d + [""] * len(col_a))[1:]) if a}
    nuevas, omitidas = [], []
    for f, fila in zip(cand, plan):
        clave = (ultimo_entero(fila[0]), norm(fila[3]))
        if not demo and clave in existentes:
            omitidas.append({"N": fila[0], "proveedor": fila[3], "presentacion": fila[4]})
            continue
        nuevas.append((f, fila))
    if not nuevas:
        return {"escrito": False, "local": local, "motivo": "todas las filas ya estaban cargadas (no se duplica)",
                "omitidas": omitidas}
    primera = len(col_a) + 1
    ultima = primera + len(nuevas) - 1
    # bloques contiguos de columnas manuales
    bloques, actual = [], []
    for i in cols["manuales"]:
        if actual and i != actual[-1] + 1:
            bloques.append(actual)
            actual = []
        actual.append(i)
    if actual:
        bloques.append(actual)
    datos = []
    for bl in bloques:
        datos.append({"range": f"'{ws.title}'!{_col_letra(bl[0])}{primera}:{_col_letra(bl[-1])}{ultima}",
                      "values": [[fila[i] for i in bl] for _, fila in nuevas]})
    if cols["sku_col"] is not None:
        sc = _col_letra(cols["sku_col"])
        datos.append({"range": f"'{ws.title}'!{sc}{primera}:{sc}{ultima}",
                      "values": [[f.get("sku") or ""] for f, _ in nuevas]})
    _asegurar_filas(ws, ultima)
    sh.values_batch_update({"valueInputOption": "USER_ENTERED", "data": datos})
    amarillo = {"backgroundColor": {"red": 1.0, "green": 0.973, "blue": 0.882}}      # #FFF8E1
    rojo = {"backgroundColor": {"red": 1.0, "green": 0.922, "blue": 0.933}}          # #FFEBEE
    ws.batch_format([{"range": f"A{primera + i}:I{primera + i}",
                      "format": rojo if f.get("accion") == "REVISAR" else amarillo}
                     for i, (f, _) in enumerate(nuevas)])
    return {"escrito": True, "local": local, "planilla": sh.title, "sheet_id": IDS_PLANILLAS[local],
            "hoja": ws.title, "desde_fila": primera, "hasta_fila": ultima, "filas": len(nuevas), "demo": demo,
            "omitidas": omitidas,
            "columnas_escritas": [_col_letra(i) for i in cols["manuales"]]
            + ([_col_letra(cols["sku_col"])] if cols["sku_col"] is not None else []),
            "columnas_con_formula_respetadas": [_col_letra(i) for i in cols["con_formula"]],
            "detalle": [{"fila": primera + i, "N": fila[0], "proveedor": fila[3], "E": fila[4], "G": fila[6], "H": fila[7],
                         "sku": f.get("sku"),
                         "color": "rojo suave #FFEBEE" if f.get("accion") == "REVISAR" else "amarillo suave #FFF8E1"}
                        for i, (f, fila) in enumerate(nuevas)]}


def deshacer_push(registro: dict) -> dict:
    """Borra de la planilla real las filas de una inyección anterior (registro devuelto por push_drive), verificando
    que cada fila siga teniendo el mismo N° y proveedor antes de borrar."""
    if not registro or not registro.get("escrito") or not registro.get("detalle") or registro.get("tipo") == "remitos_proveedor":
        return {"deshecho": False, "motivo": "no hay una inyección en vivo registrada"}
    gc, _ = _gspread()
    es_remito = registro.get("tipo") == "remitos"
    sh_ = gc.open_by_key(registro["sheet_id"])
    ws = _hoja_por_clave(sh_, "remito" if es_remito else "factura")
    a, b = registro["desde_fila"], registro["hasta_fila"]
    vals = ws.get(f"A{a}:I{b}")
    for esperado, real in zip(registro["detalle"], vals):
        if es_remito:
            if norm(real[1] if len(real) > 1 else "") != norm(esperado["PRODUCTO"]):
                return {"deshecho": False, "motivo": f"la fila {esperado['fila']} ya no coincide; no se borra"}
        elif str(real[0]) != str(esperado["N"]) or norm(real[3] if len(real) > 3 else "") != norm(esperado["proveedor"]):
            return {"deshecho": False, "motivo": f"la fila {esperado['fila']} ya no coincide (alguien la modificó); no se borra"}
    modo = "filas eliminadas"
    try:
        ws.delete_rows(a, b)
    except Exception as e:                                   # noqa: BLE001
        if "protected" not in str(e):
            raise
        # hoja protegida: no se pueden borrar filas; se vacían A:I (columnas de carga) y se quita el color
        rng = f"A{a}:E{b}" if es_remito else f"A{a}:I{b}"
        ws.batch_clear([rng])
        ws.format(rng, {"backgroundColor": {"red": 1, "green": 1, "blue": 1}})
        modo = "celdas vaciadas (hoja protegida: no admite borrar filas)"
    return {"deshecho": True, "filas": b - a + 1, "desde_fila": a, "hasta_fila": b, "planilla": registro.get("planilla"),
            "modo": modo}


# ---- remitos sin precio: valorización con lista del proveedor / último costo ----------------
def _lista_verduleria() -> list[dict]:
    p = BASE_DIR / "Verduleria" / "precios_verduleria_limpios.xlsx"
    if not p.is_file():
        return []
    d = pd.read_excel(p)
    d.columns = [str(c) for c in d.columns]
    return [{"articulo": norm(r["Articulo"]), "bulto": _num(r.get("Precio_Bulto")) or 0.0,
             "kilo": _num(r.get("Precio_Kilo")) or 0.0} for _, r in d.iterrows()]


_BULTO_U = {"CAJON", "CAJA", "BOLSA", "BULTO", "BAL", "BALDE"}


def completar_remito(reg: dict, loc: "Local", ap: dict) -> dict:
    """Remito sin precios (solo cantidades/bultos/kilos de birome): calcula precio y total de cada renglón cruzando
    contra (1) lista de verdulería en Verduleria/, (2) benchmark Grupo Free, (3) último costo vigente del local, aplicando
    BULTO_KG_DEFAULT para cajón/bolsa y los kilos reales de balanza anotados. Marca el origen de cada precio."""
    lista = _lista_verduleria()
    try:
        sys.path.insert(0, str(BASE_DIR / "Precios"))
        import benchmark_grupo_free as bg
        bench = bg.cargar()
    except Exception:
        bg, bench = None, None
    total = 0.0
    for r in reg["renglones"]:
        m = mapear_producto(r["descripcion"], None, loc, ap)
        pres = m["presentacion"] if m["confianza"] >= 0.7 else None
        r["presentacion"] = pres
        if r["anulado"]:
            r["precio_final"], r["precio_origen"], r["total_estimado"] = r["precio_unitario"] or 0.0, "anulado", 0.0
            continue
        if r["precio_unitario"]:                                    # el remito trae precio: se respeta
            base = r.get("peso_real_recibido_kg") if (r.get("unidad") or "").upper() == "KG" and r.get("peso_real_recibido_kg") else r["cantidad_facturada_remito"]
            r["precio_final"], r["precio_origen"] = r["precio_unitario"], "remito"
            r["total_estimado"] = round(base * r["precio_unitario"], 2)
            total += r["total_estimado"]
            continue
        nombre = norm(r["descripcion"])
        primera = nombre.split(" ")[0]
        por_kg = por_bulto = None
        origen = None
        mejor = max(lista, key=lambda x: fuzz.token_set_ratio(primera, x["articulo"]), default=None)
        if mejor and fuzz.token_set_ratio(primera, mejor["articulo"]) >= 90:
            por_kg, por_bulto, origen = mejor["kilo"] or None, mejor["bulto"] or None, "lista Verduleria/"
        if por_kg is None and por_bulto is None and bg and bench:
            ref = bg.precio_referencia(r["descripcion"], bench)
            if ref:
                por_kg, origen = ref["precio"], "benchmark Grupo Free"
        if por_kg is None and por_bulto is None and pres:
            vig, _ = loc.costo_vigente(pres)
            cat = loc.catalogo[loc.catalogo["PRESENTACION"] == pres].iloc[0]
            if vig:
                por_kg, origen = vig / ((_num(cat["UN_PRES"]) or 1) * (_num(cat["CANT_PRES"]) or 1)), "último costo vigente"
        kg_bulto = bulto_kg(norm(pres or primera))
        cant = r["cantidad_facturada_remito"]
        u = (r.get("unidad") or "").upper()
        peso = r.get("peso_real_recibido_kg")
        t = None
        if peso and por_kg:
            t = peso * por_kg
        elif u == "KG" and por_kg:
            t = cant * por_kg
        elif u in _BULTO_U and por_bulto:
            t = cant * por_bulto
        elif u in _BULTO_U and por_kg and kg_bulto:
            t = cant * kg_bulto * por_kg
        elif por_bulto:
            t = cant * por_bulto
        elif por_kg and u in ("UN", "UNIDAD", "PAQ", "ATADO") and peso is None:
            t = cant * por_kg                                        # sin peso: se valoriza la unidad como 1 kg (aprox.)
            origen += " (aprox.: 1 u = 1 kg)"
        if t is None:
            r["precio_final"], r["precio_origen"], r["total_estimado"] = None, "sin precio: REVISAR", None
            continue
        r["total_estimado"] = round(t, 2)
        r["precio_final"] = round(t / cant, 2) if cant else None
        r["precio_origen"] = origen
        total += r["total_estimado"]
    reg["total_estimado"] = round(total, 2)
    reg["sin_precio_en_remito"] = all(not r["precio_unitario"] for r in reg["renglones"])
    return reg


def remito_ya_cargado_como_total(reg: dict, loc: "Local", dias_antes: int = 2, dias_despues: int = 7,
                                 tolerancia: float = 0.15) -> dict | None:
    """¿Ese remito ya está en Facturas como una fila total del mismo proveedor (p. ej. 'Verduleria')? Compara fecha y
    monto estimado. Si sí, se concilia sin volver a cargar el gasto."""
    if not reg.get("total_estimado"):
        return None
    f = loc.facturas
    prov = norm(reg.get("proveedor"))
    fecha = pd.Timestamp(reg["fecha"])
    sel = f[(f["PROVEEDOR"].map(lambda p: fuzz.token_set_ratio(prov, norm(p)) >= 70))
            & (f["F_FACTURA"] >= fecha - pd.Timedelta(days=dias_antes))
            & (f["F_FACTURA"] <= fecha + pd.Timedelta(days=dias_despues))]
    for n, g in sel.groupby("N"):
        tot = float((g["CANT"] * g["PRECIO"] * (1 - g["DESC"])).sum())
        if tot and abs(tot - reg["total_estimado"]) / reg["total_estimado"] <= tolerancia:
            return {"N": int(n) if pd.notna(n) else n, "total": round(tot, 2),
                    "fecha": str(g["F_FACTURA"].iloc[0].date()), "diferencia_pct": round((tot / reg["total_estimado"] - 1) * 100, 1)}
    return None


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def _resumen_cascada(resultados: list[dict], locales: list[str]) -> None:
    for res, loc in zip(resultados, locales):
        c = res["comprobante"]
        print(f"\n▸ {c.get('archivo')} → {loc} | N° {c.get('numero_factura')} | {c.get('fecha_factura')} | "
              f"{c.get('proveedor_detectado')} → {res['proveedor_oficial'] or 'NO OFICIAL'} | "
              f"total {c.get('total_comprobante')} | cierre: {'OK' if res['validacion'].get('ok') else 'NO'}")
        for f in res["filas"]:
            marca = {"INYECTAR": "✔", "REVISAR": "✖", "DUPLICADA": "="}[f["accion"]]
            print(f"   {marca} {f['descripcion_original'][:34]:<34} → {str(f['presentacion'])[:30]:<30} "
                  f"conf {f['confianza']:.2f} {f['metodo'][:14]:<14} G={f['G']} H={f['H']} P_UN={f['P_UN']} "
                  f"{('Δ' + str(f['var_pct']) + '%') if f['var_pct'] is not None else ''} "
                  f"{'; '.join(f['motivos'])[:70]}")


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")
    ap_ = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap_.add_argument("entrada", nargs="?", help="comprobante .json/.pdf/.jpg/... o carpeta "
                                                "(default: 'Facturas Mishiguene/')")
    ap_.add_argument("--local", help="Local destino (si el comprobante no lo dice)")
    ap_.add_argument("--apply", action="store_true", help="Escribir en Copias_Gestion/ (sin esto es dry-run)")
    ap_.add_argument("--demo-historico", type=int, metavar="N", help="Prueba: reconstruye N facturas reales del local")
    ap_.add_argument("--bootstrap-aprendizaje", action="store_true")
    ap_.add_argument("--copias-dir", help="Carpeta de copias (default Copias_Gestion/); útil para pruebas")
    ap_.add_argument("--aprender", action="store_true", help="Con --apply, guarda en el aprendizaje los mapeos INYECTADOS")
    args = ap_.parse_args(argv)
    global COPIAS_DIR
    if args.copias_dir:
        COPIAS_DIR = Path(args.copias_dir)

    if args.bootstrap_aprendizaje:
        bootstrap_aprendizaje()
        if not (args.entrada or args.demo_historico):
            return 0
    ap = cargar_aprendizaje()
    if not ap["mapeos"]:
        LOG.info("aprendizaje_proveedores.json vacío: se siembra desde el Maestro y las Facturas")
        ap = bootstrap_aprendizaje()

    comprobantes: list[tuple[dict, str]] = []
    verdades: dict[str, list[dict]] = {}
    if args.demo_historico:
        import invoice_parser as ip
        nombre = (SINONIMOS_LOCAL.get((args.local or "Arcos").lower()) or args.local)
        loc = Local(nombre)
        tmp = Path(tempfile.mkdtemp(prefix="demo_comprobantes_"))      # PDFs de prueba: fuera del proyecto
        for r in reconstruir_comprobantes(loc, args.demo_historico, tmp):
            comp = ip.procesar_comprobante(r["pdf"], local=nombre, usar_cache=False)
            comprobantes.append((comp, nombre))
            verdades[r["pdf"].name] = r["verdad"]
    elif args.entrada:
        import invoice_parser as ip
        p = Path(args.entrada)
        archivos = sorted(x for x in p.iterdir() if x.suffix.lower() in ip.IMG_EXT | {".pdf", ".json"}) \
            if p.is_dir() else [p]
        for a in archivos:
            if a.suffix.lower() == ".json":
                datos = json.loads(a.read_text(encoding="utf-8"))
                for c in (datos if isinstance(datos, list) else [datos]):
                    comprobantes.append((c, c.get("local_destino") or args.local))
            else:
                c = ip.procesar_comprobante(a, local=args.local)
                comprobantes.append((c, c.get("local_destino") or args.local))
    else:
        carpeta = CARPETA_ENTRADA if CARPETA_ENTRADA.is_dir() else CARPETA_ENTRADA_VIEJA
        if not carpeta.is_dir():
            ap_.print_help()
            return 1
        return main([str(carpeta)] + [a for a in (argv if argv is not None else sys.argv[1:])])

    cache_locales: dict[str, Local] = {}
    resultados, locales = [], []
    for comp, destino in comprobantes:
        nombre = SINONIMOS_LOCAL.get(str(destino or "").lower())
        if nombre not in PLANILLAS:
            LOG.error("%s: local destino '%s' sin planilla de Facturas; indicá --local", comp.get("archivo"), destino)
            continue
        loc = cache_locales.setdefault(nombre, Local(nombre))
        resultados.append(imputar_comprobante(comp, loc, ap, omitir_duplicados=bool(args.demo_historico)))
        locales.append(nombre)

    _resumen_cascada(resultados, locales)
    todas = [f for r in resultados for f in r["filas"]]
    cuenta = {a: sum(f["accion"] == a for f in todas) for a in ("INYECTAR", "REVISAR", "DUPLICADA")}
    print(f"\nRESUMEN: {len(resultados)} comprobantes, {len(todas)} renglones → "
          f"{cuenta['INYECTAR']} inyectables, {cuenta['REVISAR']} a REVISAR, {cuenta['DUPLICADA']} duplicados")
    if verdades:
        aciertos = tot = inyect_ok = inyect = 0
        for res in resultados:
            vs = verdades[res["comprobante"]["archivo"]]
            for f, v in zip(res["filas"], vs):
                tot += 1
                bien = f["presentacion"] == v["presentacion"]
                aciertos += bien
                if f["accion"] == "INYECTAR":
                    inyect += 1
                    inyect_ok += bien
        print(f"PRECISIÓN vs. carga real: mapeo correcto {aciertos}/{tot} ({100 * aciertos / max(tot, 1):.0f} %); "
              f"de lo INYECTADO, correcto {inyect_ok}/{inyect}")
    solicitudes = registrar_solicitudes(resultados, locales, apply=args.apply)
    print(f"SOLICITUDES_PRODUCTOS a generar: {len(solicitudes)}")
    if args.apply:
        for nombre in sorted(set(locales)):
            sel = [r for r, l in zip(resultados, locales) if l == nombre]
            rep = aplicar_a_copia(sel, cache_locales[nombre])
            print(f"APLICADO en {cache_locales[nombre].copia.name}: {rep['escritas']} filas escritas, "
                  f"{rep['verificadas']} verificadas por lectura")
        for p in escribir_log(resultados, locales):
            print("log:", p.name)
        if args.aprender:
            for res, loc in zip(resultados, locales):
                for f in res["filas"]:
                    if f["accion"] == "INYECTAR" and res["proveedor_oficial"]:
                        aprender(ap, res["proveedor_oficial"], f["descripcion_original"], f["presentacion"], loc)
            guardar_aprendizaje(ap)
    else:
        print("(dry-run: no se escribió ningún archivo; usar --apply para escribir en las copias)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
