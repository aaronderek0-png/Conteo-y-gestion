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
COPIAS_DIR = BASE_DIR / "Copias_Gestion"
APRENDIZAJE_JSON = BASE_DIR / "aprendizaje_proveedores.json"
MAESTRO_XLSX = DATA_DIR / "maestro.xlsx"

# Planilla de Facturas de cada local (export local de la planilla de producción)
PLANILLAS = {"Arcos": "arcos.xlsx", "Cabello": "cabello.xlsx", "Lafinur": "lafinur.xlsx",
             "Congreso": "congreso_facturas.xlsx"}        # 'Obrador' == 'Congreso'; Chimidoha no tiene planilla propia
SINONIMOS_LOCAL = {"obrador": "Congreso", "congreso": "Congreso", "arcos": "Arcos", "cabello": "Cabello",
                   "lafinur": "Lafinur", "chimidoha": "Chimidoha", "chimi": "Chimidoha"}

CONF_MINIMA = 0.90
TOLERANCIA_PRECIO = 0.25
REVISAR_TXT = "REVISAR / NO CATALOGADO"
# Verdulería: solo estos van por bulto/cajón; el resto, por kilo (se compara por nombre de presentación)
VERDULERIA_BULTO = {"limon", "pepino", "berenjena", "tomate perita", "cebolla"}
BULTO_UNIDADES = {"BULTO", "CAJA", "CAJON", "CAJÓN", "BAL", "BALDE"}
# kg que trae un bulto cuando el comprobante no lo dice ('cajón 18 kg'). Vacío a propósito: sin dato -> REVISAR.
BULTO_KG_DEFAULT: dict[str, float] = {}

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
        fuente = self.copia if self.copia.is_file() else self.original   # si ya hay copia, se imputa sobre ella
        LOG.info("%s: leyendo %s", nombre, fuente.name)
        xl = pd.ExcelFile(fuente)
        prod = pd.read_excel(xl, "Productos", header=0).iloc[:, :8]
        prod.columns = ["SKU", "PRODUCTO_BASE", "PRESENTACION", "CAT1", "CAT2", "UN_PRES", "UN_MEDIDA", "CANT_PRES"]
        prod = prod[prod["PRESENTACION"].notna()].copy()
        prod["_norm"] = prod["PRESENTACION"].map(norm)
        self.catalogo = prod.drop_duplicates("_norm").reset_index(drop=True)
        listas = pd.read_excel(xl, "Listas", header=None).iloc[3:, 1].dropna().astype(str).str.strip()
        self.proveedores = sorted(set(listas[listas != ""]))
        fac = pd.read_excel(xl, "Facturas", header=0, usecols="A:I")
        fac.columns = ["N", "F_FACTURA", "F_INGRESO", "PROVEEDOR", "PRODUCTO", "NOTA", "CANT", "PRECIO", "DESC"]
        fac = fac[fac["N"].notna() & fac["PRODUCTO"].notna()].copy()
        fac["F_FACTURA"] = pd.to_datetime(fac["F_FACTURA"], errors="coerce")
        fac["PRECIO"] = pd.to_numeric(fac["PRECIO"], errors="coerce")
        fac["DESC"] = pd.to_numeric(fac["DESC"], errors="coerce").fillna(0.0)
        fac["P_NETO"] = fac["PRECIO"] * (1 - fac["DESC"])
        self.facturas = fac
        self.ultima_fila = None            # se completa al escribir
        self._idx_prov: dict[str, set[str]] = {}
        for p, g in fac.groupby(fac["PROVEEDOR"].map(norm)):
            self._idx_prov[p] = set(g["PRODUCTO"].map(norm))

    def costo_vigente(self, presentacion: str, antes_de: date | None = None) -> tuple[float | None, str | None]:
        """Último precio neto cargado para la presentación (anterior a `antes_de` si se indica)."""
        f = self.facturas[self.facturas["PRODUCTO"].map(norm) == norm(presentacion)]
        f = f[f["P_NETO"].notna() & (f["P_NETO"] > 0)]
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
    APRENDIZAJE_JSON.write_text(json.dumps(ap, ensure_ascii=False, indent=1, sort_keys=True), encoding="utf-8")


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
    for nombre in PLANILLAS:
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


def mapear_producto(desc: str, proveedor: str | None, loc: Local, ap: dict) -> dict:
    """Devuelve {'presentacion', 'confianza', 'metodo', 'alternativas'}; presentacion None si no hay candidato."""
    k = norm(desc)
    if proveedor:
        e = ap["mapeos"].get(f"{norm(proveedor)}|{k}")
        if e and e["presentacion"] in set(loc.catalogo["PRESENTACION"]):
            return {"presentacion": e["presentacion"], "confianza": 0.99, "metodo": f"aprendizaje (n={e['n']})",
                    "alternativas": []}
    exacto = loc.catalogo[loc.catalogo["_norm"] == k]
    if len(exacto):
        return {"presentacion": exacto.iloc[0]["PRESENTACION"], "confianza": 1.0, "metodo": "nombre idéntico",
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
            kg = float(m.group(1).replace(",", ".")) if m else BULTO_KG_DEFAULT.get(pres_norm)
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
    m = mapear_producto(desc, proveedor, loc, ap)
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
            for otro in PLANILLAS:
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
    """'0004-00012345' -> 12345 (la planilla guarda el N° como número)."""
    if not numero:
        return None
    ultimo = re.split(r"[-–]", str(numero))[-1]
    return int(ultimo) if ultimo.isdigit() else numero


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


def aplicar_a_copia(resultados: list[dict], loc: Local) -> dict:
    """Escribe las filas en la copia local de Facturas. Devuelve {'escritas': n, 'verificadas': n, 'filas': [...]}."""
    import openpyxl
    from openpyxl.styles import PatternFill
    if not loc.copia.is_file():
        COPIAS_DIR.mkdir(exist_ok=True)
        shutil.copy2(loc.original, loc.copia)
        LOG.info("copia creada: %s", loc.copia.name)
    assert loc.copia.resolve().parent == COPIAS_DIR.resolve() and loc.copia.name.endswith("_FACTURAS_COPIA.xlsx")
    wb = openpyxl.load_workbook(loc.copia)
    ws = wb["Facturas"]
    ultima = max((r for r in range(2, ws.max_row + 1) if ws.cell(r, 1).value not in (None, "")), default=1)
    rojo = PatternFill("solid", fgColor="FFC7CE")
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
                if reviso:
                    ws.cell(r, c).fill = rojo
            escritas.append((r, valores, f))
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
    ap_.add_argument("entrada", nargs="?", help="comprobante .json/.pdf/.jpg/... o carpeta")
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
        ap_.print_help()
        return 1

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
