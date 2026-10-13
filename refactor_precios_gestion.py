"""Refactor de inyección de precios en las Gestiones Mensuales (Arcos, Congreso, Cabello, Lafinur).

Trabaja SOLO sobre copias (Copias_Gestion/<Local>_COPIA.xlsx); los originales exportados quedan intactos.
Excluye Chimidoha (columna 'Precio Chimi' y hoja 'Precios Octubre LAFINURCHIMI' del mapeo).

Reglas:
  1. Ítems sin categoría, de rubro Congreso o con código OBR* -> 'Obrador'. 'REVISAR' se informa pero no se pisa.
  2. Unidad del mapeo ('GR.' / 'UN.' / 'PORC.'): gr -> precio x 1000 (se valoriza por KG/LT); un/porc -> 1:1.
  3. Precio nuevo por SKU desde la hoja de Precios directo a la columna de precio actual (G; en Cabello es F),
     sin arrastrar el precio viejo. Solo si fecha_precio >= fecha factura existente; esa pasa a ser la nueva fecha.
     Guardas: variación > +-50% contra el precio viejo/actual, unidad ambigua o precio vacío -> NO se inyecta
     (queda en el log como REVISAR).

Uso:
    python refactor_precios_gestion.py --dry-run      # imprime Dátiles y Bagel Amapola + resumen, no escribe
    python refactor_precios_gestion.py --apply        # escribe las copias *_COPIA.xlsx y los logs CSV

El parcheo es sobre el XML de la hoja (no re-serializa el libro): comentarios, hilos de comentarios, formatos
y fórmulas quedan byte a byte como en el original. Se fuerza recálculo al abrir (fullCalcOnLoad).
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import difflib
import re
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.utils import get_column_letter

from defensivo import (SchemaError, codigos_colisionados, finito, norm_header, parse_fecha_robusta,
                       parse_numero, plausibilidad_numerica, resolver_esquema, safe_div, tokens_nombre)
from motor_recetas import MotorRecetas, es_elaboracion

BASE = Path(__file__).resolve().parent
SRC_DIR = BASE / "Copias_Gestion" / "originales_export"
OUT_DIR = BASE / "Copias_Gestion"
PRECIOS_XLSX = SRC_DIR / "Precios_mapeo.xlsx"
RECETAS_XLSX = SRC_DIR / "Recetas.xlsx"
PRODUCTOS_XLSX = SRC_DIR / "MaestroProductos.xlsx"   # Drive: MAESTRO PRODUCTOS
PRODUCTOS: dict = {}   # SKU -> {precio (por UM inventario), um, unidad, fecha}
EQUIV: dict = {}       # SKU -> (unidad, FACTOR_CONVERSION) de Recetas!Equivalencias (rendimiento)
FECHA_MODO = "maestro"  # maestro: fecha de factura del Maestro de Productos (fallback ejecución) | ejecucion
TRANSICIONES: dict = {}
MOTOR: MotorRecetas | None = None   # se carga en main()
PRECIOS_SHEET = "Precios OCTUBRE AMBOS"   # A=cod, B=nombre, C=precio vivo | I=nombre mapeo, J=SKU (D:F = Chimi, ignorado)

# col_* son índices 0-based. target = columna del precio vigente ("Precio Nuevo").
LOCALES = {
    "Arcos":    dict(src="Arcos_septiembre.xlsx",    sheet="Inventario", first=2, rubro=0, medida=2, code=3, prod=4,
                     ref=5, target=6, fecha=18, fecha_fmt="Factura: {d}", obrador="OBRADOR"),
    "Congreso": dict(src="Congreso_septiembre.xlsx", sheet="Inventario", first=3, rubro=0, medida=2, code=3, prod=4,
                     ref=5, target=6, fecha=7, fecha_fmt="Factura: {d}", obrador="OBRADOR"),
    "Cabello":  dict(src="Cabello_septiembre.xlsx",  sheet="Gestion",    first=2, rubro=0, medida=1, code=2, prod=3,
                     ref=4, target=5, fecha=16, fecha_fmt="Factura: {d}", obrador="Obrador"),
    "Lafinur":  dict(src="Lafinur_septiembre.xlsx",  sheet="Gestion",    first=2, rubro=0, medida=2, code=3, prod=4,
                     ref=5, target=6, fecha=10, fecha_fmt="{d}", obrador="OBRADOR"),
}
# Homologación mundo Maxirest (códigos nativos del POS / transferencias del Obrador) -> SKU canónico del Maestro
# de Precios. Hardcodeado temporalmente: completar con las equivalencias que salgan en faltantes_mapeo.csv.
MAPEO_MAXIREST = {
    "OBR127": "BAMAPOLA",
    "BKUM": "BKUMMEL",
    "H000": "H30",     # Harina 000
    "QMO": "QMU",      # Queso muzzarella
}
# Esquema por nombre de encabezado (igualdad exacta normalizada, por orden de prioridad). Los índices de LOCALES
# quedan SOLO como referencia histórica: si el encabezado aparece en otra columna se avisa y se usa el encabezado.
SPEC_COLUMNAS = {
    "rubro": ["RUBRO", "CATEGORIA"],
    "medida": ["MEDIDA", "UN."],
    "code": ["CODIGO", "CODIGO (SKU)", "COD"],
    "prod": ["PRODUCTO", "DESCRIPCION"],
    "ref": ["$ UNITARIO ANTERIOR", "$ UNITARIO"],
    "target": ["$ UNITARIO ACTUAL", "$ UNITARIO ULTIMO PRECIO"],
    "fecha": ["FECHA ULTIMA FACTURA", "FECHA"],
}
AVISOS: list[str] = []   # avisos de schema / tipado / fechas acumulados en la corrida
MEDIDAS_UNITARIAS = {"un", "u", "unidad", "unidades", "porc", "porcion", "porción", "pack", "caja"}
UMBRAL_VARIACION = 0.5  # +-50 %
BASELINE_SIN_MAPEO = {"Arcos": 230, "Congreso": 189, "Cabello": 185, "Lafinur": 362}  # motor anterior (solo SKU directo)
PRIORIDAD_RUBROS = ("obrador", "panific", "produccion", "pasteler", "harinera")


# --------------------------------------------------------------------------- utilidades
def num(v):
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def unidad_de(nombre: str) -> str | None:
    m = re.match(r"\s*(gr|un|porc)\b\.?", str(nombre).lower())
    return m.group(1) if m else None


def parse_fecha(v, year_hint: int) -> dt.date | None:
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if not isinstance(v, str):
        return None
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", v)
    if m:
        d, mo, y = map(int, m.groups())
    else:
        m = re.search(r"\((\d{1,2})/(\d{1,2})\)", v)
        if not m:
            return None
        d, mo = map(int, m.groups())
        y = year_hint
    try:
        return dt.date(y, mo, d)
    except ValueError:
        return None


def norm_nombre(t) -> str:
    """Clave semántica: sin tildes, mayúsculas, sin prefijo de unidad (UN./GR./PORC.), sin sufijo 'X GR',
    sin puntuación y con espacios simples."""
    t = unicodedata.normalize("NFKD", str(t or ""))
    t = "".join(c for c in t if not unicodedata.combining(c)).upper()
    t = re.sub(r"^\s*(GR|UN|PORC)\.?", " ", t)
    t = re.sub(r"X\s*(GR|KG|UN|LT|L)\.?\s*$", " ", t)
    t = re.sub(r"[^A-Z0-9 ]+", " ", t)
    return " ".join(t.split())


def _bloques_precios(ws) -> tuple[int, int, int, int, int]:
    """Ubica por ETIQUETA (fila 1) los bloques 'PRECIOS vivo' (NOMBRE, PRECIO) y 'mapeo' (NOMBRE, SKU).
    Ignora 'Precio Chimi'. Si falta un ancla: SchemaError. -> (col_nombre, col_precio, col_map_nombre, col_sku, fila_datos)"""
    filas = list(ws.iter_rows(min_row=1, max_row=6, values_only=True))
    if not filas:
        raise SchemaError(f"{PRECIOS_SHEET}: hoja vacía")
    h = [norm_header(c) for c in filas[0]]
    try:
        c_vivo = h.index("PRECIOS VIVO")
        c_map = h.index("MAPEO")
    except ValueError as e:
        raise SchemaError(f"{PRECIOS_SHEET}: faltan las etiquetas 'PRECIOS vivo' / 'mapeo' en la fila 1 "
                          f"({[x for x in h if x][:8]})") from e
    for i, fila in enumerate(filas[1:], start=2):
        if norm_header(fila[c_vivo] if c_vivo < len(fila) else "") == "NOMBRE":
            return c_vivo, c_vivo + 1, c_map, c_map + 1, i + 1
    raise SchemaError(f"{PRECIOS_SHEET}: no encuentro la fila 'NOMBRE' bajo 'PRECIOS vivo'")


def cargar_precios() -> tuple[dict, list]:
    """-> (maestro, duplicados). maestro = {'sku': {SKU: item}, 'nombre': {clave_normalizada: item|None}}.
    Columnas ubicadas por etiqueta. SKU duplicado con precios en conflicto (>1%) -> sin precio + 'conflicto'."""
    ws = openpyxl.load_workbook(PRECIOS_XLSX, data_only=True)[PRECIOS_SHEET]
    c_n, c_p, c_mn, c_sku, r0 = _bloques_precios(ws)
    by_sku, by_name, dups = {}, {}, []

    def add_name(nombre, item):
        k = norm_nombre(nombre)
        if not k:
            return
        prev = by_name.get(k, False)
        if prev is False or (prev is not None and prev["precio"] is None and item["precio"] is not None):
            by_name[k] = item
        elif prev is not None and item["precio"] is not None and prev["precio"] is not None \
                and abs(prev["precio"] - item["precio"]) > 1e-9:
            by_name[k] = None  # mismo nombre, precios distintos: ambiguo, no se usa

    for r in ws.iter_rows(min_row=r0, values_only=True):
        g = lambda i: r[i] if i < len(r) else None
        nombre, sku, precio = g(c_mn), g(c_sku), parse_numero(g(c_p))[0]
        if precio is not None and precio < 0:
            precio = None   # un precio negativo es un error de carga
        if sku and nombre:
            sku = str(sku).strip().upper()
            item = dict(sku=sku, nombre=str(nombre).strip(), unidad=unidad_de(nombre), precio=precio)
            if sku in by_sku:
                ant = by_sku[sku]
                dups.append((sku, ant["nombre"], item["nombre"]))
                if ant["precio"] is not None and item["precio"] is not None and \
                        abs(safe_div(item["precio"], ant["precio"], 0) - 1) > 0.01:
                    ant["conflicto"] = [ant["nombre"], item["nombre"]]
                    ant["precio"] = None          # conflicto: nadie decide en silencio
                    add_name(nombre, ant)
                    continue
                if item["precio"] is None or ant.get("conflicto"):
                    add_name(nombre, ant)
                    continue
            by_sku[sku] = item
            add_name(nombre, item)
        elif g(1) and c_n == 1:  # fila del maestro sin SKU: solo sirve para cruce por nombre
            add_name(g(c_n), dict(sku=None, nombre=str(g(c_n)).strip(), unidad=unidad_de(g(c_n)), precio=precio))
    return {"sku": by_sku, "nombre": by_name}, dups


def cargar_productos() -> dict:
    """Maestro de Productos: precio por unidad de inventario (= P.U x UM Inventario) y fecha de factura por SKU."""
    ws = openpyxl.load_workbook(PRODUCTOS_XLSX, data_only=True, read_only=True)["Productos"]
    out = {}
    for r in ws.iter_rows(min_row=2, max_col=10, values_only=True):
        if r[0]:
            f = r[8]
            out[str(r[0]).strip().upper()] = dict(precio=num(r[4]), um=num(r[6]), unidad=r[5], pu=num(r[7]),
                                                  fecha=f.date() if isinstance(f, dt.datetime) else None)
    return out


def cargar_equivalencias() -> dict:
    ws = openpyxl.load_workbook(RECETAS_XLSX, data_only=True, read_only=True)["Equivalencias"]
    return {str(r[0]).strip().upper(): (r[4], num(r[5])) for r in ws.iter_rows(min_row=2, max_col=6, values_only=True)
            if r[0] and num(r[5])}


def fecha_de_item(item: dict, hoy: dt.date) -> tuple[dt.date, str]:
    """Fecha que se estampa al actualizar un precio: la de factura del Maestro de Productos si existe;
    si no (o en modo 'ejecucion'), la fecha de ejecución."""
    if FECHA_MODO == "maestro":
        pr = PRODUCTOS.get(str(item.get("sku") or "").upper())
        if pr and pr["fecha"]:
            return pr["fecha"], "maestro"
    return hoy, "ejecucion"


def _item_bom(res, codigo_local):
    """Adapta un Resultado del motor al formato de 'item' que consume valorizar()."""
    masa = res.unidad.lower() in ("gr", "g", "ml", "kg", "lt")
    return dict(sku=res.codigo, nombre=MOTOR.recetas[res.codigo]["nombre"], unidad="gr" if masa else "un",
                precio=res.costo_unitario, bom=res)


def _resolver(codigo: str, producto, maestro: dict, mapeo_maxirest: dict = MAPEO_MAXIREST, usar_bom: bool = True):
    """Cruce en cascada. -> (item|None, via, sugerencia).  via:
       BOM       elaboración (M*/MADRE/OBR*): explosión de la receta, nunca como insumo simple
       SKU       código del local tal cual en el Maestro de Precios
       MAXIREST  MAPEO_MAXIREST o puente COD_RECETA<->NOM_RECETA de Recetas, y nueva búsqueda
       NOMBRE    nombre normalizado (sin UN./GR./tildes) contra el Maestro
       INSUMO    precio P_UN de la hoja Insumos (solo no-elaboraciones)
       RECETA    código que existe como receta aunque no parezca elaboración (p. ej. PITA)
    Un candidato sin precio no corta la cascada. Sin cruce: sugerencia fuzzy informativa."""
    codigo = str(codigo).strip().upper()
    elab = es_elaboracion(codigo, producto)
    sin_precio = None

    def bom(cod_receta):
        r = MOTOR.explotar(cod_receta, MOTOR.recetas[MOTOR.resolver_receta(cod_receta)]["rend"] or 1.0)
        return r

    # Paso 0: elaboración -> receta (BOM)
    if MOTOR is not None and elab and usar_bom:
        cod = MOTOR.resolver_receta(codigo) or MOTOR.resolver_receta(mapeo_maxirest.get(codigo, ""))
        if cod:
            r = bom(cod)
            if r and r.completo and r.costo_unitario > 0:
                return _item_bom(r, codigo), "BOM", ""
            sin_precio = sin_precio or (dict(sku=cod, nombre=MOTOR.recetas[cod]["nombre"], unidad=None, precio=None,
                                             faltantes=r.faltantes if r else ["receta sin líneas"]), "BOM")
    # Paso 1: SKU del local tal cual
    it = maestro["sku"].get(codigo)
    if it and it["precio"]:
        return it, "SKU", ""
    sin_precio = sin_precio or ((it, "SKU") if it else None)
    # Paso 2: MAPEO_MAXIREST y puente de Recetas (COD_RECETA <-> NOM_RECETA)
    canons = []
    if codigo in mapeo_maxirest:
        canons.append(mapeo_maxirest[codigo])
    if MOTOR is not None:
        for c in (MOTOR.canonico(codigo), MOTOR.resolver_receta(codigo)):
            if c and c not in canons and c != codigo:
                canons.append(c)
    for canon in canons:
        it = maestro["sku"].get(str(canon).upper())
        if it and it["precio"]:
            return it, "MAXIREST", ""
        sin_precio = sin_precio or ((it, "MAXIREST") if it else None)
    # Paso 3: nombre normalizado
    k = norm_nombre(producto)
    it = maestro["nombre"].get(k)
    if it and it["precio"]:
        return it, "NOMBRE", ""
    sin_precio = sin_precio or ((it, "NOMBRE") if it else None)
    if MOTOR is not None:
        # Paso 4: insumo simple (no para elaboraciones)
        for c in [codigo] + canons:
            ins = MOTOR.insumos.get(str(c).upper())
            if ins and ins["precio"] and not elab:
                pr = PRODUCTOS.get(ins["sku"].upper())
                if pr and pr["precio"]:
                    # factor de presentación: Precio = P.U x UM Inventario (botella 520 ml, pote 650 g, etc.)
                    return dict(sku=ins["sku"], nombre=ins["nombre"], unidad="inv", precio=pr["precio"],
                                um=pr["um"]), "INSUMO", ""
                u = "gr" if ins["unidad"] in ("g", "gr", "ml") else "un"
                return dict(sku=ins["sku"], nombre=ins["nombre"], unidad=u, precio=ins["precio"]), "INSUMO", ""
        # Paso 5: es una receta aunque no lo parezca
        cod = MOTOR.resolver_receta(codigo)
        if cod and not elab:
            r = bom(cod)
            if r and r.completo and r.costo_unitario > 0:
                return _item_bom(r, codigo), "RECETA", ""
            sin_precio = sin_precio or (dict(sku=cod, nombre=MOTOR.recetas[cod]["nombre"], unidad=None, precio=None,
                                             faltantes=r.faltantes if r else []), "RECETA")
    if sin_precio and sin_precio[0]:
        return sin_precio[0], sin_precio[1], ""
    cand = difflib.get_close_matches(k, [n for n, v in maestro["nombre"].items() if v], n=1, cutoff=0.82)
    if cand:
        v = maestro["nombre"][cand[0]]
        return None, "", f"{v['nombre']} ({v['sku'] or 'sin SKU'})"
    return None, "", ""


def resolver(codigo: str, producto, maestro: dict, mapeo_maxirest: dict = MAPEO_MAXIREST):
    """Cascada completa. Si la explosión de receta (BOM) contradice en más de +-UMBRAL_VARIACION al precio
    vigente del Maestro de Precios, la receta se considera incompleta/mal escalada y gana el Maestro."""
    item, via, sug = _resolver(codigo, producto, maestro, mapeo_maxirest, True)
    if via == "BOM":
        alt, via2, _ = _resolver(codigo, producto, maestro, mapeo_maxirest, False)
        if alt and alt.get("precio") and via2 in ("SKU", "MAXIREST", "NOMBRE") and item["precio"]                 and abs(item["precio"] / alt["precio"] - 1) > UMBRAL_VARIACION:
            alt = dict(alt, bom=item["bom"], bom_descartado=True)
            return alt, via2, ""
    return item, via, sug


def valorizar(item, medida, ref_vals) -> tuple[float | None, str, str]:
    """-> (precio_final, unidad_usada, motivo_si_no)"""
    p = item["precio"]
    if not p:
        return None, item["unidad"] or "?", "sin precio en mapeo"
    unidad = item["unidad"]
    if unidad is None:  # sin prefijo: se infiere por escala contra el precio existente
        base = next((v for v in ref_vals if v), None)
        if base:
            ratio = base / p
            if 300 <= ratio <= 3000:
                unidad = "gr(inferida)"
            elif 0.3 <= ratio <= 3:
                unidad = "un(inferida)"
        if unidad is None:
            return None, "?", "unidad ambigua (sin prefijo y sin referencia)"
    es_gr = unidad.startswith("gr")
    if es_gr and str(medida or "").strip().lower() in MEDIDAS_UNITARIAS:
        return None, unidad, f"mapeo en gr pero el inventario cuenta por '{medida}'"
    return (p * 1000 if es_gr else p), unidad, ""


def esquema_local(ws, cfg: dict, local: str):
    """Columnas del local por nombre de encabezado. Valida plausibilidad de la columna de precio."""
    filas = list(ws.iter_rows(min_row=1, max_row=12, values_only=True))
    legacy = {k: cfg[k] for k in ("rubro", "medida", "code", "prod", "ref", "target", "fecha") if k in cfg}
    cols, hdr, avisos = resolver_esquema(filas, SPEC_COLUMNAS, obligatorios=("rubro", "code", "ref", "target", "fecha"),
                                         legacy=legacy, nombre=f"{local}/{cfg['sheet']}")
    if "prod" not in cols:
        cols["prod"] = cols["code"] + 1
        avisos.append(f"{local}: sin encabezado de producto; uso la columna siguiente al código ({cols['prod']})")
    muestra = [r[cols["target"]].value for r in ws.iter_rows(min_row=hdr + 1, max_row=hdr + 300)
               if cols["target"] < len(r)]
    if plausibilidad_numerica(muestra) < 0.6:
        raise SchemaError(f"{local}: la columna '{cfg['sheet']}'[{cols['target']}] elegida como precio no parece "
                          f"numérica (¿se movió una columna?). Se aborta sin escribir.")
    return cols, hdr, avisos


def es_formula(v) -> bool:
    return (isinstance(v, str) and v.startswith("=")) or type(v).__name__ in ("ArrayFormula", "DataTableFormula")


def evaluar(local: str, cfg: dict, precios: dict, hoy: dt.date):
    """-> (filas, esquema). Todas las columnas se resuelven por encabezado; nada se escribe aquí."""
    ws = openpyxl.load_workbook(SRC_DIR / cfg["src"])[cfg["sheet"]]
    c, hdr, avisos = esquema_local(ws, cfg, local)
    AVISOS.extend(avisos)
    celdas = list(ws.iter_rows(min_row=hdr + 1))

    def val(r, k):
        return r[c[k]].value if k in c and c[k] < len(r) else None

    colis = codigos_colisionados([(str(val(r, "code")).strip().upper(), val(r, "prod")) for r in celdas
                                  if val(r, "code") and str(val(r, "code")).strip().upper() != "NO ENCONTRADO"])
    filas = []
    for r in celdas:
        code = val(r, "code")
        if not code or not str(code).strip():
            continue
        code = str(code).strip().upper()
        rub = val(r, "rubro")
        row = dict(local=local, fila=r[0].row, codigo=code, producto=val(r, "prod"), rubro=rub,
                   medida=val(r, "medida"), F=val(r, "ref"), G=val(r, "target"), fecha_actual=val(r, "fecha"),
                   nuevo=None, unidad="", estado="", detalle="", set_rubro=None, set_fecha=None, via="",
                   sku_canonico="", sugerencia="", costo_bom=None, hoja_c_un=None, receta="", precio_maestro=None,
                   codigo_contingencia="", factor_presentacion=None, factor_equiv=None, fecha_origen="")

        def nota(t):
            row["detalle"] = (row["detalle"] + " " + t).strip()

        # 1) categoría
        rub_txt = str(rub).strip() if rub is not None else ""
        low = rub_txt.lower()
        if rub_txt == "":
            row["set_rubro"] = cfg["obrador"]
            nota("sin categoría -> Obrador")
        elif low == "revisar":
            nota("rubro 'REVISAR' (no se pisa: revisar a mano)")
        elif "congreso" in low or (code.startswith("OBR") and low != "obrador"):
            row["set_rubro"] = cfg["obrador"]   # Congreso y Obrador son el mismo centro de producción
            nota(f"origen Obrador/Congreso (rubro '{rub_txt}') -> Obrador")
        # 2) tipado de las referencias de precio (acepta '1066,75', '$ 1.250', '15 kilos' con aviso)
        g_val, g_av = parse_numero(row["G"])
        f_val, f_av = parse_numero(row["F"])
        for cel, av in (("G", g_av), ("F", f_av)):
            if av.startswith(("número guardado como texto", "texto con unidad")):
                nota(f"[tipado] {cel}: {av}")
                AVISOS.append(f"{local} fila {row['fila']} {code}: {cel} {av}")
        refs = [v for v in (g_val, f_val) if v]
        # 3) resolución del precio
        if code.startswith("NO ENCONTRADO") or code in ("", "S/C", "SIN CODIGO"):
            tmp = MotorRecetas.contingencia(row["producto"])
            row["codigo_contingencia"] = tmp
            cod_r, sim, alts = MOTOR.buscar_por_nombre(str(row["producto"] or ""))
            item = via = None
            sugerencia = ""
            if cod_r:
                r_ = MOTOR.explotar(cod_r, MOTOR.recetas[cod_r]["rend"] or 1.0)
                if r_ and r_.completo and r_.costo_unitario > 0:
                    item, via = _item_bom(r_, code), "NOMBRE_RECETA~"
                    nota(f"contingencia {tmp}: aproximación por nombre -> {cod_r} "
                         f"'{MOTOR.recetas[cod_r]['nombre']}' (similitud {sim:.2f})")
            if item is None:
                row["estado"] = "ALERTA_MAXIMA"
                row["sugerencia"] = "; ".join(f"{x}({y})" for x, y in alts[:3])
                nota(f"[ALERTA MÁXIMA] {tmp} | Producto: {row['producto']} sin costo (similitud {sim:.2f} < 0.85 o "
                     f"tokens no cubiertos). Mejores candidatos: {row['sugerencia'] or 'ninguno'} -> Proveer SKU canónico.")
                filas.append(row)
                continue
        else:
            item, via, sugerencia = resolver(code, row["producto"], precios)
        row["via"] = via
        row["sku_canonico"] = item["sku"] if item else ""
        row["sugerencia"] = sugerencia
        if item is None:
            row["estado"] = "FALTA_MAPEO"
            nota(f"[FALTA MAPEO] Código Maxirest: {code} | Producto: {row['producto']} -> Proveer SKU canónico."
                 + (f" Sugerencia: {sugerencia}" if sugerencia else ""))
            filas.append(row)
            continue
        if item.get("conflicto"):
            row["estado"] = "REVISAR"
            nota(f"SKU {item['sku']} duplicado en el Maestro con precios en conflicto: {item['conflicto']}")
            filas.append(row)
            continue
        # 3b) un código que agrupa productos distintos solo acepta un cruce cuyo nombre coincida EXACTAMENTE
        if code in colis and via not in ("NOMBRE", "NOMBRE_RECETA~"):
            if tokens_nombre(row["producto"]) != tokens_nombre(item.get("nombre")):
                row["estado"] = "REVISAR"
                nota(f"código colisionado: {code} agrupa {sorted(colis[code])[:4]}; el nombre "
                     f"'{row['producto']}' no coincide con '{item.get('nombre')}' del maestro")
                filas.append(row)
                continue
        nuevo, unidad, motivo = valorizar(item, row["medida"], refs)
        row["unidad"] = unidad
        row["factor_presentacion"] = item.get("um")
        row["factor_equiv"] = EQUIV.get(str(item.get("sku") or "").upper(), (None, None))[1]
        bomr = item.get("bom")
        if item.get("bom_descartado"):
            nota("BOM descartado: difiere >50% del Maestro (receta incompleta?)")
        if bomr is not None:
            row["costo_bom"] = round(bomr.precio_gestion, 4)
            row["hoja_c_un"] = MOTOR.recetas[bomr.codigo]["c_un_hoja"]
            row["receta"] = bomr.codigo
            mx = precios["sku"].get(MOTOR.recetas[bomr.codigo]["nom"])
            row["precio_maestro"] = mx["precio"] if mx else None
        if item.get("faltantes"):
            row["estado"] = "REVISAR"
            nota(f"receta {item['sku']} incompleta o con ciclo, falta: " + ", ".join(item["faltantes"][:6]))
            filas.append(row)
            continue
        if nuevo is None or finito(nuevo) is None or nuevo <= 0:
            row["estado"] = "REVISAR"
            nota(motivo or "precio no válido (<=0 o no finito)")
            filas.append(row)
            continue
        row["nuevo"] = round(nuevo, 6)
        if es_formula(row["G"]):
            row["estado"] = "REVISAR"
            nota("celda de precio es fórmula (no se pisa)")
            filas.append(row)
            continue
        if any(abs(safe_div(nuevo, v, 1e9) - 1) > UMBRAL_VARIACION for v in refs):
            row["estado"] = "REVISAR"
            nota(f"variación fuera de +-{UMBRAL_VARIACION:.0%} vs {', '.join(f'{v:,.2f}' for v in refs)}")
            filas.append(row)
            continue
        if g_val and abs(safe_div(nuevo, g_val, 1e9) - 1) < 0.005:
            row["estado"] = "SIN_CAMBIO"   # mismo precio: se conserva G y la fecha de factura real
            filas.append(row)
            continue
        fecha_nueva, origen_f = fecha_de_item(item, hoy)
        f_exist, f_av = parse_fecha_robusta(row["fecha_actual"], hoy)
        if f_av:
            nota(f"[fecha] {f_av}")
            AVISOS.append(f"{local} fila {row['fila']} {code}: {f_av}")
        if f_exist and fecha_nueva < f_exist:
            row["estado"] = "OMITIDO_FECHA"
            nota(f"factura existente {f_exist:%d/%m/%Y} más nueva que la del maestro {fecha_nueva:%d/%m/%Y}")
            filas.append(row)
            continue
        row["estado"] = "INYECTAR"
        row["fecha_origen"] = origen_f
        row["set_fecha"] = cfg["fecha_fmt"].format(d=f"{fecha_nueva:%d/%m/%Y}")
        filas.append(row)
    return filas, c


# --------------------------------------------------------------------------- parcheo XML
def _sheet_path(zf: zipfile.ZipFile, sheet_name: str) -> str:
    wb = zf.read("xl/workbook.xml").decode("utf8")
    name = escape(sheet_name, {'"': "&quot;"})
    rid = re.search(r'<sheet\b[^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(name), wb).group(1)
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf8")
    tgt = re.search(r'<Relationship\b[^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels) or \
        re.search(r'<Relationship\b[^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    t = tgt.group(1).lstrip("/")
    return t if t.startswith("xl/") else "xl/" + t


def _col_idx(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def _set_cell(xml: str, ref: str, kind: str, value) -> str:
    """kind: 'n' numérico | 's' texto inline. Conserva el estilo (s=) de la celda."""
    col, row_n = re.match(r"([A-Z]+)(\d+)", ref).groups()
    body = f"<v>{value!r}</v>" if kind == "n" else f'<is><t xml:space="preserve">{escape(str(value))}</t></is>'
    tattr = "" if kind == "n" else ' t="inlineStr"'
    m = re.search(r'<c r="%s"(?P<a>[^>]*?)(?:/>|>(?P<b>.*?)</c>)' % ref, xml, re.S)
    if m:
        if m.group("b") and "<f" in m.group("b"):
            raise ValueError(f"{ref} contiene una fórmula: no se sobrescribe")
        attrs = re.sub(r'\s+t="[^"]*"', "", m.group("a"))
        return xml[:m.start()] + f'<c r="{ref}"{attrs}{tattr}>{body}</c>' + xml[m.end():]
    rm = re.search(r'<row r="%s"(?P<a>[^>]*?)(?:/>|>(?P<b>.*?)</row>)' % row_n, xml, re.S)
    if not rm:
        raise ValueError(f"fila {row_n} inexistente en XML")
    new = f'<c r="{ref}"{tattr}>{body}</c>'
    if rm.group("b") is None:
        return xml[:rm.start()] + f'<row r="{row_n}"{rm.group("a")}>{new}</row>' + xml[rm.end():]
    inner_start, inner = rm.start("b"), rm.group("b")
    pos = rm.end("b")
    for cm in re.finditer(r'<c r="([A-Z]+)\d+"', inner):
        if _col_idx(cm.group(1)) > _col_idx(col):
            pos = inner_start + cm.start()
            break
    return xml[:pos] + new + xml[pos:]


def parchear(src: Path, dst: Path, sheet: str, cambios: list[tuple[str, str, object]]):
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        sp = _sheet_path(zin, sheet)
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == sp:
                xml = data.decode("utf8")
                for ref, kind, val in cambios:
                    xml = _set_cell(xml, ref, kind, val)
                data = xml.encode("utf8")
            elif item.filename == "xl/workbook.xml":
                xml = data.decode("utf8")
                if "fullCalcOnLoad" not in xml:
                    if "<calcPr" in xml:
                        xml = xml.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
                    else:
                        xml = xml.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
                data = xml.encode("utf8")
            zout.writestr(item, data)


# --------------------------------------------------------------------------- main
def verificar_escritura(dst: Path, sheet: str, cambios) -> None:
    """Read-back: reabre la copia y confirma que cada celda escrita contiene lo previsto; si no, borra y aborta."""
    ws = openpyxl.load_workbook(dst, data_only=False)[sheet]
    malas = []
    for ref, kind, val in cambios:
        got = ws[ref].value
        ok = (isinstance(got, (int, float)) and abs(got - float(val)) < 1e-6) if kind == "n" else str(got) == str(val)
        if not ok:
            malas.append((ref, val, got))
    if malas:
        dst.unlink(missing_ok=True)
        raise RuntimeError(f"verificación de escritura falló en {dst.name}: {malas[:5]} ({len(malas)} celdas)")


def main() -> int:
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=False)
    g.add_argument("--dry-run", action="store_true")
    g.add_argument("--apply", action="store_true")
    ap.add_argument("--costear", nargs=2, metavar=("CODIGO", "CANTIDAD"),
                    help="explota una receta/salida (BOM) y muestra costo e insumos primarios; no escribe nada")
    ap.add_argument("--fecha-precios", default=dt.date.today().strftime("%d/%m/%Y"),
                    help="fecha de ejecución a estampar cuando el ítem no tiene fecha propia (default: hoy)")
    ap.add_argument("--fecha-modo", choices=("maestro", "ejecucion"), default="maestro",
                    help="maestro: fecha de factura del Maestro de Productos (fallback ejecución); "
                         "ejecucion: siempre la fecha de ejecución")
    a = ap.parse_args()
    d, m, y = map(int, a.fecha_precios.split("/"))
    fecha_precio = dt.date(y, m, d)

    global MOTOR, PRODUCTOS, EQUIV, FECHA_MODO
    FECHA_MODO = a.fecha_modo
    MOTOR = MotorRecetas(RECETAS_XLSX)
    PRODUCTOS = cargar_productos()
    EQUIV = cargar_equivalencias()
    if a.costear:
        r = MOTOR.explotar(a.costear[0], float(a.costear[1].replace(",", ".")))
        if r is None:
            print(f"[FALTA MAPEO] Código '{a.costear[0]}' no existe como receta (COD_RECETA/NOM_RECETA)")
            return 1
        print(f"Receta {r.codigo} ({MOTOR.recetas[r.codigo]['nombre']}) | REND {MOTOR.recetas[r.codigo]['rend']} "
              f"{r.unidad} | cantidad {r.cantidad} -> factor {r.factor:.4f}")
        print(f"Costo total {r.costo_total:,.2f} | costo unitario {r.costo_unitario:.4f}/{r.unidad} | "
              f"precio de gestión {r.precio_gestion:,.2f} {'por KG/LT' if r.unidad.lower() in ('gr','g','ml') else 'por unidad'}")
        for k, v in sorted(r.primarios.items(), key=lambda kv: -kv[1])[:15]:
            ins = MOTOR.insumos.get(k, {})
            print(f"   {k:<12} {v:>12,.2f} {ins.get('unidad', '?'):<3} {ins.get('nombre', '')}")
        if r.faltantes:
            print("   FALTA precio:", r.faltantes)
        return 0
    if not (a.dry_run or a.apply):
        ap.error("indicá --dry-run, --apply o --costear")
    precios, dups = cargar_precios()
    print(f"Maestro: {len(precios['sku'])} SKUs + {len(precios['nombre'])} nombres indexados (Chimidoha excluido). "
          f"Fecha de precios: {fecha_precio:%d/%m/%Y}")
    if dups:
        print("SKU duplicados en mapeo (gana el último con precio):", dups)
    print("MAPEO_MAXIREST:", MAPEO_MAXIREST)

    faltantes: dict[tuple, dict] = {}
    TRANSICIONES.clear()
    auditoria: list = []
    for local, cfg in LOCALES.items():
        prev = {}
        pl = OUT_DIR / f"log_{local}.csv"
        if pl.exists():
            with open(pl, newline="", encoding="utf-8-sig") as fh:
                prev = {int(x["fila"]): x["estado"] for x in csv.DictReader(fh)}
        filas, esq = evaluar(local, cfg, precios, fecha_precio)
        trans: dict[str, int] = {}
        for f in filas:
            if prev.get(f["fila"]) == "REVISAR":
                trans[f["estado"]] = trans.get(f["estado"], 0) + 1
        if trans:
            prev_rev = sum(1 for v in prev.values() if v == "REVISAR")
            print(f"\n[{local}] REVISAR previos: {prev_rev} -> ahora: {trans}")
            TRANSICIONES[local] = (prev_rev, trans)
        cnt: dict[str, int] = {}
        vias: dict[str, int] = {}
        for f in filas:
            cnt[f["estado"]] = cnt.get(f["estado"], 0) + 1
            if f["via"]:
                vias[f["via"]] = vias.get(f["via"], 0) + 1
        base = BASELINE_SIN_MAPEO.get(local, 0)
        falta = cnt.get("FALTA_MAPEO", 0)
        print(f"\n=== {local}:{len(filas)} ítems -> {cnt} | cruce por vía: {vias} | categorías a inyectar: "
              f"{sum(1 for f in filas if f['set_rubro'])}")
        print(f"    Sin SKU: {base} (motor anterior) -> {falta} | recuperados: {base - falta}")
        for f in filas:
            if f["via"] in ("MAXIREST", "NOMBRE"):
                print(f"    +{f['via']:<8} {f['codigo']:<10} {str(f['producto'])[:34]:<34} -> {f['sku_canonico'] or 'sin SKU'}"
                      f" ({f['unidad']}) {f['nuevo']} [{f['estado']}]")
        for f in filas:
            nom = str(f["producto"]).lower()
            if f["codigo"] in ("DAT", "BAMAPOLA", "OBR127", "BKUM") or "hojaldre" in nom:
                print(f"  [{f['codigo']}] {f['producto']!r}: via={f['via'] or '-'} sku={f['sku_canonico'] or '-'} "
                      f"unidad={f['unidad']} -> nuevo={f['nuevo']} | F={f['F']} G={f['G']} => {f['estado']} {f['detalle']}")
        for f in filas:
            if f["via"] in ("BOM", "RECETA") or f["costo_bom"] is not None:
                auditoria.append([local, f["fila"], f["codigo"], f["producto"], f["via"], f["receta"], f["costo_bom"],
                                  f["hoja_c_un"], f["precio_maestro"], f["G"],
                                  f["nuevo"] if f["estado"] == "INYECTAR" else "", f["estado"], f["detalle"]])
            if f["estado"] in ("FALTA_MAPEO", "ALERTA_MAXIMA"):
                key = (f["codigo_contingencia"] or f["codigo"], str(f["producto"]).strip())
                e = faltantes.setdefault(key, dict(codigo=key[0], producto=key[1], rubro=f["rubro"], locales=[],
                                                   sugerencia=f["sugerencia"], prioridad=0))
                e["locales"].append(local)
                rl = str(f["rubro"] or "").lower()
                if f["set_rubro"] or f["codigo"].startswith("OBR") or any(k in rl for k in PRIORIDAD_RUBROS):
                    e["prioridad"] = 1
        if a.apply:
            cambios = []
            for f in filas:
                if f["set_rubro"]:
                    cambios.append((f"{get_column_letter(esq['rubro'] + 1)}{f['fila']}", "s", f["set_rubro"]))
                if f["estado"] == "INYECTAR":
                    cambios.append((f"{get_column_letter(esq['target'] + 1)}{f['fila']}", "n", f["nuevo"]))
                    cambios.append((f"{get_column_letter(esq['fecha'] + 1)}{f['fila']}", "s", f["set_fecha"]))
            dst = OUT_DIR / f"{local}_COPIA.xlsx"
            parchear(SRC_DIR / cfg["src"], dst, cfg["sheet"], cambios)
            verificar_escritura(dst, cfg["sheet"], cambios)
            with open(OUT_DIR / f"log_{local}.csv", "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh)
                w.writerow(["fila", "codigo", "producto", "rubro", "medida", "unidad_mapeo", "F_viejo", "G_antes",
                            "G_nuevo", "fecha_antes", "fecha_nueva", "estado", "via_cruce", "sku_canonico",
                            "detalle", "rubro_inyectado", "factor_presentacion", "factor_equiv_rendimiento",
                            "codigo_contingencia", "fecha_origen"])
                for f in filas:
                    w.writerow([f["fila"], f["codigo"], f["producto"], f["rubro"], f["medida"], f["unidad"], f["F"],
                                f["G"], f["nuevo"] if f["estado"] == "INYECTAR" else "", f["fecha_actual"],
                                f["set_fecha"] or "", f["estado"], f["via"], f["sku_canonico"], f["detalle"],
                                f["set_rubro"] or "", f["factor_presentacion"] or "", f["factor_equiv"] or "",
                                f["codigo_contingencia"], f["fecha_origen"]])
            print(f"  -> {dst.name}: {sum(1 for c in cambios if c[1] == 'n')} precios inyectados; log_{local}.csv")

    if a.apply:
        with open(OUT_DIR / "bom_auditoria.csv", "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["local", "fila", "codigo", "producto", "via", "receta", "precio_bom_kg_un", "hoja_c_un",
                        "precio_maestro", "G_antes", "G_nuevo", "estado", "detalle"])
            for r_ in auditoria:
                w.writerow(r_)
    if a.apply:
        st = MOTOR.sincronizar_c_un(RECETAS_XLSX, OUT_DIR / "Recetas_COPIA.xlsx", OUT_DIR / "c_un_sincronizacion.csv")
        print(f"\n[C_UN] Recetas_COPIA.xlsx: {st['recetas']}/{st['total_recetas']} recetas recalculadas "
              f"(completas), celdas escritas {st['celdas']}, difieren >2% de la hoja: {st['difieren_2pct']}")
    if AVISOS:
        print(f"\n[AVISOS defensivos] {len(AVISOS)} (schema/tipado/fechas); primeros:")
        for av in AVISOS[:12]:
            print("   -", av)
        if a.apply:
            (OUT_DIR / "avisos_defensivos.txt").write_text("\n".join(AVISOS), encoding="utf-8")
    lst = sorted(faltantes.values(), key=lambda e: (-e["prioridad"], str(e["codigo"])))
    prio = [e for e in lst if e["prioridad"]]
    print(f"\n=== FALTA MAPEO consolidado: {len(lst)} códigos únicos, {len(prio)} de Obrador/Panificados (prioridad)")
    for e in prio:
        print(f"  [FALTA MAPEO] Código Maxirest: {e['codigo']} | Producto: {e['producto']} | "
              f"{','.join(sorted(set(e['locales'])))} -> Proveer SKU canónico."
              + (f" Sugerencia: {e['sugerencia']}" if e["sugerencia"] else ""))
    if a.apply:
        with open(OUT_DIR / "faltantes_mapeo.csv", "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["prioridad_obrador", "codigo_maxirest", "producto", "rubro", "locales", "sugerencia",
                        "accion", "sku_canonico_a_completar"])
            for e in lst:
                w.writerow([e["prioridad"], e["codigo"], e["producto"], e["rubro"], ",".join(sorted(set(e["locales"]))),
                            e["sugerencia"], "Proveer SKU canónico", ""])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
