"""Capa defensiva compartida (Future-Proofing) del ecosistema Mishiguene.

Cuatro familias de riesgo que cubre:
  1. Cambio de schema    -> resolver_esquema(): columnas por NOMBRE de encabezado (normalizado), nunca por índice.
                            Si falta/duplica un encabezado obligatorio o los datos no son plausibles: SchemaError
                            (el proceso se detiene ANTES de escribir nada).
  2. Datos cíclicos      -> detectar_ciclos(): DFS iterativo sobre un grafo código->hijos (BOM).
  3. Fallos de tipado    -> parse_numero(): "15 kilos", "1.066,75", "$ 1,250.50", "NaN", "#N/A" -> (valor|None, aviso).
  4. Matemática insegura -> safe_div(), finito(): nada de ZeroDivision / NaN / inf hacia las celdas.
Más: fechas con día/mes invertidos por Sheets (parse_fecha_robusta) y detección de códigos colisionados.
"""
from __future__ import annotations

import datetime as dt
import math
import re
import unicodedata
from collections import defaultdict
from typing import Any, Iterable

# Reglas de negocio heredadas del GAS (ventas_bom_rentabilidad / congreso_consumo_teorico)
# Panificados que se COMPRAN a proveedor: son insumo final, nunca se explotan en el BOM.
PANIFICADOS_COMPRADOS = frozenset({"PITA", "BAMAPOLA", "BSESAMO", "BCEBOLLA", "BKUMMEL", "PLETZ", "CROIL", "RDCL",
                                   "PDM", "PDCAF", "BAGEL", "CROISSANT"})
# VARIOS / FLETE: gasto operativo o logística; fuera del balance de materia prima y del food cost.
LOGISTICA_RE = re.compile(r"varios|flete", re.I)


# --- Sanitización de filas de planillas de origen (subtotales, encabezados intermedios, cortes) -------------
_SECCION_RE = re.compile(
    r"^\s*(?:lista\s+de\s+precios\b|saldo\b|subtotal\b"
    r"|total(?:es)?(?:\s+(?:general|mes|local|compras|facturas|gastos))?\s*:?\s*$)", re.I)
_ENCABEZADOS = {"codigo", "código", "cod", "sku", "producto", "productos", "nombre", "item", "ítem", "unidad",
                "u. medida", "precio", "cantidad", "total", "rubro", "descripcion", "descripción"}
_UNIDAD_RE = re.compile(r"^[A-Za-zÁÉÍÓÚÑáéíóúñ][A-Za-zÁÉÍÓÚÑáéíóúñ.\s/]{0,24}$")


def _vacio(v: Any) -> bool:
    return v is None or (isinstance(v, float) and v != v) or str(v).strip() in ("", "nan", "None", "NaT")


def _es_numero(v: Any) -> bool:
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)):
        return not (isinstance(v, float) and v != v)
    return parse_numero(v)[0] is not None and re.fullmatch(r"[\s$\d.,\-+%]+", str(v)) is not None


def motivo_fila_invalida(item: Any = None, codigo: Any = None, unidad: Any = None, *,
                         exigir_codigo: bool = False, exigir_item: bool = False) -> str | None:
    """None si la fila es de datos; si no, el motivo: ítem y código vacíos o ítem numérico, encabezado o subtotal intermedio,
    unidad que no es texto. Un `código` vacío solo invalida cuando exigir_codigo=True (hay planillas donde el
    ítem se identifica por nombre)."""
    if _vacio(item) and (exigir_item or _vacio(codigo)):
        return "ítem vacío" if exigir_item else "ítem y código vacíos"
    if not _vacio(item):
        if _es_numero(item):
            return f"ítem numérico ({item})"
        t = str(item).strip()
        if _SECCION_RE.search(t):
            return f"texto de sección/subtotal ({t[:30]})"
        if t.lower() in _ENCABEZADOS:
            return f"encabezado intermedio ({t})"
    if not _vacio(codigo):
        t = str(codigo).strip()
        if _SECCION_RE.search(t) or t.lower() in _ENCABEZADOS:
            return f"código con texto de sección ({t[:30]})"
    elif exigir_codigo:
        return "código vacío"
    if unidad is not None and not _vacio(unidad):
        if _es_numero(unidad):
            return f"unidad no válida ({str(unidad)[:20]})"
    return None


def filas_validas(df, item: str | None, codigo: str | None = None, unidad: str | None = None, *,
                  exigir_codigo: bool = False, exigir_item: bool = False, origen: str = "", log=None):
    """Devuelve `df` sin las filas parásitas (subtotales, encabezados repetidos, unidad numérica). Las
    columnas que no existan se ignoran. Loguea cuántas y por qué."""
    import logging
    log = log or logging.getLogger("defensivo")
    ci = item if item in df.columns else None
    cc = codigo if codigo in df.columns else None
    cu = unidad if unidad in df.columns else None
    if df.empty or not (ci or cc or cu):
        return df
    motivos = [motivo_fila_invalida(r.get(ci) if ci else None, r.get(cc) if cc else None,
                                    r.get(cu) if cu else None, exigir_codigo=exigir_codigo,
                                    exigir_item=exigir_item and ci is not None)
               for r in df.to_dict("records")]
    malas = [m is not None for m in motivos]
    if any(malas):
        from collections import Counter
        resumen = Counter(m.split(" (")[0] for m in motivos if m)
        ej = next(m for m in motivos if m)
        solo_blancos = set(resumen) == {"ítem y código vacíos"}      # relleno en blanco de la planilla
        (log.info if solo_blancos else log.warning)(
            "%s: %d fila(s) %s descartada(s) %s | ej.: %s", origen or "planilla", sum(malas),
            "en blanco" if solo_blancos else "parásita(s)", dict(resumen), ej)
        return df.loc[[not b for b in malas]].copy()
    return df


def es_panificado_comprado(codigo: Any) -> bool:
    return str(codigo or "").strip().upper() in PANIFICADOS_COMPRADOS


def es_logistica(*textos: Any) -> bool:
    return any(LOGISTICA_RE.search(unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode())
               for t in textos)


class SchemaError(RuntimeError):
    """El libro no tiene el formato que el pipeline asume. Abortar sin escribir."""


# --------------------------------------------------------------------------- texto / encabezados
def sin_tildes(t: Any) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", str(t)) if not unicodedata.combining(c))


def norm_header(t: Any) -> str:
    """' Fecha ultima factura ' -> 'FECHA ULTIMA FACTURA'; '$ unitario  actual' -> '$ UNITARIO ACTUAL'."""
    if t is None:
        return ""
    return " ".join(sin_tildes(t).upper().replace("\n", " ").split())


def resolver_esquema(rows: Iterable[tuple], spec: dict[str, list[str]], *, obligatorios: Iterable[str],
                     max_filas_header: int = 12, legacy: dict[str, int] | None = None,
                     nombre: str = "hoja") -> tuple[dict[str, int], int, list[str]]:
    """Busca la fila de encabezados y mapea cada campo lógico -> índice 0-based.

    spec: campo -> lista de nombres aceptados, en orden de prioridad (igualdad exacta tras normalizar).
    La fila de encabezados es la primera (dentro de max_filas_header) que cubre TODOS los obligatorios.
    -> (columnas, fila_header_1based, avisos).  Si algo no cuadra: SchemaError.
    """
    obligatorios = list(obligatorios)
    avisos: list[str] = []
    candidatos = []
    for i, row in enumerate(rows, start=1):
        if i > max_filas_header:
            break
        heads = [norm_header(c) for c in row]
        cols: dict[str, int] = {}
        for campo, nombres in spec.items():
            for n in nombres:
                n = norm_header(n)
                hits = [j for j, h in enumerate(heads) if h == n]
                if hits:
                    cols[campo] = hits[0]
                    if len(hits) > 1:
                        avisos.append(f"{nombre}: encabezado '{n}' duplicado en columnas {hits}; uso la primera")
                    break
        candidatos.append((i, cols))
        if all(o in cols for o in obligatorios):
            if legacy:
                for k, v in cols.items():
                    if k in legacy and legacy[k] != v:
                        avisos.append(f"{nombre}: '{k}' está en la columna {v} y el esquema histórico decía "
                                      f"{legacy[k]} (¿se insertó/movió una columna?). Se usa el encabezado.")
            return cols, i, avisos
    mejor = max(candidatos, key=lambda c: len(c[1]), default=(0, {}))
    faltan = [o for o in obligatorios if o not in mejor[1]]
    raise SchemaError(f"{nombre}: no encuentro los encabezados obligatorios {faltan} en las primeras "
                      f"{max_filas_header} filas (mejor candidata: fila {mejor[0]} con {sorted(mejor[1])}).")


def plausibilidad_numerica(valores: Iterable[Any], minimo: float = 0.6) -> float:
    """Fracción de celdas no vacías que se pueden leer como número. Sirve para detectar una columna cambiada."""
    n = ok = 0
    for v in valores:
        if v is None or (isinstance(v, str) and not v.strip()):
            continue
        n += 1
        if parse_numero(v)[0] is not None:
            ok += 1
    return ok / n if n else 1.0


# --------------------------------------------------------------------------- números
def finito(x: Any) -> float | None:
    try:
        f = float(x)
    except (TypeError, ValueError):
        return None
    return f if math.isfinite(f) else None


def safe_div(a: Any, b: Any, default: float | None = None) -> float | None:
    a, b = finito(a), finito(b)
    if a is None or b is None or abs(b) < 1e-12:
        return default
    return a / b


_UNIDAD_TXT = re.compile(r"[a-zA-Záéíóúñ]+\.?$")


def parse_numero(v: Any) -> tuple[float | None, str]:
    """Interpreta lo que un humano puede escribir en una celda numérica.
    -> (valor|None, aviso).  '15 kilos' -> (15.0, 'texto con unidad: kilos'); '1.066,75' -> 1066.75;
    '$ 1,250.50' -> 1250.5; 'NaN', '#N/A', '' -> (None, ...). Nunca devuelve NaN/inf."""
    if v is None or isinstance(v, bool):
        return None, "vacío"
    if isinstance(v, (int, float)):
        f = finito(v)
        return (f, "") if f is not None else (None, "NaN/inf")
    s = str(v).strip()
    if not s or s.startswith("=") or s.upper() in {"#N/A", "#DIV/0!", "#REF!", "#VALUE!", "NAN", "N/A", "-", "NULL"}:
        return None, f"no numérico ({s[:20]!r})"
    unidad = ""
    m = re.match(r"^\s*(?:\$|U\$S|USD|ARS)?\s*([\-+]?[\d.,\s]+?)\s*([A-Za-zÁÉÍÓÚáéíóúÑñ].*)?$", s)
    if not m:
        return None, f"sin dígitos ({s[:20]!r})"
    num_s, resto = m.group(1).strip().replace(" ", ""), (m.group(2) or "").strip()
    if resto:
        if not _UNIDAD_TXT.match(resto):
            return None, f"texto no interpretable ({s[:20]!r})"
        unidad = resto
    if "," in num_s and "." in num_s:
        dec = "," if num_s.rfind(",") > num_s.rfind(".") else "."
        num_s = num_s.replace("." if dec == "," else ",", "").replace(dec, ".")
    elif "," in num_s:
        partes = num_s.split(",")
        num_s = num_s.replace(",", ".") if len(partes) == 2 and len(partes[1]) != 3 else num_s.replace(",", "")
    elif num_s.count(".") > 1:
        num_s = num_s.replace(".", "")
    f = finito(num_s)
    if f is None:
        return None, f"no numérico ({s[:20]!r})"
    return f, (f"texto con unidad: {unidad}" if unidad else "número guardado como texto")


# --------------------------------------------------------------------------- fechas
def parse_fecha_robusta(v: Any, referencia: dt.date, ventana_dias: int = 75) -> tuple[dt.date | None, str]:
    """Lee fechas de celdas datetime o texto 'dd/mm/yyyy' | 'Factura: dd/mm/yyyy' | 'Histórico: Mes (dd/mm)'.
    Si Sheets invirtió día/mes (ej. 2026-03-09 por '03/09/2026') y la versión invertida cae dentro de
    `ventana_dias` de la referencia mientras la original no, usa la invertida. -> (fecha|None, aviso)."""
    d: dt.date | None = None
    if isinstance(v, dt.datetime):
        d = v.date()
    elif isinstance(v, dt.date):
        d = v
    elif isinstance(v, str):
        m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", v)
        if m:
            dd, mm, yy = map(int, m.groups())
            try:
                d = dt.date(yy, mm, dd)
            except ValueError:
                return None, "fecha inválida"
        else:
            m = re.search(r"\((\d{1,2})/(\d{1,2})\)", v)
            if m:
                dd, mm = map(int, m.groups())
                try:
                    d = dt.date(referencia.year, mm, dd)
                except ValueError:
                    return None, "fecha inválida"
    if d is None:
        return None, ""
    if abs((d - referencia).days) > ventana_dias and d.day <= 12:
        try:
            sw = dt.date(d.year, d.day, d.month)
        except ValueError:
            return d, ""
        if abs((sw - referencia).days) <= ventana_dias:
            return sw, f"día/mes invertidos en origen ({d:%d/%m/%Y} -> {sw:%d/%m/%Y})"
    return d, ""


# --------------------------------------------------------------------------- grafos
def detectar_ciclos(grafo: dict[str, list[str]]) -> list[list[str]]:
    """Ciclos de un grafo dirigido (incluye autorreferencias). DFS iterativo: sin límite de recursión."""
    BLANCO, GRIS, NEGRO = 0, 1, 2
    color: dict[str, int] = defaultdict(int)
    ciclos: list[list[str]] = []
    for raiz in grafo:
        if color[raiz] != BLANCO:
            continue
        pila = [(raiz, iter(grafo.get(raiz, [])))]
        camino = [raiz]
        color[raiz] = GRIS
        while pila:
            nodo, it = pila[-1]
            avanzo = False
            for hijo in it:
                if color[hijo] == GRIS:
                    ciclos.append(camino[camino.index(hijo):] + [hijo])
                elif color[hijo] == BLANCO:
                    color[hijo] = GRIS
                    camino.append(hijo)
                    pila.append((hijo, iter(grafo.get(hijo, []))))
                    avanzo = True
                    break
            if not avanzo:
                color[nodo] = NEGRO
                pila.pop()
                camino.pop()
    return ciclos


def nodos_en_ciclo(grafo: dict[str, list[str]]) -> set[str]:
    return {n for c in detectar_ciclos(grafo) for n in c}


# --------------------------------------------------------------------------- colisiones de códigos
def tokens_nombre(t: Any) -> set[str]:
    stop = {"DE", "DEL", "LA", "EL", "Y", "CON", "UN", "GR", "PORC", "X", "EN", "A"}
    return {w for w in re.sub(r"[^A-Z0-9 ]+", " ", sin_tildes(t).upper()).split() if w not in stop and len(w) > 1}


def codigos_colisionados(pares: Iterable[tuple[str, Any]], umbral: float = 0.8) -> dict[str, set[str]]:
    """pares = (codigo, nombre). Un código está colisionado si agrupa productos cuyos nombres no son
    variantes del mismo (Jaccard de tokens < umbral): 'POL' = 'Petalo de rosa' y 'POLENTA'."""
    nombres: dict[str, list[set[str]]] = defaultdict(list)
    crudos: dict[str, set[str]] = defaultdict(set)
    for cod, nom in pares:
        if cod:
            t = tokens_nombre(nom)
            nombres[cod].append(t)
            crudos[cod].add(str(nom).strip().lower())
    out = {}
    for cod, lst in nombres.items():
        for i in range(len(lst)):
            for j in range(i + 1, len(lst)):
                u = lst[i] | lst[j]
                if u and len(lst[i] & lst[j]) / len(u) < umbral:
                    out[cod] = crudos[cod]
                    break
            if cod in out:
                break
    return out


if __name__ == "__main__":
    # auto-test mínimo
    assert parse_numero("15 kilos") == (15.0, "texto con unidad: kilos")
    assert parse_numero("1.066,75")[0] == 1066.75 and parse_numero("1066,75")[0] == 1066.75
    assert parse_numero("$ 1,250.50")[0] == 1250.5 and parse_numero("NaN")[0] is None
    assert parse_numero("abc 12 xyz")[0] is None and parse_numero(float("inf"))[0] is None
    assert safe_div(1, 0) is None and safe_div(1, float("nan")) is None and safe_div(6, 3) == 2
    assert detectar_ciclos({"A": ["B"], "B": ["C"], "C": ["A"], "D": ["D"], "E": []}) == [["A", "B", "C", "A"], ["D", "D"]]
    assert parse_fecha_robusta(dt.datetime(2026, 3, 9), dt.date(2026, 10, 12))[0] == dt.date(2026, 9, 3)
    assert parse_fecha_robusta(dt.datetime(2026, 9, 29), dt.date(2026, 10, 12))[0] == dt.date(2026, 9, 29)
    assert "POL" in codigos_colisionados([("POL", "Petalo de rosa"), ("POL", "POLENTA"), ("AZU", "Azucar"), ("AZU", "Azúcar")])
    print("defensivo.py: auto-test OK")
