"""Consolida compras y ventas de los 5 locales de Mishiguene y genera los informes que lee la web.

Qué hace, en orden:

1. FACTURAS de cada local (Congreso, Arcos, Cabello, Lafinur, Chimidoha). Fuentes, la primera con datos gana:
   a) subcarpetas del proyecto con el nombre del local con planillas .xlsx que tengan la pestaña 'Facturas';
   b) el Google Sheet del local (LOCALES, vía drive_connector);
   c) copias descargadas en Arcos_data/.
2. VENTAS: busca de forma recursiva planillas de ventas/gestión en Drive (carpetas y archivos que digan
   TABLERO, VENTAS, MENSUAL, GESTIÓN, RESUMEN, OCTUBRE...) y en carpetas locales equivalentes, y parsea todo lo que
   tenga ventas por artículo (Ventas ARCOS 09-26, Resumen Ventas, Ventas_Obrador_<Mes>, etc.). El resultado es un
   índice único que se usa para el cruce de BOM (recetas vs ventas) y para el reporte de huérfanos.
3. MESES dinámicos: se procesan los meses con datos (por defecto los 2 más recientes). Un local que solo tiene
   septiembre procesa septiembre; si aparece octubre, se incorpora solo.
4. SALIDAS (en la raíz del proyecto, o en WebTest/ si la raíz no admite archivos nuevos):
   - Dossier_Mensual_Mishiguene.xlsx: una hoja por local y mes ("Cabello Septiembre", "Arcos Octubre") y una por
     mes con la cadena ("Cadena Septiembre").
   - Analisis_Gestion_Comercial_{Local}_{MM}_{YYYY}.xlsx: completo (analisis_gestion_comercial.py) donde hay
     ventas por artículo, inventarios y recetas; si no, con facturas y el resumen de la gestión mensual.
   - reporte_items_huerfanos.txt: cada ítem sin ventas que lo expliquen, con los archivos de ventas donde se
     buscó su código y donde NO apareció.

Uso:
    python consolidador.py                      # los 2 meses más recientes con datos
    python consolidador.py --meses 9 10 --anio 2026
    python consolidador.py --sin-analisis       # solo Dossier + informes de facturas
    python consolidador.py --refresh            # volver a bajar las planillas de ventas de Drive
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
import unicodedata
from datetime import date
from pathlib import Path
from typing import Any

import pandas as pd

from defensivo import filas_validas, motivo_fila_invalida
from drive_connector import read_facturas, read_facturas_xlsx

LOG = logging.getLogger("consolidador")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "Arcos_data"
VENTAS_DIR = DATA_DIR / "ventas_drive"      # caché de las planillas de ventas bajadas de Drive
RESPALDO_SALIDA = BASE_DIR / "WebTest"      # si la raíz no admite archivos nuevos
DOSSIER = "Dossier_Mensual_Mishiguene.xlsx"
CARPETAS_IGNORADAS = {"node_modules", ".git", "__pycache__", "backup_previo", "WebTest", "Arcos_data",
                      "Archivo_Mantenimiento", "Reportes_Generados", "Documentacion"}

# ID del Google Sheet de cada local (el tramo entre /d/ y /edit en la URL).
LOCALES = {
    "Congreso": "1udLEQpQ29avK4c6fNnQig-w02xIV82IgLe85yVJXaOY",
    "Arcos": "1frh93qB8Od5mrtAfMuObKwRW93p7kev6DNgmERqz6ho",
    "Cabello": "1zAxYyNB98Sb1-D0HtfokESXc6FrDS96EbFYyMmsseIE",
    "Lafinur": "1tv0uI4KsSEPRUecYfnwcDqpGbUDelGqWwe3Hp-fikOc",
    "Chimidoha": "1gSHbDSgICeTy6lZpyIPRskFbEs0JB5mxw1e71oOkIzU",
}
# Nombres alternativos de las copias descargadas en Arcos_data/ (el nombre del local ya se prueba solo)
ALIAS_ARCHIVO = {"Congreso": ["congreso_facturas"], "Chimidoha": ["chimi"]}

MESES_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio", 7: "Julio",
            8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
MES_TOKENS = {1: ("enero", "ene"), 2: ("febrero", "feb"), 3: ("marzo", "mar"), 4: ("abril", "abr"),
              5: ("mayo",), 6: ("junio", "jun"), 7: ("julio", "jul"), 8: ("agosto", "ago"),
              9: ("septiembre", "setiembre", "sept", "sep", "set"), 10: ("octubre", "oct"),
              11: ("noviembre", "nov"), 12: ("diciembre", "dic")}

GROUP_COLS = ["Código SKU", "Productos", "Proveedor", "Categoría Gasto", "Categoría 2"]
OUTPUT_COLS = ["Código SKU", "Productos", "Categoría Gasto", "Categoría 2", "Proveedor",
               "Cantidad Inventario", "Monto Total", "Precio Ponderado"]
# Una pestaña mensual se considera "con datos" si tiene al menos este % de las filas del mes más cargado
UMBRAL_MES = 0.05

# Carpetas/archivos de Drive o locales que pueden traer ventas o gestión (búsqueda recursiva por ruta completa)
PATRON_CARPETA = re.compile(r"tablero|venta|octubre|control", re.I)   # carpetas locales (recursivo)
PATRON_CANDIDATOS = re.compile(
    r"tablero|venta|mensual|gesti[oó]n|resumen|octubre|septiembre|control|conteo|casi|obrador", re.I)
# Hojas que nunca son ventas por artículo (stock, movimientos, traza de recetas...)
HOJAS_IGNORADAS = re.compile(r"movimiento|stock|conteo|inventario|traza|maestro|equivalencia|pronostico|planific", re.I)
ALIAS_NOMBRE = {"nombre", "articulo", "producto", "descripcion", "item", "detalle", "plato"}
ALIAS_UNIDADES = {"unidades", "cantidad", "cant", "uds", "unid", "unidades vendidas", "cantidad vendida"}
ALIAS_VENTA = {"venta", "venta total", "importe", "precio total", "total", "monto", "ventas", "ventas $", "total venta"}
ALIAS_CODIGO = {"codigo", "sku", "cod", "cod producto", "codigo producto", "cod maxirest", "codigo maxirest"}
MAX_FILAS_HOJA = 50_000


def _nk(txt: object) -> str:
    t = unicodedata.normalize("NFKD", str(txt))
    return " ".join("".join(c for c in t if not unicodedata.combining(c)).lower().split())


def meses_en_texto(txt: object) -> set[int]:
    """Meses mencionados en un nombre/ruta: 'Septiembre', 'sept', 'oct', '09-26', '10/2026'."""
    t = _nk(txt)
    meses: set[int] = set()
    for m, toks in MES_TOKENS.items():
        if any(re.search(rf"(?<![a-z]){tok}(?![a-z])", t) for tok in toks):
            meses.add(m)
    for a, b in re.findall(r"(?<!\d)(\d{1,2})[-_/.](\d{2}|\d{4})(?!\d)", t):
        if 1 <= int(a) <= 12 and (len(b) == 4 and b.startswith("20") or len(b) == 2 and int(b) >= 20):
            meses.add(int(a))
    return meses


# Panificados que se cuentan SIEMPRE por unidad, diga lo que diga la columna original; hojaldre, siempre por KG
_RE_PANIFICADO_UN = re.compile(
    r"(?<![a-z])(bagels?|baguels?|pletzalej|pan de nuez|pan de campo|hogazas?|pitas?|laffas?|croissants?)(?![a-z])")
_RE_HOJALDRE_KG = re.compile(r"(?<![a-z])hojaldre")


def unidad_forzada(nombre: object) -> str | None:
    """'UN' para bagel/baguel/pletzalej/pan de nuez/pan de campo/hogaza/pita/laffa/croissant, 'KG' para hojaldre
    (la excepción gana), None si el ítem no tiene regla y vale la unidad de la planilla."""
    t = _nk(nombre)
    if _RE_HOJALDRE_KG.search(t):
        return "KG"
    return "UN" if _RE_PANIFICADO_UN.search(t) else None


def local_en_texto(txt: object) -> str | None:
    t = _nk(txt).replace("arcos_data", "")
    for local in LOCALES:
        if _nk(local) in t:
            return local
    return None


# --------------------------------------------------------------------------- #
# Carga de facturas por local
# --------------------------------------------------------------------------- #
def _tiene_facturas(path: Path) -> bool:
    try:
        return "Facturas" in pd.ExcelFile(path).sheet_names
    except Exception:  # noqa: BLE001 - archivo corrupto/abierto/no-Excel: simplemente no sirve
        return False


def _carpetas_del_local(local: str) -> list[Path]:
    """Subcarpetas del proyecto cuyo nombre es el del local (cualquier profundidad razonable)."""
    objetivo, hallazgos = _nk(local), []
    if (BASE_DIR / local).is_dir():
        hallazgos.append(BASE_DIR / local)
    for top in BASE_DIR.iterdir():
        if not top.is_dir() or top.name in CARPETAS_IGNORADAS or top.name.startswith("."):
            continue
        if _nk(top.name) == objetivo and top not in hallazgos:
            hallazgos.append(top)
        hallazgos += [d for d in top.rglob("*") if d.is_dir() and _nk(d.name) == objetivo and d not in hallazgos]
    return hallazgos


def _archivos_subcarpetas(local: str) -> list[Path]:
    archivos: list[Path] = []
    for carpeta in _carpetas_del_local(local):
        for p in carpeta.rglob("*.xlsx"):
            if p.name.startswith(("~$", "Analisis_Gestion_Comercial_", "Dossier_")):
                continue
            if _tiene_facturas(p):
                archivos.append(p)
    return sorted(archivos, key=lambda p: p.stat().st_mtime, reverse=True)


def _archivos_cache(local: str) -> list[Path]:
    claves = [_nk(local)] + [_nk(a) for a in ALIAS_ARCHIVO.get(local, [])]
    if not DATA_DIR.is_dir():
        return []
    cand = [p for p in DATA_DIR.glob("*.xlsx") if any(k in _nk(p.stem) for k in claves)]
    return sorted([p for p in cand if _tiene_facturas(p)], key=lambda p: p.stat().st_mtime, reverse=True)


def cargar_local(local: str, sheet_id: str | None) -> tuple[pd.DataFrame | None, str]:
    """Facturas filtradas (materia prima y descartables) del local y descripción de la fuente usada."""
    notas: list[str] = []
    for p in _archivos_subcarpetas(local):
        try:
            df = read_facturas_xlsx(p)
        except Exception as exc:  # noqa: BLE001
            notas.append(f"{p.name}: {exc}")
            continue
        if len(df):
            return df, f"subcarpeta {p.parent.name}/{p.name}"
    if sheet_id and not sheet_id.startswith("PEGAR_ID"):
        try:
            df = read_facturas(sheet_id)
            if len(df):
                return df, "Google Sheet del local"
        except Exception as exc:  # noqa: BLE001 - sin red / sin acceso: se sigue con las copias locales
            notas.append(f"Google Sheet: {type(exc).__name__}: {exc}")
    for p in _archivos_cache(local):
        try:
            df = read_facturas_xlsx(p)
        except Exception as exc:  # noqa: BLE001
            notas.append(f"{p.name}: {exc}")
            continue
        if len(df):
            return df, f"copia local Arcos_data/{p.name}"
    LOG.warning("%s: sin facturas en ninguna fuente (%s)", local, "; ".join(notas) or "no hay archivos")
    return None, "sin datos"


def cargar_locales() -> tuple[pd.DataFrame, dict[str, str]]:
    """DataFrame crudo con la columna 'Local' y {local: fuente}. Los 5 locales siempre figuran en el dict."""
    frames: list[pd.DataFrame] = []
    fuentes: dict[str, str] = {}
    for local, sheet_id in LOCALES.items():
        LOG.info("Facturas de %s...", local)
        df, fuente = cargar_local(local, sheet_id)
        fuentes[local] = fuente
        if df is None:
            continue
        df = df.reset_index()  # conserva 'Fila Sheet' como columna
        df.insert(0, "Local", local)
        frames.append(df)
        LOG.info("  [OK] %s: %d filas (%s)", local, len(df), fuente)
    crudo = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(
        columns=["Local", "Mes factura", "Cantidad Inventario", "Precio Unitario", *GROUP_COLS])
    crudo = filas_validas(crudo, "Productos", "Código SKU", None, exigir_item=False, origen="Facturas (todos los locales)",
                          log=LOG)
    return crudo, fuentes


def meses_disponibles(crudo: pd.DataFrame, ultimos: int) -> list[int]:
    """Los `ultimos` meses con carga significativa (descarta meses sueltos por errores de tipeo)."""
    if crudo.empty:
        return []
    cuenta = crudo["Mes factura"].value_counts()
    validos = sorted(m for m, n in cuenta.items() if n >= UMBRAL_MES * cuenta.max() and 1 <= m <= 12)
    return validos[-ultimos:]


# --------------------------------------------------------------------------- #
# Fuentes de ventas: Drive + carpetas locales (búsqueda recursiva)
# --------------------------------------------------------------------------- #
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SHEETS_MIME = "application/vnd.google-apps.spreadsheet"
FOLDER_MIME = "application/vnd.google-apps.folder"


def _slug(txt: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", _nk(txt)).strip("_")[:60] or "archivo"


def _sesion_drive():
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_file(
        str(BASE_DIR / "credentials.json"), scopes=["https://www.googleapis.com/auth/drive.readonly"])
    return AuthorizedSession(creds)


def listar_drive(ses) -> list[dict[str, Any]]:
    """Todos los archivos visibles, con la ruta completa de carpetas (recursivo vía 'parents')."""
    files: list[dict[str, Any]] = []
    token = None
    while True:
        params = {"q": "trashed=false", "pageSize": 1000, "supportsAllDrives": "true",
                  "includeItemsFromAllDrives": "true",
                  "fields": "nextPageToken,files(id,name,mimeType,parents,modifiedTime)"}
        if token:
            params["pageToken"] = token
        r = ses.get("https://www.googleapis.com/drive/v3/files", params=params)
        r.raise_for_status()
        j = r.json()
        files += j.get("files", [])
        token = j.get("nextPageToken")
        if not token:
            break
    por_id = {f["id"]: f for f in files}

    def ruta(f: dict[str, Any]) -> str:
        partes, p = [f["name"]], (f.get("parents") or [None])[0]
        while p and p in por_id:
            partes.append(por_id[p]["name"])
            p = (por_id[p].get("parents") or [None])[0]
        return "/".join(reversed(partes))

    for f in files:
        f["ruta"] = ruta(f)
    return files


def _es_planilla(f: dict[str, Any]) -> bool:
    return f["mimeType"] == SHEETS_MIME or f["name"].lower().endswith((".xlsx", ".xlsm"))


def fuentes_ventas_drive(meses: list[int], anio: int, refresh: bool) -> list[tuple[Path, str]]:
    """Planillas candidatas de Drive bajadas a VENTAS_DIR: [(archivo local, 'carpeta/nombre')].
    Se descartan las que, por su nombre, son de otro mes o de otro año."""
    try:
        ses = _sesion_drive()
        archivos = listar_drive(ses)
    except Exception as exc:  # noqa: BLE001 - sin red/credenciales: se sigue con lo local
        LOG.warning("Drive no disponible para buscar ventas (%s: %s): solo carpetas locales y caché",
                    type(exc).__name__, exc)
        return [(p, f"Arcos_data/ventas_drive/{p.name}") for p in sorted(VENTAS_DIR.glob("*.xlsx"))] \
            if VENTAS_DIR.is_dir() else []
    VENTAS_DIR.mkdir(exist_ok=True)
    salida: list[tuple[Path, str]] = []
    descartados = 0
    for f in archivos:
        if f["mimeType"] == FOLDER_MIME or not _es_planilla(f) or not PATRON_CANDIDATOS.search(f["ruta"]):
            continue
        anios = {int(a) for a in re.findall(r"(?<!\d)(20\d\d)(?!\d)", f["ruta"])}
        mes_f = meses_en_texto(f["name"])
        if (anios and anio not in anios) or (mes_f and not mes_f & set(meses)):
            descartados += 1
            continue
        destino = VENTAS_DIR / f"{_slug(f['name'])}__{f['id'][:6]}.xlsx"
        if refresh or not destino.exists():
            try:
                if f["mimeType"] == SHEETS_MIME:
                    r = ses.get(f"https://www.googleapis.com/drive/v3/files/{f['id']}/export",
                                params={"mimeType": XLSX_MIME})
                else:
                    r = ses.get(f"https://www.googleapis.com/drive/v3/files/{f['id']}", params={"alt": "media"})
                r.raise_for_status()
                destino.write_bytes(r.content)
            except Exception as exc:  # noqa: BLE001
                LOG.warning("No pude bajar '%s': %s", f["ruta"], exc)
                continue
        salida.append((destino, f["ruta"]))
    LOG.info("Drive: %d planillas candidatas a ventas/gestión (%d descartadas por otro mes/año)",
             len(salida), descartados)
    return salida


def fuentes_ventas_locales(meses: list[int], anio: int) -> list[tuple[Path, str]]:
    """Carpetas locales con nombre tipo TABLERO/VENTAS/OCTUBRE (recursivo) y las ventas cacheadas por el análisis."""
    salida: list[tuple[Path, str]] = []
    for top in BASE_DIR.iterdir():
        if not top.is_dir() or top.name in CARPETAS_IGNORADAS or top.name.startswith("."):
            continue
        for p in top.rglob("*.xlsx"):
            rel = str(p.relative_to(BASE_DIR))
            en_carpeta = any(PATRON_CARPETA.search(parte) for parte in p.relative_to(BASE_DIR).parts[:-1])
            if p.name.startswith(("~$", "Analisis_Gestion_Comercial_", "Dossier_")) or not (
                    en_carpeta or re.search("venta", p.name, re.I)):
                continue
            mes_f = meses_en_texto(p.name)
            if mes_f and not mes_f & set(meses):
                continue
            salida.append((p, rel))
    if DATA_DIR.is_dir():
        for p in DATA_DIR.glob("*.xlsx"):
            if not p.name.startswith("~$") and re.search(r"venta|resumen|conteo|casi_gesti", p.stem, re.I):
                mes_f = meses_en_texto(p.stem)
                if not mes_f or mes_f & set(meses):
                    salida.append((p, f"Arcos_data/{p.name}"))
    return salida


def _fila_encabezado(df: pd.DataFrame) -> tuple[int, dict[str, int]] | None:
    """Primera fila (de las 15 primeras) con una columna de nombre y una de unidades."""
    for i in range(min(15, len(df))):
        celdas = {_nk(v): j for j, v in enumerate(df.iloc[i].tolist()) if isinstance(v, str) and v.strip()}
        nom = next((j for k, j in celdas.items() if k in ALIAS_NOMBRE), None)
        uni = next((j for k, j in celdas.items() if k in ALIAS_UNIDADES), None)
        if nom is None or uni is None or nom == uni:
            continue
        ven = next((j for k, j in celdas.items() if k in ALIAS_VENTA and j not in (nom, uni)), None)
        cod = next((j for k, j in celdas.items() if k in ALIAS_CODIGO and j not in (nom, uni)), None)
        fec = next((j for k, j in celdas.items() if k == "fecha"), None)
        return i, {"nombre": nom, "unidades": uni, "venta": ven, "codigo": cod, "fecha": fec}
    return None


def parsear_ventas(path: Path, etiqueta: str) -> tuple[list[pd.DataFrame], list[str]]:
    """Todas las hojas del archivo que tengan ventas por artículo. Devuelve (frames, hojas_sin_ventas)."""
    frames: list[pd.DataFrame] = []
    sin: list[str] = []
    x = pd.ExcelFile(path)
    local_arch = local_en_texto(etiqueta)
    meses_arch = meses_en_texto(Path(etiqueta).name) or meses_en_texto(etiqueta)
    for hoja in x.sheet_names:
        if HOJAS_IGNORADAS.search(hoja):
            continue
        try:
            df = x.parse(hoja, header=None)
        except Exception as exc:  # noqa: BLE001
            sin.append(f"{hoja} (error: {exc})")
            continue
        if df.empty or len(df) > MAX_FILAS_HOJA:
            continue
        enc = _fila_encabezado(df)
        if enc is None:
            sin.append(hoja)
            continue
        i, cols = enc
        cuerpo = df.iloc[i + 1:]
        out = pd.DataFrame({
            "articulo": cuerpo.iloc[:, cols["nombre"]],
            "unidades": pd.to_numeric(cuerpo.iloc[:, cols["unidades"]], errors="coerce"),
            "venta": pd.to_numeric(cuerpo.iloc[:, cols["venta"]], errors="coerce") if cols["venta"] is not None else 0.0,
            "codigo": cuerpo.iloc[:, cols["codigo"]] if cols["codigo"] is not None else None,
            "fecha": pd.to_datetime(cuerpo.iloc[:, cols["fecha"]], errors="coerce") if cols["fecha"] is not None else pd.NaT,
        })
        out = out[out["articulo"].map(lambda v: isinstance(v, str) and bool(v.strip()))]
        out = out[out["unidades"].notna()]
        out = out[~out["articulo"].map(_nk).str.fullmatch(r"total.*|subtotal.*|nan")]
        if out.empty:
            sin.append(hoja)
            continue
        out["articulo"] = out["articulo"].astype(str).str.strip()
        out["k"] = out["articulo"].map(_nk)
        out["codigo"] = out["codigo"].map(lambda v: re.sub(r"\.0$", "", str(v).strip()) if pd.notna(v) else None)
        if out["fecha"].notna().any():          # historial diario: se agrega por mes para no inflar el índice
            out["mes"] = out["fecha"].dt.month
        else:
            out["mes"] = next(iter(meses_arch)) if len(meses_arch) == 1 else None
        g = out.groupby(["k", "codigo", "mes"], dropna=False, as_index=False).agg(
            articulo=("articulo", "first"), unidades=("unidades", "sum"), venta=("venta", "sum"))
        g["archivo"], g["hoja"], g["local"] = etiqueta, hoja, local_arch
        frames.append(g)
    return frames, sin


def construir_indice_ventas(fuentes: list[tuple[Path, str]]) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Índice único (archivo, hoja, local, mes, articulo, k, codigo, unidades, venta) + resumen de lo leído."""
    frames: list[pd.DataFrame] = []
    sin_ventas: list[str] = []
    con_error: list[str] = []
    vistos: set[str] = set()
    for path, etiqueta in fuentes:
        if etiqueta in vistos:
            continue
        vistos.add(etiqueta)
        try:
            fr, sin = parsear_ventas(path, etiqueta)
        except Exception as exc:  # noqa: BLE001 - un archivo roto no frena el resto
            con_error.append(f"{etiqueta}: {type(exc).__name__}: {exc}")
            continue
        frames += fr
        if not fr:
            sin_ventas.append(etiqueta)
    cols = ["archivo", "hoja", "local", "mes", "articulo", "k", "codigo", "unidades", "venta"]
    unicos: list[pd.DataFrame] = []
    firmas: dict[tuple, str] = {}
    copias: dict[str, list[str]] = {}
    for fr in frames:
        firma = (fr["hoja"].iloc[0], len(fr), round(float(fr["unidades"].sum()), 2), round(float(fr["venta"].sum()), 2),
                 tuple(sorted(fr["k"].head(25))))
        if firma in firmas:
            copias.setdefault(firmas[firma], []).append(fr["archivo"].iloc[0])
            continue
        firmas[firma] = f"{fr['archivo'].iloc[0]} › {fr['hoja'].iloc[0]}"
        unicos.append(fr)
    frames = unicos
    indice = pd.concat(frames, ignore_index=True)[cols] if frames else pd.DataFrame(columns=cols)
    resumen = (indice.groupby(["archivo", "hoja"], sort=False)
               .agg(filas=("k", "size"), local=("local", "first"), mes=("mes", lambda s: ",".join(
                   str(int(m)) for m in sorted(set(s.dropna())))))
               .reset_index()) if len(indice) else pd.DataFrame(columns=["archivo", "hoja", "filas", "local", "mes"])
    return indice, {"analizados": resumen, "sin_ventas": sin_ventas, "con_error": con_error, "copias": copias}


# --------------------------------------------------------------------------- #
# Obrador: Historial_Productos (Ventas_Obrador_<Mes>), CONTEO OBRADOR y 'Casi gestion' / Equivalencias
# --------------------------------------------------------------------------- #
LOCAL_PLANTA = "Congreso"           # el Obrador es la planta de Congreso: ahí se cuelga la producción
OUTLIER_PRECIO_X = 50.0             # precio unitario a más de X veces (o menos de 1/X) de la mediana del ítem
PREFIJO_UNIDAD = re.compile(r"^\s*(gr|g|kg|un|u|und|porc|lt|ml|cc)\.\s*(.*)$", re.I)


def partir_nombre_maxirest(nombre: Any) -> tuple[str, str]:
    """'GR. Latke' -> ('GR', 'Latke'); 'Un. Pan de Nuez' -> ('UN', 'Pan de Nuez'); sin prefijo -> ('', nombre)."""
    t = str(nombre).strip()
    m = PREFIJO_UNIDAD.match(t)
    return (m.group(1).upper(), m.group(2).strip()) if m else ("", t)


def _codigo_txt(v: Any) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    t = str(v).strip()
    return re.sub(r"\.0+$", "", t)


def _planillas(fuentes: list[tuple[Path, str]], patron: str) -> list[tuple[Path, str]]:
    """Planillas cuyo nombre/etiqueta cumple `patron`, de la más nueva a la más vieja (sin duplicar rutas)."""
    vistos, out = set(), []
    for path, etiqueta in fuentes:
        if re.search(patron, f"{etiqueta} {path.name}", re.I) and path not in vistos and path.exists():
            vistos.add(path)
            out.append((path, etiqueta))
    return sorted(out, key=lambda t: t[0].stat().st_mtime, reverse=True)


def cargar_equivalencias_maxirest(fuentes: list[tuple[Path, str]], extra: pd.DataFrame | None = None
                                  ) -> tuple[pd.DataFrame, str]:
    """Diccionario maestro 'Codigo maxirest' -> 'SKU RECETA' (pestaña Equivalencias de Casi gestion).
    `extra` (codigo, sku) completa los códigos que Casi gestion no trae. Devuelve (df, descripción)."""
    cand = _planillas(fuentes, r"casi[_ ]?gesti")
    if (DATA_DIR / "casi_gestion.xlsx").exists():
        cand.append((DATA_DIR / "casi_gestion.xlsx", "Arcos_data/casi_gestion.xlsx"))
    filas: list[dict[str, str]] = []
    usadas: list[str] = []
    for path, etiqueta in cand:
        try:
            d = pd.read_excel(path, sheet_name="Equivalencias")
        except Exception as exc:  # noqa: BLE001
            LOG.warning("Casi gestion: no pude leer 'Equivalencias' de %s (%s)", etiqueta, exc)
            continue
        cols = {_nk(c): c for c in d.columns}
        c_cod = cols.get("codigo maxirest") or d.columns[0]
        c_nom = cols.get("nombre maxi") or d.columns[1]
        c_sku = cols.get("sku receta") or d.columns[2]
        for r in d[[c_cod, c_nom, c_sku]].itertuples(index=False):
            cod, nom, sku = _codigo_txt(r[0]), str(r[1]).strip() if pd.notna(r[1]) else "", _codigo_txt(r[2])
            if motivo_fila_invalida(nom, cod, exigir_item=False) or motivo_fila_invalida(sku, exigir_item=False):
                continue                                     # encabezado intermedio / subtotal / texto de sección
            if cod and cod.lower() != "nan":
                filas.append({"codigo": cod, "nombre": nom, "sku": sku, "fuente": etiqueta})
        usadas.append(etiqueta)
    if not filas:
        LOG.warning("Casi gestion/Equivalencias no disponible: el mapeo Maxirest->receta queda incompleto")
    eq = pd.DataFrame(filas, columns=["codigo", "nombre", "sku", "fuente"])
    eq = eq.sort_values("sku", key=lambda s_: s_.eq("")).drop_duplicates("codigo", keep="first")
    if extra is not None and len(extra):
        faltan = extra[~extra["codigo"].isin(eq["codigo"]) & extra["sku"].astype(bool)]
        eq = pd.concat([eq, faltan.assign(fuente="otras equivalencias (Lista de precios Obrador / Precio viejo-nuevo)")
                        [["codigo", "nombre", "sku", "fuente"]]], ignore_index=True)
    eq = eq.sort_values("sku", key=lambda s_: s_.eq("")).drop_duplicates("codigo", keep="first")
    return eq.reset_index(drop=True), (usadas[0] if usadas else "sin datos")


def cargar_historial_obrador(fuentes: list[tuple[Path, str]], mes: int, anio: int
                             ) -> tuple[pd.DataFrame | None, str]:
    """Hoja Historial_Productos de Ventas_Obrador_<Mes> <Año>: ventas diarias del Obrador por producto.
    Devuelve (todas las filas con mes/semana, etiqueta del archivo usado) o (None, '')."""
    mejor: tuple[pd.DataFrame, str] | None = None
    for path, etiqueta in _planillas(fuentes, r"ventas[_ ]obrador"):
        try:
            d = pd.read_excel(path, sheet_name="Historial_Productos").iloc[:, :7]
        except Exception:  # noqa: BLE001 - el archivo de otro mes puede no tenerla
            continue
        d.columns = ["fecha", "codigo", "producto", "unidades", "precio", "total", "dia_sem"]
        d["fecha"] = pd.to_datetime(d["fecha"], errors="coerce")
        d = d.dropna(subset=["fecha"])
        d = filas_validas(d, "producto", "codigo", None, exigir_item=False, origen="Historial_Productos", log=LOG)
        en_mes = ((d["fecha"].dt.month == mes) & (d["fecha"].dt.year == anio)).sum()
        if en_mes and (mejor is None or en_mes > mejor[2]):
            mejor = (d, etiqueta, en_mes)
    if mejor is None:
        return None, ""
    return mejor[0], mejor[1]


def construir_produccion_obrador(hist: pd.DataFrame, eq: pd.DataFrame, mes: int, anio: int) -> pd.DataFrame:
    """Ventas/despachos del Obrador del mes, una fila por (día, código) y sin duplicados.

    No se infiere ninguna cantidad. Calidad de Historial_Productos (se informa en 'Alerta', sin corregir):
    - Días con UNIDADES = 1 y PRECIO = importe del día: la cantidad real no está cargada. Se conserva el importe
      y la cantidad queda VACÍA (el precio de referencia solo se usa para detectar el patrón, nunca para estimar).
    - Un precio unitario a más de OUTLIER_PRECIO_X veces (o menos de 1/X) del de referencia es un error de carga:
      la fila queda fuera de la valorización (importe vacío) y con su alerta."""
    h = hist.copy()
    h["codigo"] = h["codigo"].map(_codigo_txt)
    for c in ("unidades", "precio", "total"):
        h[c] = pd.to_numeric(h[c], errors="coerce")
    h = h[h["producto"].notna() & ~h["codigo"].isin(["", "0"])]
    h = h[~h["producto"].map(_nk).str.startswith("art. inexistente")]
    h["pu"] = h["total"] / h["unidades"].where(h["unidades"] > 0)
    ref = h[h["unidades"] >= 2].groupby("codigo")["pu"].median()       # referencia por código (todo el historial)
    d = h[(h["fecha"].dt.month == mes) & (h["fecha"].dt.year == anio)].copy()
    d["ref"] = d["codigo"].map(ref)
    ratio = d["pu"] / d["ref"]
    sin_cant = (d["unidades"] < 2) & (ratio > 5) & d["ref"].notna() & d["total"].notna()
    atipico = ~sin_cant & d["ref"].notna() & ((ratio > OUTLIER_PRECIO_X) | (ratio < 1 / OUTLIER_PRECIO_X))
    d["cant_ef"] = d["unidades"].where(~sin_cant)
    d["total_ef"] = d["total"].where(~atipico)
    d["alerta"] = ""
    d.loc[sin_cant, "alerta"] = "cantidad no informada (la planilla carga 1 unidad por el importe del día)"
    d.loc[atipico, "alerta"] = "precio unitario atípico: excluido de la valorización"
    g = d.groupby(["fecha", "codigo", "producto"], as_index=False).agg(
        cantidad=("cant_ef", lambda x: x.sum(min_count=1)), informada=("unidades", "sum"),
        total=("total_ef", lambda x: x.sum(min_count=1)),
        alerta=("alerta", lambda x: " | ".join(dict.fromkeys(a_ for a_ in x if a_))))
    g["precio_unit"] = g["total"] / g["cantidad"].where(g["cantidad"] > 0)

    por_cod = eq.drop_duplicates("codigo").set_index("codigo")
    nombres_eq = [(c, nm, sku, _nk(partir_nombre_maxirest(nm)[1])) for c, nm, sku in
                  zip(eq["codigo"], eq["nombre"], eq["sku"]) if sku]

    def resolver(codigo: str, producto: str) -> tuple[str, str, str]:
        """(sku, nombre maestro, criterio). Primero por código; si no, por nombre truncado (prefijo único)."""
        if codigo in por_cod.index:
            fila = por_cod.loc[codigo]
            if fila["sku"]:
                return fila["sku"], partir_nombre_maxirest(fila["nombre"])[1], "código Maxirest"
        k = _nk(partir_nombre_maxirest(producto)[1])
        if len(k) >= 5:
            hits = {(sku, nm) for _, nm, sku, kn in nombres_eq if kn.startswith(k)}
            if len({h_[0] for h_ in hits}) == 1:
                sku, nm = next(iter(hits))
                return sku, partir_nombre_maxirest(nm)[1], "nombre (prefijo único)"
        nm = partir_nombre_maxirest(por_cod.loc[codigo, "nombre"])[1] if codigo in por_cod.index else ""
        return "", nm, "sin equivalencia"

    res = [resolver(c, pr) for c, pr in zip(g["codigo"], g["producto"])]
    g["sku"], g["nombre_maestro"], g["criterio"] = [r[0] for r in res], [r[1] for r in res], [r[2] for r in res]
    dias = g["fecha"].dt.weekday
    out = pd.DataFrame({
        "Fecha": g["fecha"].dt.date,
        "Semana (lunes)": (g["fecha"] - pd.to_timedelta(dias, unit="D")).dt.date,
        "Semana ISO": g["fecha"].dt.strftime("%G-W%V"),
        "Mes": g["fecha"].dt.month,
        "Código Maxirest": g["codigo"],
        "Producto (Maxirest)": g["producto"],
        "Nombre maestro": g["nombre_maestro"].replace("", pd.NA).fillna(g["producto"].map(
            lambda t: partir_nombre_maxirest(t)[1])),
        "SKU receta": g["sku"],
        "Criterio de mapeo": g["criterio"],
        "Unidad": g["producto"].map(lambda t: partir_nombre_maxirest(t)[0]),
        "Cantidad": g["cantidad"],
        "Cantidad informada": g["informada"],
        "Precio unitario": g["precio_unit"],
        "Precio total $": g["total"],
        "Alerta": g["alerta"],
    })
    return out.sort_values(["Fecha", "Precio total $"], ascending=[True, False]).reset_index(drop=True)


# --- Balance físico del Obrador (planta Congreso): ventas Obrador + CONTEO OBRADOR (+ compras facturadas) ---------
_CANON = {"KG": ("KG", 1.0), "KGS": ("KG", 1.0), "KILO": ("KG", 1.0), "KILOS": ("KG", 1.0),
          "GR": ("KG", 0.001), "GRS": ("KG", 0.001), "G": ("KG", 0.001), "GRAMO": ("KG", 0.001),
          "GRAMOS": ("KG", 0.001),
          "LT": ("LT", 1.0), "LTS": ("LT", 1.0), "L": ("LT", 1.0), "LITRO": ("LT", 1.0), "LITROS": ("LT", 1.0),
          "ML": ("LT", 0.001), "CC": ("LT", 0.001),
          "UN": ("UN", 1.0), "UND": ("UN", 1.0), "UNI": ("UN", 1.0), "UNIDAD": ("UN", 1.0), "UNIDADES": ("UN", 1.0),
          "PORC": ("UN", 1.0), "PORCION": ("UN", 1.0), "PORCIONES": ("UN", 1.0)}


def canon_unidad(nombre: Any, unidad: Any) -> tuple[str, float]:
    """(unidad canónica, factor) de una cantidad. Panificados siempre UN; hojaldre siempre KG."""
    u = "" if unidad is None or (isinstance(unidad, float) and pd.isna(unidad)) else str(unidad).strip().upper()
    base, fac = _CANON.get(u, (u or "S/U", 1.0))
    forz = unidad_forzada(nombre)
    if forz == "UN":
        return "UN", 1.0
    if forz == "KG":
        return "KG", fac if base == "KG" else 1.0
    return base, fac


def cargar_conteo_completo(fuentes: list[tuple[Path, str]]) -> tuple[pd.DataFrame, dict[str, str], str]:
    """Todos los cierres de conteo: (dia, codigo, producto, cantidad) de CONTEO OBRADOR › Control_Semanal, completado
    con Casi gestion › Historial para los (día, código) que falten, y {código: unidad de medida} de la hoja Total."""
    cols = ["dia", "codigo", "producto", "cantidad", "fuente"]
    marcos, medidas, usado = [], {}, ""
    for path, etiqueta in _planillas(fuentes, r"conteo[_ ]obrador"):
        try:
            c = pd.read_excel(path, sheet_name="Control_Semanal").iloc[:, :4]
        except Exception:  # noqa: BLE001
            continue
        c.columns = ["fecha", "producto", "codigo", "cantidad"]
        c["fuente"] = "CONTEO OBRADOR"
        marcos.append(c)
        usado = etiqueta
        try:
            t = pd.read_excel(path, sheet_name="Total", header=None).iloc[1:, :5]
            for med, cod in zip(t[2], t[3]):
                if isinstance(cod, str) and cod.strip() and isinstance(med, str):
                    medidas.setdefault(cod.strip(), med.strip())
        except Exception:  # noqa: BLE001
            pass
        break
    for path, etiqueta in _planillas(fuentes, r"casi[_ ]?gesti"):
        try:
            c = pd.read_excel(path, sheet_name="Historial").iloc[:, :4]
        except Exception:  # noqa: BLE001
            continue
        c.columns = ["fecha", "producto", "codigo", "cantidad"]
        c["fuente"] = "Casi gestion"
        marcos.append(c)
        break
    if not marcos:
        return pd.DataFrame(columns=cols), medidas, usado
    c = pd.concat(marcos, ignore_index=True)
    c["fecha"] = pd.to_datetime(c["fecha"], errors="coerce")
    c["cantidad"] = pd.to_numeric(c["cantidad"], errors="coerce")
    c = c.dropna(subset=["fecha", "cantidad"])
    c = filas_validas(c, "producto", "codigo", None, exigir_item=False, origen="Conteos Obrador", log=LOG)
    c["dia"] = c["fecha"].dt.date
    c["codigo"] = c["codigo"].map(_codigo_txt)
    c = c[(c["codigo"] != "") & (c["codigo"].str.lower() != "no encontrado")]
    c = c.drop_duplicates(["dia", "codigo"], keep="first")        # CONTEO OBRADOR manda sobre Casi gestion
    return c[cols].reset_index(drop=True), medidas, usado


def construir_balance_obrador(prod: pd.DataFrame, conteo: pd.DataFrame, medidas: dict[str, str],
                              fact: pd.DataFrame, mes: int, anio: int) -> pd.DataFrame:
    """Una fila por SKU maestro, con cantidades FÍSICAS que salen solo de las planillas:
        despachado   = Ventas_Obrador › Historial_Productos del mes (solo filas con cantidad informada)
        stock inicial/final = último conteo (CONTEO OBRADOR) al 1° del mes / al cierre del mes
        compras      = facturas del local del mes para ese código
        producción implícita = stock final − stock inicial − compras + despachado   (si da > 0)
        consumo implícito    = lo mismo con signo contrario                          (si da < 0)
    El balance solo se calcula si hay ambos conteos y ningún día del ítem tiene cantidad sin informar;
    si no, la fila queda con sus despachos y la razón en 'Observación'. Nunca se estima una cantidad."""
    inicio = date(anio, mes, 1)
    fin = (pd.Timestamp(anio, mes, 1) + pd.offsets.MonthEnd(0)).date()
    p = prod.copy()
    p["key"] = p["SKU receta"].where(p["SKU receta"] != "", "MX:" + p["Código Maxirest"].astype(str))
    nombres = p.groupby("key")["Nombre maestro"].agg(lambda x: x.value_counts().index[0]).to_dict()
    nombres_conteo = conteo.groupby("codigo")["producto"].agg(lambda x: x.value_counts().index[0]).to_dict()

    def _base(t: Any) -> str:
        """Nombre sin ruido ('Pletzalej Obrador' == 'Pletzalej'; 'Bagel' != 'Baguel': ese lo resuelve el SKU)."""
        t = re.sub(r"\(.*?\)", " ", _nk(t))
        t = re.sub(r"\b(obrador|obr|madre|ob)\b|\b\d+\s*(gr|grs|g|kg|un)?\b", " ", t)
        return re.sub(r"[^a-z]+", " ", t).strip()

    # Productos del historial sin SKU en Casi gestion: se enlazan con el código del conteo si el nombre base es único
    por_base: dict[str, set[str]] = {}
    for cod, nm in nombres_conteo.items():
        por_base.setdefault(_base(nm), set()).add(cod)
    remap = {}
    for k in [k for k in nombres if k.startswith("MX:")]:
        cands = por_base.get(_base(nombres[k]), set())
        if len(cands) == 1:
            remap[k] = next(iter(cands))
    if remap:
        p["key"] = p["key"].replace(remap)
        nombres = {remap.get(k, k): v for k, v in nombres.items()}
    for cod, nm in nombres_conteo.items():
        nombres.setdefault(cod, nm)
    claves = sorted(set(p["key"]) | set(conteo["codigo"]))
    reg_prod = {k: g.to_dict("records") for k, g in p.groupby("key")}
    cont_por = {k: g.sort_values("dia") for k, g in conteo.groupby("codigo")}
    reg_fact: dict[str, list[dict[str, Any]]] = {}
    if len(fact):
        fm = fact[fact["Mes factura"] == mes]
        for k, g in fm.groupby(fm["Código SKU"].astype(str).str.strip()):
            reg_fact[k] = g.to_dict("records")
    filas = []
    for k in claves:
        nm = nombres.get(k, k)
        regs = reg_prod.get(k, [])
        cn = cont_por.get(k)
        unidad_cont = canon_unidad(nm, medidas.get(k, ""))[0] if k in medidas else None
        unids = [canon_unidad(nm, r["Unidad"])[0] for r in regs if str(r["Unidad"]).strip()]
        unid_hist = pd.Series(unids).value_counts().index[0] if unids else ""
        unidad = unidad_cont or unid_hist or "S/U"
        obs: list[str] = []
        despachado = valor = val_q = val_v = 0.0
        dias_sin = 0
        for r in regs:
            # sin prefijo de unidad en el nombre ('Aceite de Oliv'): se carga en la unidad de inventario del ítem
            u, fac = canon_unidad(nm, r["Unidad"]) if str(r["Unidad"]).strip() else (unidad, 1.0)
            q, v = r["Cantidad"], r["Precio total $"]
            if pd.notna(v):
                valor += v
            if pd.isna(q):
                dias_sin += 1
            elif u == unidad:
                despachado += q * fac
                if pd.notna(v):
                    val_q += q * fac
                    val_v += v
            else:
                obs.append(f"despachos en {u}, no comparables con {unidad}")
        pu = val_v / val_q if val_q > 0 else float("nan")
        ini = fin_c = d_ini = d_fin = None
        if cn is not None:
            antes = cn[cn["dia"] <= inicio]
            if len(antes):
                ini, d_ini = float(antes.iloc[-1]["cantidad"]), antes.iloc[-1]["dia"]
            dentro = cn[(cn["dia"] <= fin) & (cn["dia"] > (d_ini or date.min))]
            if len(dentro):
                fin_c, d_fin = float(dentro.iloc[-1]["cantidad"]), dentro.iloc[-1]["dia"]
        compras = 0.0
        for r in reg_fact.get(k, []):
            u, fac = canon_unidad(nm, r.get("Unidad Medida"))
            q = r.get("Cantidad Inventario")
            if pd.notna(q) and u == unidad:
                compras += q * fac
            elif pd.notna(q):
                obs.append(f"compras en {u}, no comparables con {unidad}")
        impl = None
        if ini is not None and fin_c is not None and dias_sin == 0:
            impl = fin_c - ini - compras + despachado
        if ini is None or fin_c is None:
            obs.append("sin conteo inicial/final en el mes: solo despachos")
        if dias_sin:
            obs.append(f"{dias_sin} día(s) con importe sin cantidad informada: balance no calculable")
        if not (regs or compras) and ini == fin_c:
            continue                                   # solo conteo y sin movimiento: no aporta
        filas.append({
            "SKU receta": k if not k.startswith("MX:") else "", "Ítem": nm, "Unidad": unidad,
            "Stock inicial (conteo)": ini, "Fecha conteo inicial": d_ini,
            "Compras del mes (facturas)": compras,
            "Despachado (Ventas Obrador)": despachado if regs else None,
            "Días con importe sin cantidad": dias_sin,
            "Stock final / sobrante (conteo)": fin_c, "Fecha conteo final": d_fin,
            "Producción implícita": max(impl, 0.0) if impl is not None else None,
            "Consumo implícito": max(-impl, 0.0) if impl is not None else None,
            "Valor despachado $": valor if regs else None,
            "Precio unitario ponderado": pu,
            "Balance confiable": "Sí" if impl is not None else "No",
            "Observación": "; ".join(dict.fromkeys(obs)),
        })
    out = pd.DataFrame(filas)
    if out.empty:
        return out
    return out.sort_values("Valor despachado $", ascending=False, na_position="last").reset_index(drop=True)


def cargar_conteos_obrador(fuentes: list[tuple[Path, str]], mes: int, anio: int) -> pd.DataFrame:
    """Conteos semanales: 'Control_Semanal' de CONTEO OBRADOR y 'Historial' de Casi gestion, del mes."""
    filas: list[pd.DataFrame] = []
    for path, etiqueta in _planillas(fuentes, r"conteo[_ ]obrador"):
        try:
            d = pd.read_excel(path, sheet_name="Control_Semanal").iloc[:, :4]
        except Exception:  # noqa: BLE001
            continue
        d.columns = ["fecha", "producto", "codigo", "cantidad"]
        d["fuente"] = f"CONTEO OBRADOR › Control_Semanal ({etiqueta})"
        filas.append(d)
        break
    for path, etiqueta in _planillas(fuentes, r"casi[_ ]?gesti"):
        try:
            d = pd.read_excel(path, sheet_name="Historial").iloc[:, :4]
        except Exception:  # noqa: BLE001
            continue
        d.columns = ["fecha", "producto", "codigo", "cantidad"]
        d["fuente"] = f"Casi gestion › Historial ({etiqueta})"
        filas.append(d)
        break
    cols = ["Fecha", "Semana (lunes)", "Semana ISO", "Código", "Producto", "Cantidad contada", "Fuente"]
    if not filas:
        return pd.DataFrame(columns=cols)
    c = pd.concat(filas, ignore_index=True)
    c["fecha"] = pd.to_datetime(c["fecha"], errors="coerce")
    c["cantidad"] = pd.to_numeric(c["cantidad"], errors="coerce")
    c = c.dropna(subset=["fecha", "cantidad"])
    c = filas_validas(c, "producto", "codigo", None, exigir_item=False, origen="Conteos semanales", log=LOG)
    c = c[(c["fecha"].dt.month == mes) & (c["fecha"].dt.year == anio)]
    c = c.drop_duplicates(["fecha", "codigo", "fuente"])
    out = pd.DataFrame({
        "Fecha": c["fecha"].dt.date,
        "Semana (lunes)": (c["fecha"] - pd.to_timedelta(c["fecha"].dt.weekday, unit="D")).dt.date,
        "Semana ISO": c["fecha"].dt.strftime("%G-W%V"),
        "Código": c["codigo"].map(_codigo_txt), "Producto": c["producto"], "Cantidad contada": c["cantidad"],
        "Fuente": c["fuente"]})
    return out.sort_values(["Fecha", "Producto"]).reset_index(drop=True)


def agregar_hoja_df(path: Path, nombre: str, df: pd.DataFrame, anchos: dict[str, int] | None = None) -> None:
    """Agrega (o reemplaza) una hoja con `df` en un Excel ya generado."""
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    wb = load_workbook(path)
    if nombre in wb.sheetnames:
        del wb[nombre]
    ws = wb.create_sheet(nombre)
    ws.append(list(df.columns))
    for c in ws[1]:
        c.font = Font(bold=True)
    for fila in df.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) or v is pd.NA else v for v in fila])
    for i, col in enumerate(df.columns, start=1):
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = (anchos or {}).get(col, 18)
    wb.save(path)


# --------------------------------------------------------------------------- #
# Resumen de la gestión mensual (ventas, compras, consumo, stock) por local y mes
# --------------------------------------------------------------------------- #
ETIQUETAS_RESUMEN = {
    "stock inicial": "Stock inicial", "compras": "Compras", "obrador": "Obrador", "stock final": "Stock final",
    "consumo $$$": "Consumo", "venta total sin iva": "Venta total sin IVA", "venta con iva": "Venta con IVA",
    "food cost": "Food cost", "cantidad ticket total": "Tickets", "cantidad cubierto total": "Cubiertos",
    "descuentos": "Descuentos",
}


def resumen_gestion(fuentes: list[tuple[Path, str]], local: str, mes: int) -> dict[str, Any] | None:
    """Hoja 'Resumen' de una planilla de gestión mensual del local con columna del mes pedido.
    Devuelve {'archivo': ..., etiqueta: valor} o None."""
    nombre_mes = _nk(MESES_ES[mes])
    for path, etiqueta in fuentes:
        if local_en_texto(etiqueta) != local or not re.search(r"gesti|mensual", etiqueta, re.I):
            continue
        try:
            df = pd.read_excel(path, sheet_name="Resumen", header=None)
        except Exception:  # noqa: BLE001 - la planilla puede no tener hoja Resumen
            continue
        col = None
        for i in range(min(6, len(df))):
            for j, v in enumerate(df.iloc[i].tolist()):
                if isinstance(v, str) and _nk(v) == nombre_mes:
                    col = j          # el último que figura gana (las planillas arrastran meses anteriores)
        if col is None:
            continue
        datos: dict[str, Any] = {"archivo": etiqueta}
        for _, fila in df.iterrows():
            lab = _nk(fila.iloc[0]) if isinstance(fila.iloc[0], str) else ""
            if lab in ETIQUETAS_RESUMEN and pd.notna(fila.iloc[col]):
                v = pd.to_numeric(fila.iloc[col], errors="coerce")
                if pd.notna(v):
                    datos.setdefault(ETIQUETAS_RESUMEN[lab], float(v))
        # una columna de mes existente pero vacía/en cero (planilla del mes anterior) no cuenta
        if any(datos.get(k, 0) > 0 for k in ("Venta total sin IVA", "Consumo", "Compras")):
            return datos
    return None


MESES_NK = {_nk(n): m for m, n in MESES_ES.items()}
HOJA_RESUMEN = "Resumen de Gestión"


def _a_num(v: Any) -> float | None:
    n = pd.to_numeric(v, errors="coerce")
    return float(n) if pd.notna(n) else None      # '#REF!', '#DIV/0!' y vacíos quedan sin dato


def _filas_resumen(df: pd.DataFrame, fila_enc: int, cols_mes: dict[int, int], fin: int
                   ) -> tuple[list[str], dict[int, dict[str, float]]]:
    """Etiquetas (columna A) y valores por mes de las filas (fila_enc, fin)."""
    etiquetas: list[str] = []
    datos: dict[int, dict[str, float]] = {m: {} for m in cols_mes}
    seccion = ""
    nombres = [_nk(x) for x in df.iloc[fila_enc + 1:fin, 0] if isinstance(x, str)]
    conteo = pd.Series(nombres, dtype=object).value_counts()
    for r in range(fila_enc + 1, fin):
        lab = df.iat[r, 0]
        if not isinstance(lab, str) or not lab.strip():
            continue
        lab = " ".join(lab.split())
        valores = {m: _a_num(df.iat[r, j]) for m, j in cols_mes.items()}
        if all(v is None for v in valores.values()):
            seccion = lab                              # fila de título ("Venta Mostrador", "Venta Rappi")
            continue
        if conteo.get(_nk(lab), 0) > 1 and seccion:
            lab = f"{seccion} · {lab}"                 # etiquetas repetidas ('Venta bruta $', '%')
        if lab in etiquetas:
            continue
        etiquetas.append(lab)
        for m, v in valores.items():
            if v is not None:
                datos[m][lab] = v
    return etiquetas, datos


def leer_resumen(path: Path) -> tuple[list[str], dict[int, dict[str, float]]] | None:
    """Hoja 'Resumen' de una planilla de gestión: (etiquetas en orden, {mes: {etiqueta: valor}}).
    - Tabla con los meses en el encabezado (2 o más; el último que figura es el del año en curso).
    - Bloque de un solo mes (nombre del mes en una celda, valores en la columna de al lado): se usa solo si el
      mes no está en la tabla, y únicamente con las etiquetas de la tabla."""
    try:
        df = pd.read_excel(path, sheet_name="Resumen", header=None)
    except Exception:  # noqa: BLE001 - la planilla puede no tener hoja Resumen
        return None
    if df.empty:
        return None
    tabla = None
    for i in range(min(80, len(df))):
        hallados = {MESES_NK[_nk(v)]: j for j, v in enumerate(df.iloc[i].tolist())
                    if isinstance(v, str) and _nk(v) in MESES_NK}      # último gana
        if len(hallados) >= 2:
            tabla = (i, hallados)
            break
    bloque = None
    for i in range(min(10, len(df))):
        for j in range(min(3, df.shape[1])):
            v = df.iat[i, j]
            if isinstance(v, str) and _nk(v) in MESES_NK and j + 1 < df.shape[1] and (tabla is None or i < tabla[0]):
                bloque = (i, MESES_NK[_nk(v)], j + 1)
                break
        if bloque:
            break
    if tabla is None and bloque is None:
        return None
    if tabla is not None:
        etiquetas, datos = _filas_resumen(df, tabla[0], tabla[1], len(df))
        if bloque and not datos.get(bloque[1]):
            fin = tabla[0] if tabla[0] > bloque[0] else len(df)
            et_b, dat_b = _filas_resumen(df, bloque[0], {bloque[1]: bloque[2]}, fin)
            conocidas = {_nk(e): e for e in etiquetas}
            datos.setdefault(bloque[1], {}).update(
                {conocidas[_nk(e)]: v for e, v in dat_b[bloque[1]].items() if _nk(e) in conocidas})
        return (etiquetas, datos) if etiquetas else None
    fin = len(df)
    et, dat = _filas_resumen(df, bloque[0], {bloque[1]: bloque[2]}, fin)
    return (et, dat) if et else None


def resumen_matriz(fuentes: list[tuple[Path, str]], local: str, mes: int, anio: int) -> pd.DataFrame | None:
    """Cuadro financiero del local: filas = conceptos de la gestión, columnas = mes anterior y mes actual.
    Por mes se prefiere la planilla cuyo nombre lo menciona; si no, cualquiera que tenga datos de ese mes."""
    mes_ant = mes - 1 if mes > 1 else 12
    cand: list[tuple[str, list[str], dict[int, dict[str, float]]]] = []
    for path, etiqueta in fuentes:
        if local_en_texto(etiqueta) != local or not re.search(r"gesti|mensual", etiqueta, re.I):
            continue
        r = leer_resumen(path)
        if r:
            cand.append((etiqueta, *r))
    if not cand:
        return None

    def valores_de(m: int) -> tuple[dict[str, float], str]:
        con_dato = [c for c in cand if c[2].get(m)]
        if not con_dato:
            return {}, ""
        mejor = next((c for c in con_dato if m in meses_en_texto(c[0])), con_dato[0])
        return mejor[2][m], mejor[0]

    act, fuente_act = valores_de(mes)
    ant, fuente_ant = valores_de(mes_ant)
    comunes = [e for e in act if e in ant]
    if len(comunes) >= 4 and all(abs(act[e] - ant[e]) <= max(0.5, 1e-6 * abs(ant[e])) for e in comunes):
        LOG.info("  %s %s: la planilla del mes repite los valores del mes anterior (sin actualizar): se ignora",
                 local, MESES_ES[mes])
        act, fuente_act = {}, ""
    # un mes en curso sin cerrar trae ceros de fórmula: venta/consumo/food cost en 0 no son dato
    act = {e: v for e, v in act.items()
           if not (v == 0 and ant.get(e, 0) > 0 and re.search(r"venta|consumo|food cost", e, re.I))}
    if not act and not ant:
        return None
    base = next((c for c in cand if c[0] == (fuente_act or fuente_ant)), cand[0])
    etiquetas = list(base[1]) + [e for c in cand for e in c[1] if e not in base[1]]
    filas = [(e, ant.get(e), act.get(e)) for e in dict.fromkeys(etiquetas) if e in ant or e in act]
    out = pd.DataFrame(filas, columns=["Concepto", MESES_ES[mes_ant], MESES_ES[mes]])
    out["Fuente"] = " | ".join(dict.fromkeys(f for f in (fuente_act, fuente_ant) if f))
    return out


def agregar_hoja_resumen(path: Path, matriz: pd.DataFrame) -> None:
    """Agrega (o reemplaza) la hoja 'Resumen de Gestión' en un Excel ya generado, justo después de la Introducción."""
    from openpyxl import load_workbook
    from openpyxl.styles import Font
    wb = load_workbook(path)
    if HOJA_RESUMEN in wb.sheetnames:
        del wb[HOJA_RESUMEN]
    ws = wb.create_sheet(HOJA_RESUMEN, min(1, len(wb.sheetnames)))
    ws.append(list(matriz.columns))
    for c in ws[1]:
        c.font = Font(bold=True)
    for fila in matriz.itertuples(index=False):
        ws.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in fila])
    ws.column_dimensions["A"].width = 46
    for col in "BC":
        ws.column_dimensions[col].width = 20
    wb.save(path)


# --------------------------------------------------------------------------- #
# Vistas
# --------------------------------------------------------------------------- #
def resumen_mes(df: pd.DataFrame, mes: int) -> pd.DataFrame:
    """Filtra por 'Mes factura' y agrupa por GROUP_COLS (cantidad, monto y precio ponderado)."""
    df = df[df["Mes factura"] == mes].copy()
    if df.empty:
        return pd.DataFrame(columns=OUTPUT_COLS)
    df["Monto Total"] = df["Cantidad Inventario"] * df["Precio Unitario"]
    resumen = df.groupby(GROUP_COLS, as_index=False, dropna=False).agg(
        **{"Cantidad Inventario": ("Cantidad Inventario", "sum"), "Monto Total": ("Monto Total", "sum")})
    cantidad = resumen["Cantidad Inventario"].where(resumen["Cantidad Inventario"] != 0)
    resumen["Precio Ponderado"] = resumen["Monto Total"] / cantidad
    return (resumen[OUTPUT_COLS].round({"Monto Total": 2, "Precio Ponderado": 2})
            .sort_values("Monto Total", ascending=False).reset_index(drop=True))


def consumo_neto_facturas(df: pd.DataFrame, mes: int) -> pd.DataFrame:
    """Hoja de ingresos (compras del mes) con los nombres de columna que usa la app, incluido el precio."""
    cols = ["Código", "Insumo", "Rubro", "Tipo de insumo", "Clasificación", "Unidad inventario",
            "Compras del mes (facturas)", "Monto facturado $", "Precio ponderado $"]
    d = df[df["Mes factura"] == mes].copy()
    if d.empty:
        return pd.DataFrame(columns=cols)
    d["Monto"] = d["Cantidad Inventario"] * d["Precio Unitario"]
    for c in ("Código SKU", "Categoría 2", "Unidad Medida"):
        d[c] = d[c].fillna("").astype(str)
    forzada = d["Productos"].map(unidad_forzada)
    d["Unidad Medida"] = forzada.fillna(d["Unidad Medida"])
    g = d.groupby(["Código SKU", "Productos", "Categoría 2", "Categoría Gasto", "Unidad Medida"],
                  as_index=False, dropna=False).agg(cant=("Cantidad Inventario", "sum"), monto=("Monto", "sum"))
    g["Precio"] = g["monto"] / g["cant"].where(g["cant"] != 0)
    g["Clasificación"] = g["Categoría Gasto"].map(
        lambda t: "Descartable (solo facturas)" if "descartable" in _nk(t) else "Materia prima (solo facturas)")
    out = g.rename(columns={"Código SKU": "Código", "Productos": "Insumo", "Categoría 2": "Rubro",
                            "Categoría Gasto": "Tipo de insumo", "Unidad Medida": "Unidad inventario",
                            "cant": "Compras del mes (facturas)", "monto": "Monto facturado $",
                            "Precio": "Precio ponderado $"})
    return out[cols].sort_values("Monto facturado $", ascending=False).reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Salidas
# --------------------------------------------------------------------------- #
def carpeta_salida() -> Path:
    """La raíz del proyecto si admite archivos nuevos; si no, WebTest/ (la app lo lee igual)."""
    prueba = BASE_DIR / ".__escritura_prueba"
    try:
        prueba.write_text("x")
        prueba.unlink()
        destino = BASE_DIR / "Reportes_Generados"
        destino.mkdir(exist_ok=True)
        return destino
    except OSError:
        LOG.warning("La raíz no admite archivos nuevos: las salidas van a %s", RESPALDO_SALIDA)
        return RESPALDO_SALIDA


def es_informe_completo(path: Path) -> bool:
    """Un informe completo (del análisis de gestión) tiene la hoja 'Auditoría de Desvíos'."""
    try:
        return "Auditoría de Desvíos" in pd.ExcelFile(path).sheet_names
    except Exception:  # noqa: BLE001
        return False


def escribir_informe_facturas(local: str, mes: int, anio: int, df_local: pd.DataFrame, fuente: str,
                              destino: Path, gestion: dict[str, Any] | None) -> None:
    """Informe para un local sin análisis completo: facturas + resumen de gestión mensual si existe."""
    mes_df = df_local[df_local["Mes factura"] == mes] if len(df_local) else df_local
    detalle = resumen_mes(df_local, mes) if len(df_local) else pd.DataFrame(columns=OUTPUT_COLS)
    monto = (mes_df["Cantidad Inventario"] * mes_df["Precio Unitario"]) if len(mes_df) else pd.Series(dtype=float)
    desc = mes_df["Categoría Gasto"].map(lambda t: "descartable" in _nk(t)) if len(mes_df) else pd.Series(dtype=bool)
    nota = ("Sin ventas por artículo ni recetas: no hay análisis de desvíos. Se informan facturas"
            + (" y el resumen de la gestión mensual" if gestion else ""))
    filas: list[tuple[str, Any, str]] = [
        (f"ANÁLISIS DE GESTIÓN COMERCIAL - {local.upper()}  {MESES_ES[mes].upper()} {anio}", None, nota)]
    if gestion:
        g_ = gestion
        origen = f"Gestión mensual: {g_['archivo']}"
        if "Venta total sin IVA" in g_:
            filas.append(("Venta total sin IVA (gestión mensual)", g_["Venta total sin IVA"], origen))
        if "Consumo" in g_:
            filas.append(("Consumo real neto (gestión mensual: CONSUMO $$$)", g_["Consumo"], origen))
        for et in ("Compras", "Stock inicial", "Stock final", "Obrador", "Food cost", "Tickets", "Cubiertos"):
            if et in g_:
                filas.append((f"{et} (gestión mensual)", g_[et], origen))
    filas += [
        ("Compras de materia prima (facturas)", float(monto[~desc].sum()) if len(monto) else 0.0, fuente),
        ("Compras de descartables/packaging (facturas)", float(monto[desc].sum()) if len(monto) else 0.0, ""),
        ("Líneas de factura del mes", int(len(mes_df)), ""),
        ("Ítems distintos facturados", int(len(detalle)), ""),
    ]
    intro = pd.DataFrame(filas, columns=["Concepto", "Valor", "Detalle"])
    controles = pd.DataFrame([
        ("Fuente de facturas", "Origen", fuente),
        ("Gestión mensual", "Origen", gestion["archivo"] if gestion else "no hay planilla de gestión para este mes"),
        ("Alcance", "Análisis", "Facturas" + (" + resumen de gestión" if gestion else "") + " (sin ventas por artículo)"),
        ("Generado por", "Script", "consolidador.py"),
    ], columns=["Fuente", "Dato", "Valor"])
    try:
        destino.parent.mkdir(parents=True, exist_ok=True)
        with open(destino, "ab"):                        # ¿está abierto/bloqueado por Excel o la sincronización?
            pass
    except PermissionError:
        destino = destino.with_name(destino.stem + "_nuevo.xlsx")
        LOG.warning("%s está abierto/bloqueado: el informe se guarda como %s", destino.name.replace("_nuevo", ""),
                    destino.name)
    with pd.ExcelWriter(destino) as w:
        intro.to_excel(w, sheet_name="Introducción Ejecutiva", index=False)
        consumo_neto_facturas(df_local, mes).to_excel(w, sheet_name="Consumo Neto (Solo Facturas)", index=False)
        detalle.to_excel(w, sheet_name="Facturas del mes", index=False)
        controles.to_excel(w, sheet_name="Controles", index=False)


def escribir_dossier(crudo: pd.DataFrame, meses: list[int], destino: Path) -> dict[tuple[str, int], int]:
    """Hojas 'Local Mes' solo para los locales que tienen datos ese mes, y 'Cadena Mes' con el total."""
    con_datos: dict[tuple[str, int], int] = {}
    with pd.ExcelWriter(destino) as writer:
        for mes in meses:
            for local in LOCALES:
                base = crudo[(crudo["Local"] == local) & (crudo["Mes factura"] == mes)] if len(crudo) else crudo
                if base.empty:
                    LOG.info("  [-] %s %s: sin datos, no se genera la hoja", local, MESES_ES[mes])
                    continue
                vista = resumen_mes(base, mes)
                vista.to_excel(writer, sheet_name=f"{local} {MESES_ES[mes]}", index=False)
                con_datos[(local, mes)] = len(base)
                LOG.info("  [OK] %s %s: %d SKUs | $%s", local, MESES_ES[mes], len(vista),
                         f"{vista['Monto Total'].sum():,.2f}")
            cadena = resumen_mes(crudo, mes)
            if len(cadena):
                cadena.to_excel(writer, sheet_name=f"Cadena {MESES_ES[mes]}", index=False)
                LOG.info("  [OK] Cadena %s: %d SKUs | $%s", MESES_ES[mes], len(cadena),
                         f"{cadena['Monto Total'].sum():,.2f}")
    return con_datos


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--anio", type=int, default=date.today().year)
    ap.add_argument("--meses", type=int, nargs="+", help="meses a procesar (por defecto, los más recientes)")
    ap.add_argument("--ultimos", type=int, default=2, help="cuántos meses recientes tomar si no hay --meses")
    ap.add_argument("--sin-analisis", action="store_true", help="no correr analisis_gestion_comercial.py")
    ap.add_argument("--refresh", action="store_true", help="bajar de nuevo las planillas de Drive")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")

    crudo, fuentes = cargar_locales()
    meses = args.meses or meses_disponibles(crudo, args.ultimos)
    if not meses:
        LOG.error("No se obtuvieron facturas de ningún local. No se generó nada.")
        return 1
    LOG.info("Meses a procesar: %s (año %d)", ", ".join(MESES_ES[m] for m in meses), args.anio)
    salida = carpeta_salida()

    LOG.info("Buscando planillas de ventas y gestión (Drive + carpetas locales)...")
    fuentes_v = fuentes_ventas_drive(meses, args.anio, args.refresh) + fuentes_ventas_locales(meses, args.anio)
    indice, resumen_fuentes = construir_indice_ventas(fuentes_v)
    LOG.info("Ventas: %d filas de %d archivos/hojas (%d archivos sin ventas por artículo)",
             len(indice), len(resumen_fuentes["analizados"]), len(resumen_fuentes["sin_ventas"]))

    # Diccionario maestro Maxirest -> SKU (Casi gestion); se completa con las otras equivalencias del análisis
    extra_eq = None
    try:
        import analisis_gestion_comercial as _agc
        _e = _agc.cargar_equivalencias_codigos(DATA_DIR)
        extra_eq = _e[["codigo", "nombre", "sku"]].astype(str)
    except Exception as exc:  # noqa: BLE001 - sin las dependencias del análisis se sigue solo con Casi gestion
        LOG.info("Equivalencias complementarias no disponibles (%s)", exc)
    equiv_maxirest, fuente_eq = cargar_equivalencias_maxirest(fuentes_v, extra_eq)
    LOG.info("Equivalencias Maxirest->receta: %d códigos (Casi gestion: %s)", len(equiv_maxirest), fuente_eq)

    LOG.info("Armando Dossier...")
    con_datos = escribir_dossier(crudo, meses, salida / DOSSIER)

    # Análisis completo (ventas + inventario + recetas) donde hay fuentes; el resto cae a informe de facturas
    completos: set[tuple[str, int]] = set()
    resultados: list = []
    agc = None
    if not args.sin_analisis:
        try:
            import analisis_gestion_comercial as agc
        except ImportError as exc:
            LOG.warning("No se pudo importar analisis_gestion_comercial (%s): solo informes de facturas", exc)
        if agc is not None:
            por_nombre = {v["nombre"]: k for k, v in agc.LOCALES.items()}
            soportados = [por_nombre[l] for l in LOCALES if l in por_nombre]
            for mes in meses:
                try:
                    agc.descargar(args.refresh, args.anio, mes)
                except Exception as exc:  # noqa: BLE001 - sin red/credenciales: se usa lo cacheado
                    LOG.warning("Descarga de Drive omitida (%s: %s)", type(exc).__name__, exc)
                ok, _fallidos = agc.generar(args.anio, mes, soportados, salida, indice)
                resultados += ok
                completos |= {(r["nombre_local"], mes) for r in ok}
            n = agc.escribir_huerfanos(resultados, salida / "reporte_items_huerfanos.txt", indice, resumen_fuentes)
            LOG.info("reporte_items_huerfanos.txt: %d ítems sin mapeo", n)

    LOG.info("Informes de facturas/gestión para los locales y meses sin análisis completo...")
    for (local, mes), n_lineas in con_datos.items():
        if (local, mes) in completos:
            continue
        destino = salida / f"Analisis_Gestion_Comercial_{local}_{mes:02d}_{args.anio}.xlsx"
        previos = [destino] + [p for p in BASE_DIR.glob(destino.name) if p != destino]
        if any(p.exists() and es_informe_completo(p) for p in previos):
            LOG.info("  [=] %s %s: ya existe un informe completo, no se pisa", local, MESES_ES[mes])
            continue
        gestion = resumen_gestion(fuentes_v, local, mes)
        escribir_informe_facturas(local, mes, args.anio, crudo[crudo["Local"] == local], fuentes[local],
                                  destino, gestion)
        LOG.info("  [OK] %s (%d líneas de factura%s)", destino.name, n_lineas,
                 ", con resumen de gestión mensual" if gestion else "")

    LOG.info("Resumen de gestión (mes anterior vs mes actual) en cada informe...")
    for (local, mes) in con_datos:
        destino = salida / f"Analisis_Gestion_Comercial_{local}_{mes:02d}_{args.anio}.xlsx"
        if not destino.exists():
            continue
        matriz = resumen_matriz(fuentes_v, local, mes, args.anio)
        if matriz is None or matriz.empty:
            LOG.info("  [-] %s %s: no hay planilla de gestión mensual con 'Resumen'", local, MESES_ES[mes])
            continue
        try:
            agregar_hoja_resumen(destino, matriz)
            LOG.info("  [OK] %s %s: %d conceptos (%s vs %s)", local, MESES_ES[mes], len(matriz),
                     matriz.columns[1], matriz.columns[2])
        except PermissionError:
            LOG.warning("  [!] %s está abierto: no se pudo agregar el resumen", destino.name)

    sin_equivalencia: list[pd.DataFrame] = []
    fuente_hist: dict[str, str] = {}
    LOG.info("Producción del Obrador (Historial_Productos) y conteos semanales en el informe de %s...", LOCAL_PLANTA)
    for (local, mes) in con_datos:
        if local != LOCAL_PLANTA:
            continue
        destino = salida / f"Analisis_Gestion_Comercial_{local}_{mes:02d}_{args.anio}.xlsx"
        if not destino.exists():
            continue
        hist, fuente_h = cargar_historial_obrador(fuentes_v, mes, args.anio)
        if hist is None:
            LOG.warning("  [!] %s %s: no hay Historial_Productos con ventas del mes", local, MESES_ES[mes])
            continue
        fuente_hist[f"{mes:02d}/{args.anio}"] = fuente_h.split("/")[-1]
        prod = construir_produccion_obrador(hist, equiv_maxirest, mes, args.anio)
        conteo = cargar_conteos_obrador(fuentes_v, mes, args.anio)
        conteo_all, medidas, fuente_conteo = cargar_conteo_completo(fuentes_v)
        balance = construir_balance_obrador(prod, conteo_all, medidas, crudo[crudo["Local"] == LOCAL_PLANTA],
                                            mes, args.anio)
        try:
            agregar_hoja_df(destino, "Producción Obrador", prod, {"Producto (Maxirest)": 30, "Nombre maestro": 34,
                                                                  "Criterio de mapeo": 24, "Alerta": 60})
            agregar_hoja_df(destino, "Balance Obrador", balance, {"Ítem": 34, "Observación": 70})
            agregar_hoja_df(destino, "Conteo semanal Obrador", conteo, {"Producto": 34, "Fuente": 60})
        except PermissionError:
            LOG.warning("  [!] %s está abierto: no se pudo agregar la producción", destino.name)
            continue
        pend = prod[prod["SKU receta"] == ""].groupby(["Código Maxirest", "Producto (Maxirest)"], as_index=False).agg(
            cantidad=("Cantidad", "sum"), unidad=("Unidad", "first"), importe=("Precio total $", "sum"))
        pend.insert(0, "mes", f"{mes:02d}/{args.anio}")
        sin_equivalencia.append(pend)
        sin_sku = prod.loc[prod["SKU receta"] == "", "Código Maxirest"].nunique()
        LOG.info("  [OK] %s %s: %d filas diarias, %d productos (%d sin equivalencia), $%s | conteos: %d filas | %s",
                 local, MESES_ES[mes], len(prod), prod["Código Maxirest"].nunique(), sin_sku,
                 f"{prod['Precio total $'].sum():,.0f}", len(conteo), fuente_h)
        LOG.info("      balance físico: %d SKUs (%d con balance confiable) | conteo: %s",
                 len(balance), int((balance["Balance confiable"] == "Sí").sum()) if len(balance) else 0,
                 fuente_conteo)

    if sin_equivalencia:
        # Códigos Maxirest del Historial_Productos sin SKU RECETA en 'Casi gestion / Equivalencias': el mapeo
        # manual que falta, con el archivo y la hoja donde aparecen
        todo = pd.concat(sin_equivalencia, ignore_index=True)
        lineas = ["", "#" * 78, "## CÓDIGOS MAXIREST SIN 'SKU RECETA' (Ventas_Obrador › Historial_Productos)",
                  "#" * 78,
                  "No figuran en 'Casi gestion' › Equivalencias (ni en las otras tablas de equivalencias).",
                  "Para incorporarlos: 'Codigo maxirest -> SKU RECETA'.", ""]
        for r in todo.sort_values(["mes", "importe"], ascending=[True, False]).itertuples(index=False):
            lineas.append(f"[{r.mes}] código Maxirest {r._1:<6} | {r._2:<28} | {r.cantidad:>12,.1f} {r.unidad:<5} | "
                          f"${(r.importe or 0):>14,.0f} | aparece en: {fuente_hist.get(r.mes, 'Ventas_Obrador')} › Historial_Productos")
        destino_txt = salida / "reporte_items_huerfanos.txt"
        if destino_txt.exists():
            with destino_txt.open("a", encoding="utf-8") as fh:
                fh.write("\n".join(lineas) + "\n")
            LOG.info("reporte_items_huerfanos.txt: +%d códigos Maxirest sin equivalencia", len(todo))
    LOG.info("¡Listo! Salidas en %s", salida)
    return 0


if __name__ == "__main__":
    sys.exit(main())
