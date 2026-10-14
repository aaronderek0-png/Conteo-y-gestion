"""Ingesta automática de listas de precios nuevas (Gmail -> Precios/Listas_Mail/ -> comparador).

Qué hace
  1. `descargar_adjuntos_gmail()` baja a Precios/Listas_Mail/ los adjuntos .xlsx/.xls/.csv/.pdf/.txt
     de la casilla (IMAP, GMAIL_USER + GMAIL_APP_PASSWORD en el entorno, los mismos que usa
     enviador_mails.py). No repite adjuntos: los reconoce por hash SHA-256.
  2. `leer_lista_generica()` lee cualquiera de esos archivos sin lector propio: detecta la fila de
     encabezado y las columnas de producto / precio / unidad (o, en PDF y TXT, toma el último
     número de cada renglón como precio) y devuelve los ítems ya normalizados a $/KG, $/LT o $/UN
     con el factor de empaque del nombre ('x 5kg', '4x5lt', 'cajón 16kg'). Esa normalización la
     hace el `item_fn` que inyecta comparador_precios.py (_item + parse_presentacion), así hay una
     sola lógica de unidades en todo el proyecto.
  3. Caché por hash SHA-256 en Precios/cache_listas.json: un archivo ya leído no se vuelve a parsear.

Uso:
    python Precios/ingesta_listas.py                # lista las listas nuevas / ya procesadas
    python Precios/ingesta_listas.py --gmail        # baja adjuntos nuevos antes de listar
Las listas se incorporan al comparador solas: `python comparador_precios.py [--gmail]`.
"""

from __future__ import annotations

import argparse
import email
import hashlib
import imaplib
import json
import logging
import os
import re
import sys
import unicodedata
from datetime import date, timedelta
from email.header import decode_header
from pathlib import Path
from typing import Any, Callable

import pandas as pd

LOG = logging.getLogger("ingesta_listas")

AQUI = Path(__file__).resolve().parent
CARPETA_ENTRADA = AQUI / "Listas_Mail"
CACHE_JSON = AQUI / "cache_listas.json"
EXTENSIONES = {".xlsx", ".xls", ".csv", ".pdf", ".txt"}

_HDR_PRODUCTO = ("producto", "descripcion", "articulo", "detalle", "item", "nombre", "denominacion", "concepto")
_HDR_PRECIO = ("precio", "p.lista", "p. lista", "p lista", "importe", "costo", "$", "valor", "p.unit", "p. unit")
_HDR_PRECIO_NO = ("anterior", "viejo", "total", "subtotal", "iva", "var", "%", "dto", "desc")
_HDR_UNIDAD = ("unidad", "un.", "u.m", "um", "medida", "present", "envase", "contenido", "formato", "bulto x")
_HDR_CAJA = ("bulto", "caja", "cajon", "pack", "cj")

# PDF / TXT: '[codigo] DESCRIPCION ... $ 1.234,56' (el último número del renglón es el precio)
_LINEA_PRECIO = re.compile(
    r"^\s*(?:(?P<cod>(?=[A-Z0-9\-./]*\d)[A-Z0-9][A-Z0-9\-./]{2,})\s+)?(?P<prod>[A-Za-zÁÉÍÓÚÑáéíóúñ0-9][^$]*?[A-Za-zÁÉÍÓÚÑáéíóúñ)][^$]*?)"
    r"\s+\$?\s*(?P<precio>\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?|\d+,\d{1,2}|\d+(?:\.\d{1,2})?)\s*$")
_LINEA_RUIDO = re.compile(r"\b(lista de precios|p[aá]gina|tel[eé]fono|cuit|vigencia|iva|condiciones|fecha)\b", re.I)


def _sin_acentos(t: Any) -> str:
    return unicodedata.normalize("NFKD", str(t or "")).encode("ascii", "ignore").decode().lower().strip()


def sha256_archivo(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


# --------------------------------------------------------------------------- #
# Caché por hash
# --------------------------------------------------------------------------- #
def _leer_cache() -> dict:
    try:
        return json.loads(CACHE_JSON.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _guardar_cache(cache: dict) -> None:
    CACHE_JSON.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")


def listas_pendientes(carpeta: Path = CARPETA_ENTRADA) -> list[Path]:
    """Archivos de la carpeta cuyo hash todavía no está en el caché."""
    if not carpeta.is_dir():
        return []
    cache = _leer_cache()
    return [p for p in sorted(carpeta.iterdir())
            if p.is_file() and p.suffix.lower() in EXTENSIONES and not p.name.startswith("~$")
            and sha256_archivo(p) not in cache]


# --------------------------------------------------------------------------- #
# Lectura genérica
# --------------------------------------------------------------------------- #
_PACK_NXM = re.compile(r"\b\d+\s*x\s*\d+(?:[.,]\d+)?\s*(?:kgs?|k|kilos?|grs?|gramos?|g|lts?|litros?|l|ml|cc)\b", re.I)


def _precio_por_caja(encabezado: str, nombre: str = "") -> bool:
    """El precio cubre el bulto completo si el encabezado lo dice ('Precio bulto/caja') o si el nombre
    trae '12 x 1,5 lt' sin marcar 'unitario' (así lo cotizan Alyser, Emporio, Villares...)."""
    h = _sin_acentos(encabezado)
    if any(w in h for w in _HDR_CAJA):
        return True
    if "unit" in h or re.search(r"\(\s*uni", nombre, re.I):
        return False
    return bool(_PACK_NXM.search(nombre))


def _tabla_a_items(raw: pd.DataFrame, item_fn: Callable, parse_price: Callable) -> list[dict]:
    """DataFrame sin encabezado (header=None) -> ítems. Detecta encabezado o infiere columnas."""
    raw = raw.dropna(how="all").reset_index(drop=True)
    if raw.empty:
        return []
    hdr = pcol = ucol = dcol = None
    for i in range(min(20, len(raw))):
        celdas = [_sin_acentos(c) for c in raw.iloc[i]]
        dp = next((j for j, c in enumerate(celdas) if c and any(w in c for w in _HDR_PRODUCTO)), None)
        pp = next((j for j, c in enumerate(celdas) if c and any(w in c for w in _HDR_PRECIO)
                   and not any(n in c for n in _HDR_PRECIO_NO)), None)
        if dp is not None and pp is not None and dp != pp:
            hdr, dcol, pcol = i, dp, pp
            ucol = next((j for j, c in enumerate(celdas) if j not in (dp, pp) and c
                         and any(w in c for w in _HDR_UNIDAD)), None)
            break
    if hdr is None:
        # Sin encabezado reconocible: producto = columna de texto más larga; precio = última columna numérica
        cuerpo = raw
        largo = {j: cuerpo[j].astype(str).str.len().mean() for j in cuerpo.columns
                 if cuerpo[j].map(lambda v: isinstance(v, str)).mean() > 0.6}
        numericas = [j for j in cuerpo.columns if cuerpo[j].map(lambda v: parse_price(v) is not None).mean() > 0.6]
        if not largo or not numericas:
            return []
        dcol, pcol = max(largo, key=largo.get), numericas[-1]
        hdr = -1
        encabezado_precio = ""
    else:
        encabezado_precio = str(raw.iloc[hdr, pcol])
    items = []
    for _, r in raw.iloc[hdr + 1:].iterrows():
        nombre = str(r[dcol]).strip() if pd.notna(r[dcol]) else ""
        precio = parse_price(r[pcol])
        if len(nombre) < 3 or precio is None or not re.search(r"[A-Za-z]", nombre):
            continue
        unidad = str(r[ucol]).strip() if ucol is not None and pd.notna(r[ucol]) else ""
        # Columna 'unidad' = 'KG' / 'LT' => precio por kilo/litro; otra cosa es presentación ('Caja x 12')
        u_norm = _sin_acentos(unidad).rstrip(".").upper()
        es_base = u_norm in {"KG", "KGS", "KILO", "KILOS", "LT", "L", "LTS", "LITRO", "LITROS"}
        items.append(item_fn(nombre, precio, "" if es_base else unidad, u_norm if es_base else "UN",
                             _precio_por_caja(encabezado_precio, nombre)))
    return items


def _lineas_a_items(lineas: list[str], item_fn: Callable, parse_price: Callable) -> list[dict]:
    items = []
    for ln in lineas:
        if not ln.strip() or _LINEA_RUIDO.search(ln):
            continue
        m = _LINEA_PRECIO.match(ln)
        if not m:
            continue
        prod, precio = m.group("prod").strip(), parse_price(m.group("precio"))
        if precio is None or len(prod) < 3 or not re.search(r"[A-Za-z]{3}", prod):
            continue
        items.append(item_fn(prod, precio, "", "UN", _precio_por_caja("", prod)))
    return items


def _leer_crudo(path: Path, item_fn: Callable, parse_price: Callable) -> list[dict]:
    ext = path.suffix.lower()
    if ext in {".xlsx", ".xls"}:
        items: list[dict] = []
        for _, hoja in pd.read_excel(path, sheet_name=None, header=None).items():
            items += _tabla_a_items(hoja, item_fn, parse_price)
        return items
    if ext == ".csv":
        for enc in ("utf-8-sig", "latin-1"):
            try:
                raw = pd.read_csv(path, header=None, sep=None, engine="python", encoding=enc, dtype=str)
                break
            except UnicodeDecodeError:
                continue
        else:
            return []
        return _tabla_a_items(raw, item_fn, parse_price)
    if ext == ".txt":
        texto = path.read_text(encoding="utf-8", errors="ignore")
        return _lineas_a_items(texto.splitlines(), item_fn, parse_price)
    if ext == ".pdf":
        import pdfplumber
        items, lineas = [], []
        with pdfplumber.open(path) as pdf:
            for pg in pdf.pages:
                for tabla in pg.extract_tables():
                    items += _tabla_a_items(pd.DataFrame(tabla), item_fn, parse_price)
                lineas += (pg.extract_text() or "").splitlines()
        # Si las tablas no rindieron, se cae al regex por renglón (PDF sin grilla)
        return items or _lineas_a_items(lineas, item_fn, parse_price)
    raise ValueError(f"extensión no soportada: {ext}")


def leer_lista_generica(path: Path, item_fn: Callable, parse_price: Callable) -> list[dict]:
    """Lee una lista de proveedor desconocida y devuelve ítems normalizados ($/KG, $/LT, $/UN).

    `item_fn(producto, precio, presentacion, unidad_lista, precio_por_caja)` es el normalizador
    de unidades del comparador. Resultado cacheado por SHA-256 (si el archivo no cambió no se re-lee)."""
    h = sha256_archivo(path)
    cache = _leer_cache()
    if h in cache:
        LOG.info("%s: en caché (sha256 %s…), %d ítems", path.name, h[:10], len(cache[h]["items"]))
        return cache[h]["items"]
    items = _leer_crudo(path, item_fn, parse_price)
    cache[h] = {"archivo": path.name, "items": items}
    _guardar_cache(cache)
    LOG.info("%s: %d ítems leídos y cacheados", path.name, len(items))
    return items


# --------------------------------------------------------------------------- #
# Gmail (IMAP)
# --------------------------------------------------------------------------- #
def _decodificar(valor: str | None) -> str:
    if not valor:
        return ""
    return "".join(t.decode(c or "utf-8", errors="ignore") if isinstance(t, bytes) else t
                   for t, c in decode_header(valor))


def descargar_adjuntos_gmail(destino: Path = CARPETA_ENTRADA, dias: int = 14, buzon: str = "INBOX",
                             usuario: str | None = None, clave: str | None = None) -> list[Path]:
    """Baja a `destino` los adjuntos de listas de los últimos `dias` días. No pisa ni repite:
    un adjunto cuyo contenido (SHA-256) ya está en la carpeta se saltea."""
    usuario = usuario or os.environ.get("GMAIL_USER")
    clave = clave or os.environ.get("GMAIL_APP_PASSWORD")
    if not usuario or not clave:
        raise RuntimeError("Faltan GMAIL_USER / GMAIL_APP_PASSWORD en el entorno")
    destino.mkdir(parents=True, exist_ok=True)
    conocidos = {sha256_archivo(p) for p in destino.iterdir() if p.is_file()}
    nuevas: list[Path] = []
    desde = (date.today() - timedelta(days=dias)).strftime("%d-%b-%Y")
    with imaplib.IMAP4_SSL("imap.gmail.com") as imap:
        imap.login(usuario, clave)
        imap.select(buzon, readonly=True)
        _, datos = imap.search(None, f'(SINCE "{desde}")')
        for num in datos[0].split():
            _, partes = imap.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(partes[0][1])
            for parte in msg.walk():
                nombre = _decodificar(parte.get_filename())
                if not nombre or Path(nombre).suffix.lower() not in EXTENSIONES:
                    continue
                contenido = parte.get_payload(decode=True) or b""
                h = hashlib.sha256(contenido).hexdigest()
                if h in conocidos:
                    continue
                limpio = re.sub(r'[\\/:*?"<>|]', "_", nombre)
                out = destino / limpio
                if out.exists():
                    out = destino / f"{out.stem}_{h[:8]}{out.suffix}"
                out.write_bytes(contenido)
                conocidos.add(h)
                nuevas.append(out)
                LOG.info("Gmail: adjunto guardado %s (de %s)", out.name, _decodificar(msg.get("From"))[:60])
    return nuevas


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--gmail", action="store_true", help="Bajar adjuntos nuevos de Gmail antes de listar")
    ap.add_argument("--dias", type=int, default=14)
    ap.add_argument("--carpeta", default=str(CARPETA_ENTRADA))
    args = ap.parse_args()
    carpeta = Path(args.carpeta)
    if args.gmail:
        descargar_adjuntos_gmail(carpeta, args.dias)
    pend = listas_pendientes(carpeta)
    print(f"{len(pend)} lista(s) sin procesar en {carpeta}:")
    for p in pend:
        print("  -", p.name)
    return 0


if __name__ == "__main__":
    sys.exit(main())
