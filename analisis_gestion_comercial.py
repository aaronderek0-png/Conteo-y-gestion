"""Análisis de gestión comercial Arcos: consumo real vs teórico, Obrador y precios.

Qué hace (mes por defecto: septiembre 2026):
1. Descarga de Drive (service account de `credentials.json`) las planillas de 'Arcos'
   (Gestión mensual agosto/septiembre, ARCOS = Facturas/Remitos/Desperdicios), de
   'Pendientes||//||Anotaciones globales' (Ventas ARCOS, Resumen Ventas), de 'Recetas',
   'Traspasos', 'Ventas_Obrador_<Mes> <Año>', 'Lista de precios Obrador' y
   'Precio viejo//nuevo metodo.' (equivalencias de códigos Maxirest -> SKU).
   Quedan cacheadas en `Arcos_data/` (usar --refresh para bajarlas de nuevo).
2. Balance de masa por insumo, en la unidad del inventario:
       Consumo real = Stock inicial (inventario mes anterior) + Compras del mes (Facturas por fecha
                      de ingreso) + Remitos de Obrador (los que son materia prima del local)
                      + Traspasos entrantes - Traspasos salientes - Stock final.
   Los traspasos salientes (a Cabello, Congreso...) se restan: no son consumo ni desvío del local.
3. Consumo teórico: explosión de recetas (BOM, recursiva) de lo vendido. Las recetas 'Madre'
   (prefijos M, MR, MO, OBR...) se abren hasta materia prima. Dos vistas:
     - Global: se abre también todo lo que viene de Obrador.
     - Neto Arcos: la explosión se frena en cada ítem de Obrador; esos ítems se descuentan de la
       materia prima y se valúan a la lista de precios del Obrador.
4. Auditoría del desvío (real - teórico), separado en:
     a) Desperdicio justificado (hoja 'Desperdicios' de ARCOS, explotando Madres/platos a materia prima),
     b) Merma natural en frescos (verdulería, carnes, pescados; tolerancia configurable),
     c) Gasto operativo (descartables, pastillas rational, hielo...: sin receta, no es desvío de cocina),
     d) Faltante crítico (lo que queda sin justificar).
5. Benchmark POSTA: valúa el consumo real contra la lista vigente de Posta (vía comparador) y
   calcula el ahorro potencial; también contra El Criollo y el mejor precio del mercado.
6. Arma:
     - Analisis_Gestion_Comercial_Arcos_09_2026.xlsx
     - reporte_gestion_arcos.html (resumen liviano para el mail)

Uso:
    python analisis_gestion_comercial.py                # usa Arcos_data/ si ya existe
    python analisis_gestion_comercial.py --refresh      # vuelve a bajar de Drive
No envía ningún mail.
"""

from __future__ import annotations

import argparse
import html
import logging
import math
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path
from typing import Any

import pandas as pd
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from rapidfuzz import fuzz, process

from defensivo import es_logistica, es_panificado_comprado, filas_validas, motivo_fila_invalida

LOG = logging.getLogger("analisis_gestion")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "Arcos_data"
CREDS = BASE_DIR / "credentials.json"
REPORTES_DIR = BASE_DIR / "Reportes_Ahorro"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# archivo local -> (nombre exacto en Drive, id de respaldo). El nombre se busca por clave normalizada
# (igualdad y, si no hay, 'contiene': '★ Lista de precios Obrador ★').
FUENTES_FIJAS = {
    "gestion_agosto": ("Gestión mensual Arcos_agosto2026", "1zETE2Xz3bJvPNrS8hOjYDml5-ZH-gnatwiom8pgHxkc"),
    "gestion_septiembre": ("Gestión mensual Arcos_septiembre2026", "1mzZE5ZZhz3wMpm9SB3cquxC3l3SFaAdp1iUBkMiMBAY"),
    "arcos": ("ARCOS", "1frh93qB8Od5mrtAfMuObKwRW93p7kev6DNgmERqz6ho"),
    "ventas_arcos": ("Ventas ARCOS 09-26", "1dT96HqUO4hLcGKi3oN1heScv7zk7io2u62jkYYqt-64"),
    "resumen_ventas": ("Resumen Ventas", "1NnmpFQ6ZA3rnmrdVtpruUm-QvL2u-um2fXzd8sZKLa4"),
    "recetas": ("🟢 Recetas🟢 99% gas", "1u7QSi7JyCM1J6-DNKSWcha2qIyF2Ybxxp9SVqdNpyfY"),
    "traspasos": ("Traspasos", "1FbRoJwxfZ-vLeGKosv5yoJHAtW17I9M6Hljja0hTvwk"),
    "lista_precios_obrador": ("Lista de precios Obrador", "17eC0USqk-GSCBvBrzArGHZusVLCb9mBP-qkNkv8yQtU"),
    "precio_viejo_nuevo": ("Precio viejo//nuevo metodo.", "1-62kSglNNXUhzvexf-2kifYFJq9gmQ5Zf1RqC16ZrVU"),
    "maestro": ("MAESTRO PRODUCTOS", "1lOabw9U4ilRwE5r6QXQ2Jc1nqkKUK3Hi096xga0XFnI"),
    "congreso_facturas": ("CONGRESO", "1udLEQpQ29avK4c6fNnQig-w02xIV82IgLe85yVJXaOY"),
    "conteo_obrador": ("CONTEO OBRADOR", "1i0LPYWIGqzTVr2KYJXgy5SrbL676EQxEk5qZshaztrM"),
    "lafinur": ("LAFINUR", "1tv0uI4KsSEPRUecYfnwcDqpGbUDelGqWwe3Hp-fikOc"),
    "casi_gestion": ("Casi gestion", "1bnJp0suZ8pBoE3AFMnLMdzTy5frC3e3bDZ5XVbBzrLQ"),
    "cabello": ("CABELLO", "1zAxYyNB98Sb1-D0HtfokESXc6FrDS96EbFYyMmsseIE"),
}
# Sucursales. 'planta' = Congreso: en Drive es la planta del Obrador (su gestión informa VENTA OBRADOR y se
# cuenta en Producción/Pastelería/Depósito): vende a los locales, no recibe semielaborados.
LOCALES = {
    "arcos": {"nombre": "Arcos", "corto": "Arcos", "traspaso": "ARCOS", "planta": False},
    "congreso": {"nombre": "Congreso", "corto": "Cong.", "traspaso": "CONGRESO", "planta": True},
}
# Fuentes que, si no están, no frenan el análisis (se informa en 'Controles')
OPCIONALES = {"traspasos", "lista_precios_obrador", "precio_viejo_nuevo", "maestro", "congreso_facturas",
              "conteo_obrador", "lafinur", "casi_gestion", "cabello"}


# Mes al que corresponden los nombres de FUENTES_FIJAS (ventas_arcos = "Ventas ARCOS 09-26", etc.)
MES_BASE = 9
MESES_ES = {1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio", 7: "julio",
            8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"}

# Tamaños de envase habituales (ml / g / un) a los que se ajusta un factor inferido por precio
ENVASES_TIPICOS = [1, 2, 3, 4, 5, 6, 10, 12, 20, 24, 25, 30, 40, 50, 100, 125, 150, 187.5, 200, 250, 330, 350,
                   355, 375, 473, 500, 600, 650, 700, 720, 750, 900, 946, 1000, 1500, 1750, 2000, 3000, 5000]
# Líneas de factura que agrupan muchos ítems (sin detalle por insumo): se informan como compra, sin unidad
COMPRAS_GENERICAS = {"I0275": "compra genérica de verdulería (sin detalle por ítem)"}
# Un desvío de costo real/teórico fuera de este rango se considera outlier de escala
OUTLIER_MIN, OUTLIER_MAX = 0.1, 10.0
# Diferencia de precio real vs mercado fuera de [1/x, x] = presentación/escala distinta, no ahorro
PRECIO_ESCALA_MAX = 4.0
# Merma natural (limpieza / cocción) tolerada en frescos, como % del consumo teórico. Es un SUPUESTO
# operativo (los datos no traen rendimiento): ajustarlo con cocina. Rubros sin entrada = tolerancia 0.
MERMA_NATURAL = {"frutas y verduras": 0.15, "carnes": 0.12, "pescados": 0.10}
# Artículos operativos (no comestibles) aunque la planilla los tipifique como materia prima
OPERATIVO_RE = re.compile(r"rational|pastilla|\bhielo\b|bolsas?\s+(?:de\s+)?vac|ecobandeja|eco\s?bandeja|\bfilm\b|guantes?\b")
# Una cantidad de desperdicio mayor a esto x la mediana del mismo ítem se trata como error de carga
DESP_ERROR_X = 50.0
DESP_MAX_KG_LT = 100.0
# Insumos cuyo proveedor real de compra es POSTA: el delta contra su propia lista no es ahorro; se excluyen
# de 'Ahorro vs POSTA' (p. ej. PANKO BLANCO GOCHIS, comprado a Posta en bolsa de 10 KG, siempre $/KG).
EXCLUIR_AHORRO_POSTA_SKU = {"PANK"}
EXCLUIR_AHORRO_POSTA_RE = re.compile(r"panko.*goch|goch.*panko")


# --------------------------------------------------------------------------- #
# Utilidades
# --------------------------------------------------------------------------- #
def strip_accents(t: Any) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", str(t)) if not unicodedata.combining(c))


def nk(t: Any) -> str:
    """Clave de nombre: sin tildes, minúsculas, espacios simples."""
    return " ".join(strip_accents(t).lower().split()) if pd.notna(t) else ""


def num(x: Any) -> float:
    v = pd.to_numeric(x, errors="coerce")
    return float(v) if pd.notna(v) else 0.0


def pesos(x: float) -> str:
    return f"${x:,.0f}".replace(",", ".") if pd.notna(x) else "-"


def fnum(x: float, dec: int = 1) -> str:
    return f"{x:,.{dec}f}".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notna(x) else "-"


# --------------------------------------------------------------------------- #
# Descarga
# --------------------------------------------------------------------------- #
def mes_anterior(anio: int, mes: int) -> tuple[int, int]:
    return (anio, mes - 1) if mes > 1 else (anio - 1, 12)


def sufijo_mes(mes: int) -> str:
    """'' para el mes base; '_MM' para los demás (ventas_arcos_10.xlsx)."""
    return "" if mes == MES_BASE else f"_{mes:02d}"


def fuentes_mes(anio: int, mes: int) -> dict[str, tuple[str, str]]:
    """Fuentes fijas + las que dependen del mes (patrón 'Ventas_Obrador_<Mes> <Año>')."""
    f = dict(FUENTES_FIJAS)
    # Gestión mensual de Arcos del mes y del anterior (saldo inicial); ids de respaldo solo para los fijos
    for a_, m_ in (mes_anterior(anio, mes), (anio, mes)):
        f.setdefault(f"gestion_{MESES_ES[m_]}", (f"Gestión mensual Arcos_{MESES_ES[m_]}{a_}", ""))
    if mes != MES_BASE:
        f[f"ventas_arcos{sufijo_mes(mes)}"] = (f"Ventas ARCOS {mes:02d}-{anio % 100:02d}", "")
        f[f"resumen_ventas{sufijo_mes(mes)}"] = (f"Resumen Ventas {mes:02d}-{anio % 100:02d}", "")
    f[f"ventas_obrador_{MESES_ES[mes]}"] = (f"Ventas_Obrador_{MESES_ES[mes].capitalize()} {anio}", "")
    ids_cong = {"septiembre": "1uVTec0SmdOivlKzVWJIuIVd_eiLpZ5RKUn2MXsD25jU"}
    f[f"gestion_congreso_{MESES_ES[mes]}"] = (f"Gestión Congreso {MESES_ES[mes].capitalize()}",
                                              ids_cong.get(MESES_ES[mes], ""))
    return f


def descargar(refresh: bool, anio: int, mes: int, solo: list[str] | None = None) -> None:
    """Baja de Drive (solo lectura) las planillas que faltan; con `refresh` las vuelve a bajar. `solo` limita
    la descarga a esas claves de FUENTES_FIJAS (p. ej. ['arcos', 'cabello'])."""
    DATA_DIR.mkdir(exist_ok=True)
    fuentes = fuentes_mes(anio, mes)
    faltan = [k for k in fuentes if (solo is None or k in solo)
              and (refresh or not (DATA_DIR / f"{k}.xlsx").exists())]
    if not faltan:
        return
    creds = Credentials.from_service_account_file(
        str(CREDS), scopes=["https://www.googleapis.com/auth/drive.readonly"])
    ses = AuthorizedSession(creds)
    r = ses.get("https://www.googleapis.com/drive/v3/files", params={
        "q": "trashed=false and mimeType='application/vnd.google-apps.spreadsheet'", "pageSize": 200,
        "fields": "files(id,name)", "supportsAllDrives": "true", "includeItemsFromAllDrives": "true"})
    r.raise_for_status()
    por_nombre = {nk(f["name"]): f["id"] for f in r.json().get("files", [])}
    for k in faltan:
        nombre, id_resp = fuentes[k]
        fid = por_nombre.get(nk(nombre))
        if fid is None:
            cand = sorted(i for n, i in por_nombre.items() if nk(nombre) in n)
            fid = id_resp if (id_resp and len(cand) != 1) else (cand[0] if cand else id_resp)
        if not fid:
            msg = f"No encontré '{nombre}' en Drive (ni id de respaldo)"
            if k in OPCIONALES or k.startswith(("ventas_obrador_", "gestion_", "ventas_arcos_", "resumen_ventas_")):
                LOG.warning("%s: el análisis sigue sin esa fuente", msg)
                continue
            raise FileNotFoundError(msg)
        resp = ses.get(f"https://www.googleapis.com/drive/v3/files/{fid}/export", params={"mimeType": XLSX_MIME})
        resp.raise_for_status()
        (DATA_DIR / f"{k}.xlsx").write_bytes(resp.content)
        LOG.info("Descargado '%s' (%d KB)", nombre, len(resp.content) // 1024)


# --------------------------------------------------------------------------- #
# Recetas y explosión (BOM)
# --------------------------------------------------------------------------- #
class Recetario:
    def __init__(self, path: Path):
        R = pd.read_excel(path, "Recetas").iloc[:, :14].dropna(subset=["COD_RECETA", "COD"])
        R["CANTIDAD"] = pd.to_numeric(R["CANTIDAD"], errors="coerce").fillna(0.0)
        R["REND"] = pd.to_numeric(R["REND"], errors="coerce").fillna(1.0).replace(0, 1.0)
        R["UNIDAD"] = R["UNIDAD"].astype(str).str.strip()
        self.lineas: dict[str, list[tuple[str, float, str]]] = defaultdict(list)
        self.rend: dict[str, float] = {}
        self.nombre: dict[str, str] = {}
        self.alias: dict[str, str] = {}
        self.nom_de_cod: dict[str, str] = {}
        for r in R.itertuples(index=False):
            cod = str(r.COD_RECETA).strip()
            self.lineas[cod].append((str(r.COD).strip(), r.CANTIDAD, r.UNIDAD))
            self.rend.setdefault(cod, r.REND)
            self.nombre.setdefault(cod, str(r.NOMBRE_RECETA))
            self.alias.setdefault(str(r.NOM_RECETA).strip(), cod)   # las líneas citan a veces el NOM_RECETA
            self.nom_de_cod.setdefault(cod, str(r.NOM_RECETA).strip())

        ins = pd.read_excel(path, "Insumos").dropna(subset=["SKU"]).drop_duplicates("SKU")
        ins["SKU"] = ins["SKU"].astype(str).str.strip()
        self.insumos = ins.set_index("SKU")
        for c in ("P_UN",):
            self.insumos[c] = pd.to_numeric(self.insumos[c], errors="coerce").fillna(0.0)

        obr = pd.read_excel(path, "Obrador").dropna(subset=["SKU"])
        obr["SKU"] = obr["SKU"].astype(str).str.strip()
        self.obrador = obr.drop_duplicates("SKU").set_index("SKU")
        for c in ("COSTO", "P_TRANSFERENCIA"):
            self.obrador[c] = pd.to_numeric(self.obrador[c], errors="coerce").fillna(0.0)

        # Obrador SKU (PO...) -> receta de producción cuando es única (hoja 'Equivalencias')
        eq = pd.read_excel(path, "Equivalencias")
        self.receta_obrador: dict[str, str] = {}
        for o, c in eq[["OBRADOR", "CODIGO RECETA"]].dropna().itertuples(index=False):
            cods = [x.strip() for x in str(c).split("|")]
            if len(cods) == 1 and cods[0] in self.lineas:
                self.receta_obrador[str(o).strip()] = cods[0]

    def sku_obrador_de(self, cod: str) -> str | None:
        """SKU de la hoja 'Obrador' al que corresponde un código: el propio SKU, o el NOM_RECETA de la
        receta OBR/M que lo produce ('OBR14' -> 'POBABA')."""
        if cod in self.obrador.index:
            return cod
        nom = self.nom_de_cod.get(cod)
        if nom in self.obrador.index:
            return nom
        for sku, rc in self.receta_obrador.items():
            if rc == cod:
                return sku
        return None

    def receta_de(self, cod: str) -> str | None:
        if cod in self.lineas:
            return cod
        a = self.alias.get(cod)
        return a if a in self.lineas else None

    def explotar(self, cod: str, qty: float, acc: dict[str, float], frenar_obrador: bool | frozenset,
                 ruta: tuple = ()) -> None:
        """Suma en `acc` la materia prima de `qty` unidades de `cod`.
        Con frenar_obrador=True un ítem de Obrador queda como ítem (no se abre); con un conjunto de SKUs
        solo se frena en esos (los que el Obrador realmente le provee al local: el resto lo elabora Arcos
        y su materia prima es consumo propio)."""
        if len(ruta) > 14 or cod in ruta:
            acc[f"__CICLO__{cod}"] += qty
            return
        if es_panificado_comprado(cod):      # panificado comprado: insumo final, no se abre (regla del GAS)
            acc[cod] += qty
            return
        es_obr = cod in self.obrador.index
        if es_obr and (frenar_obrador is True or (isinstance(frenar_obrador, frozenset) and cod in frenar_obrador)):
            acc[cod] += qty
            return
        rec = self.receta_de(cod)
        if rec is None and es_obr:
            rec = self.receta_obrador.get(cod)
        if rec is None:
            acc[cod] += qty
            return
        rend = self.rend.get(rec, 1.0)
        for c, q, _u in self.lineas[rec]:
            self.explotar(c, qty / rend * q, acc, frenar_obrador, ruta + (cod,))

    def precio_unit(self, cod: str) -> float:
        if cod in self.obrador.index:
            return float(self.obrador.at[cod, "P_TRANSFERENCIA"])
        if cod in self.insumos.index:
            return float(self.insumos.at[cod, "P_UN"])
        return 0.0

    def nombre_de(self, cod: str) -> str:
        if cod in self.obrador.index:
            return str(self.obrador.at[cod, "NOMBRE"])
        if cod in self.insumos.index:
            return str(self.insumos.at[cod, "Producto"])
        return self.nombre.get(cod, cod)

    def unidad_de(self, cod: str) -> str:
        if cod in self.insumos.index:
            return str(self.insumos.at[cod, "UNIDAD_M"]).strip()
        return ""


# --------------------------------------------------------------------------- #
# Ventas
# --------------------------------------------------------------------------- #
def _sku_resuelto(rec: "Recetario", sku: Any) -> bool:
    """El SKU existe como receta, insumo u ítem de Obrador (si no, la venta no explica ningún consumo)."""
    s_ = str(sku).strip() if sku is not None and pd.notna(sku) else ""
    return bool(s_) and bool(rec.receta_de(s_) or s_ in rec.insumos.index or s_ in rec.obrador.index
                             or rec.receta_obrador.get(s_))


def cruzar_ventas_con_recetas(S: pd.DataFrame, rec: "Recetario", umbral: float = 92.0) -> pd.DataFrame:
    """Cruce A:B de los artículos vendidos sin SKU utilizable contra los nombres del recetario (recetas,
    Obrador, insumos). Solo acepta coincidencias exactas por nombre normalizado o difusas >= `umbral` con un
    único código ganador. Devuelve una fila por artículo cruzado (articulo, sku, via, score)."""
    nombres: dict[str, set[str]] = defaultdict(set)
    for cod, nom in rec.nombre.items():
        nombres[nk(nom)].add(cod)
    for alias, cod in rec.alias.items():
        nombres[nk(alias)].add(cod)
    for sku_, row in rec.obrador.iterrows():
        nombres[nk(row["NOMBRE"])].add(sku_)
    for sku_, row in rec.insumos.iterrows():
        nombres[nk(row["Producto"])].add(sku_)
    nombres.pop("", None)
    claves = list(nombres)
    filas = []
    for i, r in S.iterrows():
        if _sku_resuelto(rec, r["sku"]):
            continue
        k = r["k"]
        if k in nombres and len(nombres[k]) == 1:
            filas.append((i, r["articulo"], next(iter(nombres[k])), "nombre exacto", 100.0))
            continue
        top = process.extract(k, claves, scorer=fuzz.token_sort_ratio, score_cutoff=umbral, limit=3)
        if not top:
            continue
        mejores = {c for n, sc, _ in top if sc == top[0][1] for c in nombres[n]}
        if len(mejores) == 1:
            filas.append((i, r["articulo"], next(iter(mejores)), "nombre similar", float(top[0][1])))
    return pd.DataFrame(filas, columns=["idx", "articulo", "sku", "via", "score"])


def _consolidar_por_sku(S: pd.DataFrame) -> pd.DataFrame:
    """Un único renglón por SKU (suma unidades y venta); los artículos sin SKU quedan tal cual."""
    con_sku, sin_sku = S[S["sku"].notna()], S[S["sku"].isna()]
    con_sku = con_sku.groupby("sku", as_index=False).agg(
        k=("k", "first"), articulo=("articulo", "first"), unidades=("unidades", "sum"), venta=("venta", "sum"))
    return pd.concat([con_sku, sin_sku], ignore_index=True)[["k", "articulo", "unidades", "venta", "sku"]]


def completar_ventas_con_indice(S: pd.DataFrame, rec: "Recetario", equiv: pd.DataFrame,
                                indice: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Cruce A:B de los artículos vendidos sin SKU utilizable contra TODAS las demás planillas de ventas.
    Si el mismo artículo (nombre normalizado) figura en otra planilla con un código que es una receta/insumo/
    Obrador, o que la planilla de equivalencias traduce a uno, se toma ese SKU (solo si el candidato es único).
    Después se consolida por SKU: ninguna fila queda duplicada. Devuelve (ventas, cruces aplicados)."""
    vacio = pd.DataFrame(columns=["articulo", "sku", "via", "score"])
    if indice is None or not len(indice) or "codigo" not in indice:
        return S, vacio
    sku_de_codigo: dict[str, set[str]] = defaultdict(set)
    for r in equiv[equiv["sku"] != ""].itertuples(index=False):
        if r.codigo:
            sku_de_codigo[r.codigo].add(r.sku)
    con_cod = indice[indice["codigo"].notna() & (indice["codigo"].astype(str).str.strip() != "")]
    por_k: dict[str, list[tuple[str, str]]] = defaultdict(list)

    def _sin_prefijo_k(k: str) -> str:
        return re.sub(r"^(?:gr|g|kg|un|u|und|porc|lt|ml|cc)\.\s*", "", k).strip()

    for r in con_cod.itertuples(index=False):
        por_k[_sin_prefijo_k(r.k)].append((_cod(r.codigo), r.archivo))
    filas, S = [], S.copy()
    for i, r in S.iterrows():
        if _sku_resuelto(rec, r["sku"]):
            continue
        cand: dict[str, str] = {}
        for cod, archivo in por_k.get(_sin_prefijo_k(r["k"]), []):
            if cod.upper() in {"", "SIN SKU", "0", "NAN"}:
                continue
            for c in {cod} | sku_de_codigo.get(cod, set()):
                if _sku_resuelto(rec, c):
                    cand.setdefault(c, archivo)
        if len(cand) == 1:
            sku, archivo = next(iter(cand.items()))
            S.at[i, "sku"] = sku
            filas.append((r["articulo"], sku, f"código en {archivo}", 100.0))
    if filas:
        S = _consolidar_por_sku(S)
    return S, pd.DataFrame(filas, columns=["articulo", "sku", "via", "score"])


def cargar_ventas(path: Path, path_resumen: Path, rec: "Recetario | None" = None
                  ) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Ventas del mes con el SKU de receta de cada artículo.
    Fuente de unidades: hoja 'Ventas' (todas las líneas). SKU, en este orden: 'Rentabilidad' de Ventas ARCOS
    (el que usa el costeo), 'Maestro', 'RENTABILIDAD_POR_PLATO' de Resumen Ventas y, si hay recetario, el cruce
    por nombre contra las recetas. Las líneas que terminan con el mismo SKU se consolidan (sin duplicar filas)."""
    S = pd.read_excel(path, "Ventas").iloc[:, :3]
    S.columns = ["articulo", "unidades", "venta"]
    S = S.dropna(subset=["articulo"]).copy()
    S["unidades"] = pd.to_numeric(S["unidades"], errors="coerce").fillna(0.0)
    S["venta"] = pd.to_numeric(S["venta"], errors="coerce").fillna(0.0)
    S["k"] = S["articulo"].map(nk)
    S = S.groupby("k", as_index=False).agg(articulo=("articulo", "first"), unidades=("unidades", "sum"),
                                           venta=("venta", "sum"))
    rent = pd.read_excel(path, "Rentabilidad").dropna(subset=["ARTÍCULO"])
    rent["k"] = rent["ARTÍCULO"].map(nk)
    sku_rent = rent[rent["SKU"].astype(str).str.upper() != "SIN SKU"].drop_duplicates("k").set_index("k")["SKU"]
    mae = pd.read_excel(path, "Maestro").dropna(subset=["nombre"])
    mae["k"] = mae["nombre"].map(nk)
    sku_mae = mae.drop_duplicates("k").set_index("k")["codigo"]
    S["sku"] = S["k"].map(sku_rent).fillna(S["k"].map(sku_mae)).astype("object")

    # cruce A:B contra 'Resumen Ventas' (misma venta, otra planilla): completa SKUs que A no trae
    ctrl: dict[str, Any] = {}
    try:
        S2 = pd.read_excel(path_resumen, "Ventas").iloc[:, :3]
        S2.columns = ["articulo", "unidades", "venta"]
        ctrl["resumen_unidades"] = float(pd.to_numeric(S2["unidades"], errors="coerce").sum())
        ctrl["resumen_venta"] = float(pd.to_numeric(S2["venta"], errors="coerce").sum())
        try:
            rp = pd.read_excel(path_resumen, "RENTABILIDAD_POR_PLATO").dropna(subset=["ARTÍCULO"])
            rp["k"] = rp["ARTÍCULO"].map(nk)
            sku_res = rp[rp["SKU"].astype(str).str.upper() != "SIN SKU"].drop_duplicates("k").set_index("k")["SKU"]
            falta = S["sku"].isna()
            S.loc[falta, "sku"] = S.loc[falta, "k"].map(sku_res)
            ctrl["sku_desde_resumen"] = int((falta & S["sku"].notna()).sum())
        except Exception as exc:  # noqa: BLE001 - hoja opcional
            LOG.info("Resumen Ventas sin hoja RENTABILIDAD_POR_PLATO utilizable: %s", exc)
    except Exception as exc:  # noqa: BLE001 - es solo un control
        LOG.warning("No pude leer 'Resumen Ventas' para el control cruzado: %s", exc)
    S["sku"] = S["sku"].map(lambda x: str(x).strip() if pd.notna(x) else None)

    ctrl["cruces_bom"] = pd.DataFrame(columns=["articulo", "sku", "via", "score"])
    if rec is not None:
        cr = cruzar_ventas_con_recetas(S, rec)
        if len(cr):
            S.loc[cr["idx"].values, "sku"] = cr["sku"].values
            ctrl["cruces_bom"] = cr.drop(columns="idx").reset_index(drop=True)
        # consolidación: un único renglón por SKU (sumar unidades y venta); los sin SKU quedan tal cual
        S = _consolidar_por_sku(S)
    return S, ctrl


# --------------------------------------------------------------------------- #
# Inventario, compras, Obrador
# --------------------------------------------------------------------------- #
def cargar_inventario(path: Path) -> pd.DataFrame:
    d = pd.read_excel(path, "Inventario", header=1).iloc[:, :19]
    d.columns = [str(c).strip() for c in d.columns]
    d = d.dropna(subset=["PRODUCTO"]).copy()
    d = filas_validas(d, "PRODUCTO", "CODIGO", "MEDIDA", origen=f"{path.name}/Inventario", log=LOG)
    d["codigo"] = d["CODIGO"].astype(str).str.strip().where(d["CODIGO"].notna(), None)
    d["qty"] = pd.to_numeric(d["FINAL"], errors="coerce").fillna(0.0)
    d["valor"] = pd.to_numeric(d["VALORIZADO"], errors="coerce").fillna(0.0)
    d["precio"] = pd.to_numeric(d["$ unitario actual"], errors="coerce")
    d["rubro_n"] = d["RUBRO"].map(lambda x: strip_accents(x).strip().lower())
    d["clave"] = d["codigo"].fillna(d["PRODUCTO"].map(lambda x: "NOMBRE:" + nk(x)))
    return d


def cargar_facturas(path: Path, anio: int, mes: int) -> pd.DataFrame:
    """Todas las facturas, con tres marcas de pertenencia al mes: `en_mes` (mes de FACTURA, la regla del
    sistema), `en_mes_pago` (columna 'Mes pago') y `en_mes_ingreso` (fecha de ingreso, criterio anterior)."""
    f = pd.read_excel(path, "Facturas")
    f.columns = [str(c).strip() for c in f.columns]
    f = filas_validas(f, "PRODUCTOS", "CODIGO", "UN. MEDIDA", exigir_item=False, origen=f"{path.name}/Facturas", log=LOG)
    f["fecha"] = pd.to_datetime(f["FECHA INGRESO"], errors="coerce")
    f["fecha_factura"] = pd.to_datetime(f["FECHA FACTURA"], errors="coerce")
    anio_f = f["fecha_factura"].dt.year == anio
    f["en_mes"] = (pd.to_numeric(f["Mes factura"], errors="coerce") == mes) & anio_f
    f["en_mes_pago"] = (pd.to_numeric(f["Mes pago"], errors="coerce") == mes) & anio_f
    f["en_mes_ingreso"] = (f["fecha"].dt.year == anio) & (f["fecha"].dt.month == mes)
    f["cat"] = f["CATEGORÍA GASTO"].map(lambda x: strip_accents(x).strip().lower())
    f["codigo"] = f["CODIGO"].astype(str).str.strip()
    f["qty"] = pd.to_numeric(f["CANTIDAD INVENTARIO"], errors="coerce").fillna(0.0)
    f["total"] = pd.to_numeric(f["TOTAL"], errors="coerce").fillna(0.0)
    return f


def cargar_remitos(path: Path, anio: int, mes: int) -> pd.DataFrame:
    r = pd.read_excel(path, "Remitos")
    r.columns = [str(c).strip() for c in r.columns]
    r = filas_validas(r, "PRODUCTO", None, None, origen=f"{path.name}/Remitos", log=LOG)
    r["fecha"] = pd.to_datetime(r["FECHA"], errors="coerce")
    r = r[(r["fecha"].dt.year == anio) & (r["fecha"].dt.month == mes)].copy()
    r["qty"] = pd.to_numeric(r["CANT. U/MEDIDA"], errors="coerce").fillna(0.0)
    r["total"] = pd.to_numeric(r["PRECIO TOTAL"], errors="coerce").fillna(0.0)
    m = r["PRODUCTO"].astype(str).str.extract(r"^\s*(?P<u>[A-Za-z]+)\.\s*(?P<n>.+)$")
    r["unidad"] = m["u"].str.upper().fillna("UN")
    r["nombre"] = m["n"].fillna(r["PRODUCTO"].astype(str)).str.strip()
    return r


_RUIDO_OBR = re.compile(r"\((?:[^)]*)\)|\bmadre\b|\bprod\b|\bobrador\b|\bunidad\b|\bgramo\b|\bporc\b|\bporcion\b"
                        r"|\bgrs?\b|\bun\b|\b\d+\s*(?:gr?s?|kg)\b|\b\d+u\b")


def _clave_obr(t: Any) -> str:
    s = _RUIDO_OBR.sub(" ", nk(t))
    return " ".join(re.sub(r"[^a-z0-9 ]", " ", s).split())


# Nombres de remito que no se parecen al del ítem de Obrador
ALIAS_OBRADOR = {"bagel sesamo": "BSESAMO", "croissant": "POCROIL", "masa lafa": "LAFFA"}
# ... y los que son materia prima del local que el Obrador provee
ALIAS_INSUMO = {"filet de trucha fresca": "TSAL", "aceite de oliva": "AOZU"}


def unidad_base_obrador(nombre: str, unidades_vistas: dict[str, float]) -> tuple[str, float | None]:
    """Unidad en la que se cuenta un ítem de Obrador y gramos por porción si la trae el nombre.
    'PORCION DE FRITAS (PORCION=300GRS)' -> ('PORC', 300); 'HUMMUS (GRAMO)' -> ('GR', None)."""
    t = nk(nombre)
    m = re.search(r"=\s*(\d+)\s*grs?", t)
    gpp = float(m.group(1)) if m else None
    if "porcion" in t or "porc." in t:
        return "PORC", gpp
    if re.search(r"\((?:gramo|1 grs?|gr)\)|\bgr\.", t):
        return "GR", gpp
    if re.search(r"\((?:unidad|un|1 un)\)|\bunidad\b", t):
        return "UN", gpp
    if unidades_vistas:
        return max(unidades_vistas, key=unidades_vistas.get), gpp
    return "UN", gpp


class MapaObrador:
    """Asocia nombres sueltos (inventario / remitos) con un SKU de la hoja 'Obrador'."""

    def __init__(self, rec: Recetario, equiv_inventario: dict[str, str],
                 nombre_a_sku: dict[str, str] | None = None):
        self.rec = rec
        self.claves = {sku: _clave_obr(rec.obrador.at[sku, "NOMBRE"]) for sku in rec.obrador.index}
        self.inv_codes = equiv_inventario                    # código inventario -> nombre
        self.exacto = nombre_a_sku or {}                     # nombre sin prefijo de unidad -> SKU (planilla de equivalencias)
        self._cache: dict[str, str | None] = {}

    def buscar(self, nombre: str, codigo: str | None = None) -> str | None:
        if codigo and codigo in self.rec.obrador.index:
            return codigo
        k = _clave_obr(nombre)
        if not k:
            return None
        if k in self._cache:
            return self._cache[k]
        # 0) la planilla de equivalencias (código Maxirest / id_receta) lo conoce por nombre exacto
        via_eq = self.rec.sku_obrador_de(self.exacto.get(_sin_prefijo(nombre), ""))
        if via_eq:
            self._cache[k] = via_eq
            return via_eq
        # 1) mismo nombre (orden de palabras aparte); 2) el nombre del remito está contenido en el ítem
        # (token_set 100: 'Latkes' -> 'LATKES 15'), eligiendo el más parecido en conjunto
        sku = ALIAS_OBRADOR.get(k)
        pares = [(sk, max(fuzz.token_sort_ratio(k, c), fuzz.ratio(k.replace(" ", ""), c.replace(" ", ""))),
                  fuzz.token_set_ratio(k, c)) for sk, c in self.claves.items()]
        top = max(pares, key=lambda t: (t[1], t[2]))
        if sku:
            pass
        elif top[1] >= 85:
            sku = top[0]
        else:
            cont = [t for t in pares if t[2] >= 100]
            if cont:
                sku = max(cont, key=lambda t: (t[1], -len(self.claves[t[0]])))[0]
        self._cache[k] = sku
        return sku


# --------------------------------------------------------------------------- #
# Unidades
# --------------------------------------------------------------------------- #
def clase_medida(m: Any) -> str:
    t = strip_accents(m).strip().lower()
    if t in {"kg", "kilo", "kgs", "kilos"}:
        return "kg"
    if t in {"lt", "lts", "litro", "litros", "l"}:
        return "lt"
    if t in {"un", "und", "unidad", "u"}:
        return "un"
    return "pack"      # botella, frasco, caja, pack, paquete, rollo, porc...


def clase_receta(u: Any) -> str:
    t = strip_accents(u).strip().lower()
    if t in {"g", "gr", "grs", "gramo", "gramos"}:
        return "g"
    if t in {"ml", "cc"}:
        return "ml"
    if t == "kg":
        return "kg"
    return "un"


def _tipico_cercano(x: float) -> float | None:
    mejor = min(ENVASES_TIPICOS, key=lambda e: abs(e - x) / e)
    return mejor if abs(mejor - x) / mejor <= 0.15 else None


def factor_inventario_a_receta(medida: Any, u_receta: Any, precio_inv: float, p_receta: float,
                               producto: str) -> tuple[float | None, str]:
    """Cuántas unidades de receta (g/ml/un) hay en 1 unidad de inventario.
    Devuelve (factor|None, nota). Ejemplo: KG con receta en g -> 1000; botella de gin
    con receta en ml -> 750 (no se cuenta el precio de la botella por trago)."""
    m, u = clase_medida(medida), clase_receta(u_receta)
    rho = precio_inv / p_receta if precio_inv and p_receta else None   # factor implícito por precio
    regla: float | None = None
    if (m, u) in {("kg", "g"), ("lt", "ml")}:
        regla = 1000.0
    elif (m, u) in {("kg", "kg"), ("lt", "ml_")}:
        regla = 1.0
    elif m == "un" and u == "un":
        regla = 1.0
    elif m in {"un", "pack"} and u == "ml":
        mt = re.search(r"(\d+(?:[.,]\d+)?)\s*(cc|ml|lts?|l)\b", strip_accents(producto).lower())
        if mt:
            v = float(mt.group(1).replace(",", "."))
            regla = v if mt.group(2) in {"cc", "ml"} else v * 1000
    if regla is not None:
        if rho is None or 0.4 <= rho / regla <= 2.5:
            return regla, "regla de unidades"
        nota = f"el precio sugiere factor ~{rho:,.0f}"
    else:
        nota = "sin regla de unidades"
    if rho:
        t = _tipico_cercano(rho)
        if t:
            return t, f"inferido por precio ({nota})"
    return None, f"no resuelto ({nota})"


# --------------------------------------------------------------------------- #
# Precios de mercado (comparador)
# --------------------------------------------------------------------------- #
def cargar_precios_mercado() -> pd.DataFrame:
    """Por Código SKU: precio de El Criollo, de Posta y ganador del comparador (Reportes_Ahorro).
    Solo precios con base $/KG, $/LT o $/UN (los '$/envase' no son comparables)."""
    cols = ["codigo", "base", "criollo", "mejor", "mejor_prov", "posta", "posta_base", "posta_match"]
    ruta = REPORTES_DIR / "Reporte_Ahorro_Resumen.xlsx"
    if not ruta.exists():
        LOG.warning("No existe %s: la matriz de precios queda sin mercado (correr comparador_precios.py)", ruta)
        return pd.DataFrame(columns=cols)
    d = pd.read_excel(ruta, "Comparativa general")
    d["codigo"] = d["Código SKU"].astype(str).str.strip()
    d["base"] = d["Base comparada"].astype(str).str.strip()
    d = d[d["base"].isin(["$/KG", "$/LT", "$/UN"])].copy()
    d["criollo"] = pd.to_numeric(d["El Criollo"], errors="coerce")
    d["posta"] = pd.to_numeric(d["Posta"], errors="coerce")
    d["mejor"] = pd.to_numeric(d["Precio Mejor Opción"], errors="coerce")
    d = d.dropna(subset=["mejor"])
    idx = d.groupby("codigo")["mejor"].idxmin()
    mejor = d.loc[idx, ["codigo", "base", "mejor", "Mejor Opción"]].rename(columns={"Mejor Opción": "mejor_prov"})
    crio = d.groupby("codigo")["criollo"].min().rename("criollo")
    out = mejor.merge(crio, on="codigo", how="left")
    dp = d.dropna(subset=["posta"])
    if len(dp):
        ip = dp.groupby("codigo")["posta"].idxmin()
        posta = dp.loc[ip, ["codigo", "posta", "base", "Match Posta"]].rename(
            columns={"base": "posta_base", "Match Posta": "posta_match"})
        out = out.merge(posta, on="codigo", how="left")
    else:
        out[["posta", "posta_base", "posta_match"]] = None
    return out


# --------------------------------------------------------------------------- #
# Traspasos, desperdicios, equivalencias de códigos, ventas y lista del Obrador
# --------------------------------------------------------------------------- #
LOCAL = "ARCOS"


def _sin_prefijo(nombre: Any) -> str:
    """'GR. Madre Salsa de Limon' / 'UN. ECOBANDEJA 900cc' -> nombre normalizado sin prefijo de unidad."""
    return nk(re.sub(r"^\s*[A-Za-z]{1,4}\.\s*", "", str(nombre)))


def _cod(v: Any) -> str:
    """Código como texto: 5144.0 -> '5144'."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def cargar_traspasos(path: Path, anio: int, mes: int, local: str = "ARCOS") -> pd.DataFrame:
    """Traspasos del mes donde `local` es origen (OUT = traspaso saliente) o destino (IN = entrante)."""
    cols = ["fecha", "origen", "destino", "codigo", "producto", "qty", "unidad", "total", "sentido"]
    if not path.exists():
        LOG.warning("No existe %s: el balance queda sin traspasos entrantes/salientes", path.name)
        return pd.DataFrame(columns=cols)
    t = pd.read_excel(path, "Traspasos")
    t.columns = [str(c).strip() for c in t.columns]
    t = filas_validas(t, "PRODUCTO", "CODIGO", "U. MEDIDA", origen=f"{path.name}/Traspasos", log=LOG)
    t["fecha"] = pd.to_datetime(t["FECHA"], errors="coerce")
    t = t[(t["fecha"].dt.year == anio) & (t["fecha"].dt.month == mes)].copy()
    # Regla del GAS (congreso_consumo_teorico): un traspaso no aprobado no mueve stock. El export real casi no
    # trae ESTADO (vacío = vigente), así que se descartan solo los estados explícitamente no vigentes.
    if "ESTADO" in t.columns:
        est = t["ESTADO"].map(nk)
        no_vigente = est.isin({"rechazado", "anulado", "cancelado", "pendiente"})
        if no_vigente.any():
            LOG.warning("Traspasos: %d filas con ESTADO no vigente descartadas", int(no_vigente.sum()))
        t = t[~no_vigente].copy()
    t["origen"] = t["LOCAL ORIGEN"].astype(str).str.strip().str.upper()
    t["destino"] = t["LOCAL DESTINO"].astype(str).str.strip().str.upper()
    t = t[(t["origen"] == local) | (t["destino"] == local)].copy()
    t["codigo"] = t["CODIGO"].map(_cod)
    t["producto"] = t["PRODUCTO"].astype(str).str.strip()
    t["qty"] = pd.to_numeric(t["CANTIDAD"], errors="coerce").fillna(0.0)
    t["unidad"] = t["U. MEDIDA"].astype(str).str.strip()
    t["total"] = pd.to_numeric(t["TOTAL"], errors="coerce").fillna(0.0)
    t["sentido"] = (t["origen"] == local).map({True: "OUT", False: "IN"})
    return t[cols].reset_index(drop=True)


def cargar_desperdicios(path: Path, anio: int, mes: int) -> pd.DataFrame:
    """Hoja 'Desperdicios' de ARCOS del mes. Marca como `error_carga` las cantidades que superan
    DESP_ERROR_X veces la mediana histórica del mismo ítem (p.ej. 2200 'KG' de latkes: eran gramos)."""
    cols = ["fecha", "codigo", "producto", "qty", "unidad", "motivo", "total", "error_carga"]
    d = pd.read_excel(path, "Desperdicios")
    d.columns = [str(c).strip() for c in d.columns]
    d = filas_validas(d, "PRODUCTO", "CÓDIGO", "Unidad de medida", origen=f"{path.name}/Desperdicios", log=LOG)
    d["fecha"] = pd.to_datetime(d["FECHA"], errors="coerce")
    d["codigo"] = d["CÓDIGO"].map(_cod)
    d["qty"] = pd.to_numeric(d["Cantidad"], errors="coerce")
    d = d[d["qty"].notna() & (d["codigo"] != "")].copy()
    mediana = d[d["qty"] > 0].groupby("codigo")["qty"].median()
    cuenta = d.groupby("codigo")["qty"].size()
    d["producto"] = d["PRODUCTO"].astype(str).str.strip()
    d["unidad"] = d["Unidad de medida"].astype(str).str.strip()
    d["motivo"] = d["MOTIVO"].astype(str).where(d["MOTIVO"].notna(), "")
    d["total"] = pd.to_numeric(d["*TOTAL"], errors="coerce").fillna(0.0)
    med = d["codigo"].map(mediana)
    d["error_carga"] = ((med.gt(0) & d["qty"].gt(DESP_ERROR_X * med) & d["codigo"].map(cuenta).gt(2))
                        | d["qty"].gt(500)
                        # un único descarte de >100 KG/LT es casi siempre gramos cargados como KG
                        | (d["qty"].gt(DESP_MAX_KG_LT) & d["unidad"].str.lower().isin({"kg", "kgs", "lt", "lts", "l"})))
    d = d[(d["fecha"].dt.year == anio) & (d["fecha"].dt.month == mes)]
    return d[cols].reset_index(drop=True)


def cargar_equivalencias_codigos(data_dir: Path) -> pd.DataFrame:
    """Planilla de equivalencias: código Maxirest (numérico) -> nombre -> SKU.
    Fuentes: 'Precio viejo//nuevo metodo.' (Sheet2 y 'Precios OCTUBRE AMBOS'), 'Lista de precios Obrador'
    (Cod_maxirest -> id_receta) y la hoja Equivalencias de Recetas (CODIGO M -> nombre)."""
    filas: list[tuple[str, str, str, str]] = []
    descartadas: list[str] = []

    def agregar(df: pd.DataFrame, ic: int, inom: int, isku: int | None, fuente: str) -> None:
        for r in df.itertuples(index=False):
            cod, nom = _cod(r[ic]), (str(r[inom]).strip() if pd.notna(r[inom]) else "")
            sku = _cod(r[isku]) if isku is not None else ""
            # encabezados intermedios ('NOMBRE'/'PRECIO'), subtotales y nombres numéricos no son equivalencias
            if motivo_fila_invalida(nom, sku or cod) or (sku and motivo_fila_invalida(sku)):
                descartadas.append(fuente)
                continue
            if (cod or sku) and nom and nom.lower() != "nan":
                filas.append((cod, nom, sku, fuente))

    # Casi gestion / Equivalencias ("Codigo maxirest", "Nombre maxi", "SKU RECETA"): primero la copia más nueva
    # bajada por el consolidador (Arcos_data/ventas_drive) y después la de Arcos_data
    cg = sorted((data_dir / "ventas_drive").glob("casi_gestion*.xlsx"),
                key=lambda q: q.stat().st_mtime, reverse=True) if (data_dir / "ventas_drive").is_dir() else []
    if (data_dir / "casi_gestion.xlsx").exists():
        cg.append(data_dir / "casi_gestion.xlsx")
    for q in cg:
        try:
            eq_cg = pd.read_excel(q, sheet_name="Equivalencias")
        except Exception as exc:  # noqa: BLE001 - copia dañada: se sigue con la siguiente
            LOG.warning("No pude leer Equivalencias de %s: %s", q.name, exc)
            continue
        eq_cg = eq_cg.dropna(how="all")
        if eq_cg.shape[1] >= 3:
            agregar(eq_cg, 0, 1, 2, "Casi gestion / Equivalencias")
    pvn = data_dir / "precio_viejo_nuevo.xlsx"
    if pvn.exists():
        agregar(pd.read_excel(pvn, "Sheet2").iloc[1:], 0, 1, 5, "Precio viejo//nuevo metodo. / Sheet2")
        agregar(pd.read_excel(pvn, "Precios OCTUBRE AMBOS", header=None).iloc[2:], 0, 1, 9,
                "Precio viejo//nuevo metodo. / Precios OCTUBRE AMBOS")
    lp = data_dir / "lista_precios_obrador.xlsx"
    if lp.exists():
        agregar(pd.read_excel(lp, "Sheet1"), 0, 1, 3, "Lista de precios Obrador")
    rc = data_dir / "recetas.xlsx"
    if rc.exists():
        eq = pd.read_excel(rc, "Equivalencias")
        agregar(eq.dropna(subset=["CODIGO M"]), eq.columns.get_loc("CODIGO M"), eq.columns.get_loc("NOMBRE M"),
                None, "Recetas / Equivalencias (CODIGO M)")
    if descartadas:
        LOG.warning("Equivalencias: %d fila(s) parásita(s) descartada(s) (encabezados/subtotales) %s", len(descartadas),
                    {f: descartadas.count(f) for f in dict.fromkeys(descartadas)})
    out = pd.DataFrame(filas, columns=["codigo", "nombre", "sku", "fuente"]).drop_duplicates()
    return out.reset_index(drop=True)


def cargar_ventas_obrador(path: Path, local: str = LOCAL) -> dict[str, Any]:
    """'Ventas_Obrador_<Mes> <Año>': lo que el Obrador le factura a cada local (Cocina + Pan), ingresos y
    egresos por traspaso. `cocina`/... son los del `local`; `total_*` la fila de totales."""
    if not path.exists():
        LOG.warning("No existe %s: no hay control contra las ventas del Obrador", path.name)
        return {"disponible": False}
    out: dict[str, Any] = {"disponible": True, "archivo": path.name, "clientes": {}}
    r = pd.read_excel(path, "Resumen", header=None)
    for fila in r.itertuples(index=False):
        if pd.notna(fila[2]) and isinstance(fila[2], (int, float)):
            vals = dict(cocina=num(fila[2]), ingresos=num(fila[3]), egresos=num(fila[4]), total=num(fila[5]))
            if pd.isna(fila[1]):
                out.update({f"total_{k}": v for k, v in vals.items()})
            else:
                out["clientes"][str(fila[1]).strip().lower()] = vals
    propio = out["clientes"].get(local.lower())
    if propio:
        out.update(propio)
    else:                                    # la planta: sus ventas son las del total
        out.update({k[6:]: v for k, v in out.items() if k.startswith("total_")})
    try:
        pdia = pd.read_excel(path, "Por día")
        pdia = pdia[pd.to_datetime(pdia["FECHA"], errors="coerce").notna()]       # el resto son subtotales
        vacio = pd.Series(dtype=float)
        out["por_dia_cocina"] = float(pd.to_numeric(pdia.get(f"{local.upper()} COCINA", vacio), errors="coerce").sum())
        out["por_dia_pan"] = float(pd.to_numeric(pdia.get(f"{local.upper()} PAN", vacio), errors="coerce").sum())
    except Exception as exc:  # noqa: BLE001 - es solo un control
        LOG.warning("No pude leer 'Por día' de %s: %s", path.name, exc)
    return out


def cargar_maestro(path: Path) -> dict[str, Any]:
    """MAESTRO PRODUCTOS: diccionario de códigos/nombres -> SKU y precio de referencia (fallback).
    `productos`: SKU -> {nombre, tipo, rubro, precio (por envase de `um_inv` unidades base), um, um_inv,
    pu (por unidad base), fecha, fuente}. `nombres`: nombre normalizado -> SKU (Productos, Presentación y
    Equivalencias)."""
    vacio: dict[str, Any] = {"productos": {}, "nombres": {}, "disponible": False}
    if not path.exists():
        LOG.warning("No existe %s: sin diccionario ni precio de respaldo del Maestro", path.name)
        return vacio
    x = pd.ExcelFile(path)
    productos: dict[str, dict[str, Any]] = {}
    nombres: dict[str, str] = {}
    d = x.parse("Productos").dropna(subset=["SKU"])
    for r in d.itertuples(index=False):
        sku = str(r[0]).strip()
        productos[sku] = {"nombre": str(r[1]), "tipo": str(r[2]) if pd.notna(r[2]) else "",
                          "rubro": str(r[3]) if pd.notna(r[3]) else "", "precio": num(r[4]),
                          "um": str(r[5]).strip() if pd.notna(r[5]) else "", "um_inv": num(r[6]), "pu": num(r[7]),
                          "fecha": r[8] if pd.notna(r[8]) else None,
                          "fuente": str(r[9]).strip() if pd.notna(r[9]) else ""}
        nombres.setdefault(nk(r[1]), sku)
    if "Presentacion" in x.sheet_names:
        pr = x.parse("Presentacion").dropna(subset=["SKU"])
        for r in pr.itertuples(index=False):
            for nom in (r[1], r[2]):
                if pd.notna(nom):
                    nombres.setdefault(nk(nom), str(r[0]).strip())
    if "Equivalencias" in x.sheet_names:
        eq = x.parse("Equivalencias").dropna(subset=["SKU_MISHI"])
        for r in eq.itertuples(index=False):
            for nom in (r[1], r[3]):
                if pd.notna(nom):
                    nombres.setdefault(nk(nom), str(r[0]).strip())
    return {"productos": productos, "nombres": nombres, "disponible": True}


def precio_maestro_inv(m: dict[str, Any] | None, medida_inv: Any) -> float:
    """Precio de referencia del Maestro llevado a la unidad del inventario (kg/lt -> $/kg|$/lt; envase -> $/envase)."""
    if not m or not (m["pu"] or m["precio"]):
        return 0.0
    cm = clase_medida(medida_inv)
    if cm in {"kg", "lt"}:
        return m["pu"] * 1000.0 if nk(m["um"]) in {"g", "gr", "grs", "ml", "cc"} else m["pu"]
    return m["precio"] or m["pu"]


def cargar_inventarios_planta(path_gestion: Path, path_conteo: Path, anio: int,
                              mes: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Inventarios de la planta (Congreso): stock final = hoja 'Inventario' de la gestión del mes (FINAL);
    stock inicial = último cierre del 'Control_Semanal' de CONTEO OBRADOR anterior al mes, valuado con los
    precios del inventario final. Devuelve (inv_inicial, inv_final) con el esquema de `cargar_inventario`."""
    d = pd.read_excel(path_gestion, "Inventario", header=1).iloc[:, :14]
    d.columns = [str(c).strip() for c in d.columns]
    d = d.dropna(subset=["PRODUCTO"]).copy()
    d = filas_validas(d, "PRODUCTO", "CODIGO", "MEDIDA", origen=f"{path_gestion.name}/Inventario", log=LOG)
    d["codigo"] = d["CODIGO"].map(_cod).replace("", None)
    d["qty"] = pd.to_numeric(d["FINAL"], errors="coerce").fillna(0.0)
    d["valor"] = pd.to_numeric(d["VALORIZADO"], errors="coerce").fillna(0.0)
    d["precio"] = pd.to_numeric(d["$ unitario"], errors="coerce")
    d["rubro_n"] = d["RUBRO"].map(lambda x: strip_accents(x).strip().lower())
    d["clave"] = d["codigo"].fillna(d["PRODUCTO"].map(lambda x: "NOMBRE:" + nk(x)))
    cs = pd.read_excel(path_conteo, "Control_Semanal").iloc[:, :4]
    cs.columns = ["fecha", "PRODUCTO", "codigo", "qty"]
    cs["fecha"] = pd.to_datetime(cs["fecha"], errors="coerce")
    previos = cs[cs["fecha"] < pd.Timestamp(anio, mes, 1)]
    if previos.empty:
        LOG.warning("CONTEO OBRADOR no tiene cierres anteriores a %d-%02d: stock inicial 0", anio, mes)
        return d.iloc[0:0].copy(), d
    ci = previos[previos["fecha"].dt.date == previos["fecha"].max().date()].copy()
    ci["codigo"] = ci["codigo"].map(_cod)
    ci["qty"] = pd.to_numeric(ci["qty"], errors="coerce").fillna(0.0)
    ci = ci.groupby("codigo", as_index=False).agg(qty=("qty", "sum"), PRODUCTO=("PRODUCTO", "first"))
    meta = d.dropna(subset=["codigo"]).drop_duplicates("codigo").set_index("codigo")
    ci["RUBRO"] = ci["codigo"].map(meta["RUBRO"])
    ci["MEDIDA"] = ci["codigo"].map(meta["MEDIDA"])
    ci["precio"] = ci["codigo"].map(meta["precio"])
    ci["PRODUCTO"] = ci["codigo"].map(meta["PRODUCTO"]).fillna(ci["PRODUCTO"])
    ci["valor"] = ci["qty"] * ci["precio"].fillna(0.0)
    ci["rubro_n"] = ci["RUBRO"].map(lambda x: strip_accents(x).strip().lower())
    ci["clave"] = ci["codigo"]
    ci["codigo"] = ci["codigo"].replace("", None)
    return ci, d


def cargar_lista_obrador(path: Path) -> pd.DataFrame:
    """Lista vigente de precios del Obrador: código Maxirest, nombre, id_receta y precio por unidad."""
    cols = ["codigo", "nombre", "id_receta", "precio"]
    if not path.exists():
        LOG.warning("No existe %s: el Obrador se valúa a P_TRANSFERENCIA de Recetas", path.name)
        return pd.DataFrame(columns=cols)
    try:
        d = pd.read_excel(path, "Mensual").iloc[:, :4]
    except ValueError:
        d = pd.read_excel(path, "Sheet1").iloc[:, [0, 1, 3, 5]]
    d.columns = cols
    d = filas_validas(d, "nombre", "codigo", None, origen=f"{path.name}/Lista de precios Obrador", log=LOG)
    d["codigo"] = d["codigo"].map(_cod)
    d["id_receta"] = d["id_receta"].map(_cod)
    d["precio"] = pd.to_numeric(d["precio"], errors="coerce")
    return d.dropna(subset=["nombre", "precio"]).reset_index(drop=True)


def qty_a_unidad_inv(qty: float, u_origen: Any, medida_inv: Any, producto: str) -> tuple[float, bool]:
    """Lleva una cantidad (KG/G/L/ML/UN/PORC) a la unidad del inventario. (cantidad, convertido_ok)."""
    u = nk(u_origen)
    o = {"kg": "kg", "kilo": "kg", "g": "g", "gr": "g", "grs": "g", "l": "lt", "lt": "lt", "lts": "lt",
         "litro": "lt", "ml": "ml", "cc": "ml"}.get(u, "un")
    m = clase_medida(medida_inv)
    if o == m or (o == "un" and m in {"un", "pack"}):
        return qty, True
    if (o == "g" and m == "kg") or (o == "ml" and m == "lt"):
        return qty / 1000.0, True
    if o in {"kg", "lt"} and m in {"un", "pack"}:        # botella / paquete: el tamaño viene en el nombre
        mt = re.search(r"(\d+(?:[.,]\d+)?)\s*(cc|ml|lts?|l|kg|grs?|gr|g)\b", strip_accents(producto).lower())
        if mt:
            v = float(mt.group(1).replace(",", "."))
            v = v / 1000.0 if mt.group(2) in {"cc", "ml", "g", "gr", "grs"} else v
            if v > 0:
                return qty / v, True
    return qty, False


def qty_a_base_obrador(qty: float, u_origen: Any, base: str, gpp: float | None) -> tuple[float, bool]:
    """Lleva una cantidad a la unidad base de un ítem de Obrador (GR, UN, PORC)."""
    u = nk(u_origen)
    o = {"kg": "kg", "kilo": "kg", "g": "g", "gr": "g", "grs": "g", "l": "kg", "lt": "kg", "lts": "kg",
         "ml": "g", "cc": "g", "porc": "porc", "porcion": "porc"}.get(u, "un")
    if base == "GR":
        return (qty * 1000.0, True) if o == "kg" else ((qty, True) if o == "g" else (qty, False))
    if base == "PORC":
        if o in {"porc", "un"}:
            return qty, True
        if gpp:
            return (qty * 1000.0 / gpp, True) if o == "kg" else (qty / gpp, True)
        return qty, False
    return (qty, True) if o in {"un", "porc"} else (qty, False)         # base UN


def es_gasto_operativo(tipo: Any, nombre: Any, rubro: Any) -> bool:
    """Insumo no comestible / sin receta de cocina: descartables, limpieza, pastillas rational, hielo..."""
    if es_logistica(nombre, rubro):      # VARIOS / FLETE: logística, fuera de MP y food cost
        return True
    t = nk(tipo)
    if t and t != "materia prima":
        return True
    if nk(rubro) in {"descartables", "compras limpieza", "libreria", "mantenimiento", "limpieza"}:
        return True
    return bool(OPERATIVO_RE.search(nk(nombre)))


# --------------------------------------------------------------------------- #
# Núcleo del análisis
# --------------------------------------------------------------------------- #
def analizar(anio: int, mes: int, local: str = "arcos", indice_ventas: pd.DataFrame | None = None,
             sin_movimientos: bool = False) -> dict[str, Any]:
    cfg = LOCALES[local]
    planta = cfg["planta"]
    rec = Recetario(DATA_DIR / "recetas.xlsx")
    maestro = cargar_maestro(DATA_DIR / "maestro.xlsx")
    if not planta:
        sfx = sufijo_mes(mes)
        _, mes_ant = mes_anterior(anio, mes)
        ventas, ctrl_ventas = cargar_ventas(DATA_DIR / f"ventas_arcos{sfx}.xlsx",
                                            DATA_DIR / f"resumen_ventas{sfx}.xlsx", rec)
        if indice_ventas is not None and len(indice_ventas):
            ventas, extra = completar_ventas_con_indice(ventas, rec, cargar_equivalencias_codigos(DATA_DIR),
                                                        indice_ventas)
            if len(extra):
                previos = ctrl_ventas.get("cruces_bom")
                ctrl_ventas["cruces_bom"] = (extra if previos is None or not len(previos)
                                             else pd.concat([previos, extra], ignore_index=True))
        inv_i = cargar_inventario(DATA_DIR / f"gestion_{MESES_ES[mes_ant]}.xlsx")
        inv_f = cargar_inventario(DATA_DIR / f"gestion_{MESES_ES[mes]}.xlsx")
        fac_all = cargar_facturas(DATA_DIR / "arcos.xlsx", anio, mes)
        rem = cargar_remitos(DATA_DIR / "arcos.xlsx", anio, mes)
        desp = cargar_desperdicios(DATA_DIR / "arcos.xlsx", anio, mes)
    else:
        ventas = pd.DataFrame(columns=["articulo", "unidades", "venta", "sku"])
        ctrl_ventas = {}
        inv_i, inv_f = cargar_inventarios_planta(DATA_DIR / f"gestion_congreso_{MESES_ES[mes]}.xlsx",
                                                 DATA_DIR / "conteo_obrador.xlsx", anio, mes)
        fac_all = cargar_facturas(DATA_DIR / "congreso_facturas.xlsx", anio, mes)
        rem = pd.DataFrame(columns=["fecha", "PRODUCTO", "qty", "total", "unidad", "nombre"])
        desp = cargar_desperdicios(DATA_DIR / "congreso_facturas.xlsx", anio, mes)
    fac = fac_all[fac_all["en_mes"]]                      # regla del sistema: mes de FACTURA
    if sin_movimientos:                                   # línea base para medir el efecto de Desperdicios/Traspasos
        desp = desp.iloc[0:0].copy()
    trasp = cargar_traspasos(DATA_DIR / "traspasos.xlsx", anio, mes, cfg["traspaso"])
    equiv = cargar_equivalencias_codigos(DATA_DIR)
    vobr = cargar_ventas_obrador(DATA_DIR / f"ventas_obrador_{MESES_ES[mes]}.xlsx",
                                 "Congreso" if planta else cfg["nombre"])
    lista_obr = cargar_lista_obrador(DATA_DIR / "lista_precios_obrador.xlsx")
    # nombre (sin prefijo de unidad) -> SKU / id_receta: Maestro, planilla de equivalencias y lista del Obrador
    nombre_a_sku: dict[str, str] = {}
    for r in equiv[equiv["sku"] != ""].itertuples(index=False):
        nombre_a_sku.setdefault(_sin_prefijo(r.nombre), r.sku)
    for r in lista_obr[lista_obr["id_receta"] != ""].itertuples(index=False):
        nombre_a_sku.setdefault(_sin_prefijo(r.nombre), r.id_receta)
    for nom, sku in maestro["nombres"].items():
        nombre_a_sku.setdefault(nom, sku)
    teor_cons: dict[str, float] = defaultdict(float)   # abre Obrador hasta materia prima
    teor_neto: dict[str, float] = defaultdict(float)   # frena en los ítems que provee el Obrador (más abajo)

    # ------------------ inventario / compras (por clave) ------------------
    equiv_inv = {c: n for c, n in zip(inv_f["codigo"], inv_f["PRODUCTO"]) if c}
    mapa = MapaObrador(rec, equiv_inv, nombre_a_sku)

    def es_obrador_inv(row: pd.Series) -> str | None:
        if row["rubro_n"] == "obrador" or (row["codigo"] and row["codigo"] in rec.obrador.index):
            return mapa.buscar(row["PRODUCTO"], row["codigo"]) or f"OBR?:{row['clave']}"
        return None

    def agrupar_inv(d: pd.DataFrame) -> pd.DataFrame:
        d = d.copy()
        d["obr_sku"] = d.apply(es_obrador_inv, axis=1)
        d["id"] = d["obr_sku"].fillna(d["clave"])
        d["origen"] = d["obr_sku"].map(lambda x: "Obrador" if isinstance(x, str) else "Arcos")
        return (d.groupby("id", as_index=False)
                .agg(qty=("qty", "sum"), valor=("valor", "sum"), producto=("PRODUCTO", "first"),
                     rubro=("RUBRO", "first"), medida=("MEDIDA", "first"), precio=("precio", "max"),
                     origen=("origen", "first"), codigo=("codigo", "first")))

    gi, gf = agrupar_inv(inv_i), agrupar_inv(inv_f)

    # compras de materia prima / descartables por código (en unidad de inventario)
    comp_mp = fac[fac["cat"].isin(["materia prima", "descartables/packaging"])]
    comp = comp_mp.groupby("codigo").agg(qty=("qty", "sum"), total=("total", "sum"),
                                         producto=("PRODUCTOS", "first")).reset_index()
    comp_total_materia = float(comp_mp["total"].sum())
    comp_fact = comp.set_index("codigo")[["qty", "total"]].copy()       # solo facturas
    tipo_fact = comp_mp.groupby("codigo")["CATEGORÍA GASTO"].first().to_dict()
    rubro_compras = (comp_mp.dropna(subset=["CATEGORÍA2"]).groupby("codigo")["CATEGORÍA2"].first().to_dict())

    # remitos de Obrador -> SKU
    remitos_total = float(rem["total"].sum())
    rem = rem.copy()
    rem["obr_sku"] = rem["nombre"].map(lambda n: mapa.buscar(n))

    # lo que el Obrador entrega pero es materia prima del local (no un ítem de su lista): va a las
    # compras de ese insumo, llevado a la unidad del inventario
    claves_inv = {r.id: _clave_obr(r.producto) for r in gf[gf["origen"] == "Arcos"].itertuples(index=False)}
    medida_inv = {r.id: r.medida for r in gf[gf["origen"] == "Arcos"].itertuples(index=False)}

    def buscar_insumo(nombre: str) -> str | None:
        k = _clave_obr(nombre)
        if k in ALIAS_INSUMO:
            return ALIAS_INSUMO[k]
        via_eq = nombre_a_sku.get(_sin_prefijo(nombre))        # planilla de equivalencias de códigos
        if via_eq and (via_eq in claves_inv or via_eq in rec.insumos.index):
            return via_eq
        pares = [(i, fuzz.token_sort_ratio(k, c), fuzz.token_set_ratio(k, c)) for i, c in claves_inv.items() if c]
        if not pares:
            return None
        top = max(pares, key=lambda t: (t[1], t[2]))
        if top[1] >= 88:
            return top[0]
        cont = [t for t in pares if t[2] >= 100 and len(k.split()) >= 2]
        return max(cont, key=lambda t: t[1])[0] if cont else None

    prov_obrador = pd.DataFrame(columns=["qty", "total"])      # materia prima del local entregada por Obrador
    sin_obr = rem["obr_sku"].isna()
    rem["insumo_id"] = None
    rem.loc[sin_obr, "insumo_id"] = rem.loc[sin_obr, "nombre"].map(buscar_insumo)
    es_insumo = rem["insumo_id"].notna()

    def a_unidad_inv(row: pd.Series) -> float:
        m = clase_medida(medida_inv.get(row["insumo_id"], ""))
        if m == "kg" and row["unidad"] in {"GR", "G", "GRS"}:
            return row["qty"] / 1000.0
        return row["qty"]

    if es_insumo.any():
        via = rem[es_insumo].copy()
        via["qty_inv"] = via.apply(a_unidad_inv, axis=1)
        extra = via.groupby("insumo_id").agg(qty=("qty_inv", "sum"), total=("total", "sum")).reset_index()
        extra = extra.rename(columns={"insumo_id": "codigo"})
        prov_obrador = extra.set_index("codigo")[["qty", "total"]]
        extra["producto"] = extra["codigo"].map(lambda i: next((r.producto for r in gf.itertuples() if r.id == i), i))
        comp = pd.concat([comp, extra], ignore_index=True).groupby("codigo", as_index=False).agg(
            qty=("qty", "sum"), total=("total", "sum"), producto=("producto", "first"))
    rem = rem[~es_insumo].copy()
    rem["id"] = rem["obr_sku"].fillna("OBR?:" + rem["nombre"].map(nk))
    unidades_obr: dict[str, tuple[str, float | None]] = {}
    for oid_, sub_ in rem.groupby("id"):
        vistas = sub_.groupby("unidad")["total"].sum().to_dict()
        nombre_o = rec.nombre_de(oid_) if oid_ in rec.obrador.index else str(sub_["nombre"].iloc[0])
        unidades_obr[oid_] = unidad_base_obrador(nombre_o, vistas)

    def a_base(row: pd.Series) -> tuple[float, str]:
        base, gpp = unidades_obr[row["id"]]
        u = "GR" if row["unidad"] in {"G", "GR", "GRS"} else row["unidad"]
        if u == base or u == "UN" and base in {"UN", "PORC"}:
            return row["qty"], row["id"]
        if base == "PORC" and u == "GR" and gpp:
            return row["qty"] / gpp, row["id"]
        if base == "GR" and u == "PORC" and gpp:
            return row["qty"] * gpp, row["id"]
        return row["qty"], f"{row['id']}|{u}"        # otra unidad sin conversión conocida: ítem aparte

    if len(rem):
        conv = rem.apply(a_base, axis=1, result_type="expand")
        rem["qty_base"], rem["id_u"] = conv[0], conv[1]
    else:
        rem["qty_base"], rem["id_u"] = pd.Series(dtype=float), pd.Series(dtype=str)
    rem["unidad_base"] = rem["id"].map(lambda i: unidades_obr[i][0])
    remi = rem.groupby("id_u").agg(qty=("qty_base", "sum"), total=("total", "sum"), nombre=("nombre", "first"),
                                   unidad=("unidad_base", "first"), id=("id", "first")).reset_index(drop=True)
    remi["unidad"] = [u if i == iu else uu for u, i, iu, uu in zip(
        remi["unidad"], remi["id"], remi["id"], remi["unidad"])]

    # ------------------ ventas -> teórico ------------------
    pasante_inv: dict[str, float] = defaultdict(float)     # materia prima que la planta revende tal cual
    planta_info: dict[str, Any] | None = None
    if planta:
        # ventas de la planta = lo que sus clientes recibieron por remito (Arcos y Lafinur; Cabello no está
        # disponible en Drive): ítem a ítem, en la unidad base del ítem de Obrador
        filas_v = []
        for cliente, ruta in (("Arcos", DATA_DIR / "arcos.xlsx"), ("Lafinur", DATA_DIR / "lafinur.xlsx")):
            if not ruta.exists():
                LOG.warning("Sin %s: las ventas itemizadas de la planta quedan incompletas", ruta.name)
                continue
            for r in cargar_remitos(ruta, anio, mes).itertuples(index=False):
                filas_v.append((cliente, r.nombre, r.unidad, r.qty, r.total))
        vr = pd.DataFrame(filas_v, columns=["cliente", "nombre", "unidad", "qty", "total"])
        vr["sku"] = vr["nombre"].map(lambda n: mapa.buscar(n))
        sin_ = vr["sku"].isna()
        vr["insumo"] = None
        vr.loc[sin_, "insumo"] = vr.loc[sin_, "nombre"].map(buscar_insumo)
        vr["q_base"] = vr["qty"]
        no_conv = []
        unid_v = {sku_: unidad_base_obrador(rec.nombre_de(sku_), sub_.groupby("unidad")["total"].sum().to_dict())
                  for sku_, sub_ in vr[vr["sku"].notna()].groupby("sku")}
        for i_, r_ in vr[vr["sku"].notna()].iterrows():
            base_, gpp_ = unid_v[r_["sku"]]
            q_, ok_ = qty_a_base_obrador(r_["qty"], r_["unidad"], base_, gpp_)
            vr.at[i_, "q_base"] = q_
            if not ok_:
                no_conv.append(i_)
        for i_, r_ in vr[vr["insumo"].notna()].iterrows():
            q_, ok_ = qty_a_unidad_inv(r_["qty"], r_["unidad"], medida_inv.get(r_["insumo"], ""), r_["nombre"])
            if ok_:
                pasante_inv[r_["insumo"]] += q_
            else:
                no_conv.append(i_)
        con_sku = vr[vr["sku"].notna() & ~vr.index.isin(no_conv)]
        ventas = (con_sku.groupby("sku", as_index=False).agg(unidades=("q_base", "sum"), venta=("total", "sum"))
                  .assign(articulo=lambda d: d["sku"].map(rec.nombre_de)))[["articulo", "unidades", "venta", "sku"]]
        falta = vr[(vr["sku"].isna() & vr["insumo"].isna()) | vr.index.isin(no_conv)]
        sin_receta_extra = falta.groupby("nombre", as_index=False).agg(unidades=("qty", "sum"), venta=("total", "sum"))
        sin_receta_extra["articulo"], sin_receta_extra["sku"] = sin_receta_extra["nombre"], None
        total_ref = vobr.get("total_cocina") if vobr.get("disponible") else None
        planta_info = {"venta_items": float(vr["total"].sum()), "venta_total": total_ref,
                       "por_cliente": vr.groupby("cliente")["total"].sum().to_dict(),
                       "venta_pasante": float(vr.loc[vr["insumo"].notna(), "total"].sum()),
                       "venta_sin_mapear": float(falta["total"].sum())}
        ctrl_ventas = {"resumen_venta": total_ref, "resumen_unidades": None}
        sin_receta = pd.concat([ventas[ventas["sku"].map(
            lambda s_: not (bool(rec.receta_de(s_)) or bool(rec.receta_obrador.get(s_))))],
            sin_receta_extra[["articulo", "sku", "unidades", "venta"]]], ignore_index=True)
    else:
        sin_receta = ventas[ventas["sku"].isna() | ~ventas["sku"].fillna("").map(
            lambda s_: bool(rec.receta_de(s_)) or s_ in rec.insumos.index or s_ in rec.obrador.index)]
    for r in ventas.itertuples(index=False):
        if isinstance(r.sku, str) and r.sku and r.unidades:
            rec.explotar(r.sku, r.unidades, teor_cons, frenar_obrador=False)

    # Ítems que el Obrador realmente le provee a Arcos en el mes: remitos, compras directas o stock del rubro
    # Obrador. El resto de los ítems de la hoja 'Obrador' los elabora Arcos con su materia prima: la explosión
    # del teórico "neto" no se frena en ellos.
    provistos = frozenset(
        {i for i in remi["id"] if not i.startswith("OBR?:")}
        | {c for c in comp["codigo"] if c in rec.obrador.index}
        | {i for i in gi.loc[(gi.origen == "Obrador") & (gi.qty > 0), "id"] if i in rec.obrador.index}
        | {i for i in gf.loc[(gf.origen == "Obrador") & (gf.qty > 0), "id"] if i in rec.obrador.index})
    for r in ventas.itertuples(index=False):
        if isinstance(r.sku, str) and r.sku and r.unidades:
            rec.explotar(r.sku, r.unidades, teor_neto, frenar_obrador=provistos)
    ciclos = sorted(k for k in set(teor_cons) | set(teor_neto) if k.startswith("__CICLO__"))
    for k in ciclos:
        teor_cons.pop(k, None)
        teor_neto.pop(k, None)

    # ------------------ traspasos entrantes / salientes y desperdicios registrados ------------------
    univ_ins = (set(gi.loc[gi.origen == "Arcos", "id"]) | set(gf.loc[gf.origen == "Arcos", "id"])
                | set(comp["codigo"]) | set(rec.insumos.index))
    prod_ins = {r.id: r.producto for r in gf.itertuples(index=False)}
    precio_lista: dict[str, float] = {}             # SKU de Obrador -> precio de la 'Lista de precios Obrador'
    for r in lista_obr.itertuples(index=False):
        sku_l = (rec.sku_obrador_de(r.id_receta) if r.id_receta else None) or mapa.buscar(r.nombre)
        if sku_l:
            precio_lista.setdefault(sku_l, float(r.precio))

    def resolver(codigo: str, nombre: str) -> tuple[str | None, str | None]:
        """('obr'|'ins'|'rec', id). Primero por código (SKU de Obrador, de insumo o de receta Madre/plato),
        después por nombre vía la planilla de equivalencias y el matching difuso."""
        c = _cod(codigo)
        o = rec.sku_obrador_de(c) if c else None
        if o:
            return "obr", o
        if c in univ_ins:
            return "ins", c
        if c and rec.receta_de(c):
            return "rec", rec.receta_de(c)
        o = mapa.buscar(nombre)
        if o:
            return "obr", o
        i = buscar_insumo(nombre)
        if i:
            return "ins", i
        return None, None

    flujo_ins: dict[tuple[str, str], float] = defaultdict(float)   # (flujo, id) en unidad de inventario
    flujo_obr: dict[tuple[str, str], float] = defaultdict(float)   # (flujo, id) en unidad base de Obrador
    flujo_rec: dict[tuple[str, str], float] = defaultdict(float)   # (flujo, id) en unidad de receta (g/ml/un)
    movimientos: list[dict[str, Any]] = []

    def volcar(flujo: str, codigo: str, nombre: str, qty: float, unidad: str) -> tuple[str, str]:
        kind, rid = resolver(codigo, nombre)
        estado = "ok"
        if kind == "obr":
            base, gpp = unidades_obr.get(rid) or unidad_base_obrador(rec.nombre_de(rid), {})
            q, ok = qty_a_base_obrador(qty, unidad, base, gpp)
            if ok:
                flujo_obr[(flujo, rid)] += q
            else:
                estado = f"sin conversión de unidad ({unidad} -> {base}): excluido del balance"
        elif kind == "ins":
            if rid in medida_inv:
                q, ok = qty_a_unidad_inv(qty, unidad, medida_inv[rid], prod_ins.get(rid, nombre))
            else:
                q, ok = qty, True          # insumo sin inventario: se toma la unidad del movimiento
            if ok:
                flujo_ins[(flujo, rid)] += q
            else:
                estado = f"sin conversión de unidad ({unidad} -> {medida_inv.get(rid, '?')}): excluido del balance"
        elif kind == "rec":                 # receta Madre / plato: se abre hasta materia prima
            u = nk(unidad)
            q_rec = qty * 1000.0 if u in {"kg", "kilo", "l", "lt", "lts", "litro"} else qty
            acc: dict[str, float] = defaultdict(float)
            rec.explotar(rid, q_rec, acc, frenar_obrador=provistos)
            for k, v in acc.items():
                if k.startswith("__CICLO__"):
                    continue
                (flujo_obr if k in rec.obrador.index else flujo_rec)[(flujo, k)] += v
            estado = "ok (Madre/plato explotado a materia prima)"
        else:
            estado = "sin resolver: no se pudo asociar a un insumo ni a una receta"
        return (f"{kind}:{rid}" if kind else "-"), estado

    if sin_movimientos:
        trasp = trasp.iloc[0:0].copy()
    for t in trasp.itertuples(index=False):
        res_, estado = volcar("t_in" if t.sentido == "IN" else "t_out", t.codigo, t.producto, t.qty, t.unidad)
        movimientos.append({"tipo": "Traspaso entrante" if t.sentido == "IN" else "Traspaso saliente",
                            "fecha": t.fecha, "contraparte": t.origen if t.sentido == "IN" else t.destino,
                            "codigo": t.codigo, "producto": t.producto, "cantidad": t.qty, "unidad": t.unidad,
                            "importe": t.total, "resuelto_como": res_, "estado": estado})
    for d in desp.itertuples(index=False):
        if d.error_carga:
            res_, estado = "-", "EXCLUIDO: cantidad fuera de rango (¿error de unidad al cargar?)"
        else:
            res_, estado = volcar("desp", d.codigo, d.producto, d.qty, d.unidad)
        movimientos.append({"tipo": "Desperdicio fichado", "fecha": d.fecha, "contraparte": d.motivo,
                            "codigo": d.codigo, "producto": d.producto, "cantidad": d.qty, "unidad": d.unidad,
                            "importe": d.total, "resuelto_como": res_, "estado": estado})
    mov = pd.DataFrame(movimientos, columns=["tipo", "fecha", "contraparte", "codigo", "producto", "cantidad",
                                             "unidad", "importe", "resuelto_como", "estado"])

    # ------------------ balance de masa: materia prima de Arcos ------------------
    ids = sorted(set(gi.loc[gi.origen == "Arcos", "id"]) | set(gf.loc[gf.origen == "Arcos", "id"])
                 | set(comp["codigo"]) | {k for k in teor_neto if k not in rec.obrador.index}
                 | {i for _, i in flujo_ins} | {i for _, i in flujo_rec} | set(pasante_inv))
    ids = [i for i in ids if i not in rec.obrador.index]   # lo que es de Obrador no va en la materia prima propia
    I0 = gi.set_index("id"); F0 = gf.set_index("id"); C0 = comp.set_index("codigo")
    tipo_de = {**{c: t for c, t in tipo_fact.items() if isinstance(t, str)},
               **rec.insumos["Tipo"].dropna().astype(str).to_dict()} if "Tipo" in rec.insumos.columns else dict(tipo_fact)
    for sku_m, m_ in maestro["productos"].items():          # el Maestro completa el tipo si falta
        if m_["tipo"]:
            tipo_de.setdefault(sku_m, m_["tipo"])
    rub_maestro = {sku_m: m_["rubro"] for sku_m, m_ in maestro["productos"].items() if m_["rubro"]}
    filas = []
    for cid in ids:
        si = I0.at[cid, "qty"] if cid in I0.index else 0.0
        sf = F0.at[cid, "qty"] if cid in F0.index else 0.0
        cq = C0.at[cid, "qty"] if cid in C0.index else 0.0
        ct = C0.at[cid, "total"] if cid in C0.index else 0.0
        cq_prov = prov_obrador.at[cid, "qty"] if cid in prov_obrador.index else 0.0
        ct_prov = prov_obrador.at[cid, "total"] if cid in prov_obrador.index else 0.0
        fila_inv = F0.loc[cid] if cid in F0.index else (I0.loc[cid] if cid in I0.index else None)
        producto = (fila_inv["producto"] if fila_inv is not None else
                    (C0.at[cid, "producto"] if cid in C0.index else rec.nombre_de(cid)))
        medida = fila_inv["medida"] if fila_inv is not None else ""
        rubro = fila_inv["rubro"] if fila_inv is not None else ""
        pr_inv = (fila_inv["precio"] if fila_inv is not None and pd.notna(fila_inv["precio"]) else None)
        pr_compra = (ct / cq) if cq else None
        mae = maestro["productos"].get(cid)
        pr_mae = precio_maestro_inv(mae, medida)
        precio = pr_compra or pr_inv or pr_mae or 0.0
        fuente_pr = ("compra del mes" if pr_compra else "inventario" if pr_inv else
                     "Maestro" if pr_mae else "sin precio")
        tu = rec.unidad_de(cid)
        p_rec = rec.precio_unit(cid)
        if cid in COMPRAS_GENERICAS:
            factor, nota = 1.0, COMPRAS_GENERICAS[cid]
        else:
            factor, nota = (factor_inventario_a_receta(medida, tu or "un", precio, p_rec, str(producto))
                            if cid in rec.insumos.index else (1.0, "ítem sin insumo"))
        # Madres/platos traspasados o descartados llegan en unidad de receta -> unidad de inventario
        def desde_receta(flujo: str) -> float:
            v = flujo_rec.get((flujo, cid), 0.0)
            return v / factor if (v and factor) else 0.0
        t_in = flujo_ins.get(("t_in", cid), 0.0) + desde_receta("t_in")
        t_out = flujo_ins.get(("t_out", cid), 0.0) + desde_receta("t_out")
        desp_q = flujo_ins.get(("desp", cid), 0.0) + desde_receta("desp")
        # Consumo real = stock inicial + compras + remitos de Obrador + traspasos entrantes - salientes - stock final
        consumo = si + cq + t_in - t_out - sf
        tn = teor_neto.get(cid, 0.0)
        teor_inv = (tn / factor if factor else None)
        if pasante_inv.get(cid):
            teor_inv = (teor_inv or 0.0) + pasante_inv[cid]
        filas.append({
            "codigo": cid, "insumo": producto, "rubro": rubro, "unidad_inv": medida,
            "stock_inicial": si, "compras": cq, "compras_facturas": cq - cq_prov, "remitos_obrador": cq_prov,
            "remitos_obrador_$": ct_prov, "traspasos_in": t_in, "traspasos_out": t_out,
            "stock_final": sf, "consumo_real": consumo,
            "desperdicio_qty": desp_q,
            "teorico_receta": tn, "unidad_receta": tu, "factor": factor, "nota_unidad": nota,
            "consumo_teorico": teor_inv,
            "precio_real": precio, "precio_receta": p_rec,
            "costo_real": consumo * precio,
            "costo_teorico": tn * p_rec,
            "tipo_insumo": tipo_de.get(cid, ""),
            "fuente_precio": fuente_pr, "precio_maestro": pr_mae,
            "maestro_fuente": (mae or {}).get("fuente", ""), "maestro_fecha": (mae or {}).get("fecha"),
        })
    neto = pd.DataFrame(filas)
    neto["desvio_qty"] = neto["consumo_real"] - neto["consumo_teorico"]
    neto["desvio_pct"] = neto.apply(
        lambda r: (r["desvio_qty"] / r["consumo_teorico"]) if r["consumo_teorico"] else None, axis=1)
    neto["desvio_costo"] = neto["costo_real"] - neto["costo_teorico"]

    def flags(r: pd.Series) -> str:
        f = []
        if r["consumo_real"] < -1e-9:
            f.append("consumo negativo (stock final > inicial + compras + traspasos)")
        sin_mov = not (r["stock_inicial"] or r["compras"] or r["stock_final"] or r["traspasos_in"] or r["traspasos_out"])
        if pd.isna(r["factor"]) and r["teorico_receta"] and not sin_mov:
            f.append("OUTLIER: unidades sin resolver (no hay forma segura de pasar la receta al inventario)")
        elif r["teorico_receta"] and r["precio_receta"] and r["precio_real"] and pd.notna(r["factor"]) and not sin_mov:
            esc = r["precio_real"] / (r["precio_receta"] * r["factor"])
            if not 1 / PRECIO_ESCALA_MAX <= esc <= PRECIO_ESCALA_MAX:
                f.append(f"OUTLIER: el precio de la receta está en otra escala (real/receta = {esc:,.2f}x)")
        if r["teorico_receta"] and r["consumo_real"] > 0 and r["costo_teorico"] and r["costo_real"]:
            ratio = r["costo_real"] / r["costo_teorico"]
            potencia = next((k for k in (1, 2, 3) if abs(abs(math.log10(ratio)) - k) < 0.03), None)
            if potencia:
                f.append(f"OUTLIER: relación real/teórico {ratio:,.1f}x, parece un error de escala (x10^{potencia})")
            elif not OUTLIER_MIN <= ratio <= OUTLIER_MAX:
                f.append(f"desvío extremo ({ratio:,.1f}x el teórico): revisar receta o consumo")
        if r["teorico_receta"] and not r["precio_receta"]:
            f.append("receta sin precio")
        if not r["teorico_receta"] and r["consumo_real"] > 0:
            f.append("consumo sin receta que lo explique")
        if r["teorico_receta"] and sin_mov:
            f.append("OUTLIER: la receta lo usa pero no tiene stock ni compras registradas por ítem")
        return "; ".join(f)

    neto["alerta"] = neto.apply(flags, axis=1)
    # Para totales: el outlier de escala o de datos se excluye del desvío por ítem (no es una merma)
    neto["sin_seguimiento"] = neto["alerta"].str.contains("no tiene stock ni compras", regex=False)
    neto["excl_escala"] = (neto["alerta"].str.contains("OUTLIER", regex=False) & ~neto["sin_seguimiento"])
    neto["outlier"] = (neto["alerta"].str.contains("OUTLIER", regex=False)
                       | neto["alerta"].str.contains("consumo negativo", regex=False))
    # rubro: el de los insumos de recetas; si falta, el de las facturas (la compra genérica de 'Verdulería')
    rub_ins = rec.insumos["Rubro / Categoría"].to_dict() if "Rubro / Categoría" in rec.insumos.columns else {}
    neto["rubro_gestion"] = [rubro_compras.get(c) or rub_ins.get(c) or rub_maestro.get(c)
                             or (r if isinstance(r, str) and r else "Sin rubro")
                             for c, r in zip(neto["codigo"], neto["rubro"])]
    neto["rubro_n"] = neto["rubro_gestion"].map(lambda x: nk(x))
    neto["es_operativo"] = [es_gasto_operativo(t, p, r) for t, p, r in
                            zip(neto["tipo_insumo"], neto["insumo"], neto["rubro_gestion"])]
    neto["es_verduleria_agregada"] = (
        (neto["rubro_n"] == "frutas y verduras") & ~neto["es_operativo"]
        & (neto["codigo"].isin(COMPRAS_GENERICAS) | neto["sin_seguimiento"]))

    # ------------------ Obrador: items consumidos por el local ------------------
    obr_ids = sorted(set(gi.loc[gi.origen == "Obrador", "id"]) | set(gf.loc[gf.origen == "Obrador", "id"])
                     | set(remi["id"]) | {k for k in teor_neto if k in rec.obrador.index}
                     | {i for _, i in flujo_obr})
    R0 = remi.drop_duplicates("id").set_index("id")
    obr_ids = sorted(set(obr_ids) | set(remi["id"]))
    filas_o = []
    for oid in obr_ids:
        si = I0.at[oid, "qty"] if oid in I0.index else 0.0
        sf = F0.at[oid, "qty"] if oid in F0.index else 0.0
        med = (F0.at[oid, "medida"] if oid in F0.index else (I0.at[oid, "medida"] if oid in I0.index else ""))
        nombre_o = rec.nombre_de(oid) if oid in rec.obrador.index else ""
        base, gpp = unidades_obr.get(oid) or unidad_base_obrador(nombre_o, {})
        cm = clase_medida(med)
        if (oid not in unidades_obr and base == "UN" and cm in {"kg", "lt"}
                and clase_receta(rec.unidad_de(oid)) in {"g", "ml"}):
            base = "GR"              # sin remitos que digan la unidad: la receta cuenta en g/ml y el inventario en kg/lt
        u_rem = base
        # inventario en kg/lt/un/porciones -> a la unidad base del ítem (g, un o porción)
        if cm in {"kg", "lt"} and base == "GR":
            esc = 1000.0
        elif cm == "kg" and base == "PORC" and gpp:
            esc = 1000.0 / gpp
        else:
            esc = 1.0
        rq = R0.at[oid, "qty"] if oid in R0.index else 0.0
        rt = R0.at[oid, "total"] if oid in R0.index else 0.0
        if oid in C0.index:        # el local también lo compra directo por factura (huevo, café, cajas...)
            rq += C0.at[oid, "qty"]
            rt += C0.at[oid, "total"]
        t_in_o, t_out_o = flujo_obr.get(("t_in", oid), 0.0), flujo_obr.get(("t_out", oid), 0.0)
        desp_o = flujo_obr.get(("desp", oid), 0.0)
        real = si * esc + rq + t_in_o - t_out_o - sf * esc
        p_transf = rec.precio_unit(oid) if oid in rec.obrador.index else 0.0
        p_lista = precio_lista.get(oid)
        if p_lista and p_transf and not 1 / PRECIO_ESCALA_MAX <= p_lista / p_transf <= PRECIO_ESCALA_MAX:
            fuente_extra, p_lista = (f"lista en otra escala/unidad ({p_lista:,.2f} vs {p_transf:,.2f}): se usa "
                                     "P_TRANSFERENCIA"), None
        else:
            fuente_extra = ""
        if p_lista:
            pt, fuente_pt = p_lista, "Lista de precios Obrador"
        elif p_transf:
            pt, fuente_pt = p_transf, "P_TRANSFERENCIA (Recetas)"
        else:
            pt, fuente_pt = (rt / rq if rq else 0.0), "promedio de remitos"
        if fuente_extra:
            fuente_pt = f"P_TRANSFERENCIA (Recetas); {fuente_extra}"
        teor = teor_neto.get(oid, 0.0)
        nombre = rec.nombre_de(oid) if oid in rec.obrador.index else (
            (R0.at[oid, "nombre"] if oid in R0.index else (F0.at[oid, "producto"] if oid in F0.index else oid)))
        filas_o.append({
            "codigo": oid, "item_obrador": nombre, "unidad": u_rem or "un",
            "stock_inicial": si * esc, "recibido_remitos": rq, "traspasos_in": t_in_o, "traspasos_out": t_out_o,
            "stock_final": sf * esc,
            "consumo_real": real, "consumo_teorico": teor, "desperdicio_qty": desp_o,
            "precio_transferencia": pt, "fuente_precio": fuente_pt,
            "precio_lista": precio_lista.get(oid), "precio_recetas": p_transf or None,
            "con_movimiento": bool(si or sf or rq or t_in_o or t_out_o),
            "costo_real": real * pt, "costo_teorico": teor * pt, "remitos_$": rt,
            "tiene_receta": bool(rec.receta_de(oid) or rec.receta_obrador.get(oid)),
        })
    obr = pd.DataFrame(filas_o, columns=[
        "codigo", "item_obrador", "unidad", "stock_inicial", "recibido_remitos", "traspasos_in", "traspasos_out",
        "stock_final", "consumo_real", "consumo_teorico", "desperdicio_qty", "precio_transferencia", "fuente_precio",
        "precio_lista", "precio_recetas", "con_movimiento", "costo_real", "costo_teorico", "remitos_$",
        "tiene_receta"])
    obr["desvio_qty"] = obr["consumo_real"] - obr["consumo_teorico"]
    obr["desvio_costo"] = obr["costo_real"] - obr["costo_teorico"]
    obr["alerta"] = (obr.apply(lambda r: "desvío extremo: revisar" if (r["costo_teorico"] and r["costo_real"] and not
                               (OUTLIER_MIN <= r["costo_real"] / r["costo_teorico"] <= OUTLIER_MAX))
                               else ("sin ventas que lo expliquen (ninguna receta usa este ítem o no está mapeado)"
                                     if not r["consumo_teorico"] and r["consumo_real"] > 0 else ""), axis=1)
                     if len(obr) else pd.Series(dtype=str))

    # ------------------ Consolidado: Arcos + Obrador abierto hasta materia prima ------------------
    real_obr_raw: dict[str, float] = defaultdict(float)
    for r in obr.itertuples(index=False):
        if r.consumo_real > 0:
            rec.explotar(r.codigo, r.consumo_real, real_obr_raw, frenar_obrador=False)
    cons_filas = []
    claves_c = sorted(set(neto["codigo"]) | set(teor_cons) | set(real_obr_raw))
    N0 = neto.set_index("codigo")
    for cid in claves_c:
        if cid.startswith("OBR?:"):
            continue
        en_neto = cid in N0.index
        p_rec = rec.precio_unit(cid)
        tu = rec.unidad_de(cid)
        if en_neto:
            fila = N0.loc[cid]
            factor = float(fila["factor"]) if pd.notna(fila["factor"]) else None
            real_dir_inv, costo_dir, precio_inv = float(fila["consumo_real"]), float(fila["costo_real"]), float(fila["precio_real"])
            nombre = fila["insumo"]
        else:
            factor, real_dir_inv, costo_dir, precio_inv, nombre = 1.0, 0.0, 0.0, 0.0, rec.nombre_de(cid)
        # todo en unidad de receta (g/ml/un): el real directo viene del inventario y se convierte
        p_unit = (precio_inv / factor) if (factor and precio_inv) else p_rec
        real_dir_rec = real_dir_inv * factor if factor else None
        via = real_obr_raw.get(cid, 0.0)
        cons_filas.append({
            "codigo": cid, "insumo": nombre,
            "origen": ("Arcos + Obrador" if via and real_dir_inv else ("Obrador" if via else "Arcos")),
            "unidad_receta": tu or ("un" if factor == 1.0 else ""),
            "real_directo_arcos": real_dir_rec, "real_via_obrador": via,
            "consumo_real": (real_dir_rec or 0.0) + via, "consumo_teorico": teor_cons.get(cid, 0.0),
            "precio": p_unit,
            "costo_real": costo_dir + via * (p_unit or p_rec),
            "costo_teorico": teor_cons.get(cid, 0.0) * p_rec,
        })
    cons = pd.DataFrame(cons_filas)
    cons["desvio_qty"] = cons["consumo_real"] - cons["consumo_teorico"]
    cons["desvio_costo"] = cons["costo_real"] - cons["costo_teorico"]
    cons["alerta"] = cons.apply(lambda r: "desvío extremo: revisar" if (r["costo_teorico"] and r["costo_real"] and not
                                (OUTLIER_MIN <= r["costo_real"] / r["costo_teorico"] <= OUTLIER_MAX)) else "", axis=1)

    # ------------------ precios de mercado ------------------
    mercado = cargar_precios_mercado().set_index("codigo")
    mat = neto.copy()
    mat["base_mercado"] = mat["codigo"].map(mercado["base"] if len(mercado) else {})
    mat["precio_criollo"] = mat["codigo"].map(mercado["criollo"] if len(mercado) else {})
    mat["precio_mejor"] = mat["codigo"].map(mercado["mejor"] if len(mercado) else {})
    mat["proveedor_mejor"] = mat["codigo"].map(mercado["mejor_prov"] if len(mercado) else {})
    mat["precio_posta"] = pd.to_numeric(mat["codigo"].map(mercado["posta"] if len(mercado) else {}), errors="coerce")
    mat["base_posta"] = mat["codigo"].map(mercado["posta_base"] if len(mercado) else {})
    mat["match_posta"] = mat["codigo"].map(mercado["posta_match"] if len(mercado) else {})
    clase_base = {"kg": "$/KG", "lt": "$/LT", "un": "$/UN"}
    mat["base_inv"] = mat["unidad_inv"].map(lambda m: clase_base.get(clase_medida(m), "$/UN"))
    # un outlier de escala/unidades o un ítem sin seguimiento no genera "ahorro": el dato de cantidad no es confiable
    cantidad_ok = ~mat["excl_escala"] & ~mat["sin_seguimiento"]
    comparable = (mat["base_mercado"] == mat["base_inv"]) & mat["precio_real"].gt(0) & cantidad_ok

    def escala_ok(a: pd.Series, b: pd.Series) -> pd.Series:
        r = a / b
        return b.gt(0) & r.between(1 / PRECIO_ESCALA_MAX, PRECIO_ESCALA_MAX)

    mat["comparable_mejor"] = comparable & escala_ok(mat["precio_real"], mat["precio_mejor"])
    mat["comparable_criollo"] = comparable & escala_ok(mat["precio_real"], mat["precio_criollo"])
    mat["comparable_posta"] = ((mat["base_posta"] == mat["base_inv"]) & mat["precio_real"].gt(0) & cantidad_ok
                               & escala_ok(mat["precio_real"], mat["precio_posta"]))
    mat["excl_posta"] = (mat["codigo"].isin(EXCLUIR_AHORRO_POSTA_SKU)
                         | mat["insumo"].map(lambda t: bool(EXCLUIR_AHORRO_POSTA_RE.search(nk(t)))))
    mat.loc[mat["excl_posta"], "comparable_posta"] = False
    qty_pos = mat["consumo_real"].clip(lower=0)
    mat["costo_a_posta"] = qty_pos * mat["precio_posta"].where(mat["comparable_posta"])
    mat["costo_real_comp_posta"] = qty_pos * mat["precio_real"].where(mat["comparable_posta"])
    mat["ahorro_posta"] = mat["costo_real_comp_posta"] - mat["costo_a_posta"]
    mat["dif_precio_pct_posta"] = (mat["precio_real"] / mat["precio_posta"] - 1).where(mat["comparable_posta"])
    mat["costo_a_criollo"] = qty_pos * mat["precio_criollo"].where(mat["comparable_criollo"])
    mat["costo_a_mejor"] = qty_pos * mat["precio_mejor"].where(mat["comparable_mejor"])
    mat["costo_real_comp_criollo"] = qty_pos * mat["precio_real"].where(mat["comparable_criollo"])
    mat["costo_real_comp_mejor"] = qty_pos * mat["precio_real"].where(mat["comparable_mejor"])
    mat["ahorro_criollo"] = mat["costo_real_comp_criollo"] - mat["costo_a_criollo"]
    mat["ahorro_mejor"] = mat["costo_real_comp_mejor"] - mat["costo_a_mejor"]
    mat["dif_precio_pct_mejor"] = (mat["precio_real"] / mat["precio_mejor"] - 1).where(mat["comparable_mejor"])
    mat["dif_precio_pct_criollo"] = (mat["precio_real"] / mat["precio_criollo"] - 1).where(mat["comparable_criollo"])
    # Precio de referencia por insumo: POSTA -> mejor del mercado -> El Criollo -> MAESTRO PRODUCTOS (columnas
    # 'Precio' y 'De donde extrae'). Así ningún insumo con consumo queda en 'Sin match'.
    mat["precio_ref"] = float("nan")
    mat["fuente_ref"] = ""
    for etq, col, mask in (
            ("POSTA", "precio_posta", mat["comparable_posta"]),
            ("Mejor precio del mercado", "precio_mejor", mat["comparable_mejor"]),
            ("El Criollo", "precio_criollo", mat["comparable_criollo"]),
            ("MAESTRO PRODUCTOS", "precio_maestro", cantidad_ok & mat["precio_maestro"].gt(0)
             & escala_ok(mat["precio_real"], mat["precio_maestro"]))):
        libre = mat["precio_ref"].isna() & mask
        mat.loc[libre, "precio_ref"] = mat.loc[libre, col]
        mat.loc[libre, "fuente_ref"] = etq
        if etq == "MAESTRO PRODUCTOS":
            mat.loc[libre, "fuente_ref"] = "MAESTRO PRODUCTOS (" + mat.loc[libre, "maestro_fuente"].replace(
                "", "sin dato").astype(str) + ")"
    mat["ahorro_ref"] = qty_pos * (mat["precio_real"] - mat["precio_ref"])

    # ------------------ consumo propio (sin lo provisto por Obrador) ------------------
    prov_aplicado = neto[["remitos_obrador", "consumo_real"]].apply(
        lambda r: min(r["remitos_obrador"], max(r["consumo_real"], 0.0)), axis=1)
    neto["provisto_obrador_qty"] = prov_aplicado
    neto["consumo_propio"] = neto["consumo_real"] - prov_aplicado
    neto["costo_propio"] = neto["consumo_propio"] * neto["precio_real"]

    # ------------------ auditoría del desvío ------------------
    # Todo se valúa a precio real de compra: el desvío queda en cantidad (no mezcla efecto precio).
    # real = teórico + justificado + merma natural + faltante crítico + sobrante(neg.)
    def clase_aud(r: dict[str, Any]) -> str:
        en_bom = bool(r["teorico_receta"]) or bool(pasante_inv.get(r["codigo"]))
        if r["es_operativo"] and not en_bom:       # sin receta que lo use: gasto operativo
            return "Gasto operativo"
        if r["es_verduleria_agregada"]:
            return "Verdulería (agregada)"
        if r["sin_seguimiento"]:
            return "No auditable: sin stock/compras por ítem"
        if r["consumo_real"] < -1e-9:
            return "No auditable: consumo negativo"
        if r["excl_escala"] or pd.isna(r["consumo_teorico"]):
            return "No auditable: unidades/escala"
        return "Descartable en BOM (auditado)" if r["es_operativo"] else "Materia prima auditada"

    def partir(real: float, teor: float, desp: float, tol: float) -> dict[str, float]:
        gap = real - teor
        just = min(desp, max(gap, 0.0))
        resid = gap - just
        nat = min(max(resid, 0.0), tol * teor)
        return {"justificado": just, "natural": nat, "critico": max(resid, 0.0) - nat,
                "sobrante": min(resid, 0.0), "exceso_desp": max(desp - max(gap, 0.0), 0.0)}

    filas_a: list[dict[str, Any]] = []
    vacio = {"real": 0.0, "teorico": 0.0, "justificado": 0.0, "natural": 0.0, "operativo": 0.0, "critico": 0.0,
             "sobrante": 0.0, "no_auditable": 0.0, "exceso_desp": 0.0}
    verd = {"real": 0.0, "teor": 0.0, "desp": 0.0, "n": 0}
    for r in neto.to_dict("records"):
        if not (r["consumo_real"] or r["teorico_receta"] or r["desperdicio_qty"]):
            continue
        clase = clase_aud(r)
        factor_ok = r["factor"] if pd.notna(r["factor"]) else 0.0
        p_aud = r["precio_real"] if r["precio_real"] > 0 else r["precio_receta"] * factor_ok
        real_d = max(r["consumo_real"], 0.0) * p_aud
        desp_d = r["desperdicio_qty"] * p_aud
        f = dict(vacio)
        f.update(clase=clase, codigo=r["codigo"], insumo=r["insumo"], rubro=r["rubro_gestion"],
                 unidad=r["unidad_inv"], consumo_real=r["consumo_real"], consumo_teorico=r["consumo_teorico"],
                 desperdicio_qty=r["desperdicio_qty"], alerta=r["alerta"], desperdicio_d=desp_d)
        if clase in {"Materia prima auditada", "Descartable en BOM (auditado)"}:
            teor_d = (r["consumo_teorico"] or 0.0) * p_aud
            f.update(real=real_d, teorico=teor_d)
            f.update(partir(real_d, teor_d, desp_d, MERMA_NATURAL.get(r["rubro_n"], 0.0)))
        elif clase == "Verdulería (agregada)":
            verd["real"] += real_d
            verd["teor"] += r["costo_teorico"]
            verd["desp"] += desp_d
            verd["n"] += 1
            continue
        elif clase == "Gasto operativo":
            f.update(real=real_d, teorico=0.0, operativo=real_d)
        else:
            f.update(real=real_d, teorico=0.0, no_auditable=real_d)
        filas_a.append(f)
    if verd["n"]:
        f = dict(vacio)
        f.update(clase="Verdulería (agregada)", codigo="VERDULERIA",
                 insumo=f"Verdulería: compra genérica vs explosión de {verd['n']} ítems frescos sin stock por ítem",
                 rubro="Frutas y Verduras", unidad="$", consumo_real=None, consumo_teorico=None,
                 desperdicio_qty=None, desperdicio_d=verd["desp"],
                 alerta="Se audita agregada: la factura de verdulería no trae detalle por ítem",
                 real=verd["real"], teorico=verd["teor"])
        f.update(partir(verd["real"], verd["teor"], verd["desp"], MERMA_NATURAL["frutas y verduras"]))
        filas_a.append(f)
    for r in obr.to_dict("records"):
        if not (r["consumo_real"] or r["consumo_teorico"] or r["desperdicio_qty"]):
            continue
        pt = r["precio_transferencia"] or 0.0
        real_d, teor_d, desp_d = max(r["consumo_real"], 0.0) * pt, r["consumo_teorico"] * pt, r["desperdicio_qty"] * pt
        f = dict(vacio)
        f.update(clase="Semielaborado Obrador", codigo=r["codigo"], insumo=r["item_obrador"], rubro="Obrador",
                 unidad=r["unidad"], consumo_real=r["consumo_real"], consumo_teorico=r["consumo_teorico"],
                 desperdicio_qty=r["desperdicio_qty"], alerta=r["alerta"], desperdicio_d=desp_d)
        if r["consumo_real"] < -1e-9:
            f.update(clase="No auditable: consumo negativo", real=0.0)
        elif not r["con_movimiento"] and r["consumo_teorico"]:
            f.update(clase="No auditable: sin stock/remitos por ítem", real=0.0)
        else:
            f.update(real=real_d, teorico=teor_d)
            f.update(partir(real_d, teor_d, desp_d, 0.0))
        filas_a.append(f)
    aud = pd.DataFrame(filas_a)
    aud["causa"] = ""
    sin_teo = aud["consumo_teorico"].fillna(0).eq(0) & aud["clase"].isin(["Materia prima auditada", "Descartable en BOM (auditado)", "Semielaborado Obrador"])
    aud.loc[(aud["critico"] > 0) & sin_teo, "causa"] = "Sin receta/venta que lo explique: revisar BOM o mapeo de ventas"
    aud.loc[(aud["critico"] > 0) & ~sin_teo, "causa"] = "Consumo mayor al que explican las recetas"
    aud.loc[(aud["critico"] > 0) & aud["clase"].eq("Verdulería (agregada)"), "causa"] = (
        "Compra genérica de verdulería vs teórico agregado")
    aud["chequeo"] = (aud["real"] - aud["teorico"] - aud["justificado"] - aud["natural"] - aud["critico"]
                      - aud["sobrante"] - aud["operativo"] - aud["no_auditable"]).round(4)

    # ------------------ controles ------------------
    if not planta:
        res_sept = pd.read_excel(DATA_DIR / f"gestion_{MESES_ES[mes]}.xlsx", "Resumen", header=None)
        res_ago = pd.read_excel(DATA_DIR / f"gestion_{MESES_ES[mes_anterior(anio, mes)[1]]}.xlsx", "Resumen",
                                header=None)

        def col_mes(df: pd.DataFrame, nombre: str) -> int:
            # la planilla arrastra dos años (mayo..octubre): el mes del año en curso es el último que figura
            return max(i for i, v in df.iloc[0].items() if str(v).strip().lower() == nombre.lower())

        def rs(df: pd.DataFrame, fila: str, col: int) -> float:
            m = df[df[0].astype(str).str.strip() == fila]
            return num(m.iloc[0][col]) if len(m) else float("nan")

        col_s, col_a = col_mes(res_sept, MESES_ES[mes]), col_mes(res_ago, MESES_ES[mes_anterior(anio, mes)[1]])
        gest = {
            "gest_stock_inicial": rs(res_sept, "STOCK INICIAL", col_s),
            "gest_compras": rs(res_sept, "COMPRAS", col_s),
            "gest_obrador": rs(res_sept, "OBRADOR", col_s),
            "gest_stock_final": rs(res_sept, "STOCK FINAL", col_s),
            "gest_consumo": rs(res_sept, "CONSUMO $$$", col_s),
            "gest_venta_sin_iva": rs(res_sept, "VENTA TOTAL SIN IVA", col_s),
            "gest_stock_final_ago_en_agosto": rs(res_ago, "STOCK FINAL", col_a),
        }
    else:
        rc = pd.read_excel(DATA_DIR / f"gestion_congreso_{MESES_ES[mes]}.xlsx", "Resumen", header=None)

        def rv(etq: str) -> float:
            m = rc[rc[0].astype(str).str.strip() == etq]
            return num(m.iloc[0][1]) if len(m) else float("nan")

        gest = {"gest_stock_inicial": rv("STOCK INICIAL"), "gest_compras": rv("COMPRAS"), "gest_obrador": 0.0,
                "gest_stock_final": rv("STOCK FINAL"), "gest_consumo": rv("CONSUMO $$$"),
                "gest_venta_sin_iva": rv("VENTAS NETAS"), "gest_stock_final_ago_en_agosto": float("nan")}
    f_mp = fac_all[fac_all["cat"].isin(["materia prima", "descartables/packaging"])]
    resumen = {
        **gest,
        "inv_valor_agosto": float(inv_i["valor"].sum()),
        "inv_valor_septiembre": float(inv_f["valor"].sum()),
        "compras_facturas_mp_desc": comp_total_materia,
        "compras_facturas_total": float(fac["total"].sum()),
        "fact_mes_factura": float(f_mp.loc[f_mp["en_mes"], "total"].sum()),
        "fact_mes_pago": float(f_mp.loc[f_mp["en_mes_pago"], "total"].sum()),
        "fact_mes_ingreso": float(f_mp.loc[f_mp["en_mes_ingreso"], "total"].sum()),
        "fact_solo_mp": float(f_mp.loc[f_mp["en_mes"] & (f_mp["cat"] == "materia prima"), "total"].sum()),
        "fact_solo_desc": float(f_mp.loc[f_mp["en_mes"] & (f_mp["cat"] == "descartables/packaging"), "total"].sum()),
        "remitos_obrador": remitos_total,
        "remitos_items_obrador": float(remi["total"].sum()),
        "remitos_materia_prima": float(prov_obrador["total"].sum()),
        "provistos_n": len(provistos),
    }
    return {
        "ventas": ventas, "ctrl_ventas": ctrl_ventas, "neto": neto, "obrador": obr, "consolidado": cons,
        "matriz": mat, "resumen_gestion": resumen, "sin_receta": sin_receta, "ciclos": ciclos,
        "rec": rec, "mes": mes, "anio": anio,
        "remitos_sin_mapear": remi[remi["id"].str.startswith("OBR?:")],
        "auditoria": aud, "movimientos": mov, "equivalencias": equiv, "ventas_obrador": vobr,
        "traspasos": trasp, "desperdicios": desp, "lista_obrador": lista_obr,
        "local": local, "nombre_local": cfg["nombre"], "corto_local": cfg["corto"], "planta": planta_info,
        "es_planta": planta, "maestro_disponible": maestro["disponible"],
    }


# --------------------------------------------------------------------------- #
# Salidas
# --------------------------------------------------------------------------- #
def _kpis(res: dict[str, Any]) -> dict[str, float]:
    neto, obr, cons, mat = res["neto"], res["obrador"], res["consolidado"], res["matriz"]
    aud, mov, vo = res["auditoria"], res["movimientos"], res["ventas_obrador"]
    g = res["resumen_gestion"]
    ok = ~neto["excl_escala"]
    mp = aud[aud["clase"].isin(["Materia prima auditada", "Descartable en BOM (auditado)", "Verdulería (agregada)"])]
    so = aud[aud["clase"] == "Semielaborado Obrador"]
    desp_ok = mov[(mov["tipo"] == "Desperdicio fichado") & mov["estado"].str.startswith("ok")]
    desp_ex = mov[(mov["tipo"] == "Desperdicio fichado") & mov["estado"].str.startswith("EXCLUIDO")]
    t_in = mov[mov["tipo"] == "Traspaso entrante"]
    t_out = mov[mov["tipo"] == "Traspaso saliente"]
    k = {
        "ventas_unidades": float(res["ventas"]["unidades"].sum()),
        "ventas_pesos": float(res["ventas"]["venta"].sum()),
        "costo_teorico_neto": float(neto["costo_teorico"].sum()),
        "costo_real_neto": float(neto.loc[neto["consumo_real"] > 0, "costo_real"].sum()),
        "costo_real_propio": float(neto.loc[neto["consumo_propio"] > 0, "costo_propio"].sum()),
        "prov_obrador_mp": float(neto["remitos_obrador_$"].sum()),
        "costo_teorico_neto_ok": float(neto.loc[ok, "costo_teorico"].sum()),
        "costo_real_neto_ok": float(neto.loc[ok, "costo_real"].sum()),
        "obrador_real": float(obr["costo_real"].clip(lower=0).sum()),
        "obrador_teorico": float(obr["costo_teorico"].sum()),
        "obrador_remitos": float(obr["remitos_$"].sum()),
        "consolidado_teorico": float(cons["costo_teorico"].sum()),
        "consolidado_real": float(cons.loc[cons["consumo_real"] > 0, "costo_real"].sum()),
        "ahorro_criollo": float(mat["ahorro_criollo"].clip(lower=0).sum()),
        "ahorro_mejor": float(mat["ahorro_mejor"].clip(lower=0).sum()),
        "ahorro_posta": float(mat["ahorro_posta"].clip(lower=0).sum()),
        "ahorro_posta_neto": float(mat["ahorro_posta"].sum()),
        "posta_real_comp": float(mat["costo_real_comp_posta"].sum()),
        "posta_costo": float(mat["costo_a_posta"].sum()),
        "posta_items": int(mat["comparable_posta"].sum()),
        "ref_cobertura": float(mat.loc[(mat["consumo_real"] > 0) & mat["precio_ref"].notna(), "costo_real"].sum()),
        "ref_maestro_items": int(mat["fuente_ref"].str.startswith("MAESTRO").sum()),
        "ref_sin_match": int(((mat["consumo_real"] > 0) & mat["precio_ref"].isna()).sum()),
        "descartables_bom": float(aud.loc[aud["clase"] == "Descartable en BOM (auditado)", "teorico"].sum()),
        "consumo_real_total_pos": float(mat.loc[mat["consumo_real"] > 0, "costo_real"].sum()),
        "items_comparables_mejor": int(mat["comparable_mejor"].sum()),
        "items_comparables_criollo": int(mat["comparable_criollo"].sum()),
        "n_outliers": int(neto["excl_escala"].sum()),
        "n_sin_seguimiento": int(neto["sin_seguimiento"].sum()),
        "n_negativos": int((neto["consumo_real"] < -1e-9).sum()),
        # auditoría materia prima
        "aud_real_mp": float(mp["real"].sum()), "aud_teorico_mp": float(mp["teorico"].sum()),
        "aud_justificado_mp": float(mp["justificado"].sum()), "aud_natural_mp": float(mp["natural"].sum()),
        "aud_critico_mp": float(mp["critico"].sum()), "aud_sobrante_mp": float(mp["sobrante"].sum()),
        # auditoría semielaborados Obrador
        "aud_real_obr": float(so["real"].sum()), "aud_teorico_obr": float(so["teorico"].sum()),
        "aud_justificado_obr": float(so["justificado"].sum()), "aud_critico_obr": float(so["critico"].sum()),
        "aud_sobrante_obr": float(so["sobrante"].sum()),
        "gasto_operativo": float(aud["operativo"].sum()), "no_auditable": float(aud["no_auditable"].sum()),
        "crit_sin_receta": float(mp.loc[mp["causa"].str.startswith("Sin receta"), "critico"].sum()),
        "crit_verduleria": float(mp.loc[mp["clase"] == "Verdulería (agregada)", "critico"].sum()),
        "n_operativos": int((aud["clase"] == "Gasto operativo").sum()),
        "desp_fichado_ok": float(desp_ok["importe"].sum()), "desp_excluido": float(desp_ex["importe"].sum()),
        "desp_exceso": float(aud["exceso_desp"].sum()),
        "traspasos_in": float(t_in["importe"].sum()), "traspasos_out": float(t_out["importe"].sum()),
        "obrador_ventas_arcos": float(vo.get("cocina", float("nan"))) if vo.get("disponible") else float("nan"),
    }
    k["gest_consumo"] = g["gest_consumo"]
    return k


def _clase_por_codigo(res: dict[str, Any]) -> dict[str, str]:
    aud, neto = res["auditoria"], res["neto"]
    m = {c: k for c, k in zip(aud["codigo"], aud["clase"])}
    for c in neto.loc[neto["es_verduleria_agregada"], "codigo"]:
        m[c] = "Verdulería (agregada)"
    return m


def _hojas(res: dict[str, Any]) -> dict[str, pd.DataFrame]:
    neto, obr, cons, mat = res["neto"], res["obrador"], res["consolidado"], res["matriz"]
    aud, mov, vo, eq = res["auditoria"], res["movimientos"], res["ventas_obrador"], res["equivalencias"]
    k = _kpis(res)
    g = res["resumen_gestion"]
    mes_t = f"{MESES_ES[res['mes']]} {res['anio']}"
    mes_ant = MESES_ES[res["mes"] - 1] if res["mes"] > 1 else "diciembre"
    nombre, es_p, pl = res["nombre_local"], res["es_planta"], res["planta"]
    clase_de = _clase_por_codigo(res)
    top_crit = aud[aud["critico"] > 0].nlargest(10, "critico")
    pct = lambda a, b: (a / b) if b else None  # noqa: E731
    tol_txt = ", ".join(f"{r.title()} {v:.0%}" for r, v in MERMA_NATURAL.items())
    obr_dif = (k["obrador_ventas_arcos"] - k["obrador_remitos"]) if vo.get("disponible") else None

    intro = [
        (f"ANÁLISIS DE GESTIÓN COMERCIAL - {nombre.upper()}", mes_t.upper(), ""),
        ("", "", ""),
        ("METODOLOGÍA DE CÁLCULO", "", ""),
        ("1. Consumo real (balance de masa)", "",
         f"Stock inicial (inventario fin de {mes_ant}) + compras del mes (facturas del mes de factura, rubros "
         f"Materia prima y Descartables) + remitos "
         "de Obrador que son materia prima del local + traspasos entrantes - traspasos salientes - stock final "
         "(inventario fin del mes). Los traspasos salientes a otros locales se restan: no son consumo ni desvío."),
        ("2. Consumo teórico (BOM)", "",
         "Explosión recursiva de recetas de lo vendido. Las recetas Madre (M, MR, MO, OBR, P, R...) se abren hasta "
         "materia prima; en la vista Neto Arcos la explosión se frena en cada ítem de Obrador, que se valúa aparte."),
        ("3. Equivalencias de códigos", "",
         "Los códigos Maxirest (5144 manteca, 808 aceite de girasol, 5578 bolsa vacío...) se traducen a SKU con "
         "'Precio viejo//nuevo metodo.' y la 'Lista de precios Obrador'; ver hoja 'Equivalencias Códigos'."),
        ("4. Partición del desvío (real - teórico)", "",
         "Todo a precio real de compra. a) Desperdicio justificado: lo fichado en 'Desperdicios' (las Madres y "
         "platos se explotan a materia prima). b) Merma natural en frescos hasta una tolerancia sobre el teórico. "
         "c) Gasto operativo: insumos sin receta de cocina. d) Faltante crítico: lo que queda sin justificar."),
        ("5. Obrador y Benchmark", "",
         "El consumo neto descuenta lo provisto por Obrador (valuado a la lista de precios del Obrador). "
         "El benchmark valúa el consumo real a la lista vigente de Posta, solo en ítems comparables (misma base y "
         "escala); si Posta no lo cotiza, la referencia es el Maestro (columnas 'Precio' y 'De donde extrae')."),
        ("6. Maestro y descartables", "",
         "MAESTRO PRODUCTOS es el diccionario de códigos/nombres/SKU y el precio de respaldo. Los descartables "
         "(cajas, bolsas, ecobandejas) anexados al BOM de un artículo vendido suman al teórico; solo se tratan como "
         "gasto operativo los que ninguna receta usa."),
        ("", "", ""),
        ("VENTAS", "", ""),
        ("Unidades vendidas", k["ventas_unidades"], "Hoja 'Ventas' de Ventas ARCOS"),
        ("Venta total (hoja Ventas)", k["ventas_pesos"],
         f"Referencia del Resumen de gestión: venta sin IVA {pesos(g['gest_venta_sin_iva'])}"),
        ("Artículos vendidos sin receta ni insumo asociado", len(res["sin_receta"]), "Ver hoja 'Controles'"),
        ("", "", ""),
        ("CONSUMO REAL Y TEÓRICO ($)", "", ""),
        ("Neto Arcos - materia prima propia: costo real (todo lo consumido)", k["costo_real_neto"],
         "Valuado a precio de compra del mes; incluye gasto operativo y la compra genérica de verdulería"),
        ("   de lo cual provisto por Obrador (materia prima entregada por remito)", k["prov_obrador_mp"],
         "Se descuenta para aislar el consumo neto de la sucursal"),
        ("   Consumo neto propio de Arcos (sin Obrador)", k["costo_real_propio"], ""),
        ("Neto Arcos - costo teórico por recetas", k["costo_teorico_neto"], "Materia prima propia, a precio de receta"),
        ("Obrador a lista de precios: costo real consumido", k["obrador_real"],
         "Stock + remitos + traspasos - stock final, por ítem"),
        ("Obrador a lista de precios: costo teórico (ítems con seguimiento)", k["aud_teorico_obr"],
         f"{g['provistos_n']} ítems de la lista Obrador provistos en el mes; los demás los elabora Arcos "
         "y su materia prima va al teórico propio"),
        ("Global (Arcos + Obrador abierto a materia prima): costo real", k["consolidado_real"], "Hoja 'Consumo Global'"),
        ("Global: costo teórico", k["consolidado_teorico"], ""),
        ("Consumo $$$ informado en 'Gestión mensual' (Resumen)", g["gest_consumo"],
         "Referencia: incluye todo el inventario y traspasos"),
        ("", "", ""),
        ("TRASPASOS Y OBRADOR", "", ""),
        ("Traspasos entrantes a Arcos ($)", k["traspasos_in"], "Hoja 'Movimientos'; suman al consumo"),
        ("Traspasos salientes de Arcos ($)", k["traspasos_out"],
         "Cabello / Congreso: se restan del consumo; no son desvío"),
        ("Ventas del Obrador a Arcos (planilla Ventas_Obrador)", k["obrador_ventas_arcos"],
         "Hoja 'Resumen', columna Cocina"),
        ("Remitos de Obrador en ARCOS (hoja Remitos del mes)", g["remitos_obrador"],
         f"Ítems de la lista Obrador {pesos(g['remitos_items_obrador'])} + materia prima del local "
         f"{pesos(g['remitos_materia_prima'])}. Ventas Obrador sin remito: "
         f"{pesos(obr_dif) if obr_dif is not None else '-'}"),
        ("", "", ""),
        ("AUDITORÍA DE DESVÍOS - MATERIA PRIMA ($, a precio real)", "", ""),
        ("Costo real auditado (materia prima + verdulería agregada)", k["aud_real_mp"], ""),
        ("(-) Costo teórico explicado por ventas", -k["aud_teorico_mp"], ""),
        ("(a) Desperdicio justificado (fichado en 'Desperdicios')", k["aud_justificado_mp"],
         f"{pct(k['aud_justificado_mp'], k['aud_real_mp'] - k['aud_teorico_mp']) or 0:.1%} del desvío bruto"),
        ("(b) Merma natural en frescos (verdulería, carnes, pescados)", k["aud_natural_mp"],
         f"Tolerancia supuesta sobre el teórico: {tol_txt}. Ajustable en MERMA_NATURAL"),
        ("(d) Faltante / merma crítica sin justificar", k["aud_critico_mp"], "Hoja 'Auditoría de Desvíos'"),
        ("   del faltante: ítems sin receta/venta que los explique", k["crit_sin_receta"],
         "Revisar BOM o mapeo de ventas antes de leerlo como pérdida"),
        ("   del faltante: verdulería agregada", k["crit_verduleria"], ""),
        ("Sobrante (consumo menor al teórico)", k["aud_sobrante_mp"], "Negativo: recetas que exceden lo consumido"),
        ("(c) Gasto operativo (aparte, no es desvío de cocina)", k["gasto_operativo"],
         f"{k['n_operativos']} artículos: descartables, pastillas rational, hielo, bolsas de vacío, ecobandejas..."),
        ("No auditable (unidades/escala inconsistentes)", k["no_auditable"], "Se informa, no se asigna a desvío"),
        ("", "", ""),
        ("AUDITORÍA - SEMIELABORADOS DE OBRADOR ($, lista de precios Obrador)", "", ""),
        ("Costo real consumido / teórico", k["aud_real_obr"], f"Teórico {pesos(k['aud_teorico_obr'])}"),
        ("Desperdicio justificado", k["aud_justificado_obr"], ""),
        ("Faltante crítico", k["aud_critico_obr"], "Incluye ítems que no tienen receta mapeada en ventas"),
        ("Sobrante (consumo menor al teórico)", k["aud_sobrante_obr"], ""),
        ("", "", ""),
        ("BENCHMARK POSTA", "", ""),
        ("Consumo real comparable valuado a precio actual de compra", k["posta_real_comp"],
         f"{k['posta_items']} insumos comparables ({pct(k['posta_real_comp'], k['consumo_real_total_pos']) or 0:.0%} del costo real)"),
        ("El mismo consumo valuado a la lista vigente de POSTA", k["posta_costo"], ""),
        ("Ahorro potencial global comprando a POSTA (solo ítems donde Posta es más barato)", k["ahorro_posta"],
         f"Neto compensando ítems donde Posta es más caro: {pesos(k['ahorro_posta_neto'])}"),
        ("Ahorro potencial vs El Criollo", k["ahorro_criollo"], f"{k['items_comparables_criollo']} insumos comparables"),
        ("Ahorro potencial vs mejor precio del mercado", k["ahorro_mejor"], f"{k['items_comparables_mejor']} insumos comparables"),
        ("", "", ""),
        ("TOP 10 FALTANTES CRÍTICOS ($)", "", ""),
    ]
    for r in top_crit.itertuples(index=False):
        intro.append((f"{r.insumo} [{r.codigo}] - {r.clase}", float(r.critico),
                      f"real {pesos(r.real)} vs teórico {pesos(r.teorico)}; justificado {pesos(r.justificado)}; "
                      f"merma natural {pesos(r.natural)}"))
    intro += [("", "", ""), ("SUPUESTOS Y LIMITACIONES", "", ""),
              ("Merma natural", "", f"Los datos no traen rendimiento por ítem: la tolerancia ({tol_txt}) es un supuesto operativo."),
              ("Desperdicios", "",
               f"{pesos(k['desp_fichado_ok'])} aplicados al balance. Excluidos {pesos(k['desp_excluido'])} por cantidad "
               "fuera de rango (ver 'Controles'). Lo fichado por encima del desvío observado "
               f"({pesos(k['desp_exceso'])}) no genera crédito."),
              ("Verdulería", "", "La factura genérica de verdulería no trae detalle: se audita agregada contra la explosión de los frescos sin stock por ítem."),
              ("Horeca", "", "No existe una lista 'Horeca' en Drive ni en el comparador: se compara contra Posta, El Criollo y el mejor precio del mercado."),
              ("Stock inicial", "", f"Inventario {mes_ant} valorizado {pesos(g['inv_valor_agosto'])}; la planilla informa {pesos(g['gest_stock_inicial'])} como stock inicial."),
              ("Compras", "", f"Facturas del mes de factura {mes_t} (materia prima + descartables) {pesos(g['compras_facturas_mp_desc'])} vs {pesos(g['gest_compras'])} del Resumen de gestión."),
              ]
    if es_p:
        quitar = ("Unidades vendidas", "Venta total (hoja Ventas)", "   de lo cual provisto por Obrador",
                  "   Consumo neto propio de", "Ventas del Obrador a Arcos", "Remitos de Obrador en ARCOS",
                  "Obrador a lista de precios", "Global (Arcos", "Global: costo")
        intro = [r for r in intro if not str(r[0]).startswith(quitar)]
        intro = [(str(a).replace("Arcos", nombre), b, str(c).replace("Arcos", nombre)) for a, b, c in intro]
        cob = (pl["venta_items"] / pl["venta_total"]) if pl and pl.get("venta_total") else None
        idx = next(i for i, r in enumerate(intro) if r[0] == "VENTAS") + 1
        intro[idx:idx] = [
            ("Ventas del Obrador (Resumen: Cocina + Pan)", pl["venta_total"] if pl else None,
             "Planilla Ventas_Obrador del mes; es la 'VENTA COCINA' de la gestión de Congreso"),
            ("Ventas itemizadas por remito (Arcos + Lafinur)", pl["venta_items"] if pl else None,
             f"Cobertura {cob:.0%} de las ventas; falta el detalle de Cabello (sin archivo disponible en Drive)"
             if cob else ""),
            ("   de las cuales materia prima revendida tal cual", pl["venta_pasante"] if pl else None, ""),
            ("   sin mapear a ítem de Obrador ni insumo", pl["venta_sin_mapear"] if pl else None, "Ver 'Controles'"),
            ("Teórico extrapolado al 100% de las ventas (mismo mix)",
             (k["aud_teorico_mp"] / cob) if cob else None,
             "Referencia: supone que Cabello compra el mismo mix. El faltante crítico de abajo está inflado por la "
             "parte de las ventas que no se pudo itemizar"),
            ("Aislamiento de Obrador", "", "Congreso es la planta que produce el Obrador: no recibe semielaborados "
                                         "del Obrador; no hay descuento por provisión"),
        ]
    intro_df = pd.DataFrame(intro, columns=["Concepto", "Valor", "Detalle"])

    n = neto.copy()
    n["clase"] = n["codigo"].map(clase_de).fillna("Sin movimiento")
    n = n.rename(columns={
        "codigo": "Código", "insumo": "Insumo", "rubro_gestion": "Rubro", "tipo_insumo": "Tipo de insumo",
        "clase": "Clasificación", "unidad_inv": "Unidad inventario",
        "stock_inicial": "Stock inicial (inv. mes anterior)", "compras_facturas": "Compras del mes (facturas)",
        "remitos_obrador": "Remitos Obrador (materia prima)", "traspasos_in": "Traspasos entrantes",
        "traspasos_out": "Traspasos salientes", "stock_final": "Stock final (inv. mes)",
        "consumo_real": "Consumo real", "provisto_obrador_qty": "(-) Provisto por Obrador",
        "consumo_propio": "Consumo neto propio (sin Obrador)", "desperdicio_qty": "Desperdicio fichado",
        "consumo_teorico": "Consumo teórico (neto Arcos)", "desvio_qty": "Desvío (real - teórico)",
        "desvio_pct": "Desvío %", "precio_real": "Precio real unitario", "costo_real": "Costo real $",
        "costo_propio": "Costo neto propio $", "costo_teorico": "Costo teórico $", "desvio_costo": "Desvío $",
        "nota_unidad": "Criterio de unidades", "alerta": "Alerta", "fuente_precio": "Fuente del precio real"})
    n = n[["Código", "Insumo", "Rubro", "Tipo de insumo", "Clasificación", "Unidad inventario",
           "Stock inicial (inv. mes anterior)", "Compras del mes (facturas)", "Remitos Obrador (materia prima)",
           "Traspasos entrantes", "Traspasos salientes", "Stock final (inv. mes)", "Consumo real",
           "(-) Provisto por Obrador", "Consumo neto propio (sin Obrador)", "Desperdicio fichado",
           "Consumo teórico (neto Arcos)", "Desvío (real - teórico)", "Desvío %", "Precio real unitario",
           "Fuente del precio real", "Costo real $", "Costo neto propio $", "Costo teórico $", "Desvío $",
           "Criterio de unidades", "Alerta"]].sort_values("Desvío $", ascending=False)

    o = obr.rename(columns={
        "codigo": "Código", "item_obrador": "Ítem de Obrador", "unidad": "Unidad", "stock_inicial": "Stock inicial",
        "recibido_remitos": "Recibido (remitos + compras directas)", "traspasos_in": "Traspasos entrantes",
        "traspasos_out": "Traspasos salientes", "stock_final": "Stock final", "consumo_real": "Consumo real",
        "consumo_teorico": "Consumo teórico (ventas)", "desperdicio_qty": "Desperdicio fichado",
        "precio_transferencia": "Precio aplicado", "fuente_precio": "Fuente del precio",
        "precio_lista": "Precio lista Obrador", "precio_recetas": "P_TRANSFERENCIA (Recetas)",
        "costo_real": "Costo real $", "costo_teorico": "Costo teórico $", "remitos_$": "Remitos + compras $",
        "desvio_qty": "Desvío", "desvio_costo": "Desvío $", "alerta": "Alerta"})
    o = o[["Código", "Ítem de Obrador", "Unidad", "Stock inicial", "Recibido (remitos + compras directas)",
           "Traspasos entrantes", "Traspasos salientes", "Stock final", "Consumo real", "Consumo teórico (ventas)",
           "Desperdicio fichado", "Desvío", "Precio aplicado", "Fuente del precio", "Precio lista Obrador",
           "P_TRANSFERENCIA (Recetas)", "Costo real $", "Costo teórico $", "Desvío $", "Remitos + compras $",
           "Alerta"]].sort_values("Costo real $", ascending=False)

    c = cons.rename(columns={
        "codigo": "Código", "insumo": "Insumo", "origen": "Origen del consumo", "unidad_receta": "Unidad (recetas)",
        "real_directo_arcos": "Real directo (Arcos)", "real_via_obrador": "Real vía Obrador (abierto)",
        "consumo_real": "Consumo real total", "consumo_teorico": "Consumo teórico total",
        "desvio_qty": "Desvío (real - teórico)", "precio": "Precio unitario (por unidad de receta)", "costo_real": "Costo real $",
        "costo_teorico": "Costo teórico $", "desvio_costo": "Desvío $", "alerta": "Alerta"})
    c = c[["Código", "Insumo", "Origen del consumo", "Unidad (recetas)", "Real directo (Arcos)",
           "Real vía Obrador (abierto)", "Consumo real total", "Consumo teórico total", "Desvío (real - teórico)",
           "Precio unitario (por unidad de receta)", "Costo real $", "Costo teórico $", "Desvío $",
           "Alerta"]].sort_values("Desvío $", ascending=False)

    m = mat[mat["consumo_real"] > 0].copy()
    m["clase"] = m["codigo"].map(clase_de).fillna("")

    def estado_posta(r: pd.Series) -> str:
        if r["excl_posta"]:
            return "Excluido: Posta es el proveedor real de compra (sin ahorro vs Posta)"
        if r["comparable_posta"]:
            return "Comparable"
        if r["excl_escala"] or r["sin_seguimiento"]:
            return "Cantidad no confiable (ver Auditoría)"
        if pd.isna(r["precio_posta"]):
            return ("Posta no lo cotiza: referencia MAESTRO" if str(r["fuente_ref"]).startswith("MAESTRO")
                    else "Posta no lo cotiza / sin match")
        return "Unidad o escala distinta"

    m["estado_posta"] = m.apply(estado_posta, axis=1)
    m = m.rename(columns={
        "codigo": "Código", "insumo": "Insumo", "rubro_gestion": "Rubro", "clase": "Clasificación",
        "unidad_inv": "Unidad inventario", "consumo_real": "Consumo real", "precio_real": "Precio real compra",
        "precio_posta": "Precio POSTA", "dif_precio_pct_posta": "Real vs POSTA %", "match_posta": "Producto POSTA (match)",
        "estado_posta": "Estado de la comparación", "costo_real_comp_posta": "Costo real (comparable) $",
        "costo_a_posta": "Costo a precio POSTA $", "ahorro_posta": "Ahorro vs POSTA $",
        "precio_criollo": "Precio El Criollo", "ahorro_criollo": "Ahorro vs Criollo $",
        "precio_mejor": "Precio mejor mercado", "proveedor_mejor": "Proveedor mejor mercado",
        "ahorro_mejor": "Ahorro vs mejor mercado $", "costo_real": "Costo real total $",
        "precio_ref": "Precio de referencia", "fuente_ref": "Fuente de la referencia",
        "ahorro_ref": "Ahorro vs referencia $", "precio_maestro": "Precio MAESTRO",
        "maestro_fuente": "MAESTRO: de dónde extrae", "maestro_fecha": "MAESTRO: fecha factura",
        "fuente_precio": "Fuente del precio real"})
    m = m[["Código", "Insumo", "Rubro", "Clasificación", "Unidad inventario", "Consumo real", "Precio real compra",
           "Precio POSTA", "Real vs POSTA %", "Producto POSTA (match)", "Estado de la comparación",
           "Costo real (comparable) $", "Costo a precio POSTA $", "Ahorro vs POSTA $", "Precio El Criollo",
           "Ahorro vs Criollo $", "Precio mejor mercado", "Proveedor mejor mercado", "Ahorro vs mejor mercado $",
           "Precio de referencia", "Fuente de la referencia", "Ahorro vs referencia $", "Precio MAESTRO",
           "MAESTRO: de dónde extrae", "MAESTRO: fecha factura", "Fuente del precio real", "Costo real total $"]].sort_values("Ahorro vs POSTA $", ascending=False, na_position="last")

    a = aud.rename(columns={
        "clase": "Clasificación", "codigo": "Código", "insumo": "Insumo", "rubro": "Rubro", "unidad": "Unidad",
        "consumo_real": "Consumo real", "consumo_teorico": "Consumo teórico", "desperdicio_qty": "Desperdicio fichado",
        "real": "Costo real $", "teorico": "Costo teórico $", "justificado": "(a) Desperdicio justificado $",
        "natural": "(b) Merma natural frescos $", "operativo": "(c) Gasto operativo $",
        "critico": "(d) Faltante / merma crítica $", "sobrante": "Sobrante (consumo < teórico) $",
        "no_auditable": "No auditable $", "exceso_desp": "Desperdicio fichado que excede el desvío $",
        "causa": "Causa probable del faltante", "alerta": "Alerta"})
    a = a[["Clasificación", "Código", "Insumo", "Rubro", "Unidad", "Consumo real", "Consumo teórico",
           "Desperdicio fichado", "Costo real $", "Costo teórico $", "(a) Desperdicio justificado $",
           "(b) Merma natural frescos $", "(c) Gasto operativo $", "(d) Faltante / merma crítica $",
           "Sobrante (consumo < teórico) $", "No auditable $", "Desperdicio fichado que excede el desvío $",
           "Causa probable del faltante", "Alerta"]].sort_values(["(d) Faltante / merma crítica $", "(c) Gasto operativo $"], ascending=False)

    movs = mov.rename(columns={"tipo": "Movimiento", "fecha": "Fecha", "contraparte": "Contraparte / motivo",
                               "codigo": "Código", "producto": "Producto", "cantidad": "Cantidad",
                               "unidad": "Unidad", "importe": "Importe $", "resuelto_como": "Resuelto como",
                               "estado": "Estado"}).sort_values(["Movimiento", "Fecha"])
    eqd = eq.rename(columns={"codigo": "Código Maxirest", "nombre": "Nombre", "sku": "SKU / id_receta",
                             "fuente": "Fuente"}).sort_values("Código Maxirest")

    # ---------------- Controles ----------------
    ctrl = [("Fuente", "Dato", "Valor")]
    ctrl += [("Ventas", "Unidades (hoja Ventas)", k["ventas_unidades"]),
             ("Ventas", "Unidades en Resumen Ventas (control)", res["ctrl_ventas"].get("resumen_unidades")),
             ("Ventas", "Venta $ (hoja Ventas)", k["ventas_pesos"]),
             ("Ventas", "Venta $ en Resumen Ventas (control)", res["ctrl_ventas"].get("resumen_venta")),
             ("Recetas", "Ciclos de recetas detectados", ", ".join(x.replace("__CICLO__", "") for x in res["ciclos"]) or "ninguno")]
    dif_c = g["fact_mes_factura"] - g["gest_compras"]
    ctrl += [("Integridad de compras", f"Facturas {nombre}: Materia prima + Descartables, mes de factura = "
              f"{res['mes']} (regla del sistema)", g["fact_mes_factura"]),
             ("Integridad de compras", f"COMPRAS del Resumen de la hoja 'Gestión mensual' ({MESES_ES[res['mes']]})",
              g["gest_compras"]),
             ("Integridad de compras", "Diferencia ($)", dif_c),
             ("Integridad de compras", "RESULTADO", "COINCIDE EXACTO" if abs(dif_c) < 1.0 else
              f"NO COINCIDE ({dif_c / g['gest_compras']:+.1%} sobre el Resumen)" if g["gest_compras"] else "sin Resumen"),
             ("Integridad de compras", "   detalle: solo Materia prima / solo Descartables",
              f"{pesos(g['fact_solo_mp'])} / {pesos(g['fact_solo_desc'])}"),
             ("Integridad de compras", "   contraste: mismos rubros por 'Mes pago'", g["fact_mes_pago"]),
             ("Integridad de compras", "   contraste: mismos rubros por fecha de ingreso", g["fact_mes_ingreso"]),
             ("Integridad de stock", "Stock inicial reconstruido del inventario ($) vs Resumen de gestión",
              f"{pesos(g['inv_valor_agosto'])} vs {pesos(g['gest_stock_inicial'])}"),
             ("Integridad de stock", "Stock final del inventario ($) vs Resumen de gestión",
              f"{pesos(g['inv_valor_septiembre'])} vs {pesos(g['gest_stock_final'])}")]
    ctrl += [("Maestro", "Insumos con precio real tomado del MAESTRO por no tener compra ni precio de inventario",
              int((res["neto"]["fuente_precio"] == "Maestro").sum())),
             ("Maestro", "Insumos consumidos con referencia del MAESTRO en el benchmark (no cotizados por Posta)",
              k["ref_maestro_items"]),
             ("Maestro", "Insumos consumidos que siguen sin ninguna referencia de precio", k["ref_sin_match"]),
             ("Maestro", "Descartables dentro del BOM: costo teórico incorporado ($)", k["descartables_bom"])]
    ctrl += [("Auditoría", "Máxima diferencia de la identidad real = teórico + (a) + (b) + (c) + (d) + sobrante + no auditable",
              float(aud["chequeo"].abs().max())),
             ("Auditoría", "Desperdicio fichado aplicado al balance ($)", k["desp_fichado_ok"]),
             ("Auditoría", "Desperdicio fichado excluido por cantidad fuera de rango ($)", k["desp_excluido"]),
             ("Auditoría", "Desperdicio fichado que excede el desvío observado ($, sin crédito)", k["desp_exceso"])]
    if vo.get("disponible"):
        ctrl += [("Obrador", f"Ventas del Obrador a Arcos - Cocina ({vo['archivo']})", vo.get("cocina")),
                 ("Obrador", "Ventas Obrador a Arcos - Pan", vo.get("por_dia_pan")),
                 ("Obrador", "Suma 'Por día' Arcos Cocina", vo.get("por_dia_cocina")),
                 ("Obrador", "Remitos de Obrador en ARCOS (total hoja Remitos del mes)", g["remitos_obrador"]),
                 ("Obrador", "   Remitos: ítems de la lista Obrador", g["remitos_items_obrador"]),
                 ("Obrador", "   Remitos: materia prima del local entregada por Obrador", g["remitos_materia_prima"]),
                 ("Obrador", "Ventas Obrador a Arcos sin remito en ARCOS ($)",
                  (vo.get("cocina") or 0.0) - g["remitos_obrador"]),
                 ("Obrador", "Ventas Obrador - ingresos por traspaso (Resumen)", vo.get("ingresos")),
                 ("Obrador", "Traspasos entrantes a Arcos ($, hoja Traspasos)", k["traspasos_in"]),
                 ("Obrador", "Ventas Obrador - egresos por traspaso (Resumen)", vo.get("egresos")),
                 ("Obrador", "Traspasos salientes de Arcos ($, hoja Traspasos; con signo -)", -k["traspasos_out"])]
    ref = {"5144": "MANT", "808": "AGI", "5578": "I0031"}
    for cod_ in ("5144", "808", "5578", "3904"):
        e = eq[eq["codigo"] == cod_]
        nombres = "; ".join(sorted({f"{r.nombre}" + (f" -> {r.sku}" if r.sku else "") for r in e.itertuples(index=False)}))
        sku = ref.get(cod_) or next((s_ for s_ in e["sku"] if s_), "")
        fila = neto[neto["codigo"] == sku] if sku else neto[neto["insumo"].str.contains("ecobandeja", case=False, na=False)]
        if len(fila):
            f0 = fila.iloc[0]
            det = (f"{nombres} | inventario {f0['insumo']} [{f0['codigo']}]: consumo real {fnum(f0['consumo_real'])} "
                   f"{f0['unidad_inv']}, costo {pesos(f0['costo_real'])}, clase {clase_de.get(f0['codigo'], '-')}")
        else:
            det = f"{nombres or 'sin entrada en la planilla'} | sin movimiento en el inventario del mes"
        ctrl.append(("Equivalencia de código", f"Código {cod_}: {det}"[:400], ""))
    # un mismo código con nombres distintos en la planilla de equivalencias (colisión)
    eqn = eq.assign(k=eq["nombre"].map(_sin_prefijo))
    for cod_, sub in eqn.groupby("codigo"):
        nombres_d = sorted({r.nombre for r in sub.itertuples(index=False)})
        ks_ = sorted(set(sub["k"]))     # nombres truncados ('UN. Caja MISHIGU') no son una colisión
        if cod_ and any(not (a_.startswith(b_) or b_.startswith(a_)) for i_, a_ in enumerate(ks_) for b_ in ks_[i_ + 1:]):
            ctrl.append(("Código ambiguo en equivalencias", f"{cod_}: " + " | ".join(nombres_d)[:300],
                         ", ".join(sorted(set(sub["fuente"])))[:120]))
    for r in mov[~mov["estado"].str.startswith("ok")].itertuples(index=False):
        ctrl.append((f"Movimiento no aplicado ({r.tipo})", f"{r.producto} [{r.codigo}] {r.cantidad:g} {r.unidad} - {r.estado}",
                     r.importe))
    for r in res["sin_receta"].itertuples(index=False):
        ctrl.append(("Ventas sin receta/insumo", f"{r.articulo} (SKU {r.sku if isinstance(r.sku, str) else 'sin SKU'})", r.unidades))
    if es_p and pl:
        for cli, v in pl["por_cliente"].items():
            ctrl.append(("Ventas itemizadas de la planta", f"Remitos a {cli}", v))
        ctrl.append(("Ventas itemizadas de la planta", "Cobertura sobre ventas Obrador del Resumen",
                     f"{pl['venta_items'] / pl['venta_total']:.0%}" if pl.get("venta_total") else "-"))
        for cli_, v in (vo.get("clientes") or {}).items():
            ctrl.append(("Ventas Obrador por cliente (Resumen)", cli_.title(), v.get("cocina")))
    for r in res["remitos_sin_mapear"].itertuples(index=False):
        ctrl.append(("Remitos Obrador sin ítem asociado", r.nombre, r.total))
    for r in neto[neto["outlier"]].itertuples(index=False):
        ctrl.append(("Outlier excluido del desvío (neto)", f"{r.insumo} [{r.codigo}] - {r.alerta}",
                     f"real {pesos(r.costo_real)} vs teórico {pesos(r.costo_teorico)}"))
    for r in neto[neto["alerta"].str.contains("desvío extremo", regex=False) & ~neto["outlier"]].itertuples(index=False):
        ctrl.append(("Desvío extremo (se mantiene en el desvío)", f"{r.insumo} [{r.codigo}]",
                     f"real {pesos(r.costo_real)} vs teórico {pesos(r.costo_teorico)}"))
    for r in neto[neto["alerta"].str.contains("negativo")].itertuples(index=False):
        ctrl.append(("Consumo negativo", f"{r.insumo} [{r.codigo}]", r.consumo_real))
    ctrl_df = pd.DataFrame(ctrl[1:], columns=ctrl[0])
    return {"Introducción Ejecutiva": intro_df, f"Consumo Neto {res['corto_local']}-Sin Obrador": n,
            "Consumo Global (Con Obrador)": c, "Benchmark Posta": m, "Auditoría de Desvíos": a,
            "Controles": ctrl_df, "Obrador Detalle": o, "Movimientos": movs, "Equivalencias Códigos": eqd}


def escribir_excel(res: dict[str, Any], destino: Path) -> None:
    hojas = _hojas(res)
    with pd.ExcelWriter(destino, engine="openpyxl") as xw:
        for nombre, df in hojas.items():
            df.to_excel(xw, sheet_name=nombre, index=False)
            ws = xw.sheets[nombre]
            head = PatternFill("solid", fgColor="1F3864")
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = head
                cell.alignment = Alignment(wrap_text=True, vertical="center")
            ws.freeze_panes = "A2" if nombre != "Introducción Ejecutiva" else None
            if nombre != "Introducción Ejecutiva" and len(df):
                ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(df.columns, 1):
                largo = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
                ws.column_dimensions[get_column_letter(i)].width = min(max(10, largo + 2), 60)
                nm = str(col)
                for cell in ws.iter_cols(min_col=i, max_col=i, min_row=2).__next__():
                    if isinstance(cell.value, float):
                        cell.number_format = ("0.0%" if "%" in nm else
                                              ('"$"#,##0' if ("$" in nm or "Precio" in nm or "Costo" in nm
                                                              or "Importe" in nm or "P_TRANSF" in nm) else "#,##0.00"))
            if nombre == "Introducción Ejecutiva":
                ws.column_dimensions["A"].width = 72
                ws.column_dimensions["B"].width = 20
                ws.column_dimensions["C"].width = 110
                for row in ws.iter_rows(min_row=2):
                    a, b, c = row[0], row[1], row[2]
                    if a.value and not b.value and (str(a.value).isupper() or str(a.value)[:1].isupper()
                                                    and not c.value):
                        a.font = Font(bold=True)
                        for cell in row:
                            cell.fill = PatternFill("solid", fgColor="D9E1F2")
                    c.alignment = Alignment(wrap_text=True, vertical="top")
                    if isinstance(b.value, float):
                        b.number_format = "0.0%" if abs(b.value) < 1.5 and "%" in str(a.value) else '"$"#,##0'
                    if isinstance(b.value, (int, float)) and "Unidades" in str(a.value):
                        b.number_format = "#,##0"


def escribir_html(res: dict[str, Any], destino: Path) -> None:
    k = _kpis(res)
    neto, obr, mat, aud = res["neto"], res["obrador"], res["matriz"], res["auditoria"]
    g = res["resumen_gestion"]
    mes_t = f"{MESES_ES[res['mes']].capitalize()} {res['anio']}"
    nombre, es_p, pl = res["nombre_local"], res["es_planta"], res["planta"]
    tol_txt = ", ".join(f"{r.title()} {v:.0%}" for r, v in MERMA_NATURAL.items())

    def tabla(df: pd.DataFrame, cols: list[tuple[str, str, str]]) -> str:
        th = "".join(f'<th style="text-align:{"left" if f == "t" else "right"}">{html.escape(h)}</th>' for h, _, f in cols)
        filas = []
        for r in df.itertuples(index=False):
            d = r._asdict()
            tds = []
            for _, c, f in cols:
                v = d[c]
                txt = html.escape(str(v)) if f == "t" else (pesos(v) if f == "$" else fnum(v) if f == "n" else
                                                              (f"{v * 100:+.0f}%" if pd.notna(v) else "-"))
                tds.append(f'<td style="text-align:{"left" if f == "t" else "right"}">{txt}</td>')
            filas.append("<tr>" + "".join(tds) + "</tr>")
        return f"<table><thead><tr>{th}</tr></thead><tbody>{''.join(filas)}</tbody></table>"

    crit = aud[aud["critico"] > 0].nlargest(10, "critico")
    just = aud[aud["justificado"] > 0].nlargest(6, "justificado")
    oper = aud[aud["clase"] == "Gasto operativo"].nlargest(6, "operativo")
    posta = mat.dropna(subset=["ahorro_posta"]).nlargest(8, "ahorro_posta")
    obr_top = obr.nlargest(6, "costo_real")
    kpi = [("Venta del mes", pesos(k["ventas_pesos"])),
           ("Consumo neto propio (sin Obrador)", pesos(k["costo_real_propio"])),
           ("Teórico por recetas (materia prima)", pesos(k["aud_teorico_mp"])),
           ("Faltante crítico (materia prima)", pesos(k["aud_critico_mp"])),
           ("Desperdicio justificado", pesos(k["aud_justificado_mp"] + k["aud_justificado_obr"])),
           ("Gasto operativo (aparte)", pesos(k["gasto_operativo"])),
           ("Obrador consumido (lista)", pesos(k["obrador_real"])),
           ("Ahorro potencial vs POSTA", pesos(k["ahorro_posta"]))]
    cards = "".join(f'<div class="k"><div class="v">{v}</div><div class="l">{html.escape(l)}</div></div>' for l, v in kpi)
    dif_obr = (k["obrador_ventas_arcos"] - k["obrador_remitos"]) if pd.notna(k["obrador_ventas_arcos"]) else None
    aviso_planta = ""
    if es_p and pl and pl.get("venta_total") and pl.get("venta_items"):
        cob = pl["venta_items"] / pl["venta_total"]
        extrap = f"teórico extrapolado al 100%: {pesos(k['aud_teorico_mp'] / cob)}" if cob else             "sin detalle por ítem: no se puede extrapolar el teórico"
        aviso_planta = (f"<p><b>Congreso = planta del Obrador.</b> Sus ventas son las ventas Obrador "
                        f"({pesos(pl['venta_total'])}); solo hay detalle por ítem del {cob:.0%} (remitos de Arcos y Lafinur, "
                        f"sin Cabello), por eso el teórico está subestimado y el faltante crítico sobreestimado "
                        f"({extrap}).</p>")
    doc = f"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Gestión {nombre} {mes_t}</title>
<style>body{{font-family:Segoe UI,Arial,sans-serif;color:#222;margin:0;background:#f5f6f8}}
.w{{max-width:820px;margin:auto;padding:16px}}h1{{font-size:20px;margin:8px 0}}h2{{font-size:15px;margin:22px 0 6px;color:#1F3864}}
.g{{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px}}
.k{{background:#fff;border-radius:8px;padding:10px;box-shadow:0 1px 2px #0002}}.v{{font-size:18px;font-weight:700}}.l{{font-size:11px;color:#666}}
table{{width:100%;border-collapse:collapse;background:#fff;font-size:12px}}th{{background:#1F3864;color:#fff;padding:5px}}
td{{padding:4px 6px;border-bottom:1px solid #e5e5e5}}p,li{{font-size:12px;line-height:1.45}}.n{{color:#666}}</style></head><body><div class="w">
<h1>Gestión comercial {nombre} &mdash; {mes_t}</h1>
<div class="g">{cards}</div>
<h2>Lectura rápida</h2>
{aviso_planta}
<ul>
<li>Materia prima propia auditada: costo real <b>{pesos(k['aud_real_mp'])}</b> contra <b>{pesos(k['aud_teorico_mp'])}</b> que explican las ventas.
Del desvío, <b>{pesos(k['aud_justificado_mp'])}</b> está justificado por desperdicio fichado, <b>{pesos(k['aud_natural_mp'])}</b> es merma natural
en frescos (tolerancia supuesta: {html.escape(tol_txt)}) y <b>{pesos(k['aud_critico_mp'])}</b> queda como faltante crítico sin justificar.</li>
<li>Gasto operativo (descartables, pastillas rational, hielo, bolsas de vacío...): <b>{pesos(k['gasto_operativo'])}</b>, informado aparte y fuera del desvío de cocina.
Hay {pesos(k['no_auditable'])} no auditables por unidades/escala inconsistentes.</li>
<li>Traspasos: {pesos(k['traspasos_out'])} salieron de {nombre} (no son consumo) y {pesos(k['traspasos_in'])} entraron.
{'Obrador: ' + pesos(k['obrador_ventas_arcos']) + ' vendidos a Arcos según su planilla, ' + pesos(k['obrador_remitos']) + ' itemizados en remitos (' + (pesos(dif_obr) if dif_obr is not None else '-') + ' sin detalle por ítem).' if not es_p else 'Los descartables dentro del BOM suman ' + pesos(k['descartables_bom']) + ' al teórico.'}</li>
<li>Consumo real a precios de POSTA: ahorro potencial de <b>{pesos(k['ahorro_posta'])}</b> sobre {k['posta_items']} insumos comparables
(neto {pesos(k['ahorro_posta_neto'])} compensando donde Posta es más caro).</li>
</ul>
<h2>Mayores faltantes críticos</h2>
{tabla(crit, [("Ítem", "insumo", "t"), ("Clasificación", "clase", "t"), ("Real", "real", "$"), ("Teórico", "teorico", "$"), ("Justificado", "justificado", "$"), ("Merma nat.", "natural", "$"), ("Faltante", "critico", "$")])}
<h2>Desperdicio justificado (fichado)</h2>
{tabla(just, [("Ítem", "insumo", "t"), ("Fichado", "desperdicio_d", "$"), ("Justificado", "justificado", "$")])}
<h2>Gasto operativo</h2>
{tabla(oper, [("Ítem", "insumo", "t"), ("Consumo", "consumo_real", "n"), ("Gasto", "operativo", "$")])}
<h2>Benchmark POSTA: mayor ahorro potencial</h2>
{tabla(posta, [("Insumo", "insumo", "t"), ("Precio real", "precio_real", "$"), ("POSTA", "precio_posta", "$"), ("Ahorro", "ahorro_posta", "$")])}
<h2>Obrador: ítems más consumidos</h2>
{tabla(obr_top, [("Ítem", "item_obrador", "t"), ("Consumo real", "consumo_real", "n"), ("Costo real", "costo_real", "$"), ("Teórico", "costo_teorico", "$")])}
<p class="n">Consumo real = stock inicial + compras + remitos de Obrador + traspasos entrantes &minus; traspasos salientes &minus; stock final.
Teórico = explosión de recetas (Madres incluidas) &times; ventas. Referencia de gestión: consumo $$$ informado {pesos(g['gest_consumo'])}.
Detalle y controles en Analisis_Gestion_Comercial_{nombre}_{res['mes']:02d}_{res['anio']}.xlsx.</p>
</div></body></html>"""
    destino.write_text(doc, encoding="utf-8")


# --------------------------------------------------------------------------- #
# Reporte de ítems huérfanos (sin ventas que los expliquen)
# --------------------------------------------------------------------------- #
def huerfanos(res: dict[str, Any]) -> pd.DataFrame:
    """Ítems cuyo consumo no explica ninguna venta: (local, mes, codigo, nombre, origen, detalle)."""
    filas: list[tuple[str, str, str, str, str]] = []
    mes_txt = f"{res['mes']:02d}/{res['anio']}"
    aud = res["auditoria"]
    if len(aud):
        for r in aud[aud["causa"].astype(str).str.startswith("Sin receta/venta")].itertuples(index=False):
            filas.append((str(r.codigo), str(r.insumo), "Auditoría de desvíos",
                          f"{r.clase} | consumo real {fnum(r.consumo_real, 2)} {r.unidad} | faltante {pesos(r.critico)}"))
    obr = res["obrador"]
    if len(obr):
        for r in obr[obr["alerta"].astype(str).str.startswith("sin ventas que lo expliquen")].itertuples(index=False):
            filas.append((str(r.codigo), str(r.item_obrador), "Obrador",
                          "sin receta/venta que lo use (alerta del Obrador)"))
    for r in res["sin_receta"].itertuples(index=False):
        sku = r.sku if isinstance(r.sku, str) and r.sku else "SIN SKU"
        filas.append((sku, str(r.articulo), "Venta sin receta/insumo",
                      f"{fnum(r.unidades, 2)} unidades | venta {pesos(r.venta)}"))
    df = pd.DataFrame(filas, columns=["codigo", "nombre", "origen", "detalle"])
    df.insert(0, "mes", mes_txt)
    df.insert(0, "local", res["nombre_local"])
    # una sola fila por (local, mes, código, nombre): los orígenes se unen en vez de duplicar
    return (df.groupby(["local", "mes", "codigo", "nombre"], as_index=False, sort=False)
            .agg(origen=("origen", lambda x: " + ".join(dict.fromkeys(x))),
                 detalle=("detalle", lambda x: " || ".join(dict.fromkeys(x)))))


class IndiceVentas:
    """Consulta de 'dónde aparece este código/nombre' sobre el índice de ventas (todas las planillas)."""

    def __init__(self, indice: pd.DataFrame | None, equiv: pd.DataFrame | None = None):
        self.vacio = indice is None or not len(indice)
        self.grupos: list[dict[str, Any]] = []
        self.cod_de_sku: dict[str, set[str]] = defaultdict(set)
        self.nombres_de_sku: dict[str, set[str]] = defaultdict(set)
        if equiv is not None and len(equiv):
            for r in equiv.itertuples(index=False):
                if r.sku:
                    if r.codigo:
                        self.cod_de_sku[r.sku].add(r.codigo)
                    self.nombres_de_sku[r.sku].add(nk(r.nombre))
        if self.vacio:
            return
        for (archivo, hoja), g in indice.groupby(["archivo", "hoja"], sort=False):
            codigos = {_cod(c): True for c in g["codigo"].dropna() if _cod(c)}
            self.grupos.append({"archivo": archivo, "hoja": hoja, "n": len(g), "codigos": set(codigos),
                                "g": g, "ks": sorted(set(g["k"]))})

    def evidencia(self, codigo: str, nombre: str, umbral: float = 90.0
                  ) -> tuple[list[str], list[str]]:
        """(donde aparece, donde NO aparece) como texto 'archivo › hoja'."""
        codigos = {_cod(codigo)} | self.cod_de_sku.get(codigo, set())
        claves = {nk(nombre)} | self.nombres_de_sku.get(codigo, set())
        claves.discard("")
        si, no = [], []
        for gr in self.grupos:
            etiqueta = f"{str(gr['archivo']).split('/')[-1]} › {gr['hoja']}"
            g = gr["g"]
            hit = g[g["codigo"].map(lambda c: _cod(c) in codigos if pd.notna(c) else False)]
            criterio = "por código"
            if not len(hit):
                ks = [k for k in claves if k in set(gr["ks"])]
                if not ks and gr["ks"]:
                    for k in claves:
                        m = process.extractOne(k, gr["ks"], scorer=fuzz.token_sort_ratio, score_cutoff=umbral)
                        if m:
                            ks.append(m[0])
                hit = g[g["k"].isin(ks)]
                criterio = "por nombre"
            if len(hit):
                nombres = ", ".join(dict.fromkeys(hit["articulo"].astype(str)))[:60]
                si.append(f"{etiqueta} ({criterio}: \"{nombres}\", {fnum(hit['unidades'].sum(), 1)} u., "
                          f"{pesos(hit['venta'].sum())})")
            else:
                no.append(etiqueta)
        return si, no


def escribir_huerfanos(resultados: list[dict[str, Any]], destino: Path, indice: pd.DataFrame | None = None,
                       fuentes: dict[str, Any] | None = None) -> int:
    """Vuelca a `destino` los ítems sin mapeo: para cada uno, en qué archivos de ventas SÍ aparece y en cuáles
    NO (por código o nombre), más los cruces automáticos aplicados. Devuelve la cantidad de ítems."""
    fuentes = fuentes or {}
    lineas = ["REPORTE DE ÍTEMS HUÉRFANOS - mapeo fallido",
              "=" * 78,
              "Ítems cuyo consumo no tiene ventas/recetas que lo expliquen. Para cada uno se indica en qué",
              "archivos de ventas se buscó el código o nombre y dónde NO apareció. Para corregirlo, pasame",
              "la equivalencia manual 'CÓDIGO -> SKU de receta / artículo vendido'.", ""]
    analizados = fuentes.get("analizados")
    if analizados is not None and len(analizados):
        lineas.append("ARCHIVOS DE VENTAS REVISADOS (con filas de ventas por artículo)")
        lineas += [f"  - {r.archivo} › {r.hoja}: {r.filas} filas | local: {r.local or '-'} | mes: {r.mes or '-'}"
                   for r in analizados.itertuples(index=False)]
    if fuentes.get("copias"):
        lineas.append("COPIAS IDÉNTICAS (se consultan una sola vez)")
        lineas += [f"  - {k}  ==  {', '.join(v)}" for k, v in fuentes["copias"].items()]
    for etq, clave in (("ARCHIVOS SIN VENTAS POR ARTÍCULO (no aportan al cruce)", "sin_ventas"),
                       ("ARCHIVOS CON ERROR AL LEER", "con_error")):
        if fuentes.get(clave):
            lineas.append(etq)
            lineas += [f"  - {x}" for x in fuentes[clave]]
    lineas.append("")
    total = 0
    for res in resultados:
        h = huerfanos(res)
        total += len(h)
        idx = IndiceVentas(indice, res.get("equivalencias"))
        lineas.append("#" * 78)
        lineas.append(f"## {res['nombre_local']} - {res['mes']:02d}/{res['anio']}  ({len(h)} ítems)")
        lineas.append("#" * 78)
        for r in h.itertuples(index=False):
            lineas.append(f"[{r.codigo}] {r.nombre}")
            lineas.append(f"    Origen : {r.origen} - {r.detalle}")
            if idx.vacio:
                lineas.append("    Ventas : sin índice de ventas disponible")
                continue
            si, no = idx.evidencia(r.codigo, r.nombre)
            for x in si:
                lineas.append(f"    ENCONTRADO EN: {x}")
            if no:
                lineas.append("    NO ENCONTRADO EN: " + " | ".join(no))
        cr = (res.get("ctrl_ventas") or {}).get("cruces_bom")
        if cr is not None and len(cr):
            lineas += ["", f"   Cruces automáticos aplicados ({len(cr)}): verificar que sean correctos",
                       f"   {'ARTÍCULO VENDIDO':<48}| {'SKU/RECETA':<18}| CRITERIO"]
            lineas += [f"   {str(r.articulo)[:47]:<48}| {str(r.sku)[:17]:<18}| {r.via} ({r.score:.0f})"
                       for r in cr.itertuples(index=False)]
        lineas.append("")
    try:
        destino.write_text("\n".join(lineas), encoding="utf-8")
    except PermissionError:       # la raíz puede no admitir archivos nuevos: se deja en WebTest/ (la app lo ve igual)
        alt = BASE_DIR / "WebTest" / destino.name
        LOG.warning("Sin permiso para crear %s: se guarda en %s", destino, alt)
        alt.write_text("\n".join(lineas), encoding="utf-8")
    return total


def generar(anio: int, mes: int, locales: list[str], out_dir: Path, indice_ventas: pd.DataFrame | None = None
            ) -> tuple[list[dict[str, Any]], list[str]]:
    """Analiza y escribe el Excel de cada local para el mes. Un local sin datos fuente no frena a los demás:
    queda en la lista de fallidos para que el consolidador lo cubra con lo que haya (facturas)."""
    resultados, fallidos = [], []
    for local in locales:
        nombre = LOCALES[local]["nombre"]
        LOG.info("=== %s %02d/%d ===", nombre, mes, anio)
        try:
            res = analizar(anio, mes, local, indice_ventas)
        except Exception as exc:  # noqa: BLE001 - un local sin fuentes no debe frenar a los demás
            LOG.warning("%s %02d/%d: sin fuentes suficientes para el análisis completo (%s: %s)",
                        nombre, mes, anio, type(exc).__name__, exc)
            fallidos.append(local)
            continue
        xlsx = out_dir / f"Analisis_Gestion_Comercial_{nombre}_{mes:02d}_{anio}.xlsx"
        htmlp = out_dir / f"reporte_gestion_{local}.html"
        try:
            escribir_excel(res, xlsx)
        except PermissionError:       # el Excel de destino está abierto: no se pisa, se deja al lado
            xlsx = xlsx.with_name(xlsx.stem + "_nuevo.xlsx")
            LOG.warning("El Excel de destino está abierto/bloqueado: se guarda como %s", xlsx.name)
            escribir_excel(res, xlsx)
        escribir_html(res, htmlp)
        k = _kpis(res)
        g = res["resumen_gestion"]
        LOG.info("Compras facturas %s vs Resumen %s | teórico MP %s | real MP auditado %s | justificado %s | "
                 "merma natural %s | faltante crítico %s", pesos(g["fact_mes_factura"]), pesos(g["gest_compras"]),
                 pesos(k["aud_teorico_mp"]), pesos(k["aud_real_mp"]), pesos(k["aud_justificado_mp"]),
                 pesos(k["aud_natural_mp"]), pesos(k["aud_critico_mp"]))
        LOG.info("Gasto operativo %s | traspasos salientes %s | ahorro vs POSTA %s | referencia en %s del costo real",
                 pesos(k["gasto_operativo"]), pesos(k["traspasos_out"]), pesos(k["ahorro_posta"]),
                 f"{k['ref_cobertura'] / k['consumo_real_total_pos']:.0%}" if k["consumo_real_total_pos"] else "-")
        LOG.info("Generado %s y %s", xlsx.name, htmlp.name)
        resultados.append(res)
    return resultados, fallidos


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--refresh", action="store_true", help="volver a bajar las planillas de Drive")
    ap.add_argument("--refresh-solo", nargs="+", metavar="CLAVE",
                    help="volver a bajar solo estas planillas (p. ej. arcos cabello), sin tocar el resto")
    ap.add_argument("--anio", type=int, default=2026)
    ap.add_argument("--mes", type=int, default=MES_BASE)
    ap.add_argument("--meses", type=int, nargs="+", help="varios meses (pisa --mes)")
    ap.add_argument("--out-dir", type=Path, default=BASE_DIR / "Reportes_Generados")
    ap.add_argument("--locales", nargs="+", default=list(LOCALES), choices=list(LOCALES),
                    help="sucursales a procesar (por defecto todas)")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")
    # Pasamano Obrador -> Chimidoha/locales: matriz de recargos y aprendizaje de la hoja 'Remitos Chimidoha' (si ya existe)
    import transferencias_obrador as tr
    remitos_chimi = tr.buscar_remitos_chimidoha()
    if remitos_chimi:
        apr = tr.aprender_de_remitos_chimidoha(remitos_chimi)
        LOG.info("Remitos Chimidoha: %s -> %s", [p.name for p in remitos_chimi], apr.get("por_tipo") or apr.get("estado"))
    else:
        LOG.info("Remitos Chimidoha: no hay archivo todavia; transferencias con la matriz por defecto "
                 "(panificados/descartables x1.20, carnes/pescados x1.10, almacen al costo, elaborados por BOM)")
    resultados: list[dict[str, Any]] = []
    for mes in (args.meses or [args.mes]):
        if args.refresh_solo:
            descargar(True, args.anio, mes, solo=args.refresh_solo)
        descargar(args.refresh, args.anio, mes)
        ok, _ = generar(args.anio, mes, args.locales, args.out_dir)
        resultados += ok
    args.out_dir.mkdir(parents=True, exist_ok=True)
    n = escribir_huerfanos(resultados, args.out_dir / "reporte_items_huerfanos.txt")
    LOG.info("reporte_items_huerfanos.txt: %d ítems sin mapeo", n)
    return 0


if __name__ == "__main__":
    sys.exit(main())
