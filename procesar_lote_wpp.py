"""Procesa una carpeta de comprobantes reales (Facturas_WPP/): clasifica, asigna local, cruza con el historial,
aprende, imputa en las COPIAS y concilia remitos / NC.

Niveles que ejecuta (ver CLAUDE.md):
  1 clasificación (comprobantes vs informes operativos vs listas de precios)
  2 local de destino por sello / edificio Congreso 1602 (PB Chimidoha, PA Congreso=Obrador) / trampa 'Lafinur 3368 SRL'
  3 retro-entrenamiento por N° de factura (renglón a renglón o AGRUPADO_TOTAL)
  4 remitos de verdulería con birome ('NO', kilos de balanza) + conciliación con la factura de Raíces
  5 prorrateo de impuestos internos/cargos (Quilmes) y Notas de Crédito (caso A devolución / caso B diferencia de precio)

Por defecto ESCRIBE solo en Copias_Gestion/<Local>_FACTURAS_COPIA.xlsx (Chimidoha se crea desde Congreso),
aprendizaje_proveedores.json, remitos_pendientes.json y cache_facturas.json. Nunca toca planillas reales: eso
solo con `imputacion_facturas.push_drive(..., confirmado=True)` / `--push-drive --confirmo-escribir-produccion`.

    python procesar_lote_wpp.py                  # Facturas_WPP/, escribe en copias
    python procesar_lote_wpp.py --dry-run        # solo calcula y muestra
    python procesar_lote_wpp.py --push-drive     # muestra el plan exacto para producción (no escribe)
"""

from __future__ import annotations

import argparse
import time
import hashlib
import json
import logging
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))
sys.path.insert(0, str(BASE_DIR / "Precios"))

import grupos_telegram                       # noqa: E402
import imputacion_facturas as imp          # noqa: E402
import invoice_parser as ip                 # noqa: E402

LOG = logging.getLogger("procesar_lote_wpp")
IMG_EXT = {".jpg", ".jpeg", ".png", ".webp"}
DESTINO_PRECIOS = BASE_DIR / "Precios"


def carpeta_por_defecto() -> Path:
    """'Facturas Mishiguene/' (sin guiones); si todavía no existe se lee la carpeta vieja 'Facturas_WPP/'."""
    nueva, vieja = BASE_DIR / "Facturas Mishiguene", BASE_DIR / "Facturas_WPP"
    return nueva if nueva.is_dir() or not vieja.is_dir() else vieja


def sha(path: Path) -> str:
    return ip.sha256_archivo(path)


# --------------------------------------------------------------------------- #
# Nivel 1: clasificación
# --------------------------------------------------------------------------- #
def separar_destinos_remito(comp: dict) -> list[dict]:
    """Una anotación de local por renglón no cambia el destino de todo el documento."""
    if comp.get("tipo") != "Remito" or not any(it.get("local_destino") for it in comp.get("items", [])):
        return [comp]
    destinos = {}
    for it in comp["items"]:
        local = it.get("local_destino") or comp.get("local_destino_base")
        destinos.setdefault(local, []).append(it)
    out = []
    for local, items in destinos.items():
        c = dict(comp)
        c.update({"items": items, "archivo_origen": comp["archivo"],
                  "archivo": comp["archivo"] + f" [destino {local or 'REVISAR'}]",
                  "local_destino_renglon": local})
        out.append(c)
    return out


def _asignacion_manual(numero, proveedor=None) -> str | None:
    """estado/asignaciones.json: {'<N° significativo>': {'local': 'Congreso', 'proveedor': 'rupajo'}} (comando /asignar)."""
    try:
        d = json.loads((BASE_DIR / "estado" / "asignaciones.json").read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    n = imp.ultimo_entero(numero)
    r = d.get(str(n)) if n is not None else None
    if r and (not r.get("proveedor") or r["proveedor"] in imp.norm(proveedor or "")):
        return r["local"]
    return None


def clasificar_carpeta(carpeta: Path, escribir_cache: bool = True) -> dict:
    """Devuelve {'comprobantes': [...], 'informes': [Path], 'listas': [Path], 'duplicados': [(Path, Path)]}."""
    out = {"comprobantes": [], "informes": [], "listas": [], "duplicados": [], "otros": []}
    vistos: dict[str, Path] = {}
    cache = ip._leer_cache()
    for p in sorted(carpeta.iterdir()):
        if not p.is_file() or p.name.startswith("~$"):
            continue
        ext = p.suffix.lower()
        if ext not in IMG_EXT and ext not in (".pdf", ".xlsx", ".xls"):       # archivo sin extensión (copia de WhatsApp)
            cab = p.read_bytes()[:4]
            ext = ".jpeg" if cab[:2] == bytes([255, 216]) else (".pdf" if cab == b"%PDF" else ext)
        h = sha(p)
        if h in vistos:
            out["duplicados"].append((p, vistos[h]))
            continue
        vistos[h] = p
        if ext in (".xlsx", ".xls"):
            out["informes"].append(p)
        elif ext == ".pdf" or (ext == "" and "pdf" in p.name.lower()):
            if re.search(r"grupo\s*free|lista", p.name, re.I):
                out["listas"].append(p)
                continue
            c = ip.leer_pdf_comprobante(p)
            if c["tipo"] == "Lista":
                out["listas"].append(p)
            else:
                c["sha256"] = h
                out["comprobantes"].append(c)
        elif ext in IMG_EXT:
            if escribir_cache and h not in cache and not any(k.startswith(h + "#") for k in cache):
                try:                                       # misma foto recomprimida por Telegram: match por parecido + OCR
                    reutil = ip.reutilizar_lectura_por_parecido(p)
                except Exception as e:
                    LOG.debug("parecido visual: %s", e)
                    reutil = []
                if reutil:
                    cache = ip._leer_cache()
                    LOG.info("%s: lectura reutilizada de la caché (parecido visual + OCR)", p.name)
            multi = [v for k, v in cache.items() if k.startswith(h + "#")]
            if h in cache:
                c = dict(cache[h])
                c["archivo"] = p.name
            elif multi:                                    # foto con 2+ comprobantes ya leída
                for k, v in enumerate(multi, 1):
                    cc = dict(v)
                    cc["archivo"] = f"{p.name} [comprobante {k}]"
                    cc["sha256"] = h
                    out["comprobantes"].append(cc)
                continue
            else:
                try:
                    if not escribir_cache:
                        raise RuntimeError("dry-run: imagen sin lectura en caché; requiere extracción")
                    leidos = ip.leer_imagen_vision(p)
                    LOG.info("%s: leída con visión multimodal (%d comprobante/s)", p.name, len(leidos))
                    for k, v in enumerate(leidos, 1):
                        cc = dict(v)
                        cc["archivo"] = p.name if len(leidos) == 1 else f"{p.name} [comprobante {k}]"
                        cc["sha256"] = h
                        out["comprobantes"].append(cc)
                    continue
                except Exception as e:                     # sin `claude`/sin red: queda el OCR de Tesseract
                    LOG.warning("%s: visión no disponible (%s); se usa OCR", p.name, str(e)[:100])
                LOG.warning("%s: no está en cache_facturas.json; se intenta OCR (baja confianza)", p.name)
                try:
                    c = ip.procesar_comprobante(p, usar_cache=escribir_cache)
                    c.setdefault("tipo", ip.tipo_comprobante(c.get("texto_encabezado") or ""))
                    c["advertencias"] = (c.get("advertencias") or []) + ["leída por OCR sin revisión visual"]
                except ip.OCRNoDisponible as e:
                    c = {"archivo": p.name, "tipo": "Ilegible", "items": [], "advertencias": [str(e)], "sha256": h}
            c["sha256"] = h
            out["comprobantes"].append(c)
        else:
            out["otros"].append(p)
    out["comprobantes"] = [parte for c in out["comprobantes"] for parte in separar_destinos_remito(c)]
    # duplicados semánticos (la misma factura en PDF y en foto): se conserva el PDF y se anota
    por_clave: dict[tuple, dict] = {}
    finales = []
    for c in out["comprobantes"]:
        k = (imp.ultimo_entero(c.get("numero_factura")), imp.norm(c.get("proveedor_detectado")).split(" ")[0],
             c.get("fecha_factura"), c.get("tipo"), c.get("local_destino_renglon"))
        if k[0] is not None and k in por_clave:
            previo = por_clave[k]
            for campo in ("sello", "firmas", "anotaciones_manuscritas", "direccion_entrega"):   # se queda con el que trae sello
                if not previo.get(campo) and c.get(campo):
                    previo[campo] = c[campo]
            previo.setdefault("duplicado_en", []).append(c["archivo"])
            c["duplicado_de"] = previo["archivo"]
            out["duplicados"].append((Path(c["archivo"]), Path(previo["archivo"])))
            continue
        por_clave[k] = c
        finales.append(c)
    out["comprobantes"] = finales
    return out


# --------------------------------------------------------------------------- #
# Procesamiento por comprobante
# --------------------------------------------------------------------------- #
def _alias_prov(ap: dict, detectado: str | None) -> str | None:
    return (ap.get("alias_proveedores") or {}).get(imp.norm(detectado))


def analizar(comp: dict, locales: dict, ap: dict) -> dict:
    """Asigna local y busca en el historial. No escribe nada."""
    notas = comp.get("anotaciones_manuscritas") or []
    regla = ip.detectar_local(comp.get("sello"), comp.get("direccion_entrega"), comp.get("razon_social_cliente"),
                              " | ".join(notas), None)
    matches = imp.buscar_en_historial(comp, locales) if comp.get("tipo") in ("Factura", "NC") else []
    matches.sort(key=lambda m: (-(m["modo"] in ("FILAS", "AGRUPADO_TOTAL")), -m["similitud_proveedor"]))
    m0 = matches[0] if matches else None
    if m0 and (m0["similitud_proveedor"] < 60 or not comp.get("proveedor_detectado")
               or str(comp["proveedor_detectado"]).startswith("(")):
        comp["proveedor_inferido_historial"] = m0["proveedor_planilla"]      # encabezado cortado o ilegible
        comp["proveedor_detectado"] = m0["proveedor_planilla"]
    if m0:
        comp["proveedor_planilla"] = m0["proveedor_planilla"]       # nombre con el que figura en la planilla
    if "local_destino_renglon" in comp:
        regla = {"local": comp["local_destino_renglon"], "regla": "destino por renglón: entrega base y anotación específica", "alertas": []}
    asig = _asignacion_manual(comp.get("numero_factura"), comp.get("proveedor_detectado"))
    if asig:                                                 # /asignar desde Telegram (decisión humana validada)
        regla = {"local": asig, "regla": "asignación manual por Telegram (/asignar)", "alertas": regla.get("alertas", [])}
    local, por = regla["local"], regla["regla"]
    alertas = list(regla["alertas"])
    # grupo de Telegram del local = contexto validado (cuando el sello no alcanza); si el sello contradice al grupo, alerta
    gl = grupos_telegram.local_de_archivo(comp.get("archivo"))
    comp["grupo_local"] = gl
    if gl and local is None:
        local, por = gl, "grupo Telegram del local"
        alertas = [a for a in alertas if not a.startswith("local INDETERMINADO")]
    elif gl and local and local != gl and not asig:
        alertas.append(f"ALERTA_GRUPO: el sello/anotación indica {local} pero el comprobante entró por el grupo de {gl}; "
                       f"se respeta el sello ({por})")
    if local is None and m0:
        local, por = m0["local"], f"historial: ya cargada en la planilla de {m0['local']} (el comprobante no trae sello)"
        alertas = [a for a in alertas if not a.startswith("local INDETERMINADO")]
    if local is None and "Chimidoha" in locales and "chimi" in imp.norm(comp.get("razon_social_cliente")):
        if imp.norm(comp.get("proveedor_detectado")) in locales["Chimidoha"]._idx_prov or \
                imp.norm(comp.get("proveedor_planilla")) in locales["Chimidoha"]._idx_prov:
            local, por = "Chimidoha", "proveedor con historial cargado directo en Chimidoha y razón social Chimi"
    if m0 and regla["local"] and m0["local"] != regla["local"]:
        alertas.append(f"CARGADA_EN_{m0['local'].upper()} pero {por} → corresponde {regla['local']}")
    if comp.get("remito_asociado"):
        alertas.append(f"remito asociado {comp['remito_asociado']}: el remito no vino en el lote (sin birome para conciliar)"
                       if comp["remito_asociado"] not in comp.get("_remitos_lote", ()) else "")
    extra = comp.get("extra") or {}
    if extra.get("multilocal"):
        alertas.append("MULTILOCAL (100 % en Congreso, etiqueta MULTILOCAL): " + extra["multilocal"])
    if comp.get("sello") is None and comp.get("tipo") in ("Factura", "NC") and not (comp.get("motor") or "").startswith("pdf"):
        alertas.append("sin sello de recepción")
    # aprendizaje de sellos: el local quedó validado por /asignar, grupo o historial → se guarda la variante sucia leída
    if local and comp.get("sello") and (asig or por == "grupo Telegram del local" or (m0 and m0["local"] == local)):
        try:
            if ip.local_por_patron_sello(comp.get("sello"))[0] != local:
                ip.aprender_variante_sello(local, comp.get("sello"), comp.get("firmas"), fuente=por[:40])
        except Exception:                                    # noqa: BLE001
            pass
    return {"regla": regla, "matches": matches, "m0": m0, "local": local, "por": por, "alertas": alertas}


def estado_de(comp: dict, an: dict, remito_conciliado: bool = False) -> str:
    t = comp.get("tipo")
    if t == "RemitoObrador":
        return f"REMITO_OBRADOR → {an['local']}"
    if t in ("Desconocido", "Ilegible", "ND"):
        return "TIPO_REVISAR"
    if t == "Remito":
        if an["local"] is None:
            return "REMITO_SIN_LOCAL_REVISAR"
        return "REMITO_PENDIENTE" + (" → CONCILIADO" if remito_conciliado else "")
    if t == "Captura_planilla":
        return "EVIDENCIA_PLANILLA"
    m0 = an["m0"]
    if t == "NC":
        return "NC_APLICADA" + (" (ya en planilla)" if m0 else " (nueva)")
    if m0:
        base = ("YA_CARGADA_SUMA_AGRUPADA" if m0["modo"] == "AGRUPADO_TOTAL" else
                "YA_CARGADA_FILAS" if m0["modo"] == "FILAS" else "YA_CARGADA_REVISAR_MONTO")
        if an["regla"]["local"] and an["regla"]["local"] != m0["local"]:
            base += f" [en {m0['local']}: reasignar a {an['regla']['local']}]"
        return base
    if not comp.get("numero_factura") or not comp.get("items"):
        return "NUEVA_INCOMPLETA_REVISAR"
    if an["local"] is None:
        return "NUEVA_SIN_LOCAL_REVISAR"
    return "NUEVA_IMPUTADA"


def _num_cel(x):
    return imp.numero_a_celda(x)


def escribir_carpeta_revision(filas: list[dict], analisis: dict, cl: dict) -> Path:
    """'Carpeta Revision/Revision Facturas IA.xlsx': espejo de TODAS las filas pintadas por el agente que siguen en las
    copias (amarillo #FFF8E1 = OK pre-aprobada, rojo #FFEBEE = a revisar) + alertas por comprobante."""
    import openpyxl
    from openpyxl.styles import PatternFill, Font
    carpeta = BASE_DIR / "Carpeta Revision"
    carpeta.mkdir(exist_ok=True)
    nuevas = {(f["local"], f["fila_excel"]): f for f in filas}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Filas inyectadas"
    ws.append(["Local", "Fila en la copia", "N° factura", "Fecha", "Proveedor", "Presentación / marca", "Cantidad", "Precio",
               "Desc.", "Estado", "Nota o motivo", "Precio transferencia P_UN"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for copia in sorted((imp.COPIAS_DIR).glob("*_FACTURAS_COPIA.xlsx")):
        local = copia.name.replace("_FACTURAS_COPIA.xlsx", "")
        hoja = openpyxl.load_workbook(copia)["Facturas"]
        for r in range(2, hoja.max_row + 1):
            col = (hoja.cell(r, 1).fill.fgColor.rgb or "") if hoja.cell(r, 1).fill and hoja.cell(r, 1).fill.fill_type else ""
            col = str(col)[-6:].upper()
            if col not in ("FFF8E1", "FFEBEE") or hoja.cell(r, 1).value in (None, ""):
                continue
            n = nuevas.get((local, r)) or {}
            rojo = col == "FFEBEE"
            ws.append([local, r, hoja.cell(r, 1).value, hoja.cell(r, 2).value, hoja.cell(r, 4).value, hoja.cell(r, 5).value,
                       hoja.cell(r, 7).value, hoja.cell(r, 8).value, hoja.cell(r, 9).value,
                       "REVISAR" if rojo else "OK (pre-aprobada)",
                       hoja.cell(r, 6).value or "; ".join((n.get("motivos") or "",)), n.get("transferencia")])
            for c in ws[ws.max_row]:
                c.fill = PatternFill("solid", fgColor=col)
    wa = wb.create_sheet("Alertas por comprobante")
    wa.append(["Archivo", "Tipo", "N°", "Proveedor", "Local", "Regla de local", "Alerta"])
    for c in cl["comprobantes"]:
        an = analisis.get(c["archivo"])
        if not an:
            continue
        for al in an["alertas"]:
            if al:
                wa.append([c["archivo"], c.get("tipo"), c.get("numero_factura"), c.get("proveedor_detectado"),
                           an["local"], an["por"], al])
    destino = carpeta / "Revision Facturas IA.xlsx"
    wb.save(destino)
    return destino


def main(argv: list[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s", datefmt="%H:%M:%S")
    ap_ = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap_.add_argument("carpeta", nargs="?", default=str(carpeta_por_defecto()))
    ap_.add_argument("--dry-run", action="store_true", help="No escribe ningún archivo")
    ap_.add_argument("--copias-dir", help="Carpeta de copias (pruebas)")
    ap_.add_argument("--push-drive", action="store_true", help="Muestra el plan para la planilla real (no escribe)")
    ap_.add_argument("--confirmo-escribir-produccion", action="store_true",
                     help="Junto con --push-drive: ESCRIBE en las planillas reales (solo con tu OK explícito)")
    ap_.add_argument("--json", help="Guardar el reporte en este .json")
    args = ap_.parse_args(argv)
    if args.copias_dir:
        imp.COPIAS_DIR = Path(args.copias_dir)
    carpeta = Path(args.carpeta)
    escribir = not args.dry_run
    reporte, locales, ap, cl = ejecutar(carpeta, escribir, args)
    if args.json:
        Path(args.json).write_text(json.dumps(reporte, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    return _imprimir(reporte, locales, ap, args, cl)


_ULTIMO_DHASH = 0.0


def _log_liviano(filas: list[dict], carpeta_nombre: str = "") -> None:
    """CSV liviano del camino rápido (sin abrir ni guardar ningún .xlsx): Carpeta Revision/log_camino_rapido.csv."""
    import csv
    dest = BASE_DIR / "Carpeta Revision"
    dest.mkdir(exist_ok=True)
    f = dest / "log_camino_rapido.csv"
    nuevo = not f.exists()
    with open(f, "a", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        if nuevo:
            w.writerow(["ts", "local", "N", "proveedor", "producto", "original", "G", "H", "I", "estado", "motivos"])
        ts = datetime.now().isoformat(timespec="seconds")
        for r in filas:
            w.writerow([ts, r.get("local"), r.get("N"), r.get("proveedor"), r.get("presentacion"), r.get("original"),
                        r.get("G"), r.get("H"), r.get("I"), r.get("estado"), r.get("motivos")])


def ejecutar(carpeta: Path, escribir: bool, args, indexar_informes: bool = True, forzar_nuevas: bool = False,
             demo: bool = False, escribir_xlsx_local: bool | None = None):
    """Núcleo reutilizable (lo usa también telegram_ingestor.py). `args` necesita .push_drive y
    .confirmo_escribir_produccion. Devuelve (reporte, locales, ap, clasificación).

    `escribir_xlsx_local`: False = CAMINO RÁPIDO (imputa en RAM, guarda aprendizaje + CSV liviano y escribe directo en
    Google Sheets; no abre ni guarda los Copias_Gestion/*_FACTURAS_COPIA.xlsx de ~6 MB). None = False cuando hay
    push_drive y True en el resto (CLI clásico)."""
    global _ULTIMO_DHASH
    t_ini = time.time()
    if escribir_xlsx_local is None:
        escribir_xlsx_local = not bool(getattr(args, "push_drive", False))
    rapido = not escribir_xlsx_local
    # ---- Nivel 1 ------------------------------------------------------------
    try:
        if escribir and (not rapido or time.time() - _ULTIMO_DHASH > 600):
            _ULTIMO_DHASH = time.time()
            ip.indexar_dhash(BASE_DIR / "Facturas Mishiguene")     # huella visual de las lecturas ya guardadas
    except Exception as e:
        LOG.debug("indexar_dhash: %s", e)
    cl = clasificar_carpeta(carpeta, escribir_cache=escribir)
    t_leido = time.time()
    informes_idx, bench = None, None
    if escribir and indexar_informes:
        import datos_operativos
        informes_idx = datos_operativos.indexar(carpeta)
    for pdf in cl["listas"]:
        if re.search(r"grupo\s*free", pdf.name, re.I):
            import benchmark_grupo_free as bg
            if escribir:
                shutil.copy2(pdf, DESTINO_PRECIOS / pdf.name)
                bench = bg.guardar(DESTINO_PRECIOS / pdf.name)
            else:
                bench = bg.parsear_lista(pdf)

    # ---- historial + aprendizaje --------------------------------------------
    locales = imp.locales_todos()
    ap = imp.cargar_aprendizaje()
    if not ap.get("mapeos"):
        ap = imp.bootstrap_aprendizaje()
    perfil = imp.perfil_historico_proveedores(locales)
    modos = ap.setdefault("modos_proveedor", {})
    for prov, d in perfil.items():
        e = modos.setdefault(prov, {"modo_carga": "AGRUPADO_TOTAL", "proveedor_oficial": prov, "conceptos": {},
                                     "facturas_lote": []})
        e.update({"modo_carga": "AGRUPADO_TOTAL", "concepto": d["concepto"],
                  "historico": {"facturas": d["facturas"], "una_fila": d["una_fila"]}})

    analisis: dict[str, dict] = {}
    claves_remitos = {imp._clave_remito(x.get("numero_factura")) for x in cl["comprobantes"] if x.get("tipo") == "Remito"}
    for c in cl["comprobantes"]:
        c.setdefault("tipo", ip.tipo_comprobante(c.get("texto_encabezado") or ""))
        c.setdefault("extra", {})
        c["validacion"] = c.get("validacion") or {}
        c["_remitos_lote"] = {c["remito_asociado"]} if c.get("remito_asociado") and             imp._clave_remito(c["remito_asociado"]) in claves_remitos else set()
        analisis[c["archivo"]] = analizar(c, locales, ap)
        analisis[c["archivo"]]["alertas"] = [a for a in analisis[c["archivo"]]["alertas"] if a]
        if bench and "free" in imp.norm(c.get("proveedor_detectado")) and c.get("items"):
            import benchmark_grupo_free as bg
            for it in c["items"]:
                au = bg.auditar_precio(it["descripcion_original"],
                                       it["precio_unitario"] * (1 - it.get("descuento_pct", 0) / 100), bench)
                if au["estado"] != "SIN_REFERENCIA":
                    analisis[c["archivo"]]["alertas"].append(
                        f"benchmark Grupo Free: {it['descripcion_original'][:38]} ${au['lista']:,.0f}/kg lista → "
                        f"{au['estado']} ({au['dif_pct']:+.1f} % con la bonificación)")
    aprendizajes: dict[str, dict] = {}
    for c in cl["comprobantes"]:
        an = analisis[c["archivo"]]
        if an["m0"] and c.get("tipo") in ("Factura", "NC"):
            aprendizajes[c["archivo"]] = imp.aprender_de_historial(ap, c, an["m0"])
    if escribir:
        imp.guardar_aprendizaje(ap)

    # ---- simulación sobre lo ya cargado (validación del motor contra la carga humana) -------------------
    simulaciones: dict[str, dict] = {}
    for c in ([] if rapido else cl["comprobantes"]):
        an = analisis[c["archivo"]]
        m0 = an["m0"]
        if not m0 or c.get("tipo") not in ("Factura", "NC") or not c.get("items"):
            continue
        loc = locales[m0["local"]]
        res = imp.construir_filas(c, loc, ap)
        sim_total = round(sum((f["J"] or 0) for f in res["filas"]), 2)
        real = m0["filas"]
        real_total = round(float((real["CANT"].astype(float) * real["PRECIO"].astype(float) * (1 - real["DESC"].astype(float))).sum()), 2)
        pres_sim = sorted(str(f["presentacion"]) for f in res["filas"] if (f["J"] or 0) != 0 or f["accion"] != "INYECTAR")
        pres_real = sorted(str(x) for x in real["PRODUCTO"] if True)
        simulaciones[c["archivo"]] = {
            "modo_motor": res["modo"], "modo_historial": m0["modo"], "filas_motor": len(res["filas"]),
            "filas_planilla": len(real), "total_motor": sim_total, "total_planilla": real_total,
            "coincide_total": abs(sim_total - real_total) <= max(1.0, 0.01 * abs(real_total)),
            "presentaciones_motor": [f["presentacion"] for f in res["filas"]], "presentaciones_planilla": pres_real,
            "caso_nc": next((f.get("tipo_nc") for f in res["filas"] if f.get("tipo_nc")), None),
            "H_motor": [f["H"] for f in res["filas"]], "H_planilla": [float(x) for x in real["PRECIO"]],
            "I_motor": [f["I"] for f in res["filas"]], "I_planilla": [float(x) for x in real["DESC"]]}

    # ---- remitos -------------------------------------------------------------
    remitos = imp.cargar_remitos()
    facturas_por_remito = {
        (imp._clave_remito(c.get("remito_asociado")), imp.norm(c.get("proveedor_detectado"))): c
        for c in cl["comprobantes"] if c.get("tipo") == "Factura" and c.get("remito_asociado")}
    # El número de remito no es único entre proveedores. Mantener identidad y local.
    def identidad_remito(reg):
        return (imp._clave_remito(reg.get("remito")) or reg.get("archivo"), imp.norm(reg.get("proveedor")), reg.get("local"))
    def conciliar_explicito(reg):
        fac = facturas_por_remito.get((imp._clave_remito(reg.get("remito")), imp.norm(reg.get("proveedor"))))
        if fac and reg.get("local") and analisis[fac["archivo"]]["local"] == reg["local"]:
            conc = imp.conciliar_remito_factura(reg, fac)
            reg.update({"estado": "CONCILIADO", "factura_definitiva": fac.get("numero_factura"),
                        "alertas": conc["alertas"], "conciliacion": conc})
        return fac
    estado_remitos_antes = {k_: v_.get("estado") for k_, v_ in remitos.items()}
    # También conciliar los guardados cuando la factura llega en un lote posterior.
    for previo in remitos.values():
        if previo.get("estado") == "PENDIENTE":
            conciliar_explicito(previo)
    remitos_lote: list[dict] = []
    for c in cl["comprobantes"]:
        if c.get("tipo") != "Remito":
            continue
        reg = imp.procesar_remito(c)
        reg["local"] = analisis[c["archivo"]]["local"]
        loc_r = locales.get(reg["local"])
        if loc_r is not None:
            imp.completar_remito(reg, loc_r, ap)
            reg["proveedor_oficial"] = imp.mapear_proveedor(reg["proveedor"], "", loc_r.proveedores)[0]
        else:
            reg["alertas"].append("REVISAR: local indeterminado; no se carga en un local por defecto")
        prev_key = next((k for k, v in remitos.items() if identidad_remito(v) == identidad_remito(reg) or (v.get("archivo") and v.get("archivo") == reg.get("archivo"))), None)
        if prev_key is not None and remitos[prev_key].get("estado") == "CONCILIADO":
            for campo in ("estado", "factura_definitiva", "conciliacion", "alertas"):
                if campo in remitos[prev_key]:
                    reg[campo] = remitos[prev_key][campo]
        conciliar_explicito(reg)
        if reg["estado"] == "PENDIENTE" and loc_r is not None:
            ya = imp.remito_ya_cargado_como_total(reg, loc_r)
            if ya:
                reg["posible_factura"] = ya
                reg["alertas"].append(f"REVISAR conciliación: factura {ya['N']} coincide aproximadamente en fecha/monto; falta referencia explícita")
        c["_remito"] = reg
        key = prev_key or "|".join(str(x or "") for x in identidad_remito(reg))
        remitos[key] = reg
        remitos_lote.append(reg)
    if escribir and remitos:
        imp.guardar_remitos(remitos)

    # ---- remitos del Obrador a un local (5 columnas de la pestaña 'Remitos' del receptor) -----------
    remitos_obrador_por_local: dict[str, list[dict]] = {}
    for c in cl["comprobantes"]:
        if c.get("tipo") != "RemitoObrador":
            continue
        receptor = analisis[c["archivo"]]["local"]
        if receptor:
            loc_rx = locales.get(receptor) or imp.obtener_local(receptor)
            reg = imp.procesar_remito_obrador(c, loc_rx, ap)
            c["_remito_obrador"] = reg
            remitos_obrador_por_local.setdefault(receptor, []).append(reg)

    # ---- imputación de lo NUEVO ---------------------------------------------
    plan_por_local: dict[str, list[dict]] = {}
    marcas_por_local: dict[str, list[tuple[int, str]]] = {}
    filas_por_comp: dict[str, list[dict]] = {}
    for c in cl["comprobantes"]:
        an = analisis[c["archivo"]]
        t = c.get("tipo")
        if t not in ("Factura", "NC", "Remito"):
            continue
        if t == "Remito":
            continue                           # el remito vive en remitos_pendientes (ver arriba)
        destino = an["local"]
        reasignar = bool(an["m0"] and an["regla"]["local"] and an["m0"]["local"] != an["regla"]["local"]
                         and an["regla"]["local"] == "Chimidoha" and not forzar_nuevas)
        if an["m0"] and not reasignar and not forzar_nuevas:
            continue                           # ya está en la planilla (DEMO/PRUEBA lo fuerza para mostrarlo en vivo)
        if reasignar and not forzar_nuevas and any(m["local"] == "Chimidoha" for m in an["matches"]):
            continue                           # ya reasignada a Chimidoha en una corrida anterior
        if destino is None or not c.get("items") or not c.get("numero_factura"):
            c.setdefault("advertencias", []).append(
                "no se imputa: " + ("sin local" if destino is None else "comprobante incompleto (sin N° o sin renglones)"))
            continue
        loc = imp.obtener_local(destino) if destino not in locales else locales[destino]
        res = imp.construir_filas(c, loc, ap)
        filas_por_comp[c["archivo"]] = res["filas"]
        plan_por_local.setdefault(destino, []).extend(res["filas"])
        if reasignar:
            origen = an["m0"]["local"]
            marcas_por_local.setdefault(origen, []).extend(
                (int(i) + 2, f"REASIGNAR A CHIMIDOHA: sello Café Mishiguene Núñez (factura {c['numero_factura']})")
                for i in an["m0"]["filas"].index)

    # remitos pendientes: viven en la pestaña 'Remitos' (copia y planilla real); el gasto entra con la factura

    # ---- escritura en copias -------------------------------------------------
    escrituras = {}
    revision_filas: list[dict] = []
    if escribir and not demo:
        remitos_tab_por_local: dict[str, list[dict]] = {}
        locales_por_archivo = {c["archivo"]: analisis[c["archivo"]]["local"] for c in cl["comprobantes"]}
        for reg in remitos.values():
            n = reg.get("local") or locales_por_archivo.get(reg.get("archivo"))
            if n in locales:
                reg["local"] = n
                remitos_tab_por_local.setdefault(n, []).append(reg)
        if remitos:
            imp.guardar_remitos(remitos)
        for nombre in sorted(set(plan_por_local) | set(marcas_por_local) | set(remitos_tab_por_local)
                             | set(remitos_obrador_por_local)):
            loc = locales.get(nombre) or imp.obtener_local(nombre)
            if rapido:                          # sin .xlsx: filas del plan en memoria + CSV liviano
                rep = {"escritas": len(plan_por_local.get(nombre, [])), "verificadas": 0,
                       "filas": [(None, f) for f in plan_por_local.get(nombre, [])]}
            else:
                rep = imp.aplicar_a_copia([{"filas": plan_por_local.get(nombre, [])}], loc,
                                          remitos=remitos_tab_por_local.get(nombre),
                                          marcar=marcas_por_local.get(nombre),
                                          remitos_obrador=remitos_obrador_por_local.get(nombre))
            escrituras[nombre] = {"copia": None if rapido else loc.copia.name, "filas": rep["escritas"],
                                  "verificadas": rep["verificadas"], "marcadas": len(marcas_por_local.get(nombre, []))}
            for r_, f_ in rep["filas"]:
                revision_filas.append({"local": nombre, "fila_excel": r_, "N": f_.get("A"), "fecha": f_.get("B"),
                                       "proveedor": f_.get("D"), "presentacion": f_.get("presentacion"),
                                       "original": f_.get("descripcion_original"), "G": f_.get("G") or f_.get("cant_orig"),
                                       "H": f_.get("H") or f_.get("precio_orig"), "I": f_.get("I"),
                                       "estado": "REVISAR" if f_.get("accion") == "REVISAR" else "OK (pre-aprobada)",
                                       "motivos": "; ".join((f_.get("motivos") or []) + (f_.get("avisos") or [])),
                                       "transferencia": (f_.get("transferencia") or {}).get("p_un_transferencia")})
    if escribir and not demo:
        if rapido:
            _log_liviano(revision_filas)
        else:
            escribir_carpeta_revision(revision_filas, analisis, cl)
    pushes = {}
    tiempos_push: dict = {}
    if args.push_drive and not getattr(args, "dry_run", False):
        import registro_fallos as rf
        from concurrent.futures import ThreadPoolExecutor
        proveedores_por_local = {}
        # camino rápido: solo los remitos de este lote o cuyo estado cambió (conciliación); el resto ya está en la hoja
        ids_lote = {id(r) for r in remitos_lote}
        for k_, reg in remitos.items():
            if rapido and id(reg) not in ids_lote and reg.get("estado") == estado_remitos_antes.get(k_):
                continue
            if reg.get("local") in locales:
                proveedores_por_local.setdefault(reg["local"], []).append(reg)

        def _push_local(nombre: str) -> dict:
            """Todo lo de UNA planilla (Facturas → remitos de proveedor → remitos del Obrador), en orden: cada local es
            otro Spreadsheet y corre en su propio hilo con su propio cliente gspread."""
            out: dict = {}
            if nombre in plan_por_local:
                filas = plan_por_local[nombre]
                try:
                    out[nombre] = rf.con_reintentos(          # 429/5xx: backoff; push_drive no duplica (N° + proveedor)
                        lambda n=nombre, f=filas: imp.push_drive(n, f, confirmado=args.confirmo_escribir_produccion, demo=demo),
                        trabajo=f"push:{nombre}", etapa="push_facturas_lote", intentos=4, espera=6)
                    rf.evento(f"push:{nombre}", "push_facturas", "OK" if out[nombre].get("escrito") else "NO_ESCRITO",
                              local=nombre, motivo=out[nombre].get("motivo", ""))
                except Exception as e:                       # cada destino es independiente
                    rf.evento(f"push:{nombre}", "push_facturas", "ERROR", error=e, local=nombre)
                    out[nombre] = {"escrito": False, "local": nombre, "motivo": "ERROR " + rf.detalle_excepcion(e)["tipo"],
                                   "error": rf.detalle_excepcion(e)}
            if nombre in proveedores_por_local:
                out[f"Remitos proveedores {nombre}"] = imp.push_remitos_proveedor_drive(
                    nombre, proveedores_por_local[nombre], confirmado=args.confirmo_escribir_produccion)
            if nombre in remitos_obrador_por_local:
                out[f"Remitos {nombre}"] = imp.push_remitos_drive(
                    nombre, remitos_obrador_por_local[nombre], confirmado=args.confirmo_escribir_produccion)
            return out

        destinos = sorted(set(plan_por_local) | set(proveedores_por_local) | set(remitos_obrador_por_local))
        t_push = time.time()
        if len(destinos) > 1 and not __import__("os").environ.get("MISHI_PUSH_SECUENCIAL"):
            with ThreadPoolExecutor(max_workers=min(5, len(destinos)), thread_name_prefix="push") as ex:
                for parcial in ex.map(_push_local, destinos):      # re-lanza la primera excepción al terminar
                    pushes.update(parcial)
        else:
            for n_ in destinos:
                pushes.update(_push_local(n_))
        tiempos_push = {"push_seg": round(time.time() - t_push, 2), "hilos": min(5, max(len(destinos), 1)),
                        "locales": destinos}

    # ---- reporte --------------------------------------------------------------
    filas_aud = []
    for c in cl["comprobantes"]:
        an = analisis[c["archivo"]]
        rem_conc = c.get("_remito", {}).get("estado") == "CONCILIADO" if c.get("tipo") == "Remito" else False
        # remito de Zuccardi (sin precios): conciliado con la factura por el N° de remito
        # Solo el estado persistido con referencia explícita acredita la conciliación.
        filas_aud.append({
            "archivo": c["archivo"], "tipo": c.get("tipo"), "numero": c.get("numero_factura"),
            "fecha": c.get("fecha_factura"), "proveedor": c.get("proveedor_detectado"),
            "sello_firma": (c.get("sello") or ("(PDF digital)" if str(c.get("motor", "")).startswith("pdf") else "—"))
            + ((" | firmas: " + ", ".join(c.get("firmas") or [])) if c.get("firmas") else ""),
            "local": an["local"], "regla": an["por"], "estado": estado_de(c, an, rem_conc),
            "alertas": an["alertas"], "modo_historial": an["m0"]["modo"] if an["m0"] else None,
            "local_historial": an["m0"]["local"] if an["m0"] else None,
            "aprendido": aprendizajes.get(c["archivo"], {}), "duplicado_en": c.get("duplicado_en"),
            "filas_imputadas": len(filas_por_comp.get(c["archivo"], [])),
        })
    for p in cl["informes"]:
        filas_aud.append({"archivo": p.name, "tipo": "Informe", "numero": "-", "proveedor": "Maxirest",
                          "sello_firma": "—", "local": "—", "estado": "INFORME_INDEXADO", "alertas": []})
    for p in cl["listas"]:
        filas_aud.append({"archivo": p.name, "tipo": "Lista", "numero": "-", "proveedor": "GRUPO FREE",
                          "sello_firma": "—", "local": "Benchmark", "estado": "LISTA_BENCHMARK", "alertas": []})
    for p, orig in cl["duplicados"]:
        filas_aud.append({"archivo": p.name, "tipo": "Duplicado", "numero": "-", "proveedor": "-",
                          "sello_firma": "—", "local": "—", "estado": f"DUPLICADO de {orig.name}", "alertas": []})
    cl["_plan_por_local"] = plan_por_local                  # filas completas (para recuperar_lote.py)
    reporte = {"generado": datetime.now().isoformat(timespec="seconds"), "auditoria": filas_aud,
               "remitos_obrador": [r for rs in remitos_obrador_por_local.values() for r in rs],
               "escrituras": escrituras, "push_drive": pushes, "tiempos_push": tiempos_push, "remitos": remitos_lote,
               "informes": informes_idx, "benchmark_items": len(bench["items"]) if bench else 0,
               "planes": {n: [{k: f.get(k) for k in ("A", "B", "D", "presentacion", "G", "H", "I", "accion", "motivos",
                                                    "avisos", "descripcion_original", "confianza", "metodo", "sku")}
                              for f in fs] for n, fs in plan_por_local.items()},
               "simulaciones": simulaciones,
               "modos_proveedor": ap.get("modos_proveedor"), "marcas": {n: m for n, m in marcas_por_local.items()}}
    if escribir and not demo:                               # telemetría de ROI: nunca debe romper el lote
        try:
            import telemetria_roi
            ahora = time.time()
            telemetria_roi.registrar_lote(
                cl["comprobantes"], analisis, reporte, ahora - t_ini,
                {"leer": t_leido - t_ini, "rutear_imputar": max(ahora - t_leido - tiempos_push.get("push_seg", 0), 0),
                 "escribir_sheets": tiempos_push.get("push_seg", 0)})
        except Exception as e_:                              # noqa: BLE001
            LOG.debug("telemetría: %s", e_)
    return reporte, locales, ap, cl


def _imprimir(rep: dict, locales: dict, ap: dict, args, cl: dict) -> int:
    print("\n" + "=" * 150)
    print(f"AUDITORÍA DE Facturas_WPP/  ({len(rep['auditoria'])} archivos)")
    print("=" * 150)
    print(f"{'Archivo':<44} {'Tipo':<8} {'N° comprobante':<17} {'Proveedor':<26} {'Local':<10} Estado")
    for r in rep["auditoria"]:
        print(f"{r['archivo'][-44:]:<44} {str(r['tipo'])[:8]:<8} {str(r.get('numero'))[:17]:<17} "
              f"{str(r.get('proveedor'))[:26]:<26} {str(r.get('local'))[:10]:<10} {r['estado']}")
    print("\nSELLOS / FIRMAS y alertas:")
    for r in rep["auditoria"]:
        if r.get("alertas") or r.get("tipo") in ("Factura", "NC", "Remito"):
            print(f" · {r['archivo'][-34:]:<34} sello/firma: {r['sello_firma'][:70]}")
            if r.get("regla"):
                print(f"     regla local: {r['regla']}")
            for a in r.get("alertas") or []:
                print(f"     ⚠ {a}")
    if rep["escrituras"]:
        print("\nESCRITURAS (solo copias):")
        for n, e in rep["escrituras"].items():
            print(f"  {e['copia']}: {e['filas']} filas, {e['verificadas']} verificadas por lectura, {e['marcadas']} filas marcadas")
    elif args.dry_run:
        print("\n(dry-run: no se escribió nada)")
    if rep["push_drive"]:
        print("\nPUSH A PRODUCCIÓN (plan):")
        for n, p in rep["push_drive"].items():
            print(f"  {n}: escrito={p.get('escrito')} {p.get('motivo', '')} {len(p.get('plan', []))} filas")
    return 0


if __name__ == "__main__":
    sys.exit(main())
