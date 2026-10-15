"""Telemetría de tiempo humano ahorrado y ROI (estado/telemetria_roi.json).

Qué es MEDIDO y qué es MODELO:
- MEDIDO por el agente: comprobantes, renglones, segundos de máquina, duplicados frenados, fechas futuras bloqueadas,
  desvíos de precio > ±25 %, remitos con renglones anulados / diferencias de balanza.
- MODELO (supuesto auditable, no medición): minutos humanos ahorrados = 45 s base por comprobante
  + 25 s por renglón + 60 s extra si es remito con birome/balanza o Nota de Crédito.
- Los $ de remitos usan precios impresos/de lista del propio remito: son MONTO A VERIFICAR con el proveedor, no
  dinero ya recuperado ni gasto definitivo.
- Los $ por hora humana son un parámetro (HORA_HUMANA_ARS) que el usuario debe confirmar; no es un dato del sistema.
"""
from __future__ import annotations

import json
import os
import threading
import time
from datetime import date, datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
ESTADO = Path(os.environ.get("MISHI_ESTADO_DIR") or BASE / "estado")
ARCHIVO = ESTADO / "telemetria_roi.json"
_LOCK = threading.Lock()

SEG_BASE_COMPROBANTE = 45
SEG_POR_RENGLON = 25
SEG_EXTRA_REMITO_NC = 60
HORA_HUMANA_ARS = float(os.environ.get("MISHI_HORA_HUMANA_ARS", "6000"))   # supuesto: confirmar con el dueño
ESCENARIOS = (1000, 3000, 10000)

# Hito medido el 14/10/2026: celdas reescritas por reparar_ref_planillas.py --apply en las 5 copias PRUEBA
HITO_REF = {"celdas_reescritas": 312339, "por_local": {"Lafinur": 71757, "Cabello": 59949, "Arcos": 60954,
                                                       "Congreso": 59985, "Chimidoha": 59694},
            "fuente": "reparar_ref_planillas.py --apply (14/10/2026, copias NO CARGAR - PRUEBA)"}


def _vacio() -> dict:
    return {"version": 1, "desde": datetime.now().isoformat(timespec="seconds"), "ids": {}, "lotes": 0,
            "comprobantes_procesados": 0, "comprobantes_medidos": 0, "renglones_totales": 0, "seg_maquina_total": 0.0,
            "minutos_humanos_estimados_ahorrados": 0.0, "dinero_a_verificar_remitos": 0.0,
            "errores_humanos_evitados": {"duplicados_frenados": 0, "fechas_futuras_bloqueadas": 0,
                                         "desvios_precio_mayor_25pct": 0, "celdas_ref_reparadas": HITO_REF["celdas_reescritas"]},
            "por_tipo": {}, "por_etapa_seg": {}, "ultimo_lote": None, "hitos": {"ref": HITO_REF}}


def cargar() -> dict:
    try:
        d = json.loads(ARCHIVO.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return _vacio()
    base = _vacio()
    base.update(d)
    base["errores_humanos_evitados"] = {**_vacio()["errores_humanos_evitados"], **d.get("errores_humanos_evitados", {})}
    return base


def _guardar(d: dict) -> None:
    ESTADO.mkdir(exist_ok=True)
    tmp = ARCHIVO.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    for _ in range(5):                                         # Windows: el lector puede tener el archivo abierto
        try:
            tmp.replace(ARCHIVO)
            return
        except PermissionError:
            time.sleep(0.2)
    ARCHIVO.write_text(tmp.read_text(encoding="utf-8"), encoding="utf-8")


def minutos_modelo(tipo: str | None, renglones: int, con_birome: bool = False) -> float:
    seg = SEG_BASE_COMPROBANTE + SEG_POR_RENGLON * max(int(renglones), 0)
    if tipo in ("Remito", "NC") or con_birome:
        seg += SEG_EXTRA_REMITO_NC
    return round(seg / 60, 2)


def registrar_lote(comprobantes: list[dict], analisis: dict, reporte: dict, seg_maquina: float,
                   etapas: dict | None = None) -> dict:
    """Registra un lote. Idempotente por (sha256|archivo + N° + proveedor): un reintento no cuenta dos veces."""
    hoy = date.today()
    with _LOCK:
        d = cargar()
        nuevos = 0
        ren_lote = 0
        min_lote = 0.0
        for c in comprobantes:
            tipo = c.get("tipo") or "Factura"
            items = c.get("items") or []
            clave = f"{c.get('sha256') or c.get('archivo')}|{c.get('numero_factura')}|{c.get('proveedor_detectado')}|{c.get('archivo')}"
            if clave in d["ids"]:
                continue
            if tipo not in ("Factura", "NC", "Remito", "RemitoObrador"):
                continue
            con_birome = bool(c.get("anotaciones_manuscritas")) and tipo == "Remito"
            minutos = minutos_modelo(tipo, len(items), con_birome)
            dinero = 0.0
            reg = c.get("_remito") or {}
            for r in reg.get("renglones", []) or []:
                if r.get("anulado"):
                    dinero += float(r.get("total_remito") or 0) or float(r.get("cantidad_facturada_remito") or 0) * float(r.get("precio_unitario") or 0)
                dinero += float(r.get("reclamo_por_peso") or 0)
            try:
                ff = datetime.fromisoformat(str(c.get("fecha_factura"))[:10]).date()
                if (ff - hoy).days > 3:
                    d["errores_humanos_evitados"]["fechas_futuras_bloqueadas"] += 1
            except (ValueError, TypeError):
                pass
            d["ids"][clave] = datetime.now().isoformat(timespec="seconds")
            d["comprobantes_procesados"] += 1
            d["renglones_totales"] += len(items)
            d["minutos_humanos_estimados_ahorrados"] = round(d["minutos_humanos_estimados_ahorrados"] + minutos, 2)
            d["dinero_a_verificar_remitos"] = round(d["dinero_a_verificar_remitos"] + dinero, 2)
            d["por_tipo"][tipo] = d["por_tipo"].get(tipo, 0) + 1
            nuevos += 1
            ren_lote += len(items)
            min_lote += minutos
        # duplicados frenados por el push y desvíos de precio > ±25 % (planes del lote)
        dup = 0
        for r in (reporte.get("push_drive") or {}).values():
            dup += len(r.get("omitidas") or [])
        dup += len(reporte.get("auditoria") and [a for a in reporte["auditoria"] if str(a.get("estado", "")).startswith("DUPLICADO")] or [])
        desv = 0
        for filas in (reporte.get("planes") or {}).values():
            for f in filas:
                if any("±25" in str(m) or "> ±50" in str(m) for m in (f.get("motivos") or [])):
                    desv += 1
        if nuevos:
            d["errores_humanos_evitados"]["duplicados_frenados"] += dup
            d["errores_humanos_evitados"]["desvios_precio_mayor_25pct"] += desv
            d["lotes"] += 1
            d["comprobantes_medidos"] = d.get("comprobantes_medidos", 0) + nuevos
            d["seg_maquina_total"] = round(d["seg_maquina_total"] + float(seg_maquina), 2)
            for k, v in (etapas or {}).items():
                d["por_etapa_seg"][k] = round(d["por_etapa_seg"].get(k, 0.0) + float(v), 2)
            d["ultimo_lote"] = {"ts": datetime.now().isoformat(timespec="seconds"), "comprobantes": nuevos,
                                "renglones": ren_lote, "seg_maquina": round(float(seg_maquina), 2),
                                "minutos_humanos_modelo": round(min_lote, 2), "push": reporte.get("tiempos_push")}
            _guardar(d)
        return d


def resumen(d: dict | None = None) -> dict:
    d = d or cargar()
    n = d["comprobantes_procesados"]
    seg = d["seg_maquina_total"]
    ren = d["renglones_totales"]
    prom_ren = (ren / n) if n else 5.0                          # sin datos: 5 renglones por comprobante (supuesto)
    min_por_comp = (d["minutos_humanos_estimados_ahorrados"] / n) if n else minutos_modelo("Factura", round(prom_ren))
    medidos = d.get("comprobantes_medidos", 0)
    seg_por_comp = (seg / medidos) if medidos else 0.0      # solo comprobantes con tiempo de máquina medido
    proy = {}
    for e in ESCENARIOS:
        horas = e * min_por_comp / 60
        maquina_h = e * seg_por_comp / 3600
        proy[str(e)] = {"comprobantes_mes": e, "horas_humanas_liberadas_mes": round(horas, 1),
                        "horas_humanas_liberadas_anio": round(horas * 12, 1),
                        "horas_maquina_mes": round(maquina_h, 1),
                        "ahorro_ars_mes": round(horas * HORA_HUMANA_ARS), "ahorro_ars_anio": round(horas * 12 * HORA_HUMANA_ARS)}
    return {"comprobantes_procesados": n, "renglones_totales": ren, "lotes": d["lotes"],
            "comprobantes_medidos": medidos, "seg_maquina_total": round(seg, 1), "seg_maquina_promedio": round(seg_por_comp, 2),
            "minutos_humanos_estimados_ahorrados": round(d["minutos_humanos_estimados_ahorrados"], 1),
            "horas_humanas_estimadas_ahorradas": round(d["minutos_humanos_estimados_ahorrados"] / 60, 2),
            "min_humanos_por_comprobante_modelo": round(min_por_comp, 2),
            "dinero_a_verificar_remitos": d["dinero_a_verificar_remitos"],
            "errores_humanos_evitados": d["errores_humanos_evitados"], "por_tipo": d["por_tipo"],
            "ultimo_lote": d.get("ultimo_lote"), "hora_humana_ars": HORA_HUMANA_ARS,
            "proyeccion": proy,
            "modelo": f"{SEG_BASE_COMPROBANTE} s base + {SEG_POR_RENGLON} s/renglón + {SEG_EXTRA_REMITO_NC} s si remito con birome o NC",
            "proyeccion_basada_en": "promedio medido" if n else "supuesto (sin comprobantes registrados todavía)"}


def sembrar_historico() -> dict:
    """Primera vez: toma lo ya procesado en cache_facturas.json (lecturas reales) para no empezar de cero.
    Solo cuenta si el registro está vacío; marca la fuente."""
    with _LOCK:
        d = cargar()
        if d["comprobantes_procesados"] or d.get("hitos", {}).get("sembrado"):
            return d
        try:
            cache = json.loads((BASE / "cache_facturas.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return d
        for k, c in cache.items():
            tipo = c.get("tipo") or "Factura"
            if tipo not in ("Factura", "NC", "Remito", "RemitoObrador") or not c.get("items"):
                continue
            clave = f"{c.get('sha256') or k}|{c.get('numero_factura')}|{c.get('proveedor_detectado')}|{c.get('archivo')}"
            if clave in d["ids"]:
                continue
            d["ids"][clave] = "sembrado"
            d["comprobantes_procesados"] += 1
            d["renglones_totales"] += len(c["items"])
            d["minutos_humanos_estimados_ahorrados"] = round(
                d["minutos_humanos_estimados_ahorrados"] + minutos_modelo(tipo, len(c["items"]),
                                                                          bool(c.get("anotaciones_manuscritas")) and tipo == "Remito"), 2)
            d["por_tipo"][tipo] = d["por_tipo"].get(tipo, 0) + 1
        try:                                                  # remitos ya registrados: anulados con NO + diferencias de balanza
            reg = json.loads((BASE / "remitos_pendientes.json").read_text(encoding="utf-8"))
            for v in reg.values():
                for r in v.get("renglones", []) or []:
                    if r.get("anulado"):
                        d["dinero_a_verificar_remitos"] += float(r.get("total_remito") or 0) or                             float(r.get("cantidad_facturada_remito") or 0) * float(r.get("precio_unitario") or 0)
                    d["dinero_a_verificar_remitos"] += float(r.get("reclamo_por_peso") or 0)
            d["dinero_a_verificar_remitos"] = round(d["dinero_a_verificar_remitos"], 2)
        except (OSError, ValueError):
            pass
        d["hitos"]["sembrado"] = {"fuente": "cache_facturas.json (lecturas ya procesadas)", "ts": datetime.now().isoformat(timespec="seconds")}
        _guardar(d)
        return d
