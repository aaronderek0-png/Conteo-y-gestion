"""mishi.py — punto único de entrada por consola para el ecosistema Mishiguene.

    python mishi.py estado                    estado vivo (orquestador, cola, RAM, copias PRUEBA, ROI) en <1 s
    python mishi.py lote <ruta>               procesa una carpeta o archivo por el Camino Rápido y escribe en Sheets
    python mishi.py pendientes                cola persistente: REVISAR / ERROR / en curso
    python mishi.py asignar <N°> <Local>      fija el local de un comprobante (igual que /asignar en Telegram)
    python mishi.py roi                       cuadro ejecutivo de ROI y regenera JUSTIFICACION + Panel HTML

`lote` y `asignar` escriben en las planillas del entorno activo (entornos.json / MISHI_ENTORNO). En PRODUCCION exigen
`--confirmo-escribir-produccion`. No usa Telegram ni borra nada.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent
ESTADO = BASE / "estado"
sys.path.insert(0, str(BASE))


def _json(p: Path, defecto=None):
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return defecto


def _pid_vivo(pid: int) -> bool:
    try:
        r = subprocess.run(["tasklist", "/FI", f"PID eq {pid}", "/NH"], capture_output=True, text=True, timeout=5,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        return str(pid) in r.stdout
    except Exception:                                            # noqa: BLE001
        return False


def _entorno() -> str:
    import entornos
    return entornos.entorno()


def cmd_estado() -> int:
    t0 = time.time()
    lat = _json(ESTADO / "orquestador_estado.json", {})
    pid = lat.get("pid")
    vivo = bool(pid) and _pid_vivo(int(pid))
    ent = _json(BASE / "entornos.json", {})
    print(f"ENTORNO: {_entorno()}   (activo en entornos.json: {ent.get('activo', '—')})")
    if lat:
        edad = None
        try:
            edad = round(time.time() - time.mktime(time.strptime(lat["ahora"], "%Y-%m-%d %H:%M:%S")))
        except (KeyError, ValueError):
            pass
        print(f"ORQUESTADOR: PID {pid} {'VIVO' if vivo else 'NO CORRE'} · uptime {lat.get('uptime_s')} s · "
              f"latido hace {edad} s · lote máx {lat.get('batch_max')}")
        for n, a in (lat.get("agentes") or {}).items():
            print(f"   {n:<24} {'vivo ' if a.get('vivo') else 'caído'} ciclos={a.get('ciclos')} errores={a.get('errores')}"
                  + (f" · {a.get('ultimo_error')}" if a.get("ultimo_error") else ""))
        ram = lat.get("ram") or {}
        if ram:
            print(f"RAM: {ram.get('rss_mb', '?')} MB · catálogos cargados: " +
                  (", ".join(f"{n} ({v['catalogo']} prod · {v['facturas']} filas)" for n, v in
                             (ram.get('locales_en_ram') or {}).items()) or "ninguno"))
        print(f"LECTURA (proceso): {lat.get('lectura')}")
    else:
        print("ORQUESTADOR: sin latido (¿nunca arrancó?). Iniciar: RUN_BOT.bat")
    cola = _json(ESTADO / "trabajos.json", {"trabajos": {}})["trabajos"]
    est: dict = {}
    for j in cola.values():
        est[j["estado"]] = est.get(j["estado"], 0) + 1
    print(f"COLA: {est or 'vacía'}")
    prueba = ent.get("prueba") or {}
    print("COPIAS PRUEBA: " + (", ".join(f"{k}={v[:6]}…" for k, v in prueba.items()) or "ninguna registrada"))
    try:
        import telemetria_roi as t
        r = t.resumen()
        print(f"ROI: {r['comprobantes_procesados']} comprobantes · {r['seg_maquina_promedio']} s/comprobante medido · "
              f"{r['horas_humanas_estimadas_ahorradas']} h humanas estimadas (modelo) · "
              f"${r['dinero_a_verificar_remitos']:,.0f} de remitos a verificar")
    except Exception as e:                                       # noqa: BLE001
        print("ROI: no disponible:", str(e)[:80])
    print(f"({time.time() - t0:.2f} s)")
    return 0


def cmd_lote(ruta: str, confirmo: bool) -> int:
    from types import SimpleNamespace
    import shutil
    import tempfile
    ent = _entorno()
    if ent != "PRUEBA" and not confirmo:
        print("ENTORNO PRODUCCION: agregá --confirmo-escribir-produccion si de verdad querés escribir ahí.")
        return 2
    src = Path(ruta)
    if not src.exists():
        print("No existe:", src)
        return 2
    t0 = time.time()
    import procesar_lote_wpp as lote
    import imputacion_facturas as imp
    carpeta = src
    tmp = None
    if src.is_file():
        tmp = Path(tempfile.mkdtemp(prefix="mishi_lote_"))
        shutil.copy2(src, tmp / src.name)
        carpeta = tmp
    print(f"Entorno {ent} · precargando catálogos en RAM…")
    r = imp.precargar_ram()
    print(f"   {', '.join(r['locales'])} en {r['segundos']} s")
    ns = SimpleNamespace(push_drive=True, confirmo_escribir_produccion=True, dry_run=False, json=None)
    try:
        rep = lote.ejecutar(carpeta, True, ns, indexar_informes=False, escribir_xlsx_local=False)[0]
    finally:
        if tmp:
            shutil.rmtree(tmp, ignore_errors=True)
    dt = time.time() - t0
    print(f"\n{len(rep['auditoria'])} comprobantes en {dt:.1f} s (push en paralelo: {rep.get('tiempos_push')})")
    for a in rep["auditoria"]:
        print(f"  • {a['archivo'][:46]:<46} {a.get('tipo', ''):<10} N°{a.get('numero')} {a.get('proveedor') or ''}"[:120]
              + f" → {a.get('local') or 'SIN LOCAL'} · {a['estado']}")
    for k, v in (rep.get("push_drive") or {}).items():
        print(f"  ✍ {k}: {v.get('planilla')} / {v.get('hoja')} filas {v.get('desde_fila')}–{v.get('hasta_fila')} "
              f"({v.get('filas')} escritas{', sin cambios ' + str(v['sin_cambios']) if v.get('sin_cambios') else ''})"
              if v.get("escrito") else f"  ✖ {k}: {v.get('motivo')}")
    return 0


def cmd_pendientes() -> int:
    import telegram_ingestor as ti
    print(ti.texto_pendientes())
    return 0


def cmd_asignar(n: str, local: str, confirmo: bool) -> int:
    if _entorno() != "PRUEBA" and not confirmo:
        print("ENTORNO PRODUCCION: agregá --confirmo-escribir-produccion.")
        return 2
    import telegram_ingestor as ti
    print(ti.comando_asignar(f"/asignar {n} {local}", True))
    return 0


def cmd_roi() -> int:
    import telemetria_roi as t
    t.sembrar_historico()
    r = t.resumen()
    e = r["errores_humanos_evitados"]
    print("=" * 78)
    print(" ROI Y TIEMPO HUMANO LIBERADO — Agente Mishiguene")
    print("=" * 78)
    print(f" Modelo humano (supuesto): {r['modelo']}")
    print(f" Comprobantes registrados : {r['comprobantes_procesados']}  ({r['renglones_totales']} renglones, por tipo {r['por_tipo']})")
    print(f" Tiempo de máquina        : {r['seg_maquina_promedio']} s por comprobante (medido sobre {r['comprobantes_medidos']} en {r['lotes']} lotes)")
    print(f" Horas humanas ahorradas  : {r['horas_humanas_estimadas_ahorradas']} h (MODELO: {r['min_humanos_por_comprobante_modelo']} min/comprobante)")
    print(f" Remitos a verificar      : ${r['dinero_a_verificar_remitos']:,.0f}  (anulados 'NO' + diferencias de balanza; NO es dinero recuperado)")
    print(f" Errores humanos evitados : {e['celdas_ref_reparadas']:,} celdas #REF!/derrame reescritas · {e['duplicados_frenados']} duplicados · "
          f"{e['fechas_futuras_bloqueadas']} fechas futuras · {e['desvios_precio_mayor_25pct']} desvíos > ±25 %")
    ul = r.get("ultimo_lote")
    if ul:
        print(f" Último lote              : {ul['comprobantes']} comprobantes / {ul['renglones']} renglones en {ul['seg_maquina']} s "
              f"(push: {ul.get('push')})")
    d = t.cargar()
    if d.get("por_etapa_seg") and d.get("lotes"):
        print(" Velocidad por etapa (s por lote, promedio): " +
              ", ".join(f"{k}={v / d['lotes']:.1f}" for k, v in d["por_etapa_seg"].items()))
    print(f"\n Proyección ({r['proyeccion_basada_en']}; hora humana supuesta ${r['hora_humana_ars']:,.0f}):")
    print("   comprobantes/mes | h humanas liberadas/mes | h/año    | ahorro $/mes   | ahorro $/año")
    for v in r["proyeccion"].values():
        print(f"   {v['comprobantes_mes']:>16,} | {v['horas_humanas_liberadas_mes']:>23,} | {v['horas_humanas_liberadas_anio']:>8,} | "
              f"${v['ahorro_ars_mes']:>13,} | ${v['ahorro_ars_anio']:>13,}")
    ah = None
    try:
        import pandas as pd
        res = pd.read_excel(BASE / "Reportes_Ahorro" / "Reporte_Ahorro_Resumen.xlsx", sheet_name="Resumen Cadena Octubre")
        ah = float(res["Ahorro estimado mes"].sum())
    except Exception:                                            # noqa: BLE001
        pass
    if ah is not None:
        print(f"\n Ahorro estimado con proveedores (comparador, sobre listas): ${ah:,.0f}/mes — estimación, no realizado")
    print("\nRegenerando JUSTIFICACION_EJECUTIVA_MISHIGUENE.md y Panel_Ejecutivo_Agente_Mishiguene.html …")
    p = subprocess.run([sys.executable, str(BASE / "generar_entregables.py")], capture_output=True, text=True,
                       cwd=str(BASE), creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    print("   " + ((p.stdout or "").strip().splitlines() or ["(sin salida)"])[-1] if p.returncode == 0
          else "   ✖ falló:\n" + (p.stderr or p.stdout)[-400:])
    return p.returncode


def main(argv: list[str] | None = None) -> int:
    a = list(sys.argv[1:] if argv is None else argv)
    confirmo = "--confirmo-escribir-produccion" in a
    a = [x for x in a if x != "--confirmo-escribir-produccion"]
    if not a or a[0] in ("-h", "--help", "ayuda"):
        print(__doc__)
        return 0
    c = a[0].lower()
    if c == "estado":
        return cmd_estado()
    if c == "lote" and len(a) >= 2:
        return cmd_lote(a[1], confirmo)
    if c == "pendientes":
        return cmd_pendientes()
    if c == "asignar" and len(a) >= 3:
        return cmd_asignar(a[1], a[2], confirmo)
    if c == "roi":
        return cmd_roi()
    print(__doc__)
    return 2


if __name__ == "__main__":
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:                                            # noqa: BLE001
        pass
    sys.exit(main())
