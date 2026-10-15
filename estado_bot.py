"""Estado operativo: procesos Telegram/orquestador, uptime, offset, cola, latido de agentes y conectividad.
Uso: python estado_bot.py   (solo lectura; no imprime tokens)"""
import json
import os
import subprocess
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent


def procesos() -> list[dict]:
    q = ("Get-CimInstance Win32_Process | Where-Object { $_.Name -match 'python' -and $_.CommandLine -match "
         "'telegram_ingestor.py|orquestador_mishi.py' } | Select-Object ProcessId,CreationDate,CommandLine | ConvertTo-Json -Compress")
    r = subprocess.run(["powershell", "-NoProfile", "-Command", q], capture_output=True, text=True)
    try:
        d = json.loads(r.stdout or "[]")
    except ValueError:
        return []
    return d if isinstance(d, list) else [d]


def main() -> int:
    ps = procesos()
    print(f"Procesos activos: {len(ps)}" + ("  ⚠ MÁS DE UNO: conflicto de getUpdates" if len(ps) > 1 else ""))
    for p in ps:
        print(f"  PID {p['ProcessId']} · {str(p['CreationDate'])[:20]} · {p['CommandLine'][-70:]}")
    try:
        est = json.loads((BASE / "telegram_state.json").read_text(encoding="utf-8"))
        print(f"Telegram: offset {est.get('offset')} · {len(est.get('sha256', {}))} SHA procesados · "
              f"{len(est.get('mensajes', []))} mensajes registrados")
    except OSError:
        print("Telegram: sin telegram_state.json")
    try:
        from coordinador import Cola
        print("Cola:", Cola().resumen())
    except Exception as e:
        print("Cola: error", e)
    lat = BASE / "estado" / "orquestador_estado.json"
    if lat.exists():
        d = json.loads(lat.read_text(encoding="utf-8"))
        edad = time.time() - lat.stat().st_mtime
        print(f"Orquestador PID {d['pid']} · uptime {d['uptime_s']} s · latido hace {edad:.0f} s"
              + ("  ⚠ latido viejo" if edad > 30 else ""))
        for n, a in d["agentes"].items():
            print(f"  {n}: vivo={a['vivo']} ciclos={a['ciclos']} errores={a['errores']} último={a['ultimo_ms']} ms eventos={a['eventos']}"
                  + (f" · {a['ultimo_error']}" if a.get("ultimo_error") else ""))
    ev = BASE / "logs" / "pipeline_eventos.jsonl"
    if ev.exists():
        lineas = ev.read_text(encoding="utf-8").splitlines()
        print(f"Eventos registrados: {len(lineas)}; últimos errores:")
        for l in [json.loads(x) for x in lineas if '"ERROR"' in x][-3:]:
            print("  ", l["ts"], l["etapa"], l.get("error", {}).get("tipo"), l.get("error", {}).get("http_status", ""))
    bl = BASE / "estado" / "bloqueos_permisos.json"
    if bl.exists():
        print(f"Bloqueos de permisos en cola: {len(json.loads(bl.read_text(encoding='utf-8')))}")
    try:
        import entornos
        print("Entorno:", entornos.entorno())
    except Exception as e:
        print("Entorno: error", e)
    return 0


if __name__ == "__main__":
    sys.exit(main())
