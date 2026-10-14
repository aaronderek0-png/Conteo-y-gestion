"""Verificación SOLO LECTURA contra los Sheets vivos: ¿están cargados los comprobantes de la recuperación?
Uso: python verificar_vivo.py   (no escribe nada; imprime y guarda verificacion_vivo.json)"""
from __future__ import annotations

import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
import imputacion_facturas as imp  # noqa: E402
import registro_fallos as rf       # noqa: E402

# (local, N° significativo, fragmento de proveedor) — dígitos significativos de col A
FACTURAS = [("Congreso", 97543, "resta"), ("Congreso", 35844, "mix"), ("Congreso", 19303, "buenos mares"),
            ("Congreso", 34674, "rupajo"), ("Lafinur", 34674, "rupajo"),
            ("Congreso", 421, "mi"), ("Congreso", 423, "mi"), ("Congreso", 8528, "raices")]


def main() -> dict:
    gc, _ = imp._gspread()
    out = {}
    for local in ("Congreso", "Lafinur", "Cabello"):
        try:
            sh = gc.open_by_key(imp.IDS_PLANILLAS[local])
            info = {"planilla": sh.title, "pestanas": [w.title for w in sh.worksheets()]}
            ws = imp._hoja_por_clave(sh, "factura")
            a, d = ws.col_values(1), ws.col_values(4)
            info["facturas_filas"] = len(a)
            info["facturas"] = {f"{n}:{p}": [i + 1 for i, (x, y) in enumerate(zip(a, d + [""] * len(a)))
                                              if imp.ultimo_entero(x) == n and p in imp.norm(y).lower()]
                                for l, n, p in FACTURAS if l == local}
            rp = next((w for w in sh.worksheets() if w.title == "Remitos_Pendientes"), None)
            if rp:
                v = rp.get_all_values()
                info["remitos_pendientes_filas"] = len(v) - 1
                info["remitos_pendientes_muestra"] = [r[:6] for r in v[1:4]]
            rm = imp._hoja_por_clave(sh, "remito")
            if rm:
                v = rm.get_all_values()
                info["remitos_filas"] = len(v) - 1
                info["remitos_ultimas"] = [r[:5] for r in v[-8:]]
            out[local] = info
        except Exception as e:  # noqa: BLE001
            rf.evento(f"verificar:{local}", "lectura_viva", "ERROR", error=e, local=local)
            out[local] = {"error": rf.detalle_excepcion(e)}
    (BASE / "verificacion_vivo.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    return out


if __name__ == "__main__":
    print(json.dumps(main(), ensure_ascii=False, indent=1))
