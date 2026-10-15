"""Crea (una sola vez) las copias de PRUEBA «NO CARGAR - PRUEBA - <Local>» de las 5 planillas y guarda sus IDs en
entornos.json. Usa Drive files.copy con la Service Account existente (sin compartir ni cambiar permisos de nada).

Uso: python crear_trials.py <Local>        # copia una planilla
     python crear_trials.py --estado       # muestra lo registrado
Tras copiar verifica: ID distinto del de producción, acceso editor, grilla/pestañas y fórmulas presentes.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))
import imputacion_facturas as imp  # noqa: E402

ENTORNOS = BASE / "entornos.json"


def cargar() -> dict:
    try:
        return json.loads(ENTORNOS.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"produccion": dict(imp.IDS_PLANILLAS), "prueba": {}}


def guardar(d: dict) -> None:
    tmp = ENTORNOS.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(ENTORNOS)


def copiar(local: str) -> dict:
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2.service_account import Credentials
    d = cargar()
    if local in d["prueba"]:
        return {"local": local, "ya_existe": d["prueba"][local]}
    creds = Credentials.from_service_account_file(str(BASE / "credentials.json"),
                                                  scopes=["https://www.googleapis.com/auth/drive"])
    ses = AuthorizedSession(creds)
    origen = d["produccion"][local]
    r = ses.post(f"https://www.googleapis.com/drive/v3/files/{origen}/copy",
                 params={"supportsAllDrives": "true", "fields": "id,name,capabilities(canEdit)"},
                 json={"name": f"NO CARGAR - PRUEBA - {local}"})
    if r.status_code >= 400:
        err = r.json().get("error", {})
        return {"local": local, "ok": False, "http": r.status_code, "status": err.get("status"),
                "mensaje": str(err.get("message", ""))[:300],
                "motivo_detalle": [e.get("reason") for e in err.get("errors", [])]}
    nuevo = r.json()
    assert nuevo["id"] != origen, "el ID de prueba coincide con producción"
    d["prueba"][local] = nuevo["id"]
    guardar(d)
    return {"local": local, "ok": True, "id_prueba": nuevo["id"], "nombre": nuevo["name"],
            "editable": nuevo.get("capabilities", {}).get("canEdit")}


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--estado":
        print(json.dumps(cargar(), indent=1))
    else:
        print(json.dumps(copiar(sys.argv[1]), ensure_ascii=False, indent=1))
