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


def _sesion_usuario():
    """Sesión OAuth del USUARIO humano a partir de ~/.clasprc.json (login de clasp). Nunca imprime tokens."""
    import os
    from google.auth.transport.requests import AuthorizedSession
    from google.oauth2.credentials import Credentials
    p = Path(os.path.expanduser("~/.clasprc.json"))
    t = json.loads(p.read_text(encoding="utf-8"))
    tk = (t.get("tokens") or {}).get("default") or t.get("token") or {}
    creds = Credentials(None, refresh_token=tk["refresh_token"], client_id=tk["client_id"],
                        client_secret=tk["client_secret"], token_uri="https://oauth2.googleapis.com/token")
    return AuthorizedSession(creds)


def copiar_como_usuario(local: str, compartir_con: str = "agente-mishi@black-obelisk-460914-d9.iam.gserviceaccount.com") -> dict:
    """files.copy a nombre del usuario (propietario humano) y lo comparte con la SA como writer. El scope depende de lo
    que clasp haya autorizado: con drive.file no se pueden copiar archivos preexistentes (404/403)."""
    d = cargar()
    if local in d["prueba"]:
        return {"local": local, "ya_existe": d["prueba"][local]}
    ses = _sesion_usuario()
    origen = d["produccion"][local]
    r = ses.post(f"https://www.googleapis.com/drive/v3/files/{origen}/copy",
                 params={"supportsAllDrives": "true", "fields": "id,name"}, json={"name": f"NO CARGAR - PRUEBA - {local}"})
    if r.status_code >= 400:
        err = r.json().get("error", {}) if r.headers.get("content-type", "").startswith("application/json") else {}
        return {"local": local, "ok": False, "http": r.status_code, "status": err.get("status"),
                "mensaje": str(err.get("message", ""))[:240], "razones": [e.get("reason") for e in err.get("errors", [])]}
    nuevo = r.json()
    p = ses.post(f"https://www.googleapis.com/drive/v3/files/{nuevo['id']}/permissions",
                 params={"sendNotificationEmail": "false"}, json={"type": "user", "role": "writer", "emailAddress": compartir_con})
    p.raise_for_status()
    import entornos
    entornos.registrar_prueba(local, nuevo["id"], nuevo["name"], True, ENTORNOS, d["produccion"])
    return {"local": local, "ok": True, "id_prueba": nuevo["id"], "compartido_con": compartir_con}


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
    if len(sys.argv) > 2 and sys.argv[1] == "--auth-usuario":
        print(json.dumps(copiar_como_usuario(sys.argv[2]), ensure_ascii=False, indent=1))
    elif len(sys.argv) > 1 and sys.argv[1] == "--estado":
        print(json.dumps(cargar(), indent=1))
    else:
        print(json.dumps(copiar(sys.argv[1]), ensure_ascii=False, indent=1))
