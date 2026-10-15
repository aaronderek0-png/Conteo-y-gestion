"""Coordinador de trabajos persistente (cola recuperable) del circuito comprobantes -> planillas.

Estados: RECIBIDO -> LEIDO -> VALIDADO -> (REVISAR | ESCRITO) ; cualquiera -> ERROR (reintentable).
Identidad (idempotencia): sha256 del adjunto [+ '#n' comprobante]. Un trabajo ESCRITO no se reabre salvo `forzar`.
Un fallo NUNCA marca ESCRITO ni consume el trabajo: queda ERROR con `intentos` y el último error saneado.
Un único escritor por destino: `destino_lock(local)` (archivo de bloqueo con timeout), para que workers no compitan.

Archivo: estado/trabajos.json (escritura atómica tmp+replace, bloqueo de proceso por estado/trabajos.lock).
No es un servicio: lo usan telegram_ingestor / recuperar_lote / procesar_lote_wpp como librería.
"""
from __future__ import annotations

import json
import os
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
ESTADO_DIR = Path(os.environ.get("MISHI_ESTADO_DIR") or BASE_DIR / "estado")
TRABAJOS = ESTADO_DIR / "trabajos.json"

ESTADOS = ("RECIBIDO", "LEIDO", "VALIDADO", "REVISAR", "ESCRITO", "ERROR")
_PERMITIDAS = {
    "RECIBIDO": {"LEIDO", "ERROR", "REVISAR"},
    "LEIDO": {"VALIDADO", "REVISAR", "ERROR"},
    "VALIDADO": {"ESCRITO", "REVISAR", "ERROR"},
    "REVISAR": {"VALIDADO", "ESCRITO", "ERROR", "LEIDO"},
    "ERROR": {"RECIBIDO", "LEIDO", "VALIDADO", "REVISAR", "ESCRITO"},
    "ESCRITO": set(),
}


class Bloqueado(RuntimeError):
    pass


@contextmanager
def _bloqueo(nombre: str, timeout: float = 20.0, base: Path | None = None):
    d = base or ESTADO_DIR
    d.mkdir(parents=True, exist_ok=True)
    lock = d / f"{nombre}.lock"
    t0 = time.time()
    while True:
        try:
            fd = os.open(str(lock), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.write(fd, str(os.getpid()).encode())
            os.close(fd)
            break
        except FileExistsError:
            try:                                          # bloqueo huérfano (>5 min): se libera
                if time.time() - lock.stat().st_mtime > 300:
                    lock.unlink()
                    continue
            except OSError:
                pass
            if time.time() - t0 > timeout:
                raise Bloqueado(f"no pude tomar el bloqueo {nombre}")
            time.sleep(0.1)
    try:
        yield
    finally:
        try:
            lock.unlink()
        except OSError:
            pass


def destino_lock(local: str, timeout: float = 120.0):
    """Único escritor por destino (planilla). Uso: `with destino_lock('Congreso'): push...`"""
    return _bloqueo(f"destino_{local}", timeout)


def _leer(path: Path) -> dict:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {"version": 1, "trabajos": {}}


def _guardar(path: Path, datos: dict) -> None:
    tmp = path.with_suffix(f".{os.getpid()}.tmp")
    tmp.write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    for n in range(8):                                  # Windows: otro proceso puede estar leyendo el destino
        try:
            tmp.replace(path)
            return
        except PermissionError:
            time.sleep(0.05 * (n + 1))
    tmp.replace(path)


class Cola:
    def __init__(self, path: Path = TRABAJOS):
        self.path = path

    def _mutar(self, fn):
        with _bloqueo("trabajos", base=self.path.parent):
            d = _leer(self.path)
            r = fn(d["trabajos"])
            _guardar(self.path, d)
            return r

    def registrar(self, clave: str, archivo: str = "", origen: str = "telegram", **meta) -> dict:
        """Idempotente: si ya existe, devuelve el trabajo existente sin cambiarlo."""
        def f(t):
            if clave not in t:
                t[clave] = {"clave": clave, "archivo": archivo, "origen": origen, "estado": "RECIBIDO", "intentos": 0,
                            "creado": _ahora(), "historial": [{"ts": _ahora(), "estado": "RECIBIDO"}], **meta}
            return dict(t[clave])
        return self._mutar(f)

    def transicion(self, clave: str, estado: str, evidencia: dict | None = None, error: str | None = None,
                   forzar: bool = False) -> dict:
        if estado not in ESTADOS:
            raise ValueError(estado)

        def f(t):
            j = t.get(clave)
            if j is None:
                raise KeyError(clave)
            if j["estado"] == estado:
                return dict(j)
            if estado not in _PERMITIDAS[j["estado"]] and not forzar:
                raise ValueError(f"{j['estado']} -> {estado} no permitido ({clave})")
            if estado == "ERROR":
                j["intentos"] += 1
                j["ultimo_error"] = (error or "")[:300]
            j["estado"] = estado
            j.pop("en_proceso", None)
            j.setdefault("historial", []).append({"ts": _ahora(), "estado": estado, "evidencia": evidencia or {},
                                                  **({"error": (error or "")[:300]} if error else {})})
            if evidencia:
                j.setdefault("evidencia", {}).update(evidencia)
            return dict(j)
        return self._mutar(f)

    def tomar(self, n: int = 20, max_intentos: int = 5, rancio_s: int = 900, solo_demo: bool | None = None) -> list[dict]:
        """Reclama hasta `n` trabajos (RECIBIDO, o ERROR con intentos restantes) marcándolos `en_proceso`. Un trabajo
        reclamado hace más de `rancio_s` (proceso caído) vuelve a estar disponible. Orden: más antiguos primero."""
        ahora = time.time()
        # lectura sin bloqueo ni escritura: si no hay nada que reclamar no se toca el disco (el worker consulta cada 0,5 s)
        vistos = _leer(self.path)["trabajos"].values()
        if not any((j["estado"] == "RECIBIDO" or (j["estado"] == "ERROR" and j["intentos"] < max_intentos))
                   and not (j.get("en_proceso") and ahora - j["en_proceso"] < rancio_s)
                   and (solo_demo is None or bool(j.get("demo")) == solo_demo) for j in vistos):
            return []

        def f(t):
            libres = []
            for j in t.values():
                if j.get("en_proceso") and ahora - j["en_proceso"] < rancio_s:
                    continue
                if solo_demo is not None and bool(j.get("demo")) != solo_demo:
                    continue
                if j["estado"] == "RECIBIDO" or (j["estado"] == "ERROR" and j["intentos"] < max_intentos):
                    libres.append(j)
            libres.sort(key=lambda j: j.get("creado", ""))
            elegidos = libres[:n]
            for j in elegidos:
                j["en_proceso"] = ahora
            return [dict(j) for j in elegidos]
        return self._mutar(f)

    def liberar(self, clave: str) -> None:
        def f(t):
            if clave in t:
                t[clave].pop("en_proceso", None)
        self._mutar(f)

    def actualizar(self, clave: str, **meta) -> None:
        def f(t):
            if clave in t:
                t[clave].update(meta)
        self._mutar(f)

    def obtener(self, clave: str) -> dict | None:
        return _leer(self.path)["trabajos"].get(clave)

    def pendientes(self, max_intentos: int = 5) -> list[dict]:
        """Trabajos que hay que (re)intentar: no ESCRITO, ERROR con intentos restantes o REVISAR sin resolver excluido."""
        out = []
        for j in _leer(self.path)["trabajos"].values():
            if j["estado"] in ("RECIBIDO", "LEIDO", "VALIDADO") or (j["estado"] == "ERROR" and j["intentos"] < max_intentos):
                out.append(j)
        return out

    def resumen(self) -> dict:
        r = {e: 0 for e in ESTADOS}
        for j in _leer(self.path)["trabajos"].values():
            r[j["estado"]] += 1
        return r


def _ahora() -> str:
    return datetime.now().isoformat(timespec="seconds")
