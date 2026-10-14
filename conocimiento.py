"""Base de conocimiento operativo: mapa de fuentes, grafo de dependencias e invalidación de cachés derivadas.

Archivos (en conocimiento/): fuentes.json (mapa de fuentes; ver CONOCIMIENTO.md), estado_cache.json (huella de
cada fuente cuando se calculó cada derivado). Cada fuente/derivado conserva: id, tipo, ubicación lógica, esquema_version,
fecha_consulta, vigencia, estado (observado | validado | inferencia | estimacion | pendiente).

Regla: un derivado es VÁLIDO solo si la huella de TODAS sus fuentes (transitivas) coincide con la guardada.
`registrar_fuente(id, huella)` -> devuelve los derivados invalidados (en orden topológico) para recalcular.
Una inferencia no pisa una decisión validada (`registrar_dato` lo rechaza).
"""
from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DIR = BASE_DIR / "conocimiento"
FUENTES = DIR / "fuentes.json"
ESTADO = DIR / "estado_cache.json"
JERARQUIA = {"validado": 4, "observado": 3, "inferencia": 2, "estimacion": 1, "pendiente": 0}

# derivado -> fuentes/derivados de los que depende (orden del circuito del negocio)
DEPENDENCIAS = {
    "lectura_comprobante": ["adjunto_sha"],
    "mapeo_producto": ["maestro", "aprendizaje_proveedores"],
    "filas_factura": ["lectura_comprobante", "mapeo_producto", "precios_vigentes", "planilla_facturas"],
    "escritura_verificada": [],                      # hecho observado en el destino: nunca se recalcula por precios
    "precios_vigentes": ["planilla_facturas", "listas_proveedores"],
    "costos_bom": ["recetas", "precios_vigentes", "presentaciones"],
    "gestion_margenes": ["costos_bom", "ventas_maxirest"],
    "comparativa_precios": ["precios_vigentes", "listas_proveedores", "costos_bom", "mapeo_producto"],
    "reporte_web": ["gestion_margenes", "comparativa_precios"],
}


def _ahora() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _json(path: Path, defecto):
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return defecto


def _guardar(path: Path, datos) -> None:
    DIR.mkdir(exist_ok=True)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp.replace(path)


def huella_archivo(path: Path) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()[:16]


def dependientes(fuente: str) -> list[str]:
    """Derivados afectados por un cambio en `fuente`, en orden topológico."""
    afectados: list[str] = []
    def visitar(n):
        for d, deps in DEPENDENCIAS.items():
            if n in deps and d not in afectados:
                afectados.append(d)
                visitar(d)
    visitar(fuente)
    orden = list(DEPENDENCIAS)
    return sorted(afectados, key=orden.index)


def registrar_fuente(fuente: str, huella: str, ruta: Path = ESTADO) -> list[str]:
    """Anota la huella actual de una fuente. Si cambió, devuelve los derivados a recalcular y los marca inválidos."""
    e = _json(ruta, {"fuentes": {}, "derivados": {}})
    anterior = e["fuentes"].get(fuente, {}).get("huella")
    e["fuentes"][fuente] = {"huella": huella, "consultada": _ahora()}
    invalidos = []
    if anterior is not None and anterior != huella:
        invalidos = dependientes(fuente)
        for d in invalidos:
            if d in e["derivados"]:
                e["derivados"][d]["valido"] = False
                e["derivados"][d]["motivo"] = f"cambió {fuente}"
    _guardar(ruta, e)
    return invalidos


def marcar_calculado(derivado: str, fuentes_huella: dict[str, str], version: str = "1", ruta: Path = ESTADO) -> None:
    e = _json(ruta, {"fuentes": {}, "derivados": {}})
    e["derivados"][derivado] = {"valido": True, "calculado": _ahora(), "version": version, "huellas": fuentes_huella}
    _guardar(ruta, e)


def es_valido(derivado: str, ruta: Path = ESTADO) -> bool:
    e = _json(ruta, {"fuentes": {}, "derivados": {}})
    d = e["derivados"].get(derivado)
    if not d or not d.get("valido"):
        return False
    return all(e["fuentes"].get(f, {}).get("huella") == h for f, h in d["huellas"].items())


def registrar_dato(clave: str, valor, estado: str, fuente: str, vigencia: str | None = None,
                   ruta: Path = FUENTES) -> dict:
    """Guarda un dato con procedencia. Rechaza degradar un dato de mayor jerarquía (validado) con una inferencia."""
    if estado not in JERARQUIA:
        raise ValueError(estado)
    base = _json(ruta, {"version": 1, "datos": {}})
    previo = base["datos"].get(clave)
    if previo and JERARQUIA[previo["estado"]] > JERARQUIA[estado] and previo["valor"] != valor:
        return {"aceptado": False, "motivo": f"existe {previo['estado']} para {clave}; {estado} no lo reemplaza"}
    base["datos"][clave] = {"valor": valor, "estado": estado, "fuente": fuente, "consulta": _ahora(),
                            "vigencia": vigencia, "version": (previo or {}).get("version", 0) + 1}
    _guardar(ruta, base)
    return {"aceptado": True, "version": base["datos"][clave]["version"]}
