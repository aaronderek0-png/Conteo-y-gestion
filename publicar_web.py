"""Publica la web de gestión (Streamlit en WebTest/) en GitHub, de punta a punta.

Flujo:
  1. (opción 1) corre el pipeline: comparador de precios -> preview de mail -> análisis de gestión ->
     consolidador -> movimientos explotados.
  2. Sincroniza hacia WebTest/ los últimos Excels por sucursal, el Dossier, los HTML y el app.py, y limpia
     residuos ('(1)', '_nuevo', temporales).
  3. Git en WebTest/: status -> add -> commit 'Actualización datos: <fecha hora>' -> push origin <rama>.
     Streamlit Cloud redeploya solo cuando entra el push.

Uso:
    python publicar_web.py                      # menú interactivo
    python publicar_web.py --opcion 2           # solo subir lo que hay en WebTest/
    python publicar_web.py --opcion 1 --yes     # pipeline completo + publicar, sin preguntas
    python publicar_web.py --opcion 2 --dry-run # no borra, no commitea y hace 'git push --dry-run'

Seguridad: nunca fuerza el push (no usa --force); aborta si hay archivos tipo credenciales/.env en lo que se
va a subir; avisa si el repo es público (los Excels traen costos, stocks y ventas del negocio).
"""
from __future__ import annotations

import argparse
import fnmatch
import json
import os
import re
import shutil
import subprocess
import sys
import urllib.error
import urllib.request
from datetime import datetime
from pathlib import Path

RAIZ = Path(__file__).resolve().parent
WEB = RAIZ / "WebTest"
REPORTES = RAIZ / "Reportes_Ahorro"
PY = sys.executable

# Lo que se publica desde la raíz hacia WebTest/ (patrón -> carpeta de origen)
RESIDUOS = ["* (1).*", "*_nuevo*", "~$*", ".~lock.*", "*.tmp", "*.bak", "Thumbs.db"]
# Nada de esto debe llegar a GitHub aunque alguien lo copie a WebTest/
PROHIBIDOS = ["credentials*.json", "*.pem", "*.key", ".env", ".env.*", "*clasprc*", "token*.json", "*secret*"]
GITIGNORE_WEB = [
    "__pycache__/", "*.pyc", "*.gsheet", "*.gdoc", "*.gslides", "~$*", ".~lock.*", "*.tmp",
    "* (1).*", "*_nuevo*", "credentials*.json", ".env", ".env.*",
]

PIPELINE = [
    ("Comparador de precios (listas de Drive + scrapers)", [PY, "comparador_precios.py"]),
    ("Preview del mail (HTML)", [PY, "enviador_mails.py", "--preview"]),
    ("Análisis de gestión (septiembre y octubre)", [PY, "analisis_gestion_comercial.py", "--anio", "{anio}",
                                                    "--meses", "{meses}"]),
    ("Consolidador (Dossier + informes por local)", [PY, "consolidador.py"]),
    ("Desperdicios y traspasos explotados (BOM)", [PY, "movimientos_locales.py", "--anio", "{anio}",
                                                   "--meses", "{meses}"]),
    ("Balanza de negociación (consumos + mercado)", [PY, "balanza_datos.py"]),
]


# --------------------------------------------------------------------------- #
# utilidades
# --------------------------------------------------------------------------- #
def p(msg: str = "") -> None:
    print(msg, flush=True)


def preguntar(texto: str) -> bool:
    """s/N por teclado. Sin entrada interactiva (p. ej. '!' de Claude Code o un pipe) responde NO y avisa:
    en ese caso usar --yes para aceptar todo."""
    try:
        return input(texto).strip().lower() in {"s", "si", "sí", "y", "yes"}
    except EOFError:
        p("\n   (sin entrada interactiva: se toma como NO; usá --yes para aceptar sin preguntar)")
        return False


def correr(cmd: list[str], cwd: Path = RAIZ, check: bool = False) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, cwd=cwd, capture_output=True, text=True, encoding="utf-8", errors="replace",
                          check=check)


def git(*args: str) -> subprocess.CompletedProcess:
    return correr(["git", "-C", str(WEB), *args], cwd=RAIZ)


def _coincide(nombre: str, patrones: list[str]) -> bool:
    return any(fnmatch.fnmatch(nombre.lower(), pat.lower()) for pat in patrones)


# --------------------------------------------------------------------------- #
# 1. pipeline
# --------------------------------------------------------------------------- #
def correr_pipeline(anio: int, meses: list[int], dry_run: bool) -> bool:
    p("\n=== 1/3 Pipeline ===")
    for titulo, cmd in PIPELINE:
        cmd = [c.replace("{anio}", str(anio)) for c in cmd]
        if "{meses}" in cmd:
            i = cmd.index("{meses}")
            cmd[i:i + 1] = [str(m) for m in meses]
        p(f"-> {titulo}: {' '.join(Path(c).name if c == PY else c for c in cmd)}")
        if dry_run:
            p("   (dry-run: no se ejecuta)")
            continue
        r = correr(cmd)
        if r.returncode != 0:
            cola = "\n".join((r.stderr or r.stdout).strip().splitlines()[-8:])
            p(f"   ✖ falló (código {r.returncode}). No se publica nada.\n{cola}")
            return False
        p("   ✔ ok")
    return True


# --------------------------------------------------------------------------- #
# 2. sincronización y limpieza
# --------------------------------------------------------------------------- #
def ultimos_analisis() -> dict[str, Path]:
    """Por nombre canónico (sin '_nuevo'), la versión modificada más recientemente en la raíz."""
    mejor: dict[str, Path] = {}
    for f in RAIZ.glob("Analisis_Gestion_Comercial_*_*.xlsx"):
        if f.name.startswith("~$") or " (1)" in f.name:
            continue
        canon = f.name.replace("_nuevo", "")
        if canon not in mejor or f.stat().st_mtime > mejor[canon].stat().st_mtime:
            mejor[canon] = f
    return mejor


def archivos_a_sincronizar() -> list[tuple[Path, str]]:
    """(origen, nombre en WebTest/)."""
    out: list[tuple[Path, str]] = [(src, canon) for canon, src in sorted(ultimos_analisis().items())]
    for patron in ("Movimientos_Explotados_*.xlsx", "reporte_gestion_*.html"):
        out += [(f, f.name) for f in sorted(RAIZ.glob(patron)) if "_nuevo" not in f.name and " (1)" not in f.name]
    for nombre in ("Dossier_Mensual_Mishiguene.xlsx", "Balanza_Datos.xlsx", "reporte_items_huerfanos.txt"):
        if (RAIZ / nombre).exists():
            out.append((RAIZ / nombre, nombre))
    # preview del mail: la versión más nueva entre la raíz y Reportes_Ahorro/
    prev = [f for f in (REPORTES / "preview_mail.html", RAIZ / "preview_mail.html") if f.exists()]
    if prev:
        out.append((max(prev, key=lambda f: f.stat().st_mtime), "preview_mail.html"))
    # app.py: si hubiera uno en la raíz más nuevo que el de WebTest/, se publica ese
    app_raiz = RAIZ / "app.py"
    if app_raiz.exists() and (not (WEB / "app.py").exists()
                              or app_raiz.stat().st_mtime > (WEB / "app.py").stat().st_mtime):
        out.append((app_raiz, "app.py"))
    return out


def residuos_en_web() -> list[Path]:
    """Archivos '(1)', '_nuevo' y temporales dentro de WebTest/ (no se tocan .gsheet/.gdoc legítimos)."""
    res = []
    for f in WEB.iterdir():
        if f.is_file() and _coincide(f.name, RESIDUOS):
            res.append(f)
    return sorted(res)


def sincronizar(dry_run: bool, auto_si: bool) -> None:
    p("\n=== 2/3 Sincronización hacia WebTest/ ===")
    cambios = 0
    for src, destino in archivos_a_sincronizar():
        dst = WEB / destino
        igual = dst.exists() and dst.stat().st_size == src.stat().st_size and \
            abs(dst.stat().st_mtime - src.stat().st_mtime) < 2
        if igual:
            continue
        cambios += 1
        hora = datetime.fromtimestamp(src.stat().st_mtime).strftime("%d/%m %H:%M")
        p(f"   {'(dry) ' if dry_run else ''}copia {src.relative_to(RAIZ)}  ->  WebTest/{destino}   [{hora}]")
        if not dry_run:
            shutil.copy2(src, dst)
    if not cambios:
        p("   Los archivos de WebTest/ ya están al día.")

    res = residuos_en_web()
    if res:
        p("   Residuos a limpiar en WebTest/:")
        for f in res:
            p(f"     - {f.name}")
        if dry_run:
            p("   (dry-run: no se borra nada)")
        elif auto_si or preguntar("   ¿Borrarlos? [s/N]: "):
            for f in res:
                f.unlink()
            p(f"   ✔ {len(res)} archivo(s) eliminados")
    if not dry_run:
        asegurar_gitignore()
        for f in WEB.iterdir():           # compilados que no deben viajar
            if f.name == "__pycache__":
                shutil.rmtree(f, ignore_errors=True)


def asegurar_gitignore() -> None:
    ruta = WEB / ".gitignore"
    actual = ruta.read_text(encoding="utf-8").splitlines() if ruta.exists() else []
    faltan = [l for l in GITIGNORE_WEB if l not in actual]
    if faltan:
        ruta.write_text("\n".join(actual + faltan) + "\n", encoding="utf-8")


# --------------------------------------------------------------------------- #
# 3. git
# --------------------------------------------------------------------------- #
def visibilidad_repo(url: str) -> str | None:
    m = re.search(r"github\.com[:/]([^/]+)/([^/.]+)", url)
    if not m:
        return None
    try:
        with urllib.request.urlopen(f"https://api.github.com/repos/{m[1]}/{m[2]}", timeout=10) as r:
            return "público" if not json.load(r).get("private") else "privado"
    except urllib.error.HTTPError as e:
        return "privado (o inexistente)" if e.code == 404 else None
    except Exception:  # noqa: BLE001 - sin red
        return None


def alinear_con_remoto(rama: str, dry_run: bool, auto_si: bool) -> bool:
    """El remoto puede tener commits que no están acá (subidas manuales por la web de GitHub, 'Add files via
    upload'). Si local no tiene commits propios sin subir, se mueve la rama al remoto con `reset --mixed`
    (no toca ningún archivo del disco) y se recuperan los archivos versionados que acá no existen
    (.devcontainer, etc.). Si hay divergencia real (commits en ambos lados) no se toca nada."""
    f = git("fetch", "origin")
    if f.returncode != 0:
        p(f"   ⚠ no pude consultar el remoto (git fetch): {f.stderr.strip()[:200]}")
        return True            # el push dirá si realmente hay un problema
    atras = int(git("rev-list", "--count", f"HEAD..origin/{rama}").stdout.strip() or 0)
    adelante = int(git("rev-list", "--count", f"origin/{rama}..HEAD").stdout.strip() or 0)
    if not atras:
        return True
    p(f"   El remoto tiene {atras} commit(s) que no están acá (subidas manuales desde la web de GitHub):")
    for l in git("log", "--format=      %h  %ad  %s", "--date=format:%d/%m %H:%M",
                 f"HEAD..origin/{rama}").stdout.splitlines()[:6]:
        p(l)
    if adelante:
        p(f"   ✖ Además hay {adelante} commit(s) locales sin subir: es una divergencia real. Resolvela a mano "
          "(git -C WebTest pull --rebase) y volvé a correr el script.")
        return False
    if dry_run:
        p(f"   (dry-run: se alinearía con 'git reset --mixed origin/{rama}', sin tocar tus archivos; "
          "el push simulado se omite)")
        return True
    if not auto_si and not preguntar(f"   ¿Alinear con origin/{rama} conservando tus archivos? [s/N]: "):
        p("   Cancelado.")
        return False
    r = git("reset", "--mixed", f"origin/{rama}")
    if r.returncode != 0:
        p(f"   ✖ git reset falló: {r.stderr.strip()}")
        return False
    # archivos versionados en el remoto que no existen acá: se restauran (no se borran del repo por accidente)
    faltan = [x for x in git("ls-files", "--deleted").stdout.splitlines()
              if not x.startswith("__pycache__/") and not _coincide(Path(x).name, RESIDUOS)]
    if faltan:
        git("checkout", "--", *faltan)
        p(f"   ✔ restaurados desde el remoto: {', '.join(faltan[:6])}{' ...' if len(faltan) > 6 else ''}")
    p(f"   ✔ rama alineada con origin/{rama}")
    return True


def publicar_git(dry_run: bool, auto_si: bool, rama: str | None) -> bool:
    p("\n=== 3/3 Git (WebTest/) ===")
    if git("rev-parse", "--is-inside-work-tree").returncode != 0:
        p("   ✖ WebTest/ no es un repositorio git.")
        return False
    remoto = git("remote", "get-url", "origin").stdout.strip()
    rama = rama or git("rev-parse", "--abbrev-ref", "HEAD").stdout.strip() or "main"
    p(f"   remoto: {remoto}  ·  rama: {rama}")
    vis = visibilidad_repo(remoto)
    if vis:
        p(f"   visibilidad del repo: {vis}")
        if vis == "público":
            p("   ⚠ el repo es PÚBLICO: los Excels publican costos, stock y ventas. Revisá antes de continuar.")

    if not alinear_con_remoto(rama, dry_run, auto_si):
        return False

    # 'git rm --cached' de lo que ahora es ruido pero estaba versionado (pyc, '_nuevo'...)
    if not dry_run:
        git("rm", "-r", "--cached", "--ignore-unmatch", "-q", "__pycache__")

    st = git("status", "--porcelain")
    lineas = [l for l in st.stdout.splitlines() if l.strip()]
    if not lineas:
        p("   Sin cambios que subir (el repo ya está al día).")
        # puede haber commits locales sin push
        ahead = git("rev-list", "--count", f"origin/{rama}..HEAD").stdout.strip()
        if ahead in ("", "0"):
            return True
        p(f"   Hay {ahead} commit(s) locales sin subir.")
    else:
        mod = sum(1 for l in lineas if l[:2].strip() in {"M", "MM", "AM"})
        nue = sum(1 for l in lineas if l.startswith("??"))
        bor = sum(1 for l in lineas if l[:2].strip() == "D")
        p(f"   git status: {len(lineas)} cambio(s): {mod} modificados · {nue} nuevos · {bor} borrados")
        for l in lineas[:25]:
            p(f"     {l}")
        if len(lineas) > 25:
            p(f"     ... y {len(lineas) - 25} más")
        prohib = [l[3:].strip('"') for l in lineas if _coincide(Path(l[3:].strip('"')).name, PROHIBIDOS)]
        if prohib:
            p(f"   ✖ Abortado: archivos sensibles en lo que se iba a subir: {prohib}")
            return False

        if dry_run:
            p("   (dry-run: no se hace add/commit)")
        else:
            if not auto_si and not preguntar("   ¿Hacer commit y push? [s/N]: "):
                p("   Cancelado.")
                return False
            r = git("add", "-A")
            if r.returncode != 0:
                p(f"   ✖ git add falló: {r.stderr.strip()}")
                return False
            msg = f"Actualización datos: {datetime.now():%Y-%m-%d %H:%M}"
            r = git("commit", "-m", msg)
            if r.returncode != 0:
                p(f"   ✖ git commit falló: {(r.stderr or r.stdout).strip()}")
                return False
            p(f"   ✔ commit: {msg}")

    if dry_run and int(git("rev-list", "--count", f"HEAD..origin/{rama}").stdout.strip() or 0):
        p("   (dry-run: el push simulado se omite hasta alinear con el remoto)")
        return True
    args = ["push", "origin", rama] + (["--dry-run"] if dry_run else [])
    r = git(*args)
    salida = (r.stdout + r.stderr).strip()
    if r.returncode != 0:
        p(f"   ✖ git push falló:\n{salida}")
        if "rejected" in salida or "non-fast-forward" in salida:
            p("   El remoto tiene commits que no están acá. Hacé 'git -C WebTest pull --rebase' y reintentá "
              "(el script nunca usa --force).")
        return False
    if dry_run:
        p(f"   ✔ dry-run del push OK (nada se subió):\n   {salida.replace(chr(10), chr(10) + '   ')}")
    else:
        p("   🚀 Datos subidos a GitHub. Streamlit Cloud se actualizará en 1 minuto")
    return True


# --------------------------------------------------------------------------- #
def menu() -> str:
    p("\n  Publicación de la web de gestión Mishiguene")
    p("  [1] Correr pipeline completo + Publicar a la Web")
    p("  [2] Solo subir archivos actuales de WebTest a GitHub")
    p("  [3] Salir")
    try:
        return input("\n  Opción: ").strip()
    except EOFError:
        p("\n  (sin entrada interactiva: usá --opcion 1|2)")
        return "3"


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--opcion", choices=["1", "2", "3"], help="saltea el menú")
    ap.add_argument("--dry-run", action="store_true", help="no copia/borra/commitea; hace 'git push --dry-run'")
    ap.add_argument("--yes", action="store_true", help="no pregunta confirmaciones")
    ap.add_argument("--anio", type=int, default=2026)
    ap.add_argument("--meses", type=int, nargs="+", default=[9, 10])
    ap.add_argument("--rama", help="rama destino (default: la actual de WebTest/)")
    args = ap.parse_args()

    opcion = args.opcion or menu()
    if opcion == "3" or opcion not in {"1", "2"}:
        p("Chau.")
        return 0
    if args.dry_run:
        p("\n[DRY-RUN] No se modifica nada; el push se simula con 'git push --dry-run'.")
    if opcion == "1" and not correr_pipeline(args.anio, args.meses, args.dry_run):
        return 1
    sincronizar(args.dry_run, args.yes)
    ok = publicar_git(args.dry_run, args.yes, args.rama)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
