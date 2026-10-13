"""Sincroniza TODOS los Apps Script del ecosistema Mishiguene (GAS/*) con clasp.

    python gas_sync.py pull            # baja la última versión de todos (seguro)
    python gas_sync.py status          # lista proyectos y cantidad de archivos
    python gas_sync.py push <carpeta>  # sube UN proyecto (pide confirmar; nunca masivo)
"""
import subprocess, sys, json
from pathlib import Path

RAIZ = Path(__file__).parent
GAS = RAIZ / "GAS"
CLASP = str(RAIZ / "node_modules" / ".bin" / "clasp.cmd")

def proyectos():
    return sorted(p.parent for p in GAS.glob("*/.clasp.json") if not p.parent.name.startswith("_"))

def run(d, *args):
    return subprocess.run([CLASP, *args], cwd=d, capture_output=True, text=True, shell=False)

def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    if cmd == "status":
        for d in proyectos():
            n = len([x for x in d.iterdir() if x.suffix in (".js", ".gs", ".html")])
            print(f"{d.name:32} {n} archivos")
    elif cmd == "pull":
        for d in proyectos():
            r = run(d, "pull")
            print(f"{'OK ' if r.returncode == 0 else 'ERR'} {d.name}  {r.stderr.strip()[:120]}")
    elif cmd == "push" and len(sys.argv) > 2:
        d = GAS / sys.argv[2]
        if input(f"Subir {d.name} al Apps Script real? (escribí SI) ") == "SI":
            r = run(d, "push", "-f"); print(r.stdout or r.stderr)
    else:
        print(__doc__)

if __name__ == "__main__":
    main()
