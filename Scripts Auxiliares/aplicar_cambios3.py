import os
os.chdir(r"C:\Users\EXO G12\Desktop\Agente Mishiguene")
s = open("telegram_ingestor.py", encoding="utf-8").read()


def rep(a, b):
    global s
    assert a in s, a[:60]
    s = s.replace(a, b, 1)


rep('ENTRADA = BASE_DIR / "Facturas_WPP" / "telegram"',
    'ENTRADA = BASE_DIR / "Facturas Mishiguene"          # carpeta diaria (sin guiones); lo que llega por Telegram cae acá\n'
    'ENV_FILE = BASE_DIR / ".env"\n'
    'NOMBRE_GRUPO = "Facturas Mishiguene"')
rep('tmp = ENTRADA / "_lote_actual"', 'tmp = BASE_DIR / "Carpeta Revision" / "lote actual temporal"')
rep('''    python telegram_ingestor.py --once --dry-run          # procesa sin escribir en las copias
''', '''    python telegram_ingestor.py --once --dry-run          # procesa sin escribir en las copias
    python telegram_ingestor.py --conectar <TOKEN>        # una sola vez: guarda el token en .env, detecta el grupo
                                                          # "Facturas Mishiguene", lo guarda y manda el mensaje de prueba
''')
rep('''def crear_tarea(minutos: int = 15) -> list[str]:''', '''# --------------------------------------------------------------------------- #
# .env y auto-conexión
# --------------------------------------------------------------------------- #
def cargar_env(path: Path | None = None) -> None:
    """Carga KEY=VALUE de .env en os.environ (sin pisar lo que ya esté definido)."""
    p = path or ENV_FILE
    if not p.is_file():
        return
    for linea in p.read_text(encoding="utf-8").splitlines():
        linea = linea.strip()
        if linea and not linea.startswith("#") and "=" in linea:
            k, v = linea.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def guardar_env(clave: str, valor: str, path: Path | None = None) -> None:
    p = path or ENV_FILE
    lineas = p.read_text(encoding="utf-8").splitlines() if p.is_file() else []
    nuevas, visto = [], False
    for l in lineas:
        if l.split("=", 1)[0].strip() == clave:
            nuevas.append(f"{clave}={valor}")
            visto = True
        else:
            nuevas.append(l)
    if not visto:
        nuevas.append(f"{clave}={valor}")
    p.write_text("\\n".join(nuevas) + "\\n", encoding="utf-8")
    os.environ[clave] = valor


def _norm_titulo(t: str) -> str:
    return " ".join(str(t or "").lower().replace("_", " ").replace("-", " ").split())


def conectar(token: str, http: Callable[..., Any] | None = None, escribir: bool = True) -> dict:
    """--conectar TOKEN: guarda el token en .env, detecta el chat_id del grupo 'Facturas Mishiguene' (o el último grupo
    donde hayan hablado), lo guarda como TELEGRAM_CHAT_ID, manda el mensaje de prueba y procesa lo pendiente."""
    api = TelegramAPI(token, http)
    me = api._call("getMe")
    guardar_env("TELEGRAM_BOT_TOKEN", token)
    ups = api.get_updates(None, timeout=0)          # sin confirmar: pasada() vuelve a leerlos para procesarlos
    grupos: list[dict] = []
    for u in ups:
        for k in ("message", "channel_post", "my_chat_member", "edited_message"):
            ch = (u.get(k) or {}).get("chat")
            if ch and ch.get("type") in ("group", "supergroup", "channel"):
                grupos.append(ch)
    if not grupos:
        return {"conectado": False, "bot": me.get("username"),
                "motivo": f"el bot todavía no vio ningún grupo: agregá @{me.get('username')} al grupo "
                          f"'{NOMBRE_GRUPO}', escribí cualquier mensaje y volvé a correr el comando "
                          "(si el bot tiene 'Group Privacy' activada, desactivala en @BotFather o hacelo admin)"}
    elegido = next((g for g in reversed(grupos) if _norm_titulo(g.get("title")) == _norm_titulo(NOMBRE_GRUPO)), grupos[-1])
    chat_id = elegido["id"]
    guardar_env("TELEGRAM_CHAT_ID", str(chat_id))
    api.responder(chat_id, f"✅ Agente Mishiguene conectado a {NOMBRE_GRUPO}. "
                           "Mandá fotos o PDFs acá y los proceso solos.")
    stats = pasada(api, escribir=escribir, chat_permitido=str(chat_id))
    return {"conectado": True, "bot": me.get("username"), "grupo": elegido.get("title"), "chat_id": chat_id,
            "coincide_nombre": _norm_titulo(elegido.get("title")) == _norm_titulo(NOMBRE_GRUPO), "pasada": stats}


def crear_tarea(minutos: int = 15) -> list[str]:''')
rep('''    ap.add_argument("--dry-run", action="store_true", help="Procesa sin escribir en las copias")
    args = ap.parse_args(argv)
''', '''    ap.add_argument("--dry-run", action="store_true", help="Procesa sin escribir en las copias")
    ap.add_argument("--conectar", metavar="TOKEN", help="Guarda el token en .env, detecta el grupo y manda el mensaje de prueba")
    args = ap.parse_args(argv)
    cargar_env()

    if args.conectar:
        r = conectar(args.conectar.strip(), escribir=not args.dry_run)
        if r["conectado"]:
            print(f"Conectado: bot @{r['bot']} -> grupo '{r['grupo']}' (chat_id {r['chat_id']}); .env actualizado. "
                  f"Pendientes procesados: {r['pasada']}")
            if not r["coincide_nombre"]:
                print(f"Aviso: el grupo detectado no se llama '{NOMBRE_GRUPO}' (se tomó el último grupo con actividad).")
            return 0
        print("No conectado:", r["motivo"])
        return 3
''')
open("telegram_ingestor.py", "w", encoding="utf-8").write(s)
print("ok")
