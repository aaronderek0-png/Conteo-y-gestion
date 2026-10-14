# --------------------------------------------------------------------------- #
# Comprobantes reales de WhatsApp: tipo, copias, formatos de proveedor y LOCAL DE DESTINO
# --------------------------------------------------------------------------- #
# Razón social de los 4 locales (y de Chimi Doha): NO identifica el local por sí sola.
RAZON_SOCIAL_GRUPO = re.compile(r"lafinur\s*3368", re.I)
CUIT_LAFINUR_3368 = "30-71456509-1"
CUIT_CHIMI_DOHA = "30-71567213-4"


def _plano(t: str | None) -> str:
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKD", t or "").encode("ascii", "ignore").decode().lower()).strip()


def detectar_local(sello: str | None = None, direccion_entrega: str | None = None, razon_social: str | None = None,
                   manuscrito: str | None = None, texto_cliente: str | None = None) -> dict:
    """Local de destino de un comprobante según la geografía real de Mishiguene.

    Prioridad: 1) sello (o 'RECIBIDO' manuscrito) · 2) anotación manuscrita del local · 3) dirección de entrega ·
    4) razón social propia de un local (Chimi Doha). 'LAFINUR 3368 SRL' es la razón social de los 4 locales
    (trampa fiscal): por sí sola NO asigna Lafinur y deja el local INDETERMINADO.

    Edificio Av. Congreso esq. Montañeses 1602 (Núñez): Planta Baja = Café Mishiguene Núñez (Chimidoha);
    Planta Alta = Centro de Producción (Congreso = Obrador, sello 'MISHIGUENE CERVIÑO').
    Devuelve {'local', 'regla', 'alertas'}."""
    s, d, r, m = _plano(sello), _plano(direccion_entrega), _plano(razon_social), _plano(manuscrito)
    alertas: list[str] = []
    local, regla = None, None

    def por_texto(t: str, origen: str):
        if not t:
            return None, None
        if "cervi" in t:
            return "Congreso", f"{origen}: Mishiguene Cerviño = Centro de Producción/Obrador (PA del 1602)"
        if re.search(r"cafe\s+mishiguene\s+nunez|\bnunez\b|\bchimi\b", t):
            return "Chimidoha", f"{origen}: Café Mishiguene Núñez (PB del 1602)"
        if re.search(r"cafe\s+mishiguene\s+cabello|\bcabello\b", t):
            return "Cabello", f"{origen}: Cabello"
        if re.search(r"\barcos\b", t):
            return "Arcos", f"{origen}: Arcos"
        if re.search(r"\blafi\b|\blafinur\b", t) and not RAZON_SOCIAL_GRUPO.search(t):
            return "Lafinur", f"{origen}: Lafinur"
        return None, None

    local, regla = por_texto(s, "sello")
    if local is None and m:
        local, regla = por_texto(m, "anotación manuscrita")
    if local is None and d:
        if re.search(r"cabello\s*3181", d):
            local, regla = "Cabello", "dirección de entrega: Cabello 3181"
        elif re.search(r"\barcos\b", d):
            local, regla = "Arcos", "dirección de entrega: calle Arcos"
        elif re.search(r"cervi[nñ]o", d):
            local, regla = "Congreso", "dirección de entrega: Cerviño (Obrador)"
        elif re.search(r"congreso.*1602|montanes\w*\s*1602|1602.*montanes|congreso av\.? 1602", d):
            if re.search(r"chimi|nunez", r):
                local, regla = "Chimidoha", "dirección Congreso 1602 + razón social Chimi Doha / Café Núñez"
            else:
                local, regla = "Congreso", ("dirección Av. Congreso esq. Montañeses 1602 + Lafinur 3368 SRL sin sello de "
                                            "Café Núñez → Planta Alta / Obrador")
    if local is None and re.search(r"chimi\s*doha|cafe mishiguene nunez", r):
        local, regla = "Chimidoha", "razón social Chimi Doha / Café Mishiguene Núñez"
    # alertas de conflicto
    if local == "Congreso" and re.search(r"chimi\s*doha", r + " " + _plano(texto_cliente)):
        alertas.append("RAZÓN SOCIAL CHIMI DOHA con sello/nota Cerviño: verificar si es gasto del café (Chimidoha)")
    if local == "Chimidoha" and re.search(r"\barcos\b", m):
        alertas.append("anotación 'Arcos' en un comprobante de Chimidoha: posible derivación/devolución")
    if local is None:
        alertas.append("local INDETERMINADO: 'Lafinur 3368 SRL' es la razón social de los 4 locales; sin sello ni "
                       "dirección de entrega no se asigna Lafinur")
    return {"local": local, "regla": regla, "alertas": alertas}


def separar_copias(texto: str) -> str:
    """Los PDF de AFIP traen ORIGINAL / DUPLICADO / TRIPLICADO: se usa solo la primera copia."""
    corte = re.search(r"^\s*(DUPLICADO|TRIPLICADO)\s*$", texto, re.M)
    if corte:
        return texto[:corte.start()]
    m = re.search(r"\bDUPLICADO\b", texto)
    if m and re.search(r"\bORIGINAL\b", texto[:m.start()]):
        return texto[:m.start()]
    return texto


def tipo_comprobante(texto: str) -> str:
    cab = _plano(texto[:700])
    if re.search(r"nota de credito", cab):
        return "NC"
    if re.search(r"nota de debito", cab):
        return "ND"
    if re.search(r"lista\s+(naturales|de precios)", cab):
        return "Lista"
    if re.search(r"documento no valido como factura", cab):
        return "Remito"
    return "Factura"


_FMT_AFIP = re.compile(
    r"^(?:\d+\s+(?=[A-Za-zÁÉÍÓÚÑ]))?(?P<d>.+?)\s+(?P<c>\d+(?:,\d+)?)\s+(?P<u>kg|kgs|unidades|unidad|otras unidades|lts?|u)\s+"
    r"(?P<p>\d+(?:,\d+)?)\s+(?P<b>\d+,\d{2})\s+(?:(?P<ib>\d+,\d{2})\s+)?(?P<s>\d+,\d{2})"
    r"(?:\s+(?P<iva>\d+(?:,\d+)?)%\s+(?P<si>\d+,\d{2}))?\s*$", re.I)
_FMT_RAICES = re.compile(
    r"^(?P<h>(?:\d+\s+)?\d+\s+[A-ZÁÉÍÓÚÑ0-9][A-ZÁÉÍÓÚÑ0-9 .]*?)\s+(?P<c>\d+\.\d{2})\s+(?P<p>[\d,]+\.\d{2})\s+(?P<t>[\d,]+\.\d{2})\s*$")
_FMT_FRIMAQ = re.compile(r"^(?P<c>\d+)\s+(?P<d>.+?)\s+(?P<p>\d+\.\d{2})\s+(?P<t>[\d,]+\.\d{2})\s*$")
_FMT_RODOFELI = re.compile(
    r"^(?P<cod>\S+(?:\s/\s\S+)?)\s+(?P<c>[\d,]+\.\d{2})\s+(?P<d>.+?)\s+(?P<p>[\d.]+)\s+\$\s+(?P<t>[\d,]+\.\d{2})\s*$")
_FMT_ALVEOLO = re.compile(
    r"^(?P<c>\d+(?:,\d+)?)\s+(?P<cod>\S+)\s+(?P<d>.+?)\s+(?P<p>[\d.]+,\d{2})\s+(?P<iva>[\d,]+)\s*%\s+(?P<b>[\d,]+)\s*%\s+"
    r"(?P<t>[\d.]+,\d{2})\s*$")
_FMT_NC_FREE = re.compile(
    r"^(?P<d>.+?)\s+(?P<c>[\d.,]+)\s+(?P<u>Kgs?|Un\w*)\s+(?P<p>[\d.]+,\d{2})\s+(?P<b>[\d,]+)\s+IVA\s+[\d.,]+%\s+\$\s*(?P<t>[\d.]+,\d{2})\s*$")
_FMT_NC_POSTA = re.compile(
    r"^(?P<d>.+?)\s+(?P<cod>[A-Z]{2}\d+)\s+(?P<u>UN|CAJ|KG)\s+(?P<c>[\d,]+)\s+(?P<p>[\d.]+,\d{2})\s+(?P<b>[\d,]+)%\s+(?P<t>[\d.]+,\d{2})\s*$")
_FMT_NC_VILLARES = re.compile(r"^(?P<d>NC MERC BONIFICADA.*?)\s+\([\d.]+%\)\s+(?P<p>[\d,]+\.\d{2})\s+(?P<t>[\d,]+\.\d{2})\s*$")

_UNID_RAICES = {"KILOS": "KG", "KILO": "KG", "UNIDAD": "UN", "CAJON": "CAJON", "BOLSAS": "BOLSA", "BOLSA": "BOLSA",
                "PAQUETE": "PAQ", "CAJA": "CAJA", "TORITO": "CAJON", "ATADO": "ATADO"}


def _item_pdf(d, c, u, p, b, t) -> dict:
    d = re.sub(r"\s{2,}", " ", d).strip(" .-")
    return {"descripcion_original": d, "cantidad": c, "unidad_detectada": u, "precio_unitario": p,
            "descuento_pct": b or 0.0, "total_linea": t,
            "linea_consistente": abs(c * p * (1 - (b or 0) / 100) - t) <= max(1.0, 0.02 * abs(t))}


def items_pdf_comprobante(texto: str) -> list[dict]:
    """Renglones de los formatos de proveedor vistos en Facturas_WPP/ (AFIP/ARCA genérico, Raíces, Frimaq,
    Rodofeli, Alveolo, NC de Free / Posta / Villares)."""
    items: list[dict] = []
    for ln in texto.splitlines():
        ln = ln.strip()
        if not ln:
            continue
        m = _FMT_AFIP.match(ln)
        if m:
            g = m.groupdict()
            u = {"kg": "KG", "kgs": "KG", "lt": "LT", "lts": "LT"}.get(g["u"].lower(), "UN")
            items.append(_item_pdf(g["d"], parse_number(g["c"]), u, parse_number(g["p"]), parse_number(g["b"]),
                                   parse_number(g["s"])))
            continue
        m = _FMT_RAICES.match(ln)
        if m:
            h = re.sub(r"\.{2,}", " ", m["h"])
            h = re.sub(r"^\d+\s+\d+\s+|^\d+\s+", "", h.strip(), count=1).strip()
            toks = h.split()
            u = _UNID_RAICES.get(toks[-1].upper()) if toks else None
            items.append(_item_pdf(" ".join(toks), float(m["c"]), u, float(m["p"].replace(",", "")), 0.0,
                                   float(m["t"].replace(",", ""))))
            continue
        for rx, kind in ((_FMT_RODOFELI, "rod"), (_FMT_ALVEOLO, "alv"), (_FMT_NC_FREE, "free"),
                         (_FMT_NC_POSTA, "posta"), (_FMT_NC_VILLARES, "vil"), (_FMT_FRIMAQ, "fri")):
            m = rx.match(ln)
            if not m:
                continue
            g = m.groupdict()
            if kind == "rod":
                items.append(_item_pdf(g["d"], float(g["c"].replace(",", "")), "UN", float(g["p"]), 0.0,
                                       float(g["t"].replace(",", ""))))
            elif kind == "alv":
                items.append(_item_pdf(g["d"], parse_number(g["c"]), "UN", parse_number(g["p"]), parse_number(g["b"]),
                                       parse_number(g["t"])))
            elif kind in ("free", "posta"):
                u = "KG" if g["u"].lower().startswith("k") else ("CAJA" if g["u"] == "CAJ" else "UN")
                items.append(_item_pdf(g["d"], parse_number(g["c"]), u, parse_number(g["p"]), parse_number(g["b"]),
                                       parse_number(g["t"])))
            elif kind == "vil":
                t = float(g["t"].replace(",", ""))
                items.append(_item_pdf(g["d"], 1.0, "UN", t, 0.0, t))
            else:
                t = float(g["t"].replace(",", ""))
                items.append(_item_pdf(g["d"], float(g["c"]), "UN", float(g["p"]), 0.0, t))
            break
    return items


def _numero_comprobante(texto: str) -> str | None:
    pats = [r"Punto de Venta:\s*(\d+)\s+Comp\.\s*Nro:\s*(\d+)",
            r"N[ºo°]:?\s*(\d{4})\s*-\s*(\d{8})", r"FACTURA\s*-?\s*(\d{4})/(\d{8})", r"\b(A\d{5})-(\d{8})\b",
            r"N[úu]mero:\s*(\d{5})-(\d{8})", r"NOTA DE CREDITO\s+(\d{5})-(\d{8})", r"\b(0010)-(\d{8})\b",
            r"\b(0[0-9]{3,4})-(\d{8})\b"]
    for p in pats:
        m = re.search(p, texto, re.I)
        if m:
            return f"{m.group(1)}-{m.group(2)}"
    return None


def _total_comprobante(texto: str) -> float | None:
    pats = [r"Importe Total:\s*\$\s*([\d.,]+)", r"TOTAL COMPROBANTE\.*\s*([\d,]+\.\d{2})", r"TOTAL PES\s*([\d.]+)",
            r"TOTAL Pesos\s*:\s*([\d,]+\.\d{2})", r"^Total:\s*([\d.,]+)\s*$", r"^Total\s+\$\s*([\d.,]+)\s*$",
            r"^TOTAL \$\s*([\d.,]+)", r"Importe Total:\s*\$([\d.,]+)"]
    for p in pats:
        m = re.search(p, texto, re.I | re.M)
        if m:
            raw = m.group(1)
            if re.fullmatch(r"[\d,]+\.\d{2}", raw):
                return float(raw.replace(",", ""))
            return parse_number(raw)
    return None


_PROVEEDORES_CONOCIDOS = (
    (r"RAICES|HENRIQUEZ", "RAICES"), (r"postamayorista|POSTA EXPRESS", "POSTA"), (r"Free Vegetales", "FREE"),
    (r"alimentosvillares|Villares", "VILLARES"), (r"ALVEOLO", "ALVEOLO SA (Levando)"), (r"FRIMAQ", "FRIMAQ S.R.L."),
    (r"RODOFELI", "ROBERTO O. RODOFELI Y CIA S.R.L."), (r"DISTRIBUIDORA KOSTA", "DISTRIBUIDORA KOSTA"),
    (r"CARBALLO", "CARBALLO FERNANDO ARIEL"), (r"SORRENTINO", "SORRENTINO LUCAS JAVIER"),
    (r"VILLARINO", "VILLARINO LORENA ALEJANDRA"), (r"PROCAMP", "PROCAMP (Sola Ogliese Patricio)"),
    (r"Grupo Free|grupofree", "GRUPO FREE"))


def leer_pdf_comprobante(path: str | Path) -> dict:
    """PDF digital de proveedor -> dict de comprobante (solo la copia ORIGINAL)."""
    path = Path(path)
    with pdfplumber.open(path) as pdf:
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)
    t = separar_copias(texto)
    tipo = tipo_comprobante(t)
    f = (re.search(r"(?:Fecha de Emisi[oó]n|Fecha emisi[oó]n|FECHA|Fecha)\s*:?\s*(\d{2}/\d{2}/\d{4})", t)
         or re.search(r"(\d{2}/\d{2}/\d{4})", t))
    items = items_pdf_comprobante(t)
    total = _total_comprobante(t)
    prov = next((nom for rx, nom in _PROVEEDORES_CONOCIDOS if re.search(rx, t, re.I)), None)
    entrega = None
    if re.search(r"AV\. CONGRESO ESQ\. MONTA", t):
        entrega = "AV. CONGRESO ESQ. MONTAÑESES 1602"
    elif re.search(r"Entregar:\s*CONGRESO 1602", t):
        entrega = "CONGRESO 1602 NUÑEZ"
    elif re.search(r"Direcci[oó]n de entrega:\s*Congreso 1602", t):
        entrega = "Congreso 1602 esquina Montañeses"
    else:
        cli = t.split("CUIT: 30714565091")[-1] if "CUIT: 30714565091" in t else t
        me = re.search(r"Domicilio(?: Comercial)?:\s*([^\n]+)", cli)
        entrega = me.group(1).strip() if me else None
    razon = "LAFINUR 3368 SRL" if re.search(r"LAFINUR 3368", t, re.I) else None
    ref = None
    if tipo == "NC":
        lineas = [l.strip() for l in t.splitlines()
                  if re.search(r"rechaza|error de carga|mal\s+estado|bonific|fact\.|FC\s*\d+|Reversi|Referencia", l, re.I)]
        ref = " | ".join(lineas)[:300] or None
    remito = re.search(r"REMITO:\s*(\d{4}/\d{8})", t)
    return {
        "archivo": path.name, "tipo": tipo, "numero_factura": _numero_comprobante(t),
        "fecha_factura": _fecha_iso(f.group(1)) if f else None, "proveedor_detectado": prov,
        "items": items, "total_comprobante": total, "razon_social_cliente": razon, "direccion_entrega": entrega,
        "sello": None, "firmas": [], "referencia": ref, "remito_asociado": remito.group(1) if remito else None,
        "texto_encabezado": t[:600], "motor": "pdf-texto (primera copia)",
        "validacion": validar_aritmetica(items, total) if items else {"ok": False, "motivo": "sin renglones"},
    }


