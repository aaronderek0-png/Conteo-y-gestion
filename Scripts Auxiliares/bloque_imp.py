# --------------------------------------------------------------------------- #
# Lote WhatsApp: retro-entrenamiento, carga agrupada, remitos con birome, NC, prorrateo, Chimidoha
# --------------------------------------------------------------------------- #
REMITOS_JSON = BASE_DIR / "remitos_pendientes.json"
GENERICOS_AGRUPADO = {"verduleria", "varios", "limpieza", "lavadero"}
CARGOS_RE = re.compile(r"servicio\s+log[ií]stico|\bflete\b|\benv[ií]o\b|acarreo|impuestos?\s+internos?", re.I)
PRESENTACION_NC_MP = "Nota de credito Materia Prima"       # SKU LV0218, categoría NC
NC_B_RE = re.compile(r"bonific|descuento|diferencia\s+de\s+precio|dif\.?\s+precio|ajuste|rebaja|error\s+de\s+precio", re.I)
NC_A_RE = re.compile(r"devol|rechaz|mal\s+estado|faltante|error\s+de\s+carga|no\s+fue\s+lo\s+solicitado|reversi|roto|vencid",
                     re.I)


def ultimo_entero(numero: Any) -> int | None:
    p = re.findall(r"\d+", str(numero or ""))
    return int(p[-1]) if p else None


def crear_copia_chimidoha(destino: Path) -> None:
    """Chimidoha (Café Mishiguene Núñez, PB del 1602) no tiene planilla propia: se crea desde la de Congreso con la
    misma estructura (Facturas, Productos, Listas, fórmulas) y la pestaña Facturas vacía (A:I)."""
    import openpyxl
    shutil.copy2(DATA_DIR / PLANILLAS["Congreso"], destino)
    wb = openpyxl.load_workbook(destino)
    ws = wb["Facturas"]
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value not in (None, ""):
            for c in range(1, 10):
                ws.cell(r, c).value = None if c != 9 else 0.0
    wb.save(destino)


# ---- historial -------------------------------------------------------------- #
_LOCALES_CACHE: dict[str, "Local"] = {}


def locales_todos() -> dict[str, "Local"]:
    """Los 4 locales con planilla propia (Chimidoha comparte el catálogo de Congreso)."""
    for n in ("Arcos", "Cabello", "Lafinur", "Congreso"):
        if n not in _LOCALES_CACHE:
            _LOCALES_CACHE[n] = Local(n)
    return {n: _LOCALES_CACHE[n] for n in ("Arcos", "Cabello", "Lafinur", "Congreso")}


def neto_efectivo(comp: dict) -> float:
    s = sum(i["total_linea"] for i in comp.get("items", []))
    return round(s + float((comp.get("extra") or {}).get("imp_internos_total") or 0), 2)


def _subconjunto(valores: list[float], objetivo: float, tol: float = 0.01) -> list[int] | None:
    """Índices de un subconjunto de `valores` cuya suma ≈ objetivo (fuerza bruta hasta 14 filas)."""
    from itertools import combinations
    n = len(valores)
    if n == 0:
        return None
    for k in range(1, min(n, 14) + 1):
        for comb in combinations(range(n), k):
            if abs(sum(valores[i] for i in comb) - objetivo) <= max(1.0, tol * abs(objetivo)):
                return list(comb)
    return None


def buscar_en_historial(comp: dict, locales: dict[str, "Local"]) -> list[dict]:
    """Busca el comprobante por N° (columna A) en las 4 planillas. Exige proveedor parecido o (fecha ±4 días y monto)."""
    n = ultimo_entero(comp.get("numero_factura"))
    if n is None or str(comp.get("numero_factura", "")).upper().startswith("R-"):
        return []
    fecha = pd.Timestamp(comp["fecha_factura"]) if comp.get("fecha_factura") else None
    objetivo = neto_efectivo(comp)
    prov = norm(comp.get("proveedor_detectado"))
    out = []
    for nombre, loc in locales.items():
        f = loc.facturas
        m = f[pd.to_numeric(f["N"], errors="coerce") == n]
        if m.empty:
            continue
        sim = max((fuzz.token_set_ratio(prov, norm(p)) for p in m["PROVEEDOR"].dropna().unique()), default=0) if prov else 0
        fecha_ok = fecha is not None and bool(((m["F_FACTURA"] - fecha).abs() <= pd.Timedelta(days=4)).any())
        totales = (m["CANT"].astype(float) * m["PRECIO"].astype(float) * (1 - m["DESC"].astype(float))).round(2).tolist()
        idx = _subconjunto(totales, objetivo)
        monto_ok = idx is not None
        if not ((sim >= 60 and (fecha_ok or monto_ok)) or (fecha_ok and monto_ok)):
            continue
        n_items = len(comp.get("items", []))
        if idx is None:
            modo, filas_idx = "MONTO_DISTINTO", list(range(len(m)))
        elif len(idx) == 1 and n_items > 1:
            modo, filas_idx = "AGRUPADO_TOTAL", idx
        elif len(idx) >= max(1, n_items - 1):
            modo, filas_idx = "FILAS", idx
        else:
            modo, filas_idx = "FILAS_PARCIAL", idx
        sel = m.iloc[filas_idx]
        out.append({"local": nombre, "modo": modo, "filas": sel, "proveedor_planilla": str(sel["PROVEEDOR"].iloc[0]),
                    "suma_filas": round(sum(totales[i] for i in filas_idx), 2), "objetivo": objetivo,
                    "similitud_proveedor": sim, "fecha_planilla": str(sel["F_FACTURA"].iloc[0].date())
                    if pd.notna(sel["F_FACTURA"].iloc[0]) else None})
    return out


def prorratear_cargos(comp: dict) -> list[dict]:
    """Prorrateo en memoria de impuestos internos y cargos (envío / servicio logístico) dentro del precio del producto.

    Modo exacto de la planilla de Cabello (factura Quilmes 5548-00008853): el neto de cada producto (ya con su
    descuento) más su parte de cargos e impuestos internos se expresa en la col H *bruta* y se conserva el % de
    descuento en la col I:  H = (neto + cargos prorrateados) / cantidad / (1 − desc).  Los cargos quedan como
    renglones de $0 ('envio', 'IMPUESTOS INTERNOS') solo para trazabilidad — no inflan 'materia prima perdida'."""
    items = comp.get("items", [])
    prods = [i for i in items if not CARGOS_RE.search(i["descripcion_original"])]
    cargos = [i for i in items if CARGOS_RE.search(i["descripcion_original"])]
    imp_int = float((comp.get("extra") or {}).get("imp_internos_total") or 0)
    pool = sum(c["total_linea"] for c in cargos) + imp_int
    if not prods or pool == 0:
        return [dict(i, _cargo=bool(CARGOS_RE.search(i["descripcion_original"]))) for i in items]
    base = sum(p["total_linea"] for p in prods)
    out = []
    for p in prods:
        parte = pool * p["total_linea"] / base if base else 0
        neto = p["total_linea"] + parte
        d = p.get("descuento_pct", 0.0) / 100
        h = neto / p["cantidad"] / (1 - d) if d < 1 else p["precio_unitario"]
        out.append(dict(p, precio_unitario=round(h, 2), total_linea=round(neto, 2), _prorrateo=round(parte, 2),
                        _nota=f"cargos e impuestos internos de ${pool:,.2f} prorrateados en col H"))
    out += [dict(c, precio_unitario=0.0, total_linea=0.0, cantidad=1.0, descuento_pct=0.0, _cargo=True,
                 _nota="prorrateado en el precio del producto") for c in cargos]
    if imp_int and not any(re.search(r"impuestos?\s+internos?", c["descripcion_original"], re.I) for c in cargos):
        out.append({"descripcion_original": "IMPUESTOS INTERNOS", "cantidad": 1.0, "unidad_detectada": "UN",
                    "precio_unitario": 0.0, "descuento_pct": 0.0, "total_linea": 0.0, "_cargo": True,
                    "_nota": f"${imp_int:,.2f} prorrateados en la col H del producto"})
    return out


def aprender_de_historial(ap: dict, comp: dict, match: dict) -> dict:
    """Retro-entrenamiento: aprende descripción de proveedor → presentación (+ SKU) o el modo de carga del proveedor."""
    items = prorratear_cargos(comp) if (comp.get("extra") or {}).get("imp_internos_total") else comp.get("items", [])
    filas = match["filas"]
    prov = match["proveedor_planilla"]
    res = {"modo": match["modo"], "aprendidos": [], "regla": None}
    if match["modo"] == "AGRUPADO_TOTAL":
        r = filas.iloc[0]
        modos = ap.setdefault("modos_proveedor", {})
        e = modos.setdefault(norm(prov), {"modo_carga": "AGRUPADO_TOTAL", "proveedor_oficial": prov, "conceptos": {},
                                          "facturas_lote": []})
        e["modo_carga"] = "AGRUPADO_TOTAL"
        e["conceptos"][str(r["PRODUCTO"])] = e["conceptos"].get(str(r["PRODUCTO"]), 0) + 1
        e["concepto"] = str(r["PRODUCTO"])
        e["facturas_lote"] = sorted(set(e["facturas_lote"]) | {str(comp.get("numero_factura"))})
        e["ultima"] = date.today().isoformat()
        res["regla"] = f"{prov}: AGRUPADO_TOTAL → concepto '{r['PRODUCTO']}' (1 fila por factura con el neto total)"
        return res
    if match["modo"] not in ("FILAS", "FILAS_PARCIAL"):
        return res
    libres = set(range(len(filas)))
    pares = []
    for ii, it in enumerate(items):
        if it.get("_cargo"):
            continue
        for fi in range(len(filas)):
            r = filas.iloc[fi]
            tot = float(r["CANT"]) * float(r["PRECIO"]) * (1 - float(r["DESC"]))
            ref = max(abs(it["total_linea"]), 1.0)
            pares.append((abs(tot - it["total_linea"]) / ref, ii, fi))
    usados_i: set[int] = set()
    for dif, ii, fi in sorted(pares):
        if dif > 0.02 or ii in usados_i or fi not in libres:
            continue
        usados_i.add(ii)
        libres.discard(fi)
        pres = str(filas.iloc[fi]["PRODUCTO"])
        aprender(ap, prov, items[ii]["descripcion_original"], pres, match["local"])
        res["aprendidos"].append((items[ii]["descripcion_original"], pres))
    return res


def perfil_historico_proveedores(locales: dict[str, "Local"], minimo: int = 5, umbral: float = 0.8) -> dict:
    """Proveedores cuya historia es mayoritariamente 'una fila por factura con concepto genérico' (Verdulería, Varios…)."""
    todo = pd.concat([l.facturas.assign(LOC=n) for n, l in locales.items()])
    todo["NN"] = pd.to_numeric(todo["N"], errors="coerce")
    g = todo.groupby([todo["PROVEEDOR"].map(norm), "LOC", "NN"]).agg(filas=("PRODUCTO", "size"),
                                                                    prod=("PRODUCTO", "first")).reset_index()
    out = {}
    for prov, s in g.groupby("PROVEEDOR"):
        una = s[s["filas"] == 1]
        if len(s) < minimo or len(una) / len(s) < umbral:
            continue
        top = una["prod"].map(norm).value_counts()
        if top.index[0] in GENERICOS_AGRUPADO and top.iloc[0] / len(una) >= umbral:
            out[prov] = {"facturas": int(len(s)), "una_fila": int(len(una)),
                         "concepto": str(una["prod"][una["prod"].map(norm) == top.index[0]].iloc[0])}
    return out


# ---- remitos con birome ------------------------------------------------------ #
_PESO_RE = re.compile(r"^\s*(?P<p>[A-ZÁÉÍÓÚÑa-záéíóúñ ]+?)\s+(?P<kg>\d+(?:[.,]\d{1,3})?)\s*k\s*g\s*$", re.I)
_NO_RE = re.compile(r"^(?P<p>[^:]+?)\s*:\s*['\"]?\s*NO\b", re.I)


def parse_notas_birome(notas: list[str]) -> tuple[set[str], dict[str, float]]:
    """Anotaciones manuscritas del remito → (productos anulados con 'NO' / tachados, pesos reales en kg)."""
    anulados, pesos = set(), {}
    for n in notas or []:
        m = _NO_RE.match(n)
        if m or re.search(r"tachad|anulad", n, re.I):
            clave = norm((m.group("p") if m else n)).split(" ")[0] if (m or n) else ""
            if clave:
                anulados.add(clave)
            continue
        m = _PESO_RE.match(n)
        if m:
            pesos[norm(m.group("p")).split(" ")[0]] = float(m.group("kg").replace(",", "."))
    return anulados, pesos


def procesar_remito(comp: dict) -> dict:
    """Remito de verdulería → registro de Remitos_Pendientes con renglones anulados ('NO'), kilos reales de balanza
    anotados a mano y costo real efectivo por kilo ingresado."""
    anulados, pesos = parse_notas_birome(comp.get("anotaciones_manuscritas") or [])
    renglones, total_impreso, total_efectivo = [], 0.0, 0.0
    for it in comp.get("items", []):
        clave = norm(it["descripcion_original"]).split(" ")[0]
        impreso = round(it["cantidad"] * it["precio_unitario"], 2)
        anulado = clave in anulados
        real_kg = pesos.get(clave)
        total_impreso += impreso
        cant_real = 0.0 if anulado else it["cantidad"]
        efectivo = 0.0 if anulado else impreso
        total_efectivo += efectivo
        r = {"descripcion": it["descripcion_original"], "unidad": it.get("unidad_detectada"),
             "cantidad_facturada_remito": it["cantidad"], "precio_unitario": it["precio_unitario"],
             "total_remito": impreso, "anulado": anulado, "cantidad_real": cant_real,
             "peso_real_recibido_kg": real_kg, "costo_real_por_kg": None, "reclamo_por_peso": 0.0, "nota": ""}
        if anulado:
            r["nota"] = "renglón anulado a mano con 'NO' (cantidad_real = 0)"
        elif real_kg:
            r["costo_real_por_kg"] = round(impreso / real_kg, 2)
            if (it.get("unidad_detectada") or "").upper() == "KG":
                dif = round(it["cantidad"] - real_kg, 3)
                r["reclamo_por_peso"] = round(max(dif, 0) * it["precio_unitario"], 2)
                r["nota"] = (f"remito factura {it['cantidad']:g} kg y balanza marcó {real_kg:g} kg"
                             + (f": diferencia {dif:g} kg = ${r['reclamo_por_peso']:,.2f}" if dif > 0 else " (a favor)"))
            else:
                r["nota"] = f"{it['cantidad']:g} {it.get('unidad_detectada') or 'u'} pesaron {real_kg:g} kg → ${r['costo_real_por_kg']:,.2f}/kg"
        renglones.append(r)
    return {"remito": comp.get("numero_factura"), "fecha": comp.get("fecha_factura"),
            "proveedor": comp.get("proveedor_detectado"), "archivo": comp.get("archivo"),
            "estado": "PENDIENTE", "factura_definitiva": None, "renglones": renglones,
            "total_impreso": round(total_impreso, 2), "total_efectivo": round(total_efectivo, 2),
            "anulados": sorted(anulados), "firmas": comp.get("firmas"), "alertas": []}


def cargar_remitos() -> dict:
    try:
        return json.loads(REMITOS_JSON.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def guardar_remitos(d: dict) -> None:
    REMITOS_JSON.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


def _clave_remito(r: str | None) -> str:
    p = re.findall(r"\d+", str(r or ""))
    return "-".join(p[-2:]) if len(p) >= 2 else (p[-1] if p else "")


def conciliar_remito_factura(remito: dict, factura: dict) -> dict:
    """Cruza un remito pendiente con la factura definitiva: renglones cobrados que estaban anulados ('NO'), kilos
    facturados de más contra la balanza y diferencia total. Devuelve {'alertas', 'diferencia_total', 'detalle'}."""
    alertas, detalle = [], []
    fact = {norm(i["descripcion_original"]).split(" ")[0]: i for i in factura.get("items", [])}
    for r in remito["renglones"]:
        clave = norm(r["descripcion"]).split(" ")[0]
        f = fact.get(clave)
        if r["anulado"]:
            if f:
                alertas.append(f"RECLAMO: '{r['descripcion']}' estaba anulado con 'NO' en el remito y figura cobrado "
                               f"${f['total_linea']:,.2f}")
            else:
                detalle.append(f"{r['descripcion']}: anulado con 'NO' y NO facturado ✓")
            continue
        if not f:
            alertas.append(f"'{r['descripcion']}' ${r['total_remito']:,.2f} del remito no figura en la factura")
            continue
        kg_real = r["peso_real_recibido_kg"]
        if kg_real and (f.get("unidad_detectada") or "").upper() == "KG" and f["cantidad"] - kg_real > 0.05:
            alertas.append(f"RECLAMO: {r['descripcion']} facturado {f['cantidad']:g} kg vs {kg_real:g} kg recibidos "
                           f"(${(f['cantidad'] - kg_real) * f['precio_unitario']:,.2f})")
        elif kg_real:
            detalle.append(f"{r['descripcion']}: facturado {f['cantidad']:g} {f.get('unidad_detectada') or ''} / recibido "
                           f"{kg_real:g} kg ✓")
        if abs(f["precio_unitario"] - r["precio_unitario"]) > 0.01:
            alertas.append(f"{r['descripcion']}: precio remito ${r['precio_unitario']:,.2f} ≠ factura ${f['precio_unitario']:,.2f}")
    neto_f = neto_efectivo(factura)
    return {"alertas": alertas, "detalle": detalle, "total_remito_efectivo": remito["total_efectivo"],
            "total_factura": neto_f, "diferencia_total": round(remito["total_efectivo"] - neto_f, 2)}


# ---- Notas de crédito -------------------------------------------------------- #
def clasificar_nc(comp: dict, vigente_ratio: float | None = None) -> tuple[str, str]:
    """'A' devolución física / faltante (producto específico, cantidad negativa, resta stock) o 'B' diferencia de
    precio / bonificación comercial (sin movimiento físico → 'Nota de credito Materia Prima')."""
    txt = " ".join([str(comp.get("referencia") or ""), " ".join(i["descripcion_original"] for i in comp.get("items", []))])
    if NC_B_RE.search(txt) and not NC_A_RE.search(txt):
        return "B", "bonificación / diferencia de precio"
    if NC_A_RE.search(txt):
        return "A", "devolución física / faltante / error de carga"
    if vigente_ratio is not None and vigente_ratio < 0.5:
        return "B", "importe parcial frente al precio vigente (diferencia de precio)"
    return "A", "sin motivo escrito: se trata como devolución del producto"


def filas_nc(comp: dict, loc: "Local", ap: dict, proveedor: str | None) -> list[dict]:
    items = comp.get("items", [])
    ratio = None
    if len(items) == 1:
        m = mapear_producto(items[0]["descripcion_original"], proveedor, loc, ap)
        if m["presentacion"]:
            vig, _ = loc.costo_vigente(m["presentacion"])
            cat = loc.catalogo[loc.catalogo["PRESENTACION"] == m["presentacion"]].iloc[0]
            env = (_num(cat["UN_PRES"]) or 1) * (_num(cat["CANT_PRES"]) or 1)
            if vig and env:
                ratio = items[0]["precio_unitario"] / (vig * env)
    caso, motivo = clasificar_nc(comp, ratio)
    base = {"A": comp.get("numero_factura"), "B": comp.get("fecha_factura"), "C": comp.get("fecha_factura"),
            "D": proveedor or (comp.get("proveedor_detectado") or "").upper(), "proveedor_oficial": proveedor}
    filas = []
    if caso == "B":
        total = round(sum(i["total_linea"] for i in items), 2)
        cat = loc.catalogo[loc.catalogo["PRESENTACION"] == PRESENTACION_NC_MP]
        f = dict(base, descripcion_original=f"NC {comp.get('numero_factura')} ({motivo})", presentacion=PRESENTACION_NC_MP,
                 sku=cat.iloc[0]["SKU"] if len(cat) else "LV0218", confianza=1.0, metodo="NC caso B (dif. precio/bonificación)",
                 G=-1.0, H=total, I=0.0, J=-total, P_UN=-total, vigente=None, var_pct=None, accion="INYECTAR",
                 motivos=[], avisos=[f"NC caso B: {motivo}; no resta kilos del inventario físico"], solicitud=False,
                 cant_orig=-1.0, precio_orig=total, alternativas=[], tipo_nc="B")
        filas.append(f)
        return filas
    for it in items:
        d = decidir_renglon(dict(it, cantidad=it["cantidad"]), comp, proveedor, loc, ap)
        d.update(base)
        if d["G"] is not None:
            d["G"], d["J"] = -abs(d["G"]), -abs(d["J"] or 0)
            d["P_UN"] = d["P_UN"]
        d["cant_orig"] = -abs(it["cantidad"])
        d["avisos"].append(f"NC caso A: {motivo}; cantidad negativa resta stock y gasto")
        d["tipo_nc"] = "A"
        # el precio vigente se compara contra el precio del renglón (positivo), el signo va en la cantidad
        filas.append(d)
    return filas


# ---- carga agrupada, prorrateo y orquestación por comprobante ---------------- #
def fila_simple(**k) -> dict:
    base = {"descripcion_original": "", "presentacion": None, "sku": None, "confianza": 0.0, "metodo": "",
            "G": None, "H": None, "I": 0.0, "J": None, "P_UN": None, "vigente": None, "var_pct": None,
            "accion": "REVISAR", "motivos": [], "avisos": [], "solicitud": False, "cant_orig": None,
            "precio_orig": None, "alternativas": [], "proveedor_oficial": None}
    base.update(k)
    return base


def construir_filas(comp: dict, loc: "Local", ap: dict) -> dict:
    """Decide cómo se imputa un comprobante. Devuelve {'proveedor_oficial', 'modo', 'filas', 'notas'}."""
    proveedor, conf_p = mapear_proveedor(comp.get("proveedor_detectado"), comp.get("texto_encabezado", ""), loc.proveedores)
    notas: list[str] = []
    fecha_f = comp.get("fecha_factura")
    meta = {"A": numero_a_celda(comp.get("numero_factura")), "B": fecha_f, "C": comp.get("fecha_ingreso") or fecha_f,
            "D": proveedor or (comp.get("proveedor_detectado") or "").upper(), "proveedor_oficial": proveedor}
    modos = ap.get("modos_proveedor", {})
    modo_prov = modos.get(norm(proveedor)) or modos.get(norm(comp.get("proveedor_detectado"))) if proveedor else None
    if comp.get("tipo") == "NC":
        return {"proveedor_oficial": proveedor, "modo": "NC", "filas": filas_nc(comp, loc, ap, proveedor), "notas": notas}
    if modo_prov and modo_prov.get("modo_carga") == "AGRUPADO_TOTAL" and len(comp.get("items", [])) >= 1:
        total = neto_efectivo(comp)
        cat = loc.catalogo[loc.catalogo["_norm"] == norm(modo_prov["concepto"])]
        f = fila_simple(**meta, descripcion_original=f"{len(comp['items'])} renglones agrupados",
                        presentacion=modo_prov["concepto"], sku=cat.iloc[0]["SKU"] if len(cat) else None,
                        confianza=0.99, metodo="AGRUPADO_TOTAL (regla aprendida)", G=1.0, H=total, I=0.0, J=total,
                        P_UN=total, accion="INYECTAR" if len(cat) else "REVISAR", cant_orig=1.0, precio_orig=total,
                        avisos=[f"proveedor cargado como 1 fila '{modo_prov['concepto']}' por el neto total"])
        v = comp.get("validacion") or {}
        if v.get("ok") is False:
            f["accion"], f["motivos"] = "REVISAR", [f"comprobante no cierra: {v.get('motivo')}"]
        return {"proveedor_oficial": proveedor, "modo": "AGRUPADO_TOTAL", "filas": [f], "notas": notas}
    items = prorratear_cargos(comp) if ((comp.get("extra") or {}).get("imp_internos_total")
                                        or any(CARGOS_RE.search(i["descripcion_original"]) for i in comp.get("items", []))
                                        and len(comp.get("items", [])) > 1) else comp.get("items", [])
    filas = []
    for it in items:
        if it.get("_cargo"):
            nombre = "IMPUESTOS INTERNOS" if re.search(r"impuestos?\s+internos?", it["descripcion_original"], re.I) else "Envio"
            cat = loc.catalogo[loc.catalogo["_norm"] == norm(nombre)]
            filas.append(fila_simple(**meta, descripcion_original=it["descripcion_original"], presentacion=nombre,
                                     sku=cat.iloc[0]["SKU"] if len(cat) else None, confianza=1.0, metodo="cargo prorrateado",
                                     G=1.0, H=0.0, J=0.0, P_UN=0.0, accion="INYECTAR", cant_orig=1.0, precio_orig=0.0,
                                     avisos=[it.get("_nota", "")]))
            continue
        d = decidir_renglon(it, comp, proveedor, loc, ap)
        d.update(meta)
        if it.get("_prorrateo"):
            d["avisos"].append(it["_nota"])
        filas.append(d)
    v = comp.get("validacion") or {}
    if v.get("ok") is False:
        for f in filas:
            if f["accion"] == "INYECTAR":
                f["accion"] = "REVISAR"
            f["motivos"].append(f"comprobante no cierra: {v.get('motivo')}")
    return {"proveedor_oficial": proveedor, "modo": "PRORRATEO" if items is not comp.get("items") else "FILAS",
            "filas": filas, "notas": notas}


# ---- publicación en la planilla real (solo con confirmación) ----------------- #
IDS_PLANILLAS = {"Congreso": "1udLEQpQ29avK4c6fNnQig-w02xIV82IgLe85yVJXaOY",
                 "Arcos": "1frh93qB8Od5mrtAfMuObKwRW93p7kev6DNgmERqz6ho",
                 "Lafinur": "1tv0uI4KsSEPRUecYfnwcDqpGbUDelGqWwe3Hp-fikOc",
                 "Cabello": "1zAxYyNB98Sb1-D0HtfokESXc6FrDS96EbFYyMmsseIE"}


def push_drive(local: str, filas: list[dict], confirmado: bool = False) -> dict:
    """Agrega las filas A:I a la pestaña Facturas de la planilla REAL del local (Google Sheets, credentials.json).

    NUNCA se llama sola: requiere `--push-drive` Y `--confirmo-escribir-produccion`. Sin confirmación devuelve el
    plan exacto (qué fila y qué valores) sin escribir nada. Chimidoha no tiene planilla real todavía."""
    if local not in IDS_PLANILLAS:
        return {"escrito": False, "motivo": f"{local} no tiene planilla real de Facturas; usar la copia local"}
    inyectables = [f for f in filas if f.get("accion") == "INYECTAR"]
    plan = [[f["A"], f["B"], f["C"], f["D"], f["presentacion"], "", f["G"], f["H"], f["I"]] for f in inyectables]
    if not confirmado:
        return {"escrito": False, "motivo": "falta --confirmo-escribir-produccion", "plan": plan,
                "planilla": IDS_PLANILLAS[local]}
    import gspread
    from google.oauth2.service_account import Credentials
    creds = Credentials.from_service_account_file(
        str(BASE_DIR / "credentials.json"), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    ws = gspread.authorize(creds).open_by_key(IDS_PLANILLAS[local]).worksheet("Facturas")
    col_a = ws.col_values(1)
    primera = len(col_a) + 1                      # primera fila libre después de la última con N°
    fmt = [[str(v) if isinstance(v, datetime) else v for v in r] for r in plan]
    for r in fmt:
        for k in (1, 2):
            if isinstance(r[k], str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", r[k]):
                y, m, d = r[k].split("-")
                r[k] = f"{d}/{m}/{y}"
    ws.update(f"A{primera}:I{primera + len(fmt) - 1}", fmt, value_input_option="USER_ENTERED")
    return {"escrito": True, "desde_fila": primera, "filas": len(fmt)}


