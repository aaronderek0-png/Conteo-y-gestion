# Lecturas visuales (Claude, multimodal) de las fotos de Facturas_WPP/. Una entrada por archivo.
def it(d, q, u, p, desc=0.0, tot=None):
    t = round(q * p * (1 - desc / 100), 2) if tot is None else tot
    return {"descripcion_original": d, "cantidad": q, "unidad_detectada": u, "precio_unitario": p,
            "descuento_pct": desc, "total_linea": t}

W = "WhatsApp Image "
L = []

def add(archivo, tipo, num, fecha, prov, items, total, sello, direccion, razon, firmas=None, notas=None,
        ref=None, cuit_cliente="30-71456509-1", remito=None, extra=None):
    L.append(dict(archivo=archivo, tipo=tipo, numero_factura=num, fecha_factura=fecha, proveedor_detectado=prov,
                  items=items, total_comprobante=total, sello=sello, direccion_entrega=direccion,
                  razon_social_cliente=razon, firmas=firmas or [], anotaciones_manuscritas=notas or [],
                  referencia=ref, remito_asociado=remito, extra=extra or {}))

CERV = "MISHIGUENE CERVIÑO - RECIBIDO"
add(W+"2026-08-20 at 09.23.11.jpeg", "Factura", "00002-00001393", "2026-08-18", "DISTRIBUIDORA KOSTA (Lukides Cleopatra Catalina)",
    [it("Caja 3 pouch hojas de Parra x 3 kg LA EMILIA", 1, "UN", 57000.0)], 68970.0, CERV, "Lafinur 3368 (domicilio fiscal)",
    "LAFINUR 3368 SRL", ["Aaron Gaitán", "(otra firma ilegible)"])
add(W+"2026-08-20 at 09.24.06.jpeg", "Factura", "00006-00024099", "2026-08-18", "LACTEOS LA DELFINA S.A.",
    [it("Queso Crema Tipico x 4 Kg", 4, "UN", 27469.0), it("Gastos de Envio", 1, "UN", 9900.0)], 144928.96, CERV,
    "Lafinur esquina Av Libertador 3368", "Lafinur 3368 SRL", ["Aaron Gaitán", "(otra firma ilegible)"])
add(W+"2026-08-22 at 11.24.33.jpeg", "NC", "A00010-00091309", "2026-08-19", "ALYSER S.A.",
    [it("MANTECA CORLASA X 25KG", 1, "CAJA", 18529.42)], 22420.6, CERV, "Av Congreso 1602, Belgrano",
    "LAFINUR 3368 SRL (CERVIÑO)", ["Aaron Gaitán"], ref="sin referencia escrita; la factura A00010-00924396 (06/08) fue por $251.470,70")
add(W+"2026-08-22 at 14.20.53.jpeg", "Factura", "01030-00281610", "2026-08-18", "VILLARES S.A.C.",
    [it("GARBANZO SALTEÑO X 25 KG. SAUCO COMUN", 25, "KG", 1425.60, tot=35640.0),
     it("HINOJO X 1 KG.", 2, "UN", 5310.36, tot=10620.72),
     it("PIMIENTA CAYENA MOLIDA X 1 KG.", 2, "UN", 8942.40, tot=17884.80),
     it("TAHINA AL KANATER X 3.6 KG. S/GLUTEN", 5, "CAJA", 38234.59, tot=191172.96)],
    None, CERV, "Lafinur 3368", "LAFINUR 3368 S.R.L.", ["Aaron Gaitán", "20/08"],
    extra={"bonificacion_pct": 28.0, "neto_105": 35640.0, "neto_21": 219678.48, "iva_21": 46132.48,
           "nota": "Total no visible en la foto; bonificación 28 % ya aplicada en Pr. Neto"})
add(W+"2026-08-25 at 10.19.33.jpeg", "Remito", "0029-00007411", "2026-08-18", "FAMILIA ZUCCARDI (La Agrícola S.A.)",
    [it("FZ GENOVESA 12X250 NUEVO ARTE CAJ 12X250 ML", 1, "CAJA", 0.0, tot=0.0),
     it("ZUELO ORIGINAL BIB 2X5000 CAJ 2X5000 ML", 5, "CAJA", 0.0, tot=0.0)],
    None, None, "LAFINUR 3368 SRL - CONGRESO 1602 (1428) BELGRANO", "LAFINUR 3368 SRL",
    notas=["DUPLICADO sin sello ni firma del receptor"], ref="Ref. S0-353377; facturado en 0039-00143503",
    extra={"valor_declarado": 713636.36})
add(W+"2026-08-25 at 17.00.14.jpeg", "Factura", "0039-00143503", "2026-08-18", "FAMILIA ZUCCARDI (La Agrícola S.A.)",
    [it("FZ GENOVESA 12X250 NUEVO ARTE CAJ 12X250 ML", 1, "CAJA", 138842.98, desc=100.0, tot=0.0),
     it("ZUELO ORIGINAL BIB 2X5000 CAJ 2X5000 ML", 5, "CAJA", 259504.13, desc=45.0, tot=713636.36)],
    920590.91, "CERVIÑO RECIBIDO (manuscrito, encuadrado)", "LAFINUR 3368 SRL - CONGRESO 1602 (1428) BELGRANO",
    "LAFINUR 3368 SRL", ["A. Hatrick"], remito="0029-00007411",
    extra={"perc_iva": 21409.09, "perc_iibb": 35681.82, "iva": 149863.64})
add(W+"2026-08-26 at 09.03.19.jpeg", "Factura", "0004-00000106", "2026-08-25", "EMPAPELAR (Daich Marcos Roberto)",
    [it("FOLEX 25X35 X KILO", 1, "UN", 8350.0), it("P.B CONSORCIO 80X100X30 NEGRA BULTO X 50UDS", 1, "UN", 9800.0),
     it("GUANTES NITRILO NEGRO BULTO X 100UDS", 1, "UN", 8400.0), it("BOBINA 20 X 400 X 2 SEDAPEL", 1, "UN", 23900.0),
     it("ALUMINIO X KILO", 1, "UN", 8900.0), it("PAPEL FILM 45X1000", 1, "UN", 48000.0), it("ETIQUETA 44X55", 2, "UN", 3000.0)],
    137153.5, CERV, "Congreso 1602 - 1406", "LAFINUR 3368 SRL", ["Aaron Gaitán", "(otra firma ilegible)"])
add(W+"2026-08-26 at 17.02.54.jpeg", "Factura", "0002-00034822", "2026-08-26", "RUPAJO S.A.",
    [it("PAN ARABE CLASICO BOLSA CRISTAL X300G", 42, "UN", 650.0)], 33852.0, "MISHIGUENE ARCOS - RECIBIDO (+ 'Arcos' manuscrito)",
    "Lafinur 3368", "LAFINUR 3368 SRL", ["Alexis? (26/8)"], extra={"iva": 5733.0, "otros": 819.0})
add(W+"2026-08-27 at 11.30.48.jpeg", "NC", "0015-00000128", "2026-08-26", "CHEF SUPPLY S.A. (Master Supply)",
    [it("MBF8113 Freezer Vertical Atosa EU MBF8113", 1, "UN", 2887782.81)], 3191000.0, CERV, "Cerviño 4417, Palermo",
    "LAFINUR 3368 SRL", ["Aaron Gaitán"], extra={"iva_105": 303217.19})
add(W+"2026-08-28 at 15.32.29.jpeg", "Factura", "0024-00097543", "2026-08-28", "RESTA SANITARIOS (Sucesores de Domingo Resta y Cía S.A.)",
    [it("TERMOTANQUE SAIAR MULTIGAS DE PIE 120 LT", 1, "UN", 731138.36, desc=30.4, tot=508872.3)], 615735.48, CERV,
    "Lafinur 3368 (Abasto)", "LAFINUR 3368 SRL", ["Aaron Gaitán"],
    extra={"descuento_financiacion": -127218.08, "iva_21": 106863.18})
add(W+"2026-08-31 at 12.31.03.jpeg", "Factura", "A00010-00924396", "2026-08-06", "ALYSER S.A.",
    [it("MANTECA CORLASA X 25KG", 1, "CAJA", 251470.70)], 304279.55, CERV, "Ent: Av Congreso 1602, Belgrano",
    "LAFINUR 3368 SRL (CERVIÑO)", ["Aaron Gaitán", "Conforme cliente: (ilegible)"])
add(W+"2026-09-02 at 10.10.53.jpeg", "Factura", "0011-00003118", "2026-08-31", "ROBERTO O. RODOFELI Y CIA S.R.L.",
    [it("BOLSAS VACIO 20X30 / MP", 1000, "UN", 154.0)], 186340.0, CERV, "Lafinur 3368", "LAFINUR 3368 SRL",
    ["Aaron Gaitán"], extra={"duplicado_de_pdf": "FAE001100003118-LAFINUR 3368 SR_260831_130415.pdf"})
add(W+"2026-09-02 at 13.49.26.jpeg", "Factura", "0006-00098160", "2026-09-02", "JMS PACKAGING SA",
    [it("Pote Aderezos 60 cc BAMBOO Natural (x 50 unidades)", 4, "UN", 891.25),
     it("Tapa Cristal Pote NEGRO Darnel 45-60-75 cc (x 100 unidades)", 2, "UN", 3472.0)],
    12715.89, CERV, "CABA", "Lafinur 3368 SRL", ["Aaron Gaitán"])
add(W+"2026-09-07 at 11.12.39.jpeg", "Factura", "00002-00000572", "2026-09-07", "MIÑO GASTON ALEJANDRO",
    [it("beguel", 120, "UN", 1155.0), it("pletsalej", 96, "UN", 2650.0), it("pita", 160, "UN", 450.0),
     it("lafa", 100, "UN", 550.0), it("beguel amapola", 6, "UN", 2000.0), it("beguel cebolla", 6, "UN", 2000.0)],
    658240.0, CERV + " (sello tenue)", "Lafinur 3368", "LAFINUR 3368 SRL", ["Aaron Gaitán", "David Gobbo(?)"],
    notas=["tildes de control en cada renglón de cantidad"])
add(W+"2026-09-29 at 10.51.07.jpeg", "Factura", "00003-00025412", "2026-09-28", "QUESOS JUAN GRANDE S.A.",
    [it("Queso Feta Horma (3 cajas)", 9.25, "KG", 20414.0)], 228483.7, CERV, "Congreso 1602 entrada x Montañeses",
    "LAFINUR 3368 SRL - Nombre fantasía: Obrador Mishiguene", ["Aaron Gaitán"], extra={"iva_21": 39654.2, "neto": 188829.5})
add(W+"2026-09-29 at 10.55.45.jpeg", "Factura", "00002-00000903", "2026-09-16", "FERREYRA JUAN CARLOS",
    [it("Correspondiente a 50% de anticipo sobre nuestra oferta por cortinas roller sun screen", 1, "UN", 757603.31)],
    916700.01, CERV, "Lafinur 3368", "LAFINUR 3368 SRL", ["Aaron Gaitán"], notas=["Abonar 50% faltante"])
add(W+"2026-09-30 at 12.15.24.jpeg", "Remito", "R-0001/00009194", "2026-09-27", "RAICES (Juan Sebastian Henriquez)",
    [it("LIMON CAJON", 1, "CAJON", 16200.0), it("BERENJENA TORITO", 3, "UN", 19200.0),
     it("JALAPEÑO KILOS", 4, "KG", 8200.0), it("MORRON ROJO KILOS", 3, "KG", 6910.0),
     it("ZANAHORIA INDUS KILOS", 5, "KG", 1700.0), it("PUERRO UNIDAD", 2, "UN", 3800.0),
     it("APIO UNIDAD", 2, "UN", 1900.0), it("CEBOLLON BOLSAS", 1, "BOLSA", 42200.0),
     it("PEREJIL PAQUETE", 1, "PAQ", 1900.0), it("CILANTRO PAQUETE", 5, "PAQ", 1250.0),
     it("MENTA PAQUETE", 2, "PAQ", 3750.0)],
    None, None, "AV. CONGRESO ESQ. MONTAÑESES 1602 (Núñez)", "LAFINUR 3368 S.R.L.", ["David Gobbo(?)"],
    notas=["BERENJENA TORITO: 'NO' escrito en birome con flecha (anulada)", "APIO 1,090 KG", "JALAPEÑO 3,645 KG",
           "ZANAHORIA 5,200 KG", "PUERRO 2,600 KG", "tildes de control en cada renglón"],
    extra={"codigos": {"LIMON": 536, "BERENJENA": 202, "JALAPEÑO": 998, "MORRON ROJO": "1-13", "ZANAHORIA": "4-2",
                       "PUERRO": 186, "APIO": 191, "CEBOLLON": 51, "PEREJIL": 294, "CILANTRO": 490, "MENTA": 545}})
add(W+"2026-09-30 at 12.46.21.jpeg", "Factura", "0001/00063554", "2026-09-30", "RAICES (Juan Sebastian Henriquez)",
    [it("LIMON CAJON", 1, "CAJON", 16200.0), it("JALAPEÑO KILOS", 3.6, "KG", 8200.0, tot=29520.0),
     it("MORRON ROJO KILOS", 3, "KG", 6910.0), it("ZANAHORIA INDUS KILOS", 5, "KG", 1700.0),
     it("PUERRO UNIDAD", 2, "UN", 3800.0), it("APIO UNIDAD", 2, "UN", 1900.0), it("CEBOLLON BOLSAS", 1, "BOLSA", 42200.0),
     it("PEREJIL PAQUETE", 1, "PAQ", 1900.0), it("CILANTRO PAQUETE", 5, "PAQ", 1250.0), it("MENTA PAQUETE", 2, "PAQ", 3750.0)],
    159341.0, CERV, "AV. CONGRESO ESQ. MONTAÑESES 1602", "LAFINUR 3368 S.R.L.", ["Aaron Gaitán"],
    remito="0001/00009194",
    notas=["+ ingreso 2,6 kg ... el kg → Precio por kg (puerro)", "- ingreso 1,09 kg ... el kg (apio)",
           "cifras por kg ilegibles", "renglones PUERRO y APIO encuadrados a mano"],
    extra={"neto": 144200.0, "iva_105": 15141.0, "entregado": "2026-09-28"})
add(W+"2026-09-30 at 13.27.13.jpeg", "Factura", None, None, "CIDEL ORGANICA SRL",
    [it("acelga roja", 1, "UN", 2999.0), it("acelga amarilla", 1, "UN", 2999.0), it("esparrago grueso", 1, "UN", 6999.0),
     it("pera roja", 2, "UN", 3199.0), it("mini acedera sang", 1, "UN", 5999.0), it("mini armuelle rojo", 1, "UN", 5999.0),
     it("MINI MIZUNA", 1, "UN", 5999.0), it("MINI espinaca tetragonia", 1, "UN", 5999.0), it("BROTE DE ARVEJA", 2, "UN", 7999.0),
     it("FLOR DE CLAVELINA", 1, "UN", 7499.0), it("FLOR DE BORRAJA", 1, "UN", 7499.0), it("FLOR DE NABO", 1, "UN", 7499.0),
     it("PENSAMIENTOS", 1, "UN", 7499.0)], None, None, "Lafinur 3368 (domicilio fiscal)", "LAFINUR 3368 SRL",
    ["firma ilegible"], notas=["foto recortada: sin N°, fecha ni totales; tilde azul en cada renglón"])
add(W+"2026-10-01 at 10.40.23.jpeg", "Factura", "5548-00008853", "2026-09-22", "CERVECERIA Y MALTERIA QUILMES S.A.I.C.A. y G.",
    [it("STELLA ARTOIS PM SLIM KEG RET B", 1, "BARR", 115123.43, desc=30.0, tot=80586.40),
     it("SERVICIO LOGISTICO", 1, "UNID", 3900.0, tot=3900.0)], 123428.73,
    "CAFE MISHIGUENE CABELLO (sello con firma de responsable)", "CABELLO 3181 (nombre fantasía CAFE MISHIGUENE)",
    "LAFINUR 3368 SRL", ["responsable (firma)"],
    extra={"imp_internos": [13118.72, 634.88], "imp_internos_total": 13753.60, "iva": 17742.14, "perc_iibb": 4912.0,
           "perc_iva": 2534.59, "subtotal_neto": 84486.40, "subtotal2": 115982.14, "factura_asociada": "5548A 88852"})
add(W+"2026-10-01 at 10.40.48.jpeg", "Captura_planilla", "8853", "2026-09-22", "QUILMES", [], None, None, None, None,
    notas=["Cabello fila 835: 'Stella Artois (cerveza tirada)' cant 1, H=$140.343, desc 30 %",
           "fila 836: 'envio' cant 1, $0", "fila 837: 'IMPUESTOS INTERNOS'",
           "comentario: 'Impuesto Interno de $13.753,6 distribuido por prorrateo general entre 1 producto(s) en Columna H'"])
add(W+"2026-10-01 at 11.13.56.jpeg", "Factura", "00009-00026358", "2026-09-28", "NOGADA S.A.",
    [it("Aceite Girasol Alto Oleico Areco Caj. x 4 Un x 5 lts Pet", 4, "CAJA", 66648.07, desc=2.0, tot=261260.43)],
    323962.93, "(sin sello; razón social CHIMI DOHA, 'Café Mishiguene Núñez')", "CONGRESO 1602 (1428) Belgrano",
    "CHIMI DOHA - CUIT 30-71567213-4", notas=["'DEVO ... ARCOS' manuscrito arriba (posible devolución/derivación a Arcos)"],
    cuit_cliente="30-71567213-4", extra={"perc_iva": 7837.81, "iva_21": 54864.69})
add(W+"2026-10-01 at 11.17.04.jpeg", "Factura", "0002-00035451", "2026-09-25", "RUPAJO S.A.",
    [it("PAN ARABE CLASICO BOLSA CRISTAL X300G", 36, "UN", 650.0)], 29016.0, "(sin sello; 'CHIMI' manuscrito)",
    "Lafinur 3368", "LAFINUR 3368 SRL", ["David Gobbo(?)"], notas=["CHIMI"], extra={"iva": 4914.0, "perc_3": 702.0})
add(W+"2026-10-01 at 15.41.21.jpeg", "Factura", "0001/00063617", "2026-10-01", "RAICES (Juan Sebastian Henriquez)",
    [it("COLIFLOR UNIDAD", 4, "UN", 2500.0), it("PEPINO CAJON", 3, "CAJON", 20200.0), it("RABANITO KILOS", 3, "KG", 4065.0)],
    91488.48, None, "AV. CONGRESO ESQ. MONTAÑESES 1602", "LAFINUR 3368 S.R.L.", [], remito="0001/00009232",
    notas=["foto recortada: sin sello visible; total leído parcial"], extra={"neto": 82795.0, "entregado": "2026-10-01"})
add(W+"2026-10-02 at 16.22.50.jpeg", "Factura", "00003-00004462", "2026-09-30", "AGUILERA LUIS ALBERTO",
    [it("servilletas", 54, "UN", 346.85), it("repasador", 11, "UN", 619.41), it("cristal", 6, "UN", 435.31)],
    34067.88, CERV, "Lafinur 3368 (domicilio fiscal)", "LAFINUR 3368 SRL", ["Aaron Gaitán", "(otra firma)"],
    extra={"periodo": "16/09/2026-30/09/2026", "neto": 28155.27, "iva": 5912.61})
add(W+"2026-10-02 at 19.58.18.jpeg", "Factura", "A00011-00001488", "2026-08-25", "MEGAENVASES SA",
    [it("Frasco Morronero 200ml Conserv ades Dorado", 2, "UN", 40727.27)], 98560.0,
    "MISHIGUENE ARCOS - RECIBIDO", "Lafinur 3368", "LAFINUR 3368 SRL", ["A. Hatrick"],
    extra={"iva_21": 17105.45})
add(W+"2026-10-05 at 09.12.56.jpeg", "Factura", "00001-00000058", "2026-10-05", "SORRENTINO LUCAS JAVIER",
    [it("Lunes 28/9 (O;A;L;C)", 2, "UN", 25000.0), it("Martes 29/9 (O;A;L;C)", 2, "UN", 25000.0),
     it("Miercoles 30/9 (O;A;L;C)", 2, "UN", 25000.0), it("Jueves 01/10 (DARWING 389;O;A;L;C)", 4.02, "UN", 25000.0, tot=100500.0),
     it("Viernes 2/10 (O;A;L;C)", 2.5, "UN", 25000.0), it("Sabado 3/10 (L;O;A;L;C)", 2, "UN", 25000.0),
     it("Domingo 4/10 (O;A;L;C)", 3, "UN", 25000.0)], 438000.0, "'RECIBIDO CERVIÑO' manuscrito encuadrado",
    "Lafinur 3368", "LAFINUR 3368 SRL", ["A. Hatrick"],
    extra={"multilocal": "O=Obrador/Congreso, A=Arcos, L=Lafinur, C=Cabello (envíos por día)"})
add(W+"2026-10-05 at 09.45.37.jpeg", "Factura", "00002-00001440", "2026-10-03", "DISTRIBUIDORA KOSTA (Lukides Cleopatra Catalina)",
    [it("Caja Hojas de Parra 3 x 3 kg LA EMILIA", 1, "UN", 57000.0), it("Trigo grueso 10 x 500 g ARARAT", 1, "UN", 37476.0)],
    114315.96, "'CERVIÑO RECIBIDO' manuscrito encuadrado", "Congreso Av. 1602", "CHIMI DOHA SRL (CUIT 30-71567213-4)",
    ["A. Hatrick"], cuit_cliente="30-71567213-4", extra={"iva_21": 19839.96})
add(W+"2026-10-06 at 17.01.45.jpeg", "Factura", "00006-00025341", "2026-10-05", "LACTEOS LA DELFINA S.A.",
    [it("Queso Crema TP x 4 Kg", 6, "UN", 28843.0)], 173058.0, CERV, "Lafinur esquina Av Libertador 3368",
    "Lafinur 3368 srl", ["Aaron Gaitán", "David Gobbo(?)"])
add(W+"2026-10-07 at 11.28.40.jpeg", "Factura", "00006-00025148", "2026-09-28", "LACTEOS LA DELFINA S.A.",
    [it("Queso Crema TP x 4 Kg", 5, "UN", 28843.0), it("Gastos de Envio", 1, "UN", 10395.0)], 187078.10, CERV,
    "Lafinur esquina Av Libertador 3368", "Lafinur 3368 srl", ["Aaron Gaitán", "David Gobbo(?)"])
add(W+"2026-10-12 at 09.44.01.jpeg", "Factura", "00004-00001610", "2026-10-06", "ECOBAIRES AMBIENTAL (Soiza Reilly Juan Martín)",
    [it("DESINSECTACION SEPTIEMBRE 2026", 1, "UN", 95000.0)], 114950.0, CERV, "Congreso Av. 1602",
    "CHIMI DOHA SRL (CUIT 30-71567213-4)", ["Aaron Gaitán"], cuit_cliente="30-71567213-4", extra={"iva_21": 19950.0})
add(W+"2026-10-12 at 10.08.15.jpeg", "Factura", "0001-00008101", "2026-10-12", "(proveedor de huevos; encabezado cortado - Darwin 1745, Ezeiza)",
    [it("HUEVOS X CAJON (12 MAPLES)", 0.5, "CAJON", 89000.0)], 49172.5,
    "CAFE MISHIGUENE NUÑEZ - RECIBIDO", "Lafinur 3368, Palermo", "LAFINUR 3368 SRL",
    ["David Gobbo(?) 12/10/26", "Macarena Venego(?) 12/10"], extra={"nota_de_venta": "0002-00007892", "iva_105": 4672.5})
add(W+"2026-10-12 at 10.09.39.jpeg", "Factura", "0007-00014643", "2026-10-12", "ALVEOLO SA (Levando - despacho de pan)",
    [it("CROISSANT (PS00001)", 50, "UN", 1990.95)], 110000.0, None, "Lafinur 3368", "LAFINUR 3368 SRL",
    ["David Gobbo(?) 12/10/26"], notas=["sin sello"], extra={"iva_105": 10452.49, "neto": 99547.51})
add(W+"2026-10-12 at 10.19.02.jpeg", "Factura", "0001-00036428", "2026-10-05", "MIX-57 S.A.S.",
    [it("ETIQUETA IL VIP 50X50 (1000U) B25 1B CIRC IMP LOGO CELESTE (P. 284C)", 5000, "UN", 19.83, tot=99173.56),
     it("ENVIO ENVIAMELO SSP (PP)", 1, "UN", 8264.46)], 130000.0, CERV, "Montañeses 2794 (Belgrano)", "Lafinur 3368 SRL",
    ["Aaron Gaitán"], notas=["'Abonar el ser…' tachado"], extra={"iva_21": 22561.98})
