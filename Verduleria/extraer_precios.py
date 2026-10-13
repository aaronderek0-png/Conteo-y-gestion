import os
import re
import pandas as pd
import pdfplumber

print("==================================================")
print("⚙️  EXTRAYENDO AMBAS COLUMNAS (IZQUIERDA Y DERECHA)")
print("==================================================\n")

def limpiar_numero(val):
    if not val:
        return 0.0
    s = re.sub(r'[^\d,.]', '', str(val))
    if not s:
        return 0.0
    if ',' in s and '.' in s:
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        s = s.replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return 0.0

def es_encabezado_o_vacio(celdas):
    texto_linea = " ".join(celdas).lower()
    if not any(celdas):
        return True
    if "articulo" in texto_linea or "bulto" in texto_linea or "kilo" in texto_linea and "kilo" == celdas[0].strip().lower():
        pass
    palabras_header = ["articulo", "precio*", "heavy", "raices", "group"]
    if any(p in texto_linea for p in palabras_header):
        return True
    return False

def extraer_filas_de_tabla_completa(tabla):
    """
    La tabla tiene 9 columnas: 
    [Articulo1, Cantidad1, Bulto1, Kilo1, '' (gutter), Articulo2, Cantidad2, Bulto2, Kilo2]
    A veces alguna fila puede tener 8 columnas si el gutter no se detectó -> lo manejamos.
    """
    filas = []
    for fila in tabla:
        if not fila:
            continue
        celdas = [str(c).strip() if c is not None else "" for c in fila]

        # Normalizar a 9 columnas rellenando si faltan
        while len(celdas) < 9:
            celdas.append("")

        bloque_izq = celdas[0:4]
        bloque_der = celdas[5:9]

        for bloque in (bloque_izq, bloque_der):
            articulo = bloque[0]
            if not articulo:
                continue
            texto_linea = " ".join(bloque).lower()
            if any(p in texto_linea for p in ["articulo", "precio*", "cantidad" ]) and bloque[0].lower() in ("articulo", "producto"):
                continue
            if bloque[0].strip().lower() in ("articulo", "producto", "heavy", ""):
                continue
            
            # Filtra encabezados de sub-tablas (ej: "CONGELADOS", "PRODUCTO")
            if len(bloque) > 1 and bloque[1].strip().lower() == "cantidad":
                continue

            cantidad = bloque[1] if len(bloque) > 1 else ""
            precio_bulto_raw = bloque[2] if len(bloque) > 2 else ""
            precio_kilo_raw = bloque[3] if len(bloque) > 3 else ""

            filas.append({
                'Articulo': articulo,
                'Cantidad_Presentacion': cantidad,
                'Precio_Bulto': limpiar_numero(precio_bulto_raw),
                'Precio_Kilo': limpiar_numero(precio_kilo_raw)
            })
    return filas

# --- AHORA BUSCA SIEMPRE EL NOMBRE ESTÁNDAR ---
pdf_path = "raices.pdf"

if not os.path.exists(pdf_path):
    print(f"❌ No se encontró '{pdf_path}' en la carpeta. Acordate de renombrar el PDF de la verdulería a 'raices.pdf'.")
else:
    print(f"📄 Procesando: '{pdf_path}'...")
    todas_las_filas = []

    with pdfplumber.open(pdf_path) as pdf:
        for num_pag, pagina in enumerate(pdf.pages, start=1):
            tablas = pagina.extract_tables(table_settings={
                "vertical_strategy": "lines",
                "horizontal_strategy": "lines",
                "intersection_tolerance": 5,
            })

            for tabla in tablas:
                filas = extraer_filas_de_tabla_completa(tabla)
                todas_las_filas.extend(filas)

    if todas_las_filas:
        df_final = pd.DataFrame(todas_las_filas)
        df_final = df_final.drop_duplicates(subset=['Articulo', 'Precio_Bulto'])

        excel_salida = "precios_verduleria_limpios.xlsx"
        try:
            df_final.to_excel(excel_salida, index=False)
            print(f"\n✅ ¡CONSEGUIDO PERFECTO!")
            print(f"📊 Se extrajeron {len(df_final)} artículos en total (Lado Izq + Lado Der).")
            print(f"📁 Guardado en '{excel_salida}'")
        except PermissionError:
            excel_salida_alt = "precios_verduleria_completo.xlsx"
            df_final.to_excel(excel_salida_alt, index=False)
            print(f"\n⚠️ Tenías el Excel abierto. Se guardó como: '{excel_salida_alt}'")
    else:
        print("⚠️ No se pudieron extraer datos de las dos columnas.")

print("\n==================================================")