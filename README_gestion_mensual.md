# Gestión mensual — generador local

Flujo mensual (todo local, nada se escribe en Drive):

1. Exportar en solo lectura a `Copias_Gestion/octubre_trabajo/` (gestión del mes anterior, conteos, Maestro, Recetas, PRECIOS Auto) y refrescar `cache/facturas_<Local>.pkl`.
2. `python gestion_mensual.py [--mes 10 --anio 2026]` → `Copias_Gestion/octubre_VF/Gestion_<Local>_<mes>.xlsx`
   (`--prueba` rellena los conteos con los del cierre anterior para validar precios y fórmulas).
3. `python conteos_octubre.py` → `Conteo_<Local>_<mes>.xlsx` (formato Arcos, conteos en blanco, SKU unificado).
4. `python verificar_gestion.py [carpeta]` → chequeo independiente de fórmulas, duplicados y reglas de Agus.

Reglas de precio, correcciones de Agus y límites (±50 % general; lista oficial del Obrador hasta +300 % / −70 %): ver el docstring de `gestion_mensual.py`
y `Copias_Gestion/octubre_VF/LEEME_octubre.md`. Ajuste de columnas de los GAS para los libros nuevos: `GAS/AJUSTE_COLUMNAS_OCTUBRE.md`.
Dependencias: `pandas`, `openpyxl`, `rapidfuzz` (ver `requirements_agente.txt`).
