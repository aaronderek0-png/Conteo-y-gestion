# Agente Mishiguene

Ecosistema de gestión de costos/inventario de Mishiguene (locales: Arcos, Congreso/Obrador, Lafinur, Cabello, Chimidoha).

## Piezas
- `GAS/` — 15 proyectos de Apps Script clonados con clasp. **Leer `GAS/ECOSISTEMA.md` primero.** Registro en `GAS/gas_registry.json`.
- `*.py` (raíz) — pipeline Python: `drive_connector`, `consolidador`, `comparador_precios`, `analisis_gestion_comercial`, `motor_recetas`, `defensivo`, `ecosistema_scan`.
- `Copias_Gestion/` — copias de las Gestiones y logs de auditoría; se trabaja sobre copias, no sobre originales.
- `WebTest/app.py` — dashboard Streamlit (a conectar con datos reales).
- `credentials.json` — Service Account (no subir a repos).

## Comandos
- `python gas_sync.py pull` — actualiza todo el código GAS desde Google.
- `python gas_sync.py push <carpeta>` — sube UN proyecto, con confirmación.
- clasp: `node_modules\.bin\clasp.cmd` (login ya hecho con aaron.derek0@gmail.com).

## Reglas
- Nunca `push` a GAS ni escribir en Sheets de producción sin que el usuario lo pida explícitamente; primero probar en copia.
- No imprimir ni commitear la API key/teléfono de CallMeBot ni `credentials.json`.
- El Maestro de Productos (`1lOabw9U…`) es la fuente única de SKU/rubros.
