# Conocimiento operativo, cachés y recálculo (Agente Mishiguene)

La memoria de la conversación orienta; la **verdad operativa** sale de fuentes verificadas. Cada dato guarda
fuente, fecha de consulta, vigencia, versión y estado: `validado > observado > inferencia > estimacion > pendiente`.
Una inferencia/estimación nunca pisa un dato validado (`conocimiento.registrar_dato`).

## Mapa de fuentes (`conocimiento/fuentes.json`)
| Fuente | Dónde | Refresco | Se invalida cuando |
|---|---|---|---|
| Maestro (SKU, presentaciones, rubros) | Sheet `1lOabw9U…` / `Arcos_data/maestro.xlsx` | `drive_connector` (export) | cambia la huella del export |
| Facturas por local | 5 planillas vivas (IDS en `imputacion_facturas.IDS_PLANILLAS`) / `Arcos_data/*.xlsx` | export + lectura viva antes de escribir | cualquier escritura nueva |
| Lecturas de comprobantes | `cache_facturas.json` (por SHA; dHash para recompresión) | al leer | nunca caduca; se corrige por SHA |
| Aprendizaje proveedor→producto | `aprendizaje_proveedores.json` | `--aprender` (validado) | una inferencia no lo reemplaza |
| Remitos de proveedor | `remitos_pendientes.json` + pestaña `Remitos_Pendientes` | por lote | factura con referencia explícita |
| Transferencias Obrador | pestaña `Remitos` (5 col.) | por remito | matriz `transferencias_obrador.py` |
| Listas de precios | `Precios/` + `cache_listas.json` | ingesta SHA-256 | lista vencida/sin stock |
| Trabajos | `estado/trabajos.json` (`coordinador.py`) | continuo | — |
| Eventos/errores | `logs/pipeline_eventos.jsonl` (`registro_fallos.py`) | continuo | rotación manual |

## Cadena de dependencias (`conocimiento.DEPENDENCIAS`)
precio nuevo → filas_factura → costos_bom → gestion_margenes → comparativa_precios → reporte_web.
`registrar_fuente(fuente, huella)` devuelve en orden los derivados a recalcular; `es_valido(derivado)` exige que la huella
de **todas** sus fuentes coincida con la usada al calcular. `escritura_verificada` es un hecho observado y no se recalcula.

## Cola y escritura
`coordinador.Cola`: RECIBIDO→LEIDO→VALIDADO→(REVISAR|ESCRITO), ERROR reintentable. Idempotente por SHA. Un solo escritor
por destino (`destino_lock`). El offset de Telegram y el SHA procesado solo avanzan tras éxito.

## Qué requiere que la PC siga encendida
El bot (`Iniciar_Bot_Vivo.bat` / `telegram_ingestor.py --watch`) corre en esta PC; no depende de la sesión de Claude.
Detener: cerrar su ventana o `Stop-Process -Id <PID>`; verificar con `Get-CimInstance Win32_Process | ? CommandLine -like *telegram_ingestor*`.
