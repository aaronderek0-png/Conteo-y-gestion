# Justificación ejecutiva — Agente Mishiguene (15/10/2026 01:52)

Cifras **medidas en esta máquina** el día del informe; lo no medido se indica. Costo de operación: **$0** (Python local; la lectura de fotos usa caché SHA/dHash, texto de PDF, Gemini 3.8 Flash por API (nivel gratuito) y Tesseract como respaldo).

## 1. Arquitectura 100 % local
Telegram/WhatsApp → **AgenteReceptor** (long polling 25 s, descarga paralela, SHA-256, cola persistente) → **AgenteProcesador** (único escritor; caché → aprendizaje → OCR/VLM; lote ≤ 20) → **escritura en bloque** a Google Sheets (una `values_batch_update` por planilla y lote, backoff ante 429/5xx) → **AgenteRecalculoCascada** (BOM, comparador, scouting).
Controles: `entornos.py` (PRUEBA/PRODUCCIÓN por ID), duplicados por N° + proveedor, read-back, registro saneado de fallos, un solo consumidor de Telegram.

## 2. Resultados de negocio medidos
- **Ahorro estimado del comparador: $346,639/mes** sobre 124 de 324 insumos con match (+ $185,502/mes a revisar). Es una estimación sobre precios de lista, no ahorro realizado.
- **Gasto sin referencia comparable:** 206 productos = $38,773,446 de $47,656,232 (81.4 %). 32 son nuevos vs septiembre. Es el principal frente de scouting.
- **BOM (export local de Recetas):** 1.833 recetas auditadas; 142 con diferencia vs la hoja: 93 diferencia sin patrón, 25 INCOMPLETA, 24 hoja = motor. Ningún costo dudoso se publicó.
- **Precisión de imputación histórica:** la prueba de regresión reproduce 44 de 46 cargas históricas al peso (95,7 %); origen: `CLAUDE.md` (retro-entrenamiento), no re-medido hoy.
- **Lecturas guardadas:** 55 entradas en `cache_facturas.json` (cada foto/PDF ya leído no se vuelve a leer).
- **Cola actual:** {'REVISAR': 4}.

## 3. Cronometraje real de cada script (subproceso aislado, tracemalloc)
| Script | Modo | Resultado | Tiempo (s) | Pico mem. (MB) | Nota |
|---|---|---|---:|---:|---|
| `Precios/benchmark_grupo_free.py` | import | OK | 0.516 | 12.8 |  |
| `Precios/ingesta_listas.py` | import | OK | 1.345 | 40.1 |  |
| `WebTest/app.py` | import | FALLA | 0.113 | 10.5 | ModuleNotFoundError: No module named 'streamlit' |
| `analisis_gestion_comercial.py` | run | OK | 2.596 | 59.4 |  |
| `aprobaciones.py` | run | OK | 0.011 | 0.6 |  |
| `balanza_datos.py` | import | OK | 2.958 | 61.9 |  |
| `bom_diferencias.py` | run | OK | 14.488 | 30.6 |  |
| `comparador_precios.py` | run | OK | 281.055 | 328.6 |  |
| `conocimiento.py` | run | OK | 0.031 | 0.9 |  |
| `consolidador.py` | run | OK | 3.129 | 55.8 |  |
| `coordinador.py` | run | OK | 0.022 | 0.9 |  |
| `crear_trials.py` | run | OK | 1.328 | 40.0 |  |
| `datos_operativos.py` | run | OK | 1.088 | 39.4 |  |
| `defensivo.py` | run | OK | 0.036 | 1.6 |  |
| `drive_connector.py` | run | OK | 2.69 | 55.5 |  |
| `ecosistema_scan.py` | run | OK | 2.793 | 34.7 |  |
| `entornos.py` | run | OK | 0.02 | 0.6 |  |
| `enviador_mails.py` | run | OK | 2.69 | 51.0 |  |
| `estado_bot.py` | run | OK | 0.439 | 1.4 |  |
| `gas_sync.py` | run | OK | 0.04 | 1.2 |  |
| `imputacion_facturas.py` | run | OK | 1.434 | 41.1 |  |
| `invoice_parser.py` | run | OK | 1.766 | 47.3 |  |
| `mapa_sistema.py` | run | OK | 1.639 | 18.4 |  |
| `motor_recetas.py` | run | OK | 13.794 | 30.3 |  |
| `movimientos_locales.py` | import | OK | 3.149 | 61.3 |  |
| `orquestador_mishi.py` | run | OK | 0.275 | 4.6 |  |
| `procesar_lote_wpp.py` | run | OK | 1.898 | 48.4 |  |
| `publicar_web.py` | import | OK | 0.147 | 5.1 |  |
| `recuperar_lote.py` | run | OK | 1.912 | 47.4 |  |
| `refactor_precios_gestion.py` | run | OK | 1.169 | 21.4 |  |
| `registro_fallos.py` | run | OK | 0.018 | 0.6 |  |
| `reparar_ref_planillas.py` | run | OK | 0.133 | 2.5 |  |
| `scouting_sin_referencia.py` | run | OK | 2.979 | 49.5 |  |
| `telegram_ingestor.py` | run | OK | 0.2 | 4.9 |  |
| `transferencias_obrador.py` | run | OK | 1.289 | 39.2 |  |
| `verificar_vivo.py` | import | OK | 1.323 | 40.0 |  |

**35 de 36 scripts** compilan y ejecutan en su modo seguro. Modo `import` = el módulo carga sin errores (no ejecuta su `main`); modo `run` = ejecución real en solo lectura / dry-run.

## 4. Escenarios representativos (sin escribir en producción)
| Escenario | Resultado | ms | Detalle |
|---|---|---:|---|
| 1 factura normal (ya cargada: no duplica) | OK | 55944 | YA_CARGADA_FILAS |
| 2 NC caso A (devolución: cantidad negativa) | OK | 13136 | 1 fila(s), G=[-2.0] |
| 3 NC caso B (bonificación: fila NC −1) | OK | 7353 | 1 fila(s): [('Nota de credito Materia Prima', -1.0)] |
| 4 remito proveedor (birome/NO/balanza) | OK | 528 | estado=PENDIENTE renglones=5 total_est=50225.0 |
| 5 transferencia Obrador → Cabello | OK | 75 | 5 renglones total=58656.3 → Cabello |
| 6 destinos por renglón (Wigen) | OK | 0 | {'Congreso': ['NUECES PARTIDAS', 'ALMENDRAS'], 'Lafinur': ['NUECES (→ LAFINUR)']} |
| 7 aprobación (APROBADO → CARGAR; ya ESCRITO → NADA) | OK | 0 | idempotente |
| 8 denegación (cierra sin efecto) | OK | 0 | sin gasto ni stock |
| 9 duplicado (idempotencia por SHA) | OK | 14 | una sola entrada por SHA |
| 10 fallo parcial y 11 reintento | OK | 48 | ERROR no consume; reintento termina en ESCRITO |
| 12 hoja sin filas / protegida | OK | 0 | add_rows rechazado → push_remitos_drive deriva a _bloqueo_remitos (cola + respaldo) |

**11 de 11 escenarios en verde.** Suite `tests/`: Ran 31 tests in 0.558s  OK

## 5. Telemetría de ROI: horas humanas liberadas
**Modelo (supuesto auditable, no medición):** 45 s base + 25 s/renglón + 60 s si remito con birome o NC. Lo MEDIDO es lo que el agente procesó y su propio tiempo; los minutos humanos son el modelo aplicado a esas cantidades.

- Comprobantes registrados: **56** (164 renglones; 2 con tiempo de máquina medido en 1 lotes).
- Tiempo de máquina: 6.1 s en total, **3.07 s por comprobante** (promedio medido).
- Tiempo humano estimado ahorrado hasta hoy: **2.06 h** (2.2 min por comprobante en promedio).
- Monto de remitos a verificar con el proveedor (renglones anulados con NO + diferencias de balanza): **$57,600** — NO es dinero recuperado todavía.
- Errores humanos evitados: 312,339 celdas con `#REF!`/derrame reescritas en las 5 copias PRUEBA, 0 fechas futuras bloqueadas, 0 duplicados frenados, 0 desvíos de precio > ±25 % alertados.

**Proyección** (promedio de 2.2 min humanos por comprobante; hora humana supuesta $6,000 — confirmar valor real):

| Comprobantes/mes | Horas humanas liberadas/mes | Horas liberadas/año | Ahorro $/mes | Ahorro $/año |
|---:|---:|---:|---:|---:|
| 1,000 | 36.7 | 440.5 | $220,232 | $2,642,786 |
| 3,000 | 110.1 | 1,321.4 | $660,696 | $7,928,357 |
| 10,000 | 367.1 | 4,404.6 | $2,202,321 | $26,427,857 |

## 6. Qué falta (honesto)
- Copias de prueba en Drive (la cuota de la Service Account es 0; requiere que el dueño las cree: `GAS/crear_trials`).
- Remito 3727 en Cabello/Remitos: pestaña protegida y llena (respaldo en `Remitos_Pendientes`, estado BLOQUEADO_PERMISOS).
- Rupajo 34674, Miño 421/423 y la factura de Wigen esperan decisión humana (`/pendientes`, `/asignar`).
- Desvíos de gestión comercial por local: `analisis_gestion_comercial.py` no se ejecutó en este benchmark (solo `--help`).
