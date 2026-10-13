/**
 * SISTEMA INTEGRAL DE GESTIÓN SEMANAL + EXPLOSIÓN DE RECETAS Y DESVÍOS
 * Archivo: Copy of Trial / TRIAL
 */

const GESTION_GLOBAL = {
  // IDs de Spreadsheets Remotos
  ID_FACTURAS: '1udLEQpQ29avK4c6fNnQig-w02xIV82IgLe85yVJXaOY',
  ID_CONTEO_INICIAL: '1i0LPYWIGqzTVr2KYJXgy5SrbL676EQxEk5qZshaztrM',
  HOJA_CONTEO_INICIAL: 'Historial', 
  ID_TRASPASOS: '1FbRoJwxfZ-vLeGKosv5yoJHAtW17I9M6Hljja0hTvwk',
  ID_LOGS_RECETAS: '1SXFt1KdsYblVVea3Ms0gQ7teuknbxNqrOBeF1Kfisx0',
  HOJA_LOGS_RECETAS: 'Historico_Maestro',
  
  // Archivos de Recetas y Salidas Maxirest
  ID_RECETAS: '1u7QSi7JyCM1J6-DNKSWcha2qIyF2Ybxxp9SVqdNpyfY',
  HOJA_RECETAS: 'Recetas',
  ID_VENTAS_MAXI: '1IRRh0wfv8oUoeQxqjZ0jSFKQ_xPUxpEqxLZCiYKpyGE',
  HOJA_VENTAS_MAXI: 'Historial_Productos',
  HOJA_EQUIVALENCIAS: 'Equivalencias',

  LOCAL_OBJETIVO: 'CONGRESO',
  
  HOJAS_SEMANAS: ['Semana1', 'Semana2', 'Semana3', 'Semana4', 'Semana5'],
  FILA_FECHAS_HEADER: 2,
  FILA_INICIO_DATOS: 3,

  // Índices base 0 en Facturas
  FAC_COL_FECHA_FAC: 1,   // Col B
  FAC_COL_FECHA_ING: 2,   // Col C
  FAC_COL_PROD: 4,        // Col E
  FAC_COL_CANT_INV: 13,   // Col N
  FAC_COL_PRECIO_INV: 14, // Col O
  FAC_COL_CODIGO_X: 23,   // Col X
  FAC_COL_CODIGO_W: 22,   // Col W

  // Índices base 0 en Historial_Productos (Maxirest)
  VENTAS_COL_FECHA: 0,    // Col A
  VENTAS_COL_CODIGO: 1,   // Col B
  VENTAS_COL_CANT: 3,     // Col D (UNIDADES)

  // Mapeo columnas Gestión Semanal en TRIAL (Base 1)
  COL_RUBRO: 1,           // A
  COL_MEDIDA: 2,          // B
  COL_CODIGO: 3,          // C
  COL_DESCRIPCION: 4,     // D
  COL_DIA_INICIO: 5,      // E (Día 1)
  COL_DIA_FIN: 11,        // K (Día 7 - Domingo de corte)
  COL_TOTAL_MOV: 12,      // L
  COL_PRECIO_UN: 13,      // M
  COL_TOTAL_PESOS: 14,    // N
  COL_SALDO_INI: 15,      // O
  COL_INI_MAS_COMP: 16,   // P
  COL_SALDO_FIN: 17,      // Q
  COL_CONSUMO: 18,        // R (Consumo Real Balanza)
  COL_CONSUMO_VAL: 19,    // S
  COL_MERC_VAL: 20,       // T
  
  // Columnas de Control Teórico y Desvío
  COL_TEORICO_VENTA: 21,  // U (Salida Teórica Recetas - Cantidad)
  COL_DESVIO_CANT: 22,    // V (R - U - Cantidad)
  COL_DESVIO_VAL: 23      // W (V * M en $)
};

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('⚙️ CONTROL GESTIÓN')
      .addItem('🚀 Sync Total + Explosión Teórica', 'ejecutarSincronizacionTotal')
      .addToUi();
  } catch (e) {
    console.log("Contexto sin UI: " + e.message);
  }
}

function ejecutarSincronizacionTotal() {
  const ssDestino = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = GESTION_GLOBAL;

  // =========================================================================
  // 1. CARGA DE CONTEO FÍSICO (HISTORIAL / CONTEO)
  // =========================================================================
  const mapaConteoPorFecha = new Map();
  let conteoDisponible = false;

  try {
    let shConteo = ssDestino.getSheetByName('Historial') || ssDestino.getSheetByName(cfg.HOJA_CONTEO_INICIAL);
    if (!shConteo) {
      const ssRemoto = SpreadsheetApp.openById(cfg.ID_CONTEO_INICIAL);
      shConteo = ssRemoto.getSheetByName('Historial') || ssRemoto.getSheetByName(cfg.HOJA_CONTEO_INICIAL) || ssRemoto.getSheets()[0];
    }

    if (shConteo && shConteo.getLastRow() >= 2) {
      const uFilaC = shConteo.getLastRow();
      const datosC = shConteo.getRange(2, 1, uFilaC - 1, 4).getValues();

      datosC.forEach(r => {
        const fechaKey = claveFechaStockInicial_(r[0]);
        const prod = String(r[1] || "").trim();
        const cod = String(r[2] || "").trim().toUpperCase(); // SKU en Col C
        const cant = parseNumero_(r[3]);                     // Cantidad en Col D

        if (!fechaKey || !cod || cod === "NO ENCONTRADO" || esExcluido_(prod) || esExcluido_(cod)) return;

        const obj = { sku: cod, prod: prod, cant: cant };
        if (!mapaConteoPorFecha.has(fechaKey)) {
          mapaConteoPorFecha.set(fechaKey, new Map());
        }
        mapaConteoPorFecha.get(fechaKey).set(cod, obj);
      });
      conteoDisponible = true;
    }
  } catch (err) {
    console.error("Error al cargar Conteo Historial: " + err.message);
  }

  // =========================================================================
  // 2. CARGA DE FACTURAS (COMPRAS Y PRECIOS)
  // =========================================================================
  const comprasFacturasArray = [];
  try {
    const ssFacturas = SpreadsheetApp.openById(cfg.ID_FACTURAS);
    const shFac = ssFacturas.getSheetByName('Facturas');
    if (shFac && shFac.getLastRow() > 1) {
      const uFilaF = shFac.getLastRow();
      const uColF = shFac.getLastColumn();
      const datosF = shFac.getRange(2, 1, uFilaF - 1, uColF).getValues();

      for (let i = 0; i < datosF.length; i++) {
        const row = datosF[i];
        const prodNombre = String(row[cfg.FAC_COL_PROD] || "").trim();
        
        let sku = "";
        if (row.length > cfg.FAC_COL_CODIGO_X && row[cfg.FAC_COL_CODIGO_X]) {
          sku = String(row[cfg.FAC_COL_CODIGO_X]).trim().toUpperCase();
        }
        if (!sku && row.length > cfg.FAC_COL_CODIGO_W && row[cfg.FAC_COL_CODIGO_W]) {
          sku = String(row[cfg.FAC_COL_CODIGO_W]).trim().toUpperCase();
        }

        if (esExcluido_(prodNombre) || esExcluido_(sku)) continue;
        if (!sku || sku === "SIN SKU" || sku === "NO ENCONTRADO") continue;

        const fVal = row[cfg.FAC_COL_FECHA_ING] || row[cfg.FAC_COL_FECHA_FAC];
        const fechaObj = parsearFechaCompleta_(fVal);
        const dKey = normalizarClaveDia_(fVal);
        if (!fechaObj || !dKey) continue;

        const cantInv = parseNumero_(row[cfg.FAC_COL_CANT_INV]);
        
        let precioInv = null;
        const celdaPrecio = row[cfg.FAC_COL_PRECIO_INV];
        if (celdaPrecio !== "" && celdaPrecio !== null && celdaPrecio !== undefined) {
          precioInv = parseNumeroConCero_(celdaPrecio);
        }

        comprasFacturasArray.push({
          sku: sku,
          dKey: dKey,
          cant: cantInv,
          fechaTime: fechaObj.getTime(),
          precio: precioInv
        });
      }
    }
  } catch (err) {
    console.error("Error al cargar Facturas: " + err.message);
  }

  // =========================================================================
  // 3. CARGA DE LOG RECETAS / PLATOS (Historico_Maestro)
  // =========================================================================
  const preciosLogRecetasArray = [];
  try {
    const ssLogs = SpreadsheetApp.openById(cfg.ID_LOGS_RECETAS);
    const shLogs = ssLogs.getSheetByName(cfg.HOJA_LOGS_RECETAS) || ssLogs.getSheets()[0];
    const uFilaL = shLogs.getLastRow();
    const uColL = shLogs.getLastColumn();

    if (uFilaL > 1) {
      const datosL = shLogs.getRange(2, 1, uFilaL - 1, Math.max(uColL, 6)).getValues();
      for (let i = 0; i < datosL.length; i++) {
        const seg = normTexto_(datosL[i][0]);
        if (seg !== "RECETA" && seg !== "PLATO") continue;

        const skuLog = String(datosL[i][1] || "").trim().toUpperCase();
        if (!skuLog || esExcluido_(skuLog)) continue;

        const celdaPrecio = datosL[i][3];
        if (celdaPrecio === "" || celdaPrecio === null || celdaPrecio === undefined) continue;
        const precioLog = parseNumeroConCero_(celdaPrecio);

        const fechaLogVal = datosL[i][5];
        const fLogObj = parsearFechaCompleta_(fechaLogVal);
        if (!fLogObj) continue;

        preciosLogRecetasArray.push({
          sku: skuLog,
          fechaTime: fLogObj.getTime(),
          precio: precioLog
        });
      }
    }
  } catch (err) {
    console.error("Error al cargar Historico_Maestro: " + err.message);
  }

  // =========================================================================
  // 4. CARGA DE TRASPASOS APROBADOS
  // =========================================================================
  const traspasosMap = {};
  try {
    const ssTraspasos = SpreadsheetApp.openById(cfg.ID_TRASPASOS);
    let shTras = ssTraspasos.getSheetByName('Traspasos') || ssTraspasos.getSheetByName('Unico');
    if (!shTras) {
      shTras = ssTraspasos.getSheets().find(sheet => {
        const lastColumn = sheet.getLastColumn();
        if (lastColumn < 1) return false;
        const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(h => normTexto_(h));
        return ["ORIGEN", "DESTINO", "CANTIDAD", "ESTADO"].every(header => headers.includes(header));
      });
    }
    if (!shTras) throw new Error("No se encontró hoja válida de traspasos");
    const uFilaT = shTras.getLastRow();
    const uColT = shTras.getLastColumn();

    if (uFilaT > 1) {
      const headersT = shTras.getRange(1, 1, 1, uColT).getValues()[0].map(h => normTexto_(h));
      let idxFecha = headersT.findIndex(h => h.includes("FECHA"));
      let idxOrigen = headersT.findIndex(h => h.includes("ORIGEN"));
      let idxDestino = headersT.findIndex(h => h.includes("DESTINO"));
      let idxSku = headersT.findIndex(h => h.includes("SKU") || h.includes("CODIGO"));
      let idxCant = headersT.findIndex(h => h.includes("CANTIDAD") || h.includes("CANT"));
      let idxEstado = headersT.findIndex(h => h.includes("ESTADO"));

      const datosT = shTras.getRange(2, 1, uFilaT - 1, uColT).getValues();
      for (let i = 0; i < datosT.length; i++) {
        const row = datosT[i];
        if (normTexto_(row[idxEstado]) !== "APROBADO") continue;

        const origen = normTexto_(row[idxOrigen]);
        const destino = normTexto_(row[idxDestino]);
        const objetivo = normTexto_(cfg.LOCAL_OBJETIVO);

        let factor = 0;
        if (origen === objetivo) factor = -1;
        else if (destino === objetivo) factor = 1;
        if (factor === 0) continue;

        const skuT = String(row[idxSku] || "").trim().toUpperCase();
        if (!skuT || esExcluido_(skuT)) continue;

        const dKeyT = normalizarClaveDia_(row[idxFecha]);
        if (!dKeyT) continue;

        const cantT = parseNumero_(row[idxCant]) * factor;
        if (!traspasosMap[skuT]) traspasosMap[skuT] = {};
        if (!traspasosMap[skuT][dKeyT]) traspasosMap[skuT][dKeyT] = 0;
        traspasosMap[skuT][dKeyT] += cantT;
      }
    }
  } catch (err) {
    console.error("Error al cargar Traspasos: " + err.message);
  }

  // =========================================================================
  // 5. MOTOR DE EXPLOSIÓN: EQUIVALENCIAS, RECETAS Y VENTAS MAXIREST
  // =========================================================================
  const mapaEquivalencias = cargarMapaEquivalencias_(ssDestino);
  const recetasMap = cargarArbolRecetas_();
  const ventasMaxiArray = cargarVentasMaxirest_(mapaEquivalencias);

  // =========================================================================
  // 6. PROCESAMIENTO SEMANAL CON SALDOS INICIAL/FINAL Y CÁLCULOS
  // =========================================================================
  let totalSemanasProcesadas = 0;

  cfg.HOJAS_SEMANAS.forEach(nomHoja => {
    const sh = ssDestino.getSheetByName(nomHoja);
    if (!sh) return;

    let uFila = sh.getLastRow();
    if (uFila < cfg.FILA_FECHAS_HEADER) return;

    // Encabezados
    sh.getRange(cfg.FILA_FECHAS_HEADER, cfg.COL_TEORICO_VENTA).setValue("Venta Teórica");
    sh.getRange(cfg.FILA_FECHAS_HEADER, cfg.COL_DESVIO_CANT).setValue("Desvío (Cant)");
    sh.getRange(cfg.FILA_FECHAS_HEADER, cfg.COL_DESVIO_VAL).setValue("Desvío ($)");

    const encabezadosDias = sh.getRange(cfg.FILA_FECHAS_HEADER, cfg.COL_DIA_INICIO, 1, 7).getValues()[0];
    const clavesDiasSemana = encabezadosDias.map(d => normalizarClaveDia_(d));

    const fechaInicioSemana = parsearFechaCompleta_(encabezadosDias[0]);
    const fechaDomingoCorte = parsearFechaCompleta_(encabezadosDias[6]);

    const fechaInicioKey = claveFechaStockInicial_(fechaInicioSemana);
    const fechaDomingoPrevioKey = fechaInicioSemana ? claveFechaStockInicial_(new Date(fechaInicioSemana.getTime() - 24 * 60 * 60 * 1000)) : "";
    
    const fechaFinKey = claveFechaStockInicial_(fechaDomingoCorte);
    const fechaLunesSiguienteKey = fechaDomingoCorte ? claveFechaStockInicial_(new Date(fechaDomingoCorte.getTime() + 24 * 60 * 60 * 1000)) : "";

    const stockInicialMapa = mapaConteoPorFecha.get(fechaInicioKey) || mapaConteoPorFecha.get(fechaDomingoPrevioKey) || new Map();
    const stockFinalMapa = mapaConteoPorFecha.get(fechaFinKey) || mapaConteoPorFecha.get(fechaLunesSiguienteKey) || new Map();

    let timestampCorte = Infinity;
    if (fechaDomingoCorte) {
      fechaDomingoCorte.setHours(23, 59, 59, 999);
      timestampCorte = fechaDomingoCorte.getTime();
    }

    // Filtrar Compras y Precios al corte
    const comprasValidasMap = {};
    const precioUltimoFac = {};
    const maxFechaFac = {};

    comprasFacturasArray.forEach(reg => {
      if (reg.fechaTime <= timestampCorte) {
        if (!comprasValidasMap[reg.sku]) comprasValidasMap[reg.sku] = {};
        if (!comprasValidasMap[reg.sku][reg.dKey]) comprasValidasMap[reg.sku][reg.dKey] = 0;
        comprasValidasMap[reg.sku][reg.dKey] += reg.cant;

        if (reg.precio !== null) {
          if (!maxFechaFac[reg.sku] || reg.fechaTime >= maxFechaFac[reg.sku]) {
            maxFechaFac[reg.sku] = reg.fechaTime;
            precioUltimoFac[reg.sku] = reg.precio;
          }
        }
      }
    });

    const precioUltimoLog = {};
    const maxFechaLog = {};

    preciosLogRecetasArray.forEach(reg => {
      if (reg.fechaTime <= timestampCorte) {
        if (!maxFechaLog[reg.sku] || reg.fechaTime >= maxFechaLog[reg.sku]) {
          maxFechaLog[reg.sku] = reg.fechaTime;
          precioUltimoLog[reg.sku] = reg.precio;
        }
      }
    });

    // EXPLOSIÓN DE VENTAS DE LA SEMANA
    const consumoSemanaMP = {};
    ventasMaxiArray.forEach(v => {
      if (v.fechaTime <= timestampCorte && clavesDiasSemana.includes(v.dKey)) {
        explotarRecetaRecursivo_(v.skuReceta, v.cant, recetasMap, consumoSemanaMP);
      }
    });

    let cantFilas = uFila >= cfg.FILA_INICIO_DATOS ? (uFila - cfg.FILA_INICIO_DATOS + 1) : 0;
    let matrizDatos = cantFilas > 0 ? sh.getRange(cfg.FILA_INICIO_DATOS, 1, cantFilas, cfg.COL_DESVIO_VAL).getValues() : [];

    for (let r = 0; r < matrizDatos.length; r++) {
      const row = matrizDatos[r];
      const sku = String(row[cfg.COL_CODIGO - 1] || "").trim().toUpperCase();
      const desc = String(row[cfg.COL_DESCRIPCION - 1] || "").trim();
      const descNorm = normTexto_(desc);
      const medida = normTexto_(row[cfg.COL_MEDIDA - 1]);

      if (esExcluido_(desc) || esExcluido_(sku)) continue;

      // A) Saldo Inicial (Col O)
      if (conteoDisponible && sku && stockInicialMapa.has(sku)) {
        row[cfg.COL_SALDO_INI - 1] = stockInicialMapa.get(sku).cant;
      }

      // B) Saldo Final (Col Q)
      if (conteoDisponible && sku && stockFinalMapa.has(sku)) {
        row[cfg.COL_SALDO_FIN - 1] = stockFinalMapa.get(sku).cant;
      }

      // C) Días E a K: Compras + Traspasos
      let totalMovimientosSemana = 0;
      for (let d = 0; d < 7; d++) {
        const dKey = clavesDiasSemana[d];
        let movDia = 0;
        if (dKey) {
          if (sku && comprasValidasMap[sku] && comprasValidasMap[sku][dKey]) {
            movDia += comprasValidasMap[sku][dKey];
          }
          if (sku && traspasosMap[sku] && traspasosMap[sku][dKey]) {
            movDia += traspasosMap[sku][dKey];
          }
        }
        row[cfg.COL_DIA_INICIO - 1 + d] = movDia !== 0 ? Math.round(movDia * 1000) / 1000 : "";
        totalMovimientosSemana += movDia;
      }

      // Col L: Total
      totalMovimientosSemana = Math.round(totalMovimientosSemana * 1000) / 1000;
      row[cfg.COL_TOTAL_MOV - 1] = totalMovimientosSemana;

      // Col M: Precio Unitario
      let precioFinalFila = null;
      if (sku && precioUltimoFac.hasOwnProperty(sku)) {
        precioFinalFila = precioUltimoFac[sku];
      } else if (sku && precioUltimoLog.hasOwnProperty(sku)) {
        precioFinalFila = precioUltimoLog[sku];
      } else {
        const valActual = row[cfg.COL_PRECIO_UN - 1];
        if (valActual !== "" && valActual !== null && valActual !== undefined) {
          precioFinalFila = parseNumeroConCero_(valActual);
        }
      }

      row[cfg.COL_PRECIO_UN - 1] = precioFinalFila !== null ? precioFinalFila : "";
      const precioCalculo = precioFinalFila !== null ? precioFinalFila : 0;

      // Col N: Total $
      row[cfg.COL_TOTAL_PESOS - 1] = totalMovimientosSemana * precioCalculo;

      // Col P: Saldo Inicial + Compras
      const saldoIni = parseNumero_(row[cfg.COL_SALDO_INI - 1]);
      const iniMasComp = saldoIni + totalMovimientosSemana;
      row[cfg.COL_INI_MAS_COMP - 1] = iniMasComp;

      // Col R, S, T: Consumo Real Físico
      const celdaFinal = row[cfg.COL_SALDO_FIN - 1];
      const tieneSaldoFinal = celdaFinal !== "" && celdaFinal != null;
      const saldoFin = parseNumero_(celdaFinal);

      let consumoRealFisico = null;
      if (tieneSaldoFinal) {
        consumoRealFisico = iniMasComp - saldoFin;
        row[cfg.COL_CONSUMO - 1] = Math.round(consumoRealFisico * 1000) / 1000;
        row[cfg.COL_CONSUMO_VAL - 1] = Math.round((consumoRealFisico * precioCalculo) * 100) / 100;
        row[cfg.COL_MERC_VAL - 1] = Math.round((saldoFin * precioCalculo) * 100) / 100;
      } else {
        row[cfg.COL_CONSUMO - 1] = "";
        row[cfg.COL_CONSUMO_VAL - 1] = "";
        row[cfg.COL_MERC_VAL - 1] = "";
      }

      // Col U: Venta Teórica Explotada (CANTIDAD PURA)
      let consumoTeorico = consumoSemanaMP[sku] || 0;
      const esMedidaMayor = ["KG", "KILOS", "L", "LT", "LITROS"].includes(medida);
      if (consumoTeorico > 0 && esMedidaMayor) {
        consumoTeorico = consumoTeorico / 1000;
      }
      if (consumoTeorico > 0 && medida === "UN" && consumoTeorico > 500) {
        consumoTeorico = consumoTeorico / 1000;
      }

      if (consumoTeorico > 0) {
        row[cfg.COL_TEORICO_VENTA - 1] = Math.round(consumoTeorico * 1000) / 1000;
      } else {
        row[cfg.COL_TEORICO_VENTA - 1] = "";
      }

      // Col V y W: Desvíos (Cant y $)
      if (tieneSaldoFinal && consumoRealFisico !== null) {
        const desvioCant = consumoRealFisico - consumoTeorico;
        row[cfg.COL_DESVIO_CANT - 1] = Math.round(desvioCant * 1000) / 1000; // Col V: Cantidad
        row[cfg.COL_DESVIO_VAL - 1] = Math.round((desvioCant * precioCalculo) * 100) / 100; // Col W: $
      } else {
        row[cfg.COL_DESVIO_CANT - 1] = "";
        row[cfg.COL_DESVIO_VAL - 1] = "";
      }
    }

    if (matrizDatos.length > 0) {
      // 1. Escribir valores
      sh.getRange(cfg.FILA_INICIO_DATOS, 1, matrizDatos.length, cfg.COL_DESVIO_VAL).setValues(matrizDatos);

      // 2. FORMATEO FORZADO DE CELDAS:
      // Col U (21) y Col V (22): Cantidades (Numérico con 3 decimales, SIN signo $)
      sh.getRange(cfg.FILA_INICIO_DATOS, cfg.COL_TEORICO_VENTA, matrizDatos.length, 2).setNumberFormat("#,##0.000");

      // Col W (23): Valorizado (Formato Moneda $)
      sh.getRange(cfg.FILA_INICIO_DATOS, cfg.COL_DESVIO_VAL, matrizDatos.length, 1).setNumberFormat("$ #,##0.00");
    }

    totalSemanasProcesadas++;
  });

  const msg = `Sincronización Total finalizada en ${totalSemanasProcesadas} semanas.\n` +
              `• Columna U (Venta Teórica): Medida física corregida (sin $).\n` +
              `• Columna V (Desvío Cant): Diferencia física corregida (sin $).\n` +
              `• Columna W (Desvío $): Valorizado en pesos ($).`;
  console.log(msg);
  mostrarAlerta_("🚀 CONTROL TOTAL EJECUTADO", msg);
}

// ===================== MÓDULOS DE EXPLOSIÓN Y EQUIVALENCIAS =====================

function cargarMapaEquivalencias_(ss) {
  const cfg = GESTION_GLOBAL;
  const sh = ss.getSheetByName(cfg.HOJA_EQUIVALENCIAS);
  const mapa = new Map();
  if (!sh) return mapa;

  const uFila = sh.getLastRow();
  if (uFila < 2) return mapa;

  const datos = sh.getRange(2, 1, uFila - 1, 3).getValues();
  for (let i = 0; i < datos.length; i++) {
    const codMaxiRaw = datos[i][0];
    const skuReceta = String(datos[i][2] || "").trim().toUpperCase();

    if (!codMaxiRaw || !skuReceta) continue;

    const codLimpio = String(codMaxiRaw).replace(/\./g, "").trim();
    mapa.set(codLimpio, skuReceta);
  }
  return mapa;
}

function cargarArbolRecetas_() {
  const cfg = GESTION_GLOBAL;
  const ss = SpreadsheetApp.openById(cfg.ID_RECETAS);
  const sh = ss.getSheetByName(cfg.HOJA_RECETAS);
  const recetasMap = new Map();

  if (!sh) return recetasMap;
  const uFila = sh.getLastRow();
  if (uFila < 2) return recetasMap;

  const datos = sh.getRange(2, 1, uFila - 1, 11).getValues();

  for (let i = 0; i < datos.length; i++) {
    const codReceA = String(datos[i][0] || "").trim().toUpperCase();
    const nomReceB = String(datos[i][1] || "").trim().toUpperCase();
    const codInsumo = String(datos[i][3] || "").trim().toUpperCase();
    const cant = parseNumero_(datos[i][5]);
    const unidad = normTexto_(datos[i][6]);
    const rend = parseNumero_(datos[i][10]) || 1;

    if (!codInsumo || cant <= 0) continue;

    const ingrediente = { cod: codInsumo, cant: cant, unidad: unidad };

    [codReceA, nomReceB].forEach(clave => {
      if (clave) {
        if (!recetasMap.has(clave)) {
          recetasMap.set(clave, { rend: rend, ingredientes: [] });
        }
        recetasMap.get(clave).ingredientes.push(ingrediente);
      }
    });
  }
  return recetasMap;
}

function cargarVentasMaxirest_(mapaEquivalencias) {
  const cfg = GESTION_GLOBAL;
  const ss = SpreadsheetApp.openById(cfg.ID_VENTAS_MAXI);
  const sh = ss.getSheetByName(cfg.HOJA_VENTAS_MAXI) || ss.getSheets()[0];
  const ventas = [];

  if (!sh) return ventas;
  const uFila = sh.getLastRow();
  if (uFila < 2) return ventas;

  const datos = sh.getRange(2, 1, uFila - 1, 4).getValues();

  for (let i = 0; i < datos.length; i++) {
    const fVal = datos[i][cfg.VENTAS_COL_FECHA];
    const codMaxiRaw = datos[i][cfg.VENTAS_COL_CODIGO];
    const cantVendida = parseNumero_(datos[i][cfg.VENTAS_COL_CANT]);

    if (!fVal || !codMaxiRaw || cantVendida <= 0) continue;

    const codLimpio = String(codMaxiRaw).replace(/\./g, "").trim();
    const skuReceta = mapaEquivalencias.get(codLimpio);

    if (!skuReceta) continue;

    const fechaObj = parsearFechaCompleta_(fVal);
    const dKey = normalizarClaveDia_(fVal);
    if (!fechaObj || !dKey) continue;

    ventas.push({
      skuReceta: skuReceta,
      cant: cantVendida,
      dKey: dKey,
      fechaTime: fechaObj.getTime()
    });
  }
  return ventas;
}

function explotarRecetaRecursivo_(codigo, cantidadRequerida, recetasMap, acumuladorMP, visitados) {
  if (!visitados) visitados = new Set();
  
  if (visitados.has(codigo)) return;
  visitados.add(codigo);

  if (recetasMap.has(codigo)) {
    const receta = recetasMap.get(codigo);
    const rend = receta.rend > 0 ? receta.rend : 1;
    const factor = cantidadRequerida / rend;

    receta.ingredientes.forEach(ing => {
      explotarRecetaRecursivo_(ing.cod, ing.cant * factor, recetasMap, acumuladorMP, new Set(visitados));
    });
  } else {
    acumuladorMP[codigo] = (acumuladorMP[codigo] || 0) + cantidadRequerida;
  }
}

// ===================== FUNCIONES AUXILIARES =====================

function parseNumeroConCero_(v) {
  if (v === 0 || v === "0") return 0;
  if (v === null || v === undefined || v === "") return null;
  if (typeof v === "number") return isNaN(v) ? null : v;
  const s = v.toString().trim().replace("$", "").replace(/\s/g, "").replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function parseNumero_(v) {
  const n = parseNumeroConCero_(v);
  return n === null ? 0 : n;
}

function parsearFechaCompleta_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }
  const s = String(v).trim();
  
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) {
    let dia = parseInt(m[1], 10);
    let mes = parseInt(m[2], 10);
    let anio = parseInt(m[3], 10);
    if (anio < 100) anio += 2000;
    
    if (dia > 12 && mes <= 12) {
      return new Date(anio, mes - 1, dia);
    } else if (dia <= 12 && mes > 12) {
      return new Date(anio, dia - 1, mes);
    }
    return new Date(anio, mes - 1, dia);
  }
  
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})/);
  if (m) {
    let dia = parseInt(m[1], 10);
    let mes = parseInt(m[2], 10);
    return new Date(2026, mes - 1, dia);
  }
  return null;
}

function claveFechaStockInicial_(valor) {
  const fecha = parsearFechaCompleta_(valor);
  if (!fecha) return "";
  const d = ("0" + fecha.getDate()).slice(-2);
  const m = ("0" + (fecha.getMonth() + 1)).slice(-2);
  return `${fecha.getFullYear()}-${m}-${d}`;
}

function normalizarClaveDia_(v) {
  const fecha = parsearFechaCompleta_(v);
  if (!fecha) return "";
  return fecha.getDate() + "/" + (fecha.getMonth() + 1);
}

function normTexto_(t) {
  if (!t) return "";
  return t.toString().trim().toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w\s]/gi, "")
    .replace(/\s+/g, " ");
}

function esExcluido_(texto) {
  if (!texto) return false;
  const t = normTexto_(texto);
  return t === "VARIOS" || t.includes("VARIOS") ||
         t === "FLETE" || t === "FLETES" || t.includes("FLETE");
}

function mostrarAlerta_(titulo, mensaje) {
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui) {
      ui.alert(titulo, mensaje, ui.ButtonSet.OK);
      return;
    }
  } catch (e) {}
  console.log(`[${titulo}] ${mensaje}`);
}