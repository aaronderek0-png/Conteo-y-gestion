/**
 * SISTEMA INTEGRAL DE GESTIÓN DE COSTOS Y ESCANDALLOS (MISHIGUENE)
 * v20.1 - Maestro con F (Lafinur) y G (Café) + Alertas UI, Mail y WhatsApp
 */

const CONFIG = {
  EMAIL_DESTINO: "mishiguene.cervino@gmail.com",
  ORIGEN_INSUMOS_ID: "1lOabw9U4ilRwE5r6QXQ2Jc1nqkKUK3Hi096xga0XFnI",
  ORIGEN_INSUMOS_HOJA: "PRODUCTOS",

  LOCAL_HOJA_INSUMOS: "Insumos",
  LOCAL_HOJA_RECETAS: "Recetas",
  LOCAL_HOJA_COSTOS: "Costos",
  LOCAL_HOJA_MAESTRO: "Maestro",

  TARGET_LOGS_ID: "1SXFt1KdsYblVVea3Ms0gQ7teuknbxNqrOBeF1Kfisx0",
  TARGET_LOGS_HOJA: "Historico_Maestro",
  
  UMBRAL_INSUMOS: 0.15, // ±15%
  UMBRAL_RECETAS: 0.10  // ±10%
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Gestión Costos')
    .addItem('⚡ 1. Recalcular Recetas, Costos y Maestro', 'ejecutarRecalculoCompletoLocal')
    .addItem('🔄 2. Sincronizar Insumos (Desde Origen)', 'menuSincronizarInsumos')
    .addItem('⚡🔄 3. Pipeline Rápido (Sincronizar + Recálculo + Alertas)', 'ejecutarPipelineRapido')
    .addItem('📸 4. Generar Snapshot Histórico (Logs)', 'generarSnapshotMaestro')
    .addSeparator()
    .addItem('🚀 5. Pipeline Total (+ Histórico Logs)', 'ejecutarPipelineTotal')
    .addToUi();
}

/**
 * Validador booleano simple para determinar si un ítem pasa al Maestro
 */
function esActivo(val) {
  if (!val) return false;
  const str = val.toString().trim().toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  return str === "SI" || str === "TRUE" || str === "1";
}

/**
 * Limpia el texto para preservar SI, NO o dejar celda vacía exacta
 */
function limpiarTextoExacto(val) {
  if (val === null || val === undefined) return "";
  const str = val.toString().trim();
  if (str === "") return "";
  const norm = str.toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (norm === "SI" || norm === "TRUE" || norm === "1") return "SI";
  if (norm === "NO" || norm === "FALSE" || norm === "0") return "NO";
  return str;
}

/**
 * 1. RECALCULAR RECETAS, COSTOS Y MAESTRO
 */
function ejecutarRecalculoCompletoLocal() {
  const t0 = new Date();
  recalcularRecetasYCostos();
  recalcularHojaMaestro();
  const duracion = ((new Date() - t0) / 1000).toFixed(1);
  
  try {
    SpreadsheetApp.getUi().alert("✅ Sistema Actualizado", `Recetas, Costos y Maestro actualizados en ${duracion}s.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

function menuSincronizarInsumos() {
  const alertasInsumos = sincronizarInsumos(false);
  if (alertasInsumos.length > 0) {
    enviarReporteConsolidado(alertasInsumos, []);
    mostrarAlertaUI(alertasInsumos, []);
    dispararAlertaWhatsApp(alertasInsumos, []);
  } else {
    try {
      SpreadsheetApp.getUi().alert("✅ Insumos Actualizados", "Insumos sincronizados sin variaciones críticas (±15%).", SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {}
  }
}

function ejecutarPipelineRapido() {
  const t0 = new Date();
  const alertasInsumos = sincronizarInsumos(true);
  const alertasRecetas = recalcularRecetasYCostos();
  recalcularHojaMaestro();
  const duracion = ((new Date() - t0) / 1000).toFixed(1);

  if (alertasInsumos.length > 0 || alertasRecetas.length > 0) {
    enviarReporteConsolidado(alertasInsumos, alertasRecetas);
    mostrarAlertaUI(alertasInsumos, alertasRecetas);
    dispararAlertaWhatsApp(alertasInsumos, alertasRecetas);
  } else {
    try {
      SpreadsheetApp.getUi().alert("🚀 Pipeline Rápido Finalizado", `Insumos, Recetas, Costos y Maestro actualizados en ${duracion}s sin variaciones críticas.`, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {}
  }
}

function ejecutarPipelineTotal() {
  const t0 = new Date();
  const alertasInsumos = sincronizarInsumos(true);
  const alertasRecetas = recalcularRecetasYCostos();
  recalcularHojaMaestro();
  generarSnapshotMaestro();
  const duracion = ((new Date() - t0) / 1000).toFixed(1);

  if (alertasInsumos.length > 0 || alertasRecetas.length > 0) {
    enviarReporteConsolidado(alertasInsumos, alertasRecetas);
    mostrarAlertaUI(alertasInsumos, alertasRecetas);
    dispararAlertaWhatsApp(alertasInsumos, alertasRecetas);
  } else {
    try {
      SpreadsheetApp.getUi().alert("🚀 Pipeline Total Finalizado", `Circuito completo y Logs generados en ${duracion}s sin variaciones críticas.`, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch(e) {}
  }
}

/**
 * 2. SINCRONIZAR INSUMOS
 */
function sincronizarInsumos(silencioso = false) {
  const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInsumos = ssLocal.getSheetByName(CONFIG.LOCAL_HOJA_INSUMOS);
  if (!hojaInsumos) throw new Error(`No se encontró "${CONFIG.LOCAL_HOJA_INSUMOS}"`);

  const ultFilaLocal = hojaInsumos.getLastRow();
  const mapaPreciosViejos = {};

  if (ultFilaLocal >= 2) {
    const datosViejos = hojaInsumos.getRange(2, 1, ultFilaLocal - 1, 6).getValues();
    for (let i = 0; i < datosViejos.length; i++) {
      const cod = datosViejos[i][0]?.toString().trim();
      const precio = Number(datosViejos[i][5]) || 0;
      if (cod) mapaPreciosViejos[cod] = precio;
    }
  }

  const ssOrigen = SpreadsheetApp.openById(CONFIG.ORIGEN_INSUMOS_ID);
  const hojaOrigen = ssOrigen.getSheetByName(CONFIG.ORIGEN_INSUMOS_HOJA);
  const ultFilaOrigen = hojaOrigen.getLastRow();
  if (ultFilaOrigen < 2) return [];

  const datosOrigen = hojaOrigen.getRange(2, 1, ultFilaOrigen - 1, 8).getValues();
  const filasFiltradas = [];
  const alertasInsumos = [];

  for (let i = 0; i < datosOrigen.length; i++) {
    const fila = datosOrigen[i];
    const cod = fila[0]?.toString().trim();
    const nombre = fila[1];
    const prov = fila[2];
    const rubro = fila[3];
    const um = fila[5];
    const pNuevo = Number(fila[7]) || 0;

    if (!cod) continue;

    filasFiltradas.push([cod, nombre, prov, rubro, um, pNuevo]);

    const pViejo = mapaPreciosViejos[cod] || 0;
    if (pViejo > 0 && pNuevo > 0) {
      const variacion = (pNuevo - pViejo) / pViejo;
      if (Math.abs(variacion) >= CONFIG.UMBRAL_INSUMOS) {
        alertasInsumos.push({
          cod: cod,
          nombre: nombre,
          viejo: pViejo.toFixed(2),
          nuevo: pNuevo.toFixed(2),
          esAumento: variacion > 0,
          porcentaje: (variacion > 0 ? "+" : "") + (variacion * 100).toFixed(1) + "%"
        });
      }
    }
  }

  if (ultFilaLocal >= 2) {
    hojaInsumos.getRange(2, 1, ultFilaLocal - 1, 6).clearContent();
  }
  if (filasFiltradas.length > 0) {
    hojaInsumos.getRange(2, 1, filasFiltradas.length, 6).setValues(filasFiltradas);
  }

  return alertasInsumos;
}

/**
 * 3. RECALCULAR RECETAS Y COSTOS (Actualizado para sincronizar U. Medida de Costos H:H a Recetas G:G)
 */
function recalcularRecetasYCostos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInsumos = ss.getSheetByName(CONFIG.LOCAL_HOJA_INSUMOS);
  const hojaRecetas = ss.getSheetByName(CONFIG.LOCAL_HOJA_RECETAS);
  const hojaCostos = ss.getSheetByName(CONFIG.LOCAL_HOJA_COSTOS);

  const mapaCostosViejos = {};
  const mapaActivosCostos = {};
  const mapaUnidadesCostos = {}; // Diccionario para guardar U. MEDIDA (Columna H)
  const ultFilaCostos = hojaCostos.getLastRow();
  
  if (ultFilaCostos >= 2) {
    // Leemos 8 columnas: COD_REC, NOM, NOMBRE, C_UN, ACTIVO, LAFINUR, CAFE, U. MEDIDA
    const datosCostosPrevios = hojaCostos.getRange(2, 1, ultFilaCostos - 1, 8).getValues();
    for (let i = 0; i < datosCostosPrevios.length; i++) {
      const codRec = datosCostosPrevios[i][0]?.toString().trim().toUpperCase();
      const nomRec = datosCostosPrevios[i][1]?.toString().trim().toUpperCase();
      const cUnPrevio = Number(datosCostosPrevios[i][3]) || 0;
      const actCostos = datosCostosPrevios[i][4];
      const umCostos = datosCostosPrevios[i][7]?.toString().trim(); // Col H

      if (codRec) {
        mapaCostosViejos[codRec] = cUnPrevio;
        mapaActivosCostos[codRec] = esActivo(actCostos);
        if (umCostos) mapaUnidadesCostos[codRec] = umCostos;
      }
      if (nomRec && umCostos) {
        mapaUnidadesCostos[nomRec] = umCostos;
      }
    }
  }

  const datosInsumos = hojaInsumos.getDataRange().getValues();
  const dictInsumosPorNombre = {}; 
  const dictInsumosPorCod = {}; 

  for (let i = 1; i < datosInsumos.length; i++) {
    const cod = datosInsumos[i][0]?.toString().trim();
    const nombre = datosInsumos[i][1]?.toString().trim();
    const um = datosInsumos[i][4]?.toString().trim() || "GR";
    const precio = Number(datosInsumos[i][5]) || 0;

    if (nombre) dictInsumosPorNombre[nombre.toUpperCase()] = { cod, um, precio };
    if (cod) dictInsumosPorCod[cod.toUpperCase()] = { um, precio };
  }

  const ultFilaRecetas = hojaRecetas.getLastRow();
  if (ultFilaRecetas < 2) return [];

  const numFilas = ultFilaRecetas - 1;
  const matrizRecetas = hojaRecetas.getRange(2, 1, numFilas, 14).getValues();

  const dictRecetasPorNombre = {};
  for (let i = 0; i < matrizRecetas.length; i++) {
    const codA = matrizRecetas[i][0]?.toString().trim();
    const nomB = matrizRecetas[i][1]?.toString().trim();
    const descC = matrizRecetas[i][2]?.toString().trim();

    const codigoSalida = nomB || codA;
    if (codigoSalida) {
      if (descC) dictRecetasPorNombre[descC.toUpperCase()] = codigoSalida;
      if (nomB) dictRecetasPorNombre[nomB.toUpperCase()] = codigoSalida;
      if (codA) dictRecetasPorNombre[codA.toUpperCase()] = codigoSalida;
    }
  }

  const dictCostosRecetas = {};

  for (let pasada = 0; pasada < 6; pasada++) {
    const totalesPorReceta = {};

    for (let i = 0; i < matrizRecetas.length; i++) {
      const fila = matrizRecetas[i];
      const nomInsumo = fila[4]?.toString().trim();
      let codInsumo = fila[3]?.toString().trim();
      const cant = Number(fila[5]) || 0;
      let pun = 0;
      let um = fila[6]?.toString().trim() || "GR";

      if (nomInsumo) {
        const nomUpper = nomInsumo.toUpperCase();
        if (dictInsumosPorNombre[nomUpper]) {
          codInsumo = dictInsumosPorNombre[nomUpper].cod;
          um = dictInsumosPorNombre[nomUpper].um;
          pun = dictInsumosPorNombre[nomUpper].precio;
        } else if (dictRecetasPorNombre[nomUpper]) {
          codInsumo = dictRecetasPorNombre[nomUpper];
          pun = dictCostosRecetas[codInsumo.toUpperCase()] || 0;
          // Asigna la unidad definida en Costos para esa subreceta
          um = mapaUnidadesCostos[codInsumo.toUpperCase()] || "GR";
        } else if (codInsumo) {
          const codUpper = codInsumo.toUpperCase();
          if (dictCostosRecetas[codUpper] !== undefined) {
            pun = dictCostosRecetas[codUpper];
            um = mapaUnidadesCostos[codUpper] || "GR";
          } else if (dictInsumosPorCod[codUpper]) {
            pun = dictInsumosPorCod[codUpper].precio;
            um = dictInsumosPorCod[codUpper].um;
          }
        }
      } else if (codInsumo) {
        const codUpper = codInsumo.toUpperCase();
        if (dictCostosRecetas[codUpper] !== undefined) {
          pun = dictCostosRecetas[codUpper];
          um = mapaUnidadesCostos[codUpper] || "GR";
        } else if (dictInsumosPorCod[codUpper]) {
          pun = dictInsumosPorCod[codUpper].precio;
          um = dictInsumosPorCod[codUpper].um;
        }
      }

      fila[3] = codInsumo;
      fila[6] = um;
      fila[7] = pun;
      fila[8] = cant * pun;

      const claveB = fila[1]?.toString().trim();
      const claveA = fila[0]?.toString().trim();
      
      if (claveB) {
        totalesPorReceta[claveB.toUpperCase()] = (totalesPorReceta[claveB.toUpperCase()] || 0) + fila[8];
      }
      if (claveA && claveA !== claveB) {
        totalesPorReceta[claveA.toUpperCase()] = (totalesPorReceta[claveA.toUpperCase()] || 0) + fila[8];
      }
    }

    for (let i = 0; i < matrizRecetas.length; i++) {
      const fila = matrizRecetas[i];
      const codA = fila[0]?.toString().trim();
      const nomB = fila[1]?.toString().trim();
      const clavePrincipal = nomB || codA;

      if (!clavePrincipal) continue;

      const cTotal = totalesPorReceta[clavePrincipal.toUpperCase()] || 0;
      const rend = Number(fila[10]) || 0;
      const cUn = rend > 0 ? (cTotal / rend) : 0;
      const cParc = Number(fila[8]) || 0;

      fila[9] = cTotal;
      fila[11] = cUn;
      fila[13] = cTotal > 0 ? (cParc / cTotal) : 0;

      if (nomB) dictCostosRecetas[nomB.toUpperCase()] = cUn;
      if (codA) dictCostosRecetas[codA.toUpperCase()] = cUn;
    }
  }

  const colD = [], colG = [], colH = [], colI = [], colJ = [], colL = [], colM = [], colN = [];
  
  for (let i = 0; i < matrizRecetas.length; i++) {
    const fila = matrizRecetas[i];
    const codRec = (fila[0] || "").toString().trim().toUpperCase();
    
    const activoEnRecetas = esActivo(fila[12]);
    const activoEnCostos = mapaActivosCostos[codRec] === true;
    const estadoDefinitivo = (activoEnRecetas || activoEnCostos) ? "SI" : "No";

    fila[12] = estadoDefinitivo;

    colD.push([fila[3]]);
    colG.push([fila[6]]);
    colH.push([fila[7]]);
    colI.push([fila[8]]);
    colJ.push([fila[9]]);
    colL.push([fila[11]]);
    colM.push([estadoDefinitivo]);
    colN.push([fila[13]]);
  }

  hojaRecetas.getRange(2, 4, numFilas, 1).setValues(colD);
  hojaRecetas.getRange(2, 7, numFilas, 1).setValues(colG);
  hojaRecetas.getRange(2, 8, numFilas, 1).setValues(colH);
  hojaRecetas.getRange(2, 9, numFilas, 1).setValues(colI);
  hojaRecetas.getRange(2, 10, numFilas, 1).setValues(colJ);
  hojaRecetas.getRange(2, 12, numFilas, 1).setValues(colL);
  hojaRecetas.getRange(2, 13, numFilas, 1).setValues(colM);
  hojaRecetas.getRange(2, 14, numFilas, 1).setValues(colN);

  const alertasRecetas = actualizarColumnasBasicasEnCostos(hojaCostos, matrizRecetas, mapaCostosViejos);
  return alertasRecetas;
}

/**
 * 4. SINCRONIZAR Y ACTUALIZAR RECETAS EN HOJA COSTOS
 * Agrega recetas nuevas de Recetas (A, B, C, L, M) a Costos y sincroniza costos/activos existentes
 */
function actualizarColumnasBasicasEnCostos(hojaCostos, matrizRecetas, mapaCostosViejos) {
  if (!hojaCostos) return [];

  const regexMadres = /^(M|MR|MO|OBR)/i;

  // 1. Mapear recetas únicas calculadas en matrizRecetas
  // Matriz Recetas: Col A [0]=COD, Col B [1]=NOM, Col C [2]=NOMBRE, Col L [11]=C_UN, Col M [12]=ACTIVO
  const mapaRecetasCalculadas = new Map();
  for (let i = 0; i < matrizRecetas.length; i++) {
    const codRec = (matrizRecetas[i][0] || "").toString().trim().toUpperCase();
    if (codRec && !mapaRecetasCalculadas.has(codRec)) {
      mapaRecetasCalculadas.set(codRec, {
        codOriginal: (matrizRecetas[i][0] || "").toString().trim(),
        nomB: (matrizRecetas[i][1] || "").toString().trim(),
        nombreC: (matrizRecetas[i][2] || "").toString().trim(),
        cUn: Number(matrizRecetas[i][11]) || 0,
        activo: matrizRecetas[i][12] || "SI",
        desc: matrizRecetas[i][2] || matrizRecetas[i][1] || matrizRecetas[i][0]
      });
    }
  }

  // 2. Leer estado actual de hoja Costos (8 columnas: COD_REC, NOM, NOMBRE, C_UN, ACTIVO, LAFINUR, CAFE, U.MEDIDA)
  const ultFilaCostos = hojaCostos.getLastRow();
  let datosCostos = [];
  const codigosExistentesEnCostos = new Set();

  if (ultFilaCostos >= 2) {
    datosCostos = hojaCostos.getRange(2, 1, ultFilaCostos - 1, 8).getValues();
    for (let i = 0; i < datosCostos.length; i++) {
      const codRec = (datosCostos[i][0] || "").toString().trim().toUpperCase();
      if (codRec) codigosExistentesEnCostos.add(codRec);
    }
  }

  // 3. Detectar recetas nuevas que están en 'Recetas' pero faltan en 'Costos'
  const filasNuevasParaCostos = [];
  for (const [codUpper, item] of mapaRecetasCalculadas.entries()) {
    if (!codigosExistentesEnCostos.has(codUpper)) {
      const esMadre = regexMadres.test(codUpper);
      const umDefault = esMadre ? "GR" : "UN";

      // Estructura fila: [A: COD_REC, B: NOM_RECETA, C: NOMBRE_RECETA, D: C_UN, E: ACTIVO, F: LAFINUR, G: CAFE, H: U. MEDIDA]
      filasNuevasParaCostos.push([
        item.codOriginal,
        item.nomB,
        item.nombreC,
        item.cUn,
        item.activo,
        "",
        "",
        umDefault
      ]);
      codigosExistentesEnCostos.add(codUpper);
    }
  }

  // Si hay nuevas, se insertan al final de Costos
  if (filasNuevasParaCostos.length > 0) {
    const filaInsercion = hojaCostos.getLastRow() + 1;
    hojaCostos.getRange(filaInsercion, 1, filasNuevasParaCostos.length, 8).setValues(filasNuevasParaCostos);
  }

  // 4. Volver a leer Costos completo para actualizar valores y calcular alertas en memoria
  const totalFilasCostos = hojaCostos.getLastRow() - 1;
  if (totalFilasCostos < 1) return [];

  const matrizActualizada = hojaCostos.getRange(2, 1, totalFilasCostos, 5).getValues();
  const updatesD = [];
  const updatesE = [];
  const alertasRecetas = [];

  for (let i = 0; i < matrizActualizada.length; i++) {
    const codRec = (matrizActualizada[i][0] || "").toString().trim().toUpperCase();
    const cUnActual = Number(matrizActualizada[i][3]) || 0;
    const actActual = matrizActualizada[i][4];

    if (mapaRecetasCalculadas.has(codRec)) {
      const item = mapaRecetasCalculadas.get(codRec);
      updatesD.push([item.cUn]);
      updatesE.push([item.activo]);

      const cViejo = mapaCostosViejos[codRec] || 0;
      if (cViejo > 0 && item.cUn > 0 && esActivo(item.activo)) {
        const variacion = (item.cUn - cViejo) / cViejo;
        if (Math.abs(variacion) >= CONFIG.UMBRAL_RECETAS) {
          alertasRecetas.push({
            cod: codRec,
            nombre: item.desc,
            viejo: cViejo.toFixed(2),
            nuevo: item.cUn.toFixed(2),
            esAumento: variacion > 0,
            porcentaje: (variacion > 0 ? "+" : "") + (variacion * 100).toFixed(1) + "%"
          });
        }
      }
    } else {
      updatesD.push([cUnActual]);
      updatesE.push([actActual]);
    }
  }

  // Escritura por bloque en Costos (Columnas D y E)
  hojaCostos.getRange(2, 4, totalFilasCostos, 1).setValues(updatesD);
  hojaCostos.getRange(2, 5, totalFilasCostos, 1).setValues(updatesE);

  return alertasRecetas;
}

/**
 * 5. RECALCULAR HOJA MAESTRO
 * Incluye override de U. Medida desde la columna H de Costos
 */
function recalcularHojaMaestro() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaInsumos = ss.getSheetByName(CONFIG.LOCAL_HOJA_INSUMOS);
  const hojaCostos = ss.getSheetByName(CONFIG.LOCAL_HOJA_COSTOS);
  const hojaMaestro = ss.getSheetByName(CONFIG.LOCAL_HOJA_MAESTRO);

  if (!hojaMaestro) return;

  const filasMaestro = [];

  // Insumos locales
  const ultFilaInsumos = hojaInsumos.getLastRow();
  if (ultFilaInsumos >= 2) {
    const datosInsumos = hojaInsumos.getRange(2, 1, ultFilaInsumos - 1, 6).getValues();
    for (let i = 0; i < datosInsumos.length; i++) {
      const cod = datosInsumos[i][0]?.toString().trim();
      const desc = datosInsumos[i][1];
      const um = datosInsumos[i][4] || "UN";
      const pun = Number(datosInsumos[i][5]) || 0;

      if (cod) {
        filasMaestro.push(["Insumo", cod, desc, pun, um, "", ""]);
      }
    }
  }

  // Costos: Recetas y Platos
  const ultFilaCostos = hojaCostos.getLastRow();
  if (ultFilaCostos >= 2) {
    // Leemos hasta la columna H (8 columnas: COD_REC, NOM, NOMBRE, C_UN, ACTIVO, LAFINUR, CAFE, U. MEDIDA)
    const datosCostos = hojaCostos.getRange(2, 1, ultFilaCostos - 1, 8).getValues();
    const regexMadres = /^(M|MR|MO|OBR)/i;

    for (let i = 0; i < datosCostos.length; i++) {
      const codRec = datosCostos[i][0]?.toString().trim();
      const nomRec = datosCostos[i][1]?.toString().trim();
      const descRec = datosCostos[i][2];
      const cUn = Number(datosCostos[i][3]) || 0;
      const activo = datosCostos[i][4];
      
      const lafiExacto = limpiarTextoExacto(datosCostos[i][5]);
      const cafeExacto = limpiarTextoExacto(datosCostos[i][6]);
      const umCustom = datosCostos[i][7]?.toString().trim(); // Columna H

      if (!codRec || !esActivo(activo)) continue;

      const codigoIdentificador = nomRec || codRec;
      const tipo = regexMadres.test(codRec) ? "Receta" : "Plato";
      
      // Si la columna H tiene valor explícito, se prioriza ese; caso contrario usa el default
      const umDefecto = (tipo === "Receta") ? "GR" : "UN";
      const um = umCustom ? umCustom : umDefecto;

      filasMaestro.push([tipo, codigoIdentificador, descRec, cUn, um, lafiExacto, cafeExacto]);
    }
  }

  const ultFilaMaestro = hojaMaestro.getLastRow();
  if (ultFilaMaestro >= 2) {
    hojaMaestro.getRange(2, 1, ultFilaMaestro - 1, 7).clearContent();
  }

  if (filasMaestro.length > 0) {
    hojaMaestro.getRange(2, 1, filasMaestro.length, 7).setValues(filasMaestro);
  }
}

/**
 * 6. SNAPSHOT HISTÓRICO HACIA LOGS EXTERNOS
 */
function generarSnapshotMaestro() {
  const ssOrigen = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = ssOrigen.getSheetByName(CONFIG.LOCAL_HOJA_MAESTRO);
  if (!hojaOrigen) return;

  const ultFila = hojaOrigen.getLastRow();
  if (ultFila < 2) return;

  const datosOrigen = hojaOrigen.getRange(2, 1, ultFila - 1, 7).getValues();
  const ahora = new Date();
  const timeZone = Session.getScriptTimeZone();
  const fechaCorte = Utilities.formatDate(ahora, timeZone, "yyyy-MM-dd");
  const semanaId = obtenerSemanaISO(ahora);
  const timestamp = Utilities.formatDate(ahora, timeZone, "yyyy-MM-dd HH:mm:ss");

  const filasLog = [];
  for (let i = 0; i < datosOrigen.length; i++) {
    const segmento = datosOrigen[i][0];
    const cod = datosOrigen[i][1];
    const desc = datosOrigen[i][2];
    const pun = datosOrigen[i][3];
    const um = datosOrigen[i][4];
    const lafinur = datosOrigen[i][5] || "";
    const cafe = datosOrigen[i][6] || "";

    if (!cod || cod.toString().trim() === "") continue;

    filasLog.push([
      segmento,        // Col A: SEGMENTO
      cod,             // Col B: CODIGO
      desc,            // Col C: DESCRIPCION
      pun,             // Col D: P_UNITARIO
      um,              // Col E: U_MEDIDA
      fechaCorte,      // Col F: FECHA_CORTE
      semanaId,        // Col G: SEMANA_ID
      timestamp,       // Col H: TIMESTAMP
      lafinur,         // Col I: LAFINUR (SI / NO / "")
      cafe             // Col J: CAFE (SI / NO / "")
    ]);
  }

  if (filasLog.length === 0) return;

  const ssDestino = SpreadsheetApp.openById(CONFIG.TARGET_LOGS_ID);
  let hojaDestino = ssDestino.getSheetByName(CONFIG.TARGET_LOGS_HOJA);
  
  if (!hojaDestino) {
    hojaDestino = ssDestino.insertSheet(CONFIG.TARGET_LOGS_HOJA);
    const cabeceras = [
      "SEGMENTO", "CODIGO", "DESCRIPCION", "P_UNITARIO", "U_MEDIDA", 
      "FECHA_CORTE", "SEMANA_ID", "TIMESTAMP", "LAFINUR", "CAFE"
    ];
    hojaDestino.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras]).setFontWeight("bold").setBackground("#d9ead3");
    hojaDestino.setFrozenRows(1);
  } else {
    if (hojaDestino.getLastColumn() < 10) {
      hojaDestino.getRange(1, 9, 1, 2).setValues([["LAFINUR", "CAFE"]]).setFontWeight("bold").setBackground("#d9ead3");
    }
  }

  const ultFilaDestino = hojaDestino.getLastRow();
  hojaDestino.getRange(ultFilaDestino + 1, 1, filasLog.length, 10).setValues(filasLog);

  try {
    SpreadsheetApp.getUi().alert("📸 Snapshot Exitoso", `Se enviaron ${filasLog.length} filas a Logs respetando SI, NO y celdas vacías.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch(e) {}
}

function obtenerSemanaISO(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return `${date.getUTCFullYear()}-W${weekNo < 10 ? '0' + weekNo : weekNo}`;
}

/**
 * 7. REPORTE CONSOLIDADO POR EMAIL
 */
function enviarReporteConsolidado(alertasInsumos = [], alertasRecetas = []) {
  if (alertasInsumos.length === 0 && alertasRecetas.length === 0) return;

  const asunto = `⚠️ ALERTA: ${alertasInsumos.length} Insumos (±15%) | ${alertasRecetas.length} Recetas Afectadas (±10%)`;
  
  let cuerpoHtml = `
    <div style="font-family: Arial, sans-serif; color: #333; line-height: 1.5;">
      <h2 style="color: #2c3e50; border-bottom: 2px solid #ecf0f1; padding-bottom: 8px;">Reporte de Variaciones Críticas de Costos</h2>`;

  if (alertasInsumos.length > 0) {
    cuerpoHtml += `
      <h3 style="color: #c0392b; margin-top: 20px;">📦 Insumos con Variación ≥ ±15% (${alertasInsumos.length})</h3>
      <table style="border-collapse: collapse; width: 100%; max-width: 650px; font-size: 13px; margin-bottom: 25px;">
        <thead>
          <tr style="background-color: #f2f2f2; text-align: left;">
            <th style="padding: 8px; border: 1px solid #ddd;">Código</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Insumo</th>
            <th style="padding: 8px; border: 1px solid #ddd;">P. Anterior</th>
            <th style="padding: 8px; border: 1px solid #ddd;">P. Nuevo</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Variación</th>
          </tr>
        </thead>
        <tbody>`;

    alertasInsumos.forEach(a => {
      const color = a.esAumento ? "#c0392b" : "#27ae60";
      const bg = a.esAumento ? "#fde8e8" : "#e8f8f0";
      cuerpoHtml += `
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>${a.cod}</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${a.nombre}</td>
          <td style="padding: 8px; border: 1px solid #ddd;">$${a.viejo}</td>
          <td style="padding: 8px; border: 1px solid #ddd; color: ${color};"><strong>$${a.nuevo}</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd; background-color: ${bg}; font-weight: bold; color: ${color};">${a.porcentaje}</td>
        </tr>`;
    });
    cuerpoHtml += `</tbody></table>`;
  }

  if (alertasRecetas.length > 0) {
    cuerpoHtml += `
      <h3 style="color: #d35400; margin-top: 20px;">🍽️ Platos / Recetas Afectadas con Variación ≥ ±10% (${alertasRecetas.length})</h3>
      <table style="border-collapse: collapse; width: 100%; max-width: 650px; font-size: 13px;">
        <thead>
          <tr style="background-color: #f2f2f2; text-align: left;">
            <th style="padding: 8px; border: 1px solid #ddd;">Código</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Plato / Receta</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Costo Anterior</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Costo Nuevo</th>
            <th style="padding: 8px; border: 1px solid #ddd;">Variación</th>
          </tr>
        </thead>
        <tbody>`;

    alertasRecetas.forEach(r => {
      const color = r.esAumento ? "#c0392b" : "#27ae60";
      const bg = r.esAumento ? "#fde8e8" : "#e8f8f0";
      cuerpoHtml += `
        <tr>
          <td style="padding: 8px; border: 1px solid #ddd;"><strong>${r.cod}</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd;">${r.nombre}</td>
          <td style="padding: 8px; border: 1px solid #ddd;">$${r.viejo}</td>
          <td style="padding: 8px; border: 1px solid #ddd; color: ${color};"><strong>$${r.nuevo}</strong></td>
          <td style="padding: 8px; border: 1px solid #ddd; background-color: ${bg}; font-weight: bold; color: ${color};">${r.porcentaje}</td>
        </tr>`;
    });
    cuerpoHtml += `</tbody></table>`;
  }

  cuerpoHtml += `</div>`;

  MailApp.sendEmail({
    to: CONFIG.EMAIL_DESTINO,
    subject: asunto,
    htmlBody: cuerpoHtml
  });
}

/**
 * 8. CARTEL EN PANTALLA (UI)
 */
function mostrarAlertaUI(alertasInsumos = [], alertasRecetas = []) {
  let msg = `Resumen de Alertas:\n\n`;

  if (alertasInsumos.length > 0) {
    msg += `📦 Insumos (±15%): ${alertasInsumos.length}\n`;
    alertasInsumos.slice(0, 6).forEach(a => {
      msg += ` • [${a.cod}] ${a.nombre}: $${a.viejo} ➔ $${a.nuevo} (${a.porcentaje})\n`;
    });
    if (alertasInsumos.length > 6) {
      msg += ` ...y ${alertasInsumos.length - 6} insumos más.\n`;
    }
  }

  if (alertasRecetas.length > 0) {
    msg += `\n🍽️ Recetas Afectadas (±10%): ${alertasRecetas.length}\n`;
    alertasRecetas.slice(0, 6).forEach(r => {
      msg += ` • [${r.cod}] ${r.nombre}: $${r.viejo} ➔ $${r.nuevo} (${r.porcentaje})\n`;
    });
    if (alertasRecetas.length > 6) {
      msg += ` ...y ${alertasRecetas.length - 6} recetas más.\n`;
    }
  }

  msg += `\n✉️ Se envió el reporte detallado a ${CONFIG.EMAIL_DESTINO}`;

  try {
    SpreadsheetApp.getUi().alert("⚠️ Alerta de Costos Críticos", msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {}
}

// ==========================================================================
// 🟢 INICIO MÓDULO WHATSAPP (BORRABLE)
// Si querés quitarlo, borrá desde acá hasta el final del archivo.
// ==========================================================================

// Credenciales en Script Properties (Configuración del proyecto > Propiedades del script):
//   WPP_API_KEY, WPP_NUMERO_DESTINO
function getConfigWpp_() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("WPP_API_KEY");
  const numero = props.getProperty("WPP_NUMERO_DESTINO");
  return {
    ACTIVO: !!(apiKey && numero),
    API_KEY: apiKey,
    NUMERO_DESTINO: numero,
    URL_ENDPOINT: "https://api.callmebot.com/whatsapp.php"
  };
}

function enviarMensajeWhatsApp(mensajeTexto) {
  const CONFIG_WPP = getConfigWpp_();
  if (!CONFIG_WPP.ACTIVO) {
    Logger.log("WhatsApp desactivado: faltan WPP_API_KEY / WPP_NUMERO_DESTINO en Script Properties");
    return;
  }

  try {
    const url = `${CONFIG_WPP.URL_ENDPOINT}?phone=${encodeURIComponent(CONFIG_WPP.NUMERO_DESTINO)}&text=${encodeURIComponent(mensajeTexto)}&apikey=${encodeURIComponent(CONFIG_WPP.API_KEY)}`;
    
    const opciones = {
      method: "get",
      muteHttpExceptions: true
    };

    const respuesta = UrlFetchApp.fetch(url, opciones);
    Logger.log("Respuesta WhatsApp: " + respuesta.getContentText());
  } catch (err) {
    Logger.log("Error al despachar WhatsApp: " + err.toString());
  }
}

function dispararAlertaWhatsApp(alertasInsumos = [], alertasRecetas = []) {
  if (alertasInsumos.length === 0 && alertasRecetas.length === 0) return;

  let texto = `⚠️ *ALERTA DE COSTOS MISHIGUENE*\n\n`;

  if (alertasInsumos.length > 0) {
    texto += `📦 *Insumos con variación (±15%):* ${alertasInsumos.length}\n`;
    alertasInsumos.slice(0, 5).forEach(a => {
      texto += `• ${a.nombre}: $${a.viejo} ➔ $${a.nuevo} (${a.porcentaje})\n`;
    });
    if (alertasInsumos.length > 5) {
      texto += `_...y ${alertasInsumos.length - 5} más._\n`;
    }
    texto += `\n`;
  }

  if (alertasRecetas.length > 0) {
    texto += `🍽️ *Recetas afectadas (±10%):* ${alertasRecetas.length}\n`;
    alertasRecetas.slice(0, 5).forEach(r => {
      texto += `• ${r.nombre}: $${r.viejo} ➔ $${r.nuevo} (${r.porcentaje})\n`;
    });
    if (alertasRecetas.length > 5) {
      texto += `_...y ${alertasRecetas.length - 5} más._\n`;
    }
  }

  texto += `\n✉️ El detalle completo se envió por correo.`;

  enviarMensajeWhatsApp(texto);
}
// ==========================================================================
// 🔴 FIN MÓDULO WHATSAPP
// ==========================================================================