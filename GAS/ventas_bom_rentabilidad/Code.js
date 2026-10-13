/**
 * SISTEMA INTEGRAL MISHIGUENE - VERSIÓN FINAL 3.0 (ANTI-DUPLICADOS)
 * Autor: AI Solution para Mishiguene
 * 
 * 1. BOM: Explosión de Insumos (Kilos/Unidades) -> REPORTE_TEORICO_SEPT
 * 2. RENTABILIDAD: Ventas vs Costos ($) -> RENTABILIDAD_POR_PLATO
 */

const GLOBAL_CONFIG = {
  ID_RECETAS: '1u7QSi7JyCM1J6-DNKSWcha2qIyF2Ybxxp9SVqdNpyfY',
  HOJA_MAESTRO: 'Maestro',
  HOJA_VENTAS: 'Ventas',
  // Nombres de hojas de salida
  OUT_REPORTE_TEORICO: 'REPORTE_TEORICO_SEPT',
  OUT_TRAZA: 'DETALLE_TRAZA_COMPLETA',
  OUT_FALTANTES: 'FALTANTES_PARA_REVISAR',
  OUT_RENTABILIDAD: 'RENTABILIDAD_POR_PLATO'
};

// SKUs de panadería que se compran a proveedores (no se explotan)
const EXCLUIDOS = ["PITA", "BAMAPOLA", "BSESAMO", "BCEBOLLA", "BKUMMEL", "PLETZ", "CROIL", "RDCL", "PDM", "PDCAF", "BAGEL", "CROISSANT"];

/**
 * FUNCIÓN 1: EXPLOSIÓN DE INSUMOS (KG/UN)
 * Genera el consumo teórico de materia prima.
 */
function generarBOM_Septiembre_Final() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const datos = cargarDatosMishiguene();
  
  if (!datos) return;

  const totalInsumosFinales = {}; 
  const trazaDetallada = [];
  const faltantes = [];

  datos.rawVentas.slice(1).forEach(row => {
    const nomVenta = row[0];
    const cantVenta = parseFloat(row[1]);
    if (!nomVenta || !cantVenta) return;

    const match = buscarMatchManual(nomVenta, datos.maestroMap);
    if (match) {
      explodirRecursivo(match.sku, cantVenta, nomVenta, cantVenta, datos.recetasMap, datos.insumosMap, totalInsumosFinales, trazaDetallada, 0, "", new Set());
    } else {
      faltantes.push([nomVenta, cantVenta, "[NO ENCONTRADO EN MAESTRO]"]);
    }
  });

  escribirResultadosBOM(ss, totalInsumosFinales, trazaDetallada, faltantes, datos.insumosMap);
}

/**
 * FUNCIÓN 2: ANÁLISIS DE RENTABILIDAD ($)
 * Genera el informe de Margen Bruto por plato.
 */
function ejecutarAnalisisRentabilidad() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const datos = cargarDatosMishiguene();
  
  if (!datos) return;

  // Crear Mapa de Costos Unitarios Finales (C_UN de la hoja Recetas del archivo remoto)
  const costosFinalesMap = {};
  datos.rawRecetas.forEach(row => {
    const skuA = String(row[0] || "").trim().toUpperCase();
    const skuB = String(row[1] || "").trim().toUpperCase();
    const costoElab = parseNumMishi(row[11]); // Col L (C_UN)
    
    if (skuA && skuA !== "COD_RECE" && costoElab > 0 && !costosFinalesMap[skuA]) costosFinalesMap[skuA] = costoElab;
    if (skuB && skuB !== "NOM_RECETA" && costoElab > 0 && !costosFinalesMap[skuB]) costosFinalesMap[skuB] = costoElab;
  });

  const resultadosRentab = [];
  let tVenta = 0, tCosto = 0;

  datos.rawVentas.slice(1).forEach(row => {
    const nomVenta = row[0];
    const unidades = parseNumMishi(row[1]);
    const ventaTotal = parseNumMishi(row[2]); // Columna C de Ventas
    if (!nomVenta || unidades === 0) return;

    const match = buscarMatchManual(nomVenta, datos.maestroMap);
    const sku = match ? match.sku : null;
    
    // Si no hay receta, busca precio base en Insumos
    const costoUnit = sku ? (costosFinalesMap[sku] || datos.insumosMap[sku]?.precio || 0) : 0;
    const costoTotalFab = unidades * costoUnit;
    const margen = ventaTotal - costoTotalFab;
    const porc = ventaTotal > 0 ? (margen / ventaTotal) : 0;

    tVenta += ventaTotal;
    tCosto += costoTotalFab;

    resultadosRentab.push([nomVenta, sku || "SIN SKU", unidades, ventaTotal, costoUnit, costoTotalFab, margen, porc]);
  });

  escribirHojaRentab(ss, resultadosRentab, tVenta, tCosto);
}

// --- MOTOR DE CARGA (CORREGIDO PARA EVITAR DUPLICADOS) ---
function cargarDatosMishiguene() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shV = ss.getSheetByName(GLOBAL_CONFIG.HOJA_VENTAS);
  const shM = ss.getSheetByName(GLOBAL_CONFIG.HOJA_MAESTRO);
  
  if (!shV || !shM) {
    SpreadsheetApp.getUi().alert("Error: Faltan las hojas 'Ventas' o 'Maestro'.");
    return null;
  }

  const db = SpreadsheetApp.openById(GLOBAL_CONFIG.ID_RECETAS);
  const rawVentas = shV.getDataRange().getValues();
  const rawMaestro = shM.getDataRange().getValues();
  const rawRecetas = db.getSheetByName("Recetas").getDataRange().getValues();
  const rawInsumos = db.getSheetByName("Insumos").getDataRange().getValues();

  // 1. Maestro
  const maestroMap = {};
  rawMaestro.slice(1).forEach(row => {
    const n = normMishi(row[2]);
    if (row[0] && n) maestroMap[n] = { sku: String(row[0]).trim().toUpperCase() };
  });

  // 2. Recetas (Sin duplicar ingredientes si A y B son iguales)
  const recetasMap = {};
  rawRecetas.slice(1).forEach(row => {
    const cA = String(row[0] || "").trim().toUpperCase();
    const cB = String(row[1] || "").trim().toUpperCase();
    const ing = String(row[3] || "").trim().toUpperCase();
    if (!ing || ing === "UNDEFINED") return;

    const data = { cod: ing, nombre: row[4], cant: parseFloat(row[5]) || 0 };
    const rend = parseFloat(row[10]) || 1;

    // Usamos un Set para no procesar el mismo código dos veces en una misma fila
    let keys = new Set();
    if (cA && cA !== "") keys.add(cA);
    if (cB && cB !== "") keys.add(cB);

    keys.forEach(k => {
      if (!recetasMap[k]) recetasMap[k] = { rend: rend, ingredientes: [] };
      recetasMap[k].ingredientes.push(data);
    });
  });

  // 3. Insumos
  const insumosMap = {};
  rawInsumos.slice(1).forEach(row => {
    const sku = String(row[0] || "").trim().toUpperCase();
    if (sku) insumosMap[sku] = { nombre: row[1], unidad: row[4], precio: parseNumMishi(row[5]) };
  });

  return {rawVentas, rawRecetas, maestroMap, recetasMap, insumosMap};
}

// --- LÓGICA RECURSIVA ---
function explodirRecursivo(cod, cantReq, vOrig, cVenta, recetasMap, insumosMap, acumulado, traza, nivel, path, visitados) {
  if (visitados.has(cod) || nivel > 10) return;
  
  const currentPath = path === "" ? cod : path + " > " + cod;
  const nuevosVisitados = new Set(visitados); nuevosVisitados.add(cod);
  
  if (recetasMap[cod] && !EXCLUIDOS.includes(cod)) {
    const receta = recetasMap[cod];
    const factor = cantReq / (receta.rend || 1);
    receta.ingredientes.forEach(ing => {
      explodirRecursivo(ing.cod, ing.cant * factor, vOrig, cVenta, recetasMap, insumosMap, acumulado, traza, nivel + 1, currentPath, nuevosVisitados);
    });
  } else {
    acumulado[cod] = (acumulado[cod] || 0) + cantReq;
    const info = insumosMap[cod];
    traza.push([vOrig, cVenta, nivel, currentPath, cod, info ? info.nombre : "DESCONOCIDO", cantReq, info ? info.precio : 0, cantReq * (info ? info.precio : 0)]);
  }
}

// --- UTILIDADES ---
function normMishi(t) {
  return t ? String(t).toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^A-Z0-9]/g, "").trim() : "";
}

function parseNumMishi(v) {
  if (typeof v === 'number') return v;
  let n = parseFloat(String(v || "").replace(/[^0-9,.-]/g, "").replace(",", "."));
  return isNaN(n) ? 0 : n;
}

function buscarMatchManual(nombre, mapa) {
  const n = normMishi(nombre);
  if (mapa[n]) return mapa[n];
  for (let key in mapa) { if (n.includes(key) || key.includes(n)) return mapa[key]; }
  return null;
}

function escribirResultadosBOM(ss, consolidado, traza, faltantes, insumosMap) {
  let shT = actualizarHojaMishi(ss, GLOBAL_CONFIG.OUT_TRAZA);
  shT.getRange(1,1,1,9).setValues([["VENTA ORIG", "CANT", "NIVEL", "JERARQUÍA", "SKU FINAL", "NOMBRE INSUMO", "CANT. TOTAL", "PRECIO UN", "SUBTOTAL $$"]]).setFontWeight("bold").setBackground("#dcfce7");
  if (traza.length > 0) shT.getRange(2,1, traza.length, 9).setValues(traza);

  let shC = actualizarHojaMishi(ss, GLOBAL_CONFIG.OUT_REPORTE_TEORICO);
  const rows = [];
  for (let sku in consolidado) {
    const info = insumosMap[sku];
    rows.push([sku, info ? info.nombre : "REVISAR SKU", info ? info.unidad : "---", consolidado[sku], info ? info.precio : 0, consolidado[sku] * (info ? info.precio : 0)]);
  }
  shC.getRange(1,1,1,6).setValues([["SKU", "INSUMO PURO", "UNIDAD", "CANT. TEORICA", "COSTO UN.", "TOTAL $$"]]).setFontWeight("bold").setBackground("#dcfce7");
  if (rows.length > 0) shC.getRange(2,1, rows.length, 6).setValues(rows);
  
  let shF = actualizarHojaMishi(ss, GLOBAL_CONFIG.OUT_FALTANTES);
  shF.getRange(1,1,1,3).setValues([["PRODUCTO", "CANTIDAD", "ESTADO"]]).setFontWeight("bold").setBackground("#fee2e2");
  if (faltantes.length > 0) shF.getRange(2,1, faltantes.length, 3).setValues(faltantes);
  
  SpreadsheetApp.getUi().alert("BOM Actualizado correctamente.");
}

function escribirHojaRentab(ss, datos, v, c) {
  let sh = actualizarHojaMishi(ss, GLOBAL_CONFIG.OUT_RENTABILIDAD);
  sh.getRange(1,1,1,8).setValues([["ARTÍCULO", "SKU", "UNIDADES", "VENTA TOTAL", "COSTO UNIT FAB", "COSTO TOTAL FAB", "MARGEN ($)", "MARGEN %"]]).setFontWeight("bold").setBackground("#cfe2f3");
  if (datos.length > 0) {
    sh.getRange(2,1,datos.length,8).setValues(datos);
    sh.getRange(2,4,datos.length,4).setNumberFormat("$#,##0");
    sh.getRange(2,8,datos.length,1).setNumberFormat("0.00%");
    sh.getRange(datos.length + 3, 3, 1, 6).setValues([["TOTALES:", v, "", c, v-c, v > 0 ? (v-c)/v : 0]]).setFontWeight("bold");
  }
  SpreadsheetApp.getUi().alert("Rentabilidad Actualizada.");
}

function actualizarHojaMishi(ss, n) {
  let sh = ss.getSheetByName(n);
  if (sh) sh.clear(); else sh = ss.insertSheet(n);
  return sh;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🚀 CONTROL MISHIGUENE')
    .addItem('📦 Calcular BOM (Ingredientes)', 'generarBOM_Septiembre_Final')
    .addItem('💰 Calcular Rentabilidad ($)', 'ejecutarAnalisisRentabilidad')
    .addToUi();
}