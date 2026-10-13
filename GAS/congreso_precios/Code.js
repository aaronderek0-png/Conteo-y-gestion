// ==========================================================================
// CONFIGURACIÓN GLOBAL DEL SISTEMA (v17.0 - CRUCE ESTRICTO POR SKU X2:X)
// ==========================================================================
const ID_FACTURAS = "1udLEQpQ29avK4c6fNnQig-w02xIV82IgLe85yVJXaOY";

const CONFIG_PRECIOS = {
  HOJA_GESTION_INV: "Inventario",
  FILA_INICIO_INV: 3,
  
  COL_CODIGO_INV: 4,          // Columna D (Código / SKU)
  COL_PRODUCTO_INV: 5,        // Columna E (Nombre del producto)
  COL_PRECIO_VIEJO_INV: 6,    // Columna F (Precio anterior / Base de respaldo)
  COL_PRECIO_NUEVO_INV: 7,    // Columna G (Precio unitario actual)
  COL_ORIGEN_PRECIO_INV: 8,   // Columna H (Auditoría: Fecha / Histórico)
  
  COL_FINAL_INV: 12,          // Columna L (Conteo final)
  COL_VALORIZADO_INV: 13,     // Columna M (FINAL * PRECIO G)
  COL_PORCENTAJE_INV: 14      // Columna N (% sobre el total valorizado)
};

// Índices base 0 en la hoja Facturas remota
const FACTURAS_COLS = {
  COL_FECHA_FACTURA: 1,      // Columna B (Fecha Factura)
  COL_FECHA_INGRESO: 2,      // Columna C (Fecha Ingreso - Prioridad)
  COL_PRODUCTO: 4,           // Columna E (Nombre en factura)
  COL_PRECIO_UN_INV: 14,     // Columna O (PRECIO UN. INVENTARI)
  COL_CODIGO_SKU: 23         // Columna X (CODIGO SKU)
};

// ──────────────────────────────────────────────────────────────────────────────
// FUNCIÓN AUXILIAR: DETECTA EL MES Y AÑO AUTOMÁTICAMENTE DESDE EL TÍTULO
// ──────────────────────────────────────────────────────────────────────────────
function obtenerPeriodoDesdeNombre() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaLocal = ss.getSheetByName(CONFIG_PRECIOS.HOJA_GESTION_INV) || ss.getActiveSheet();
  const textoCompleto = (ss.getName() + " " + hojaLocal.getName()).toLowerCase();
  
  const meses = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12
  };
  
  let mesEncontrado = null;
  for (let nombreMes in meses) {
    if (textoCompleto.includes(nombreMes)) {
      mesEncontrado = meses[nombreMes];
      break;
    }
  }
  
  let anoEncontrado = 2026; 
  const matchAno = textoCompleto.match(/(20\d{2})/);
  if (matchAno) {
    anoEncontrado = parseInt(matchAno[1]);
  }
  
  if (!mesEncontrado) {
    mesEncontrado = 9; // Fallback Septiembre
  }
  
  return { mes: mesEncontrado, ano: anoEncontrado, texto: textoCompleto };
}

function parseFecha(fechaIngresoRaw, fechaFacturaRaw, anoReferencia) {
  let raw = (fechaIngresoRaw !== "" && fechaIngresoRaw !== null && fechaIngresoRaw !== undefined)
            ? fechaIngresoRaw
            : fechaFacturaRaw;

  if (!raw) return null;

  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return new Date(raw.getFullYear(), raw.getMonth(), raw.getDate());
  }

  const str = String(raw).trim();
  const partes = str.split(/[\/\-]/);
  if (partes.length >= 2) {
    let dia = parseInt(partes[0], 10);
    let mes = parseInt(partes[1], 10) - 1;
    let anio = partes[2] ? parseInt(partes[2], 10) : anoReferencia;
    if (anio < 100) anio += 2000;
    let f = new Date(anio, mes, dia);
    if (!isNaN(f.getTime())) return f;
  }

  let f = new Date(raw);
  return !isNaN(f.getTime()) ? f : null;
}

function parseNum(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const clean = val.toString().trim().replace(/\$/g, "").replace(/\s/g, "").replace(/\./g, "").replace(",", ".");
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
}

// ──────────────────────────────────────────────────────────────────────────────
// SCRIPT PRINCIPAL: CRUCE ESTRICTO SKU (D vs X) Y AUDITORÍA
// ──────────────────────────────────────────────────────────────────────────────
function actualizarUltimoPrecioMensual() {
  const ssGestion = SpreadsheetApp.getActiveSpreadsheet();
  const periodo = obtenerPeriodoDesdeNombre();

  let ssFacturas;
  try {
    ssFacturas = SpreadsheetApp.openById(ID_FACTURAS);
  } catch(e) {
    SpreadsheetApp.getUi().alert("❌ No se pudo abrir el archivo remoto de Facturas: " + e.message);
    return;
  }

  const hojaFacturas = ssFacturas.getSheetByName("Facturas") || ssFacturas.getSheetByName("9 Facturas");
  if (!hojaFacturas) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la pestaña 'Facturas' en el archivo remoto.");
    return;
  }

  const uFilaFacturas = hojaFacturas.getLastRow();
  if (uFilaFacturas < 2) {
    SpreadsheetApp.getUi().alert("❌ La hoja Facturas remota no tiene datos.");
    return;
  }

  // Traer columnas de Facturas (hasta Columna X = 24 cols)
  const datosFacturas = hojaFacturas.getRange(2, 1, uFilaFacturas - 1, 24).getValues();

  // Ventana de búsqueda: mes actual y hasta 3 meses atrás
  const mesesValidos = [];
  for (let m = 0; m < 4; m++) {
    let mesTarget = periodo.mes - m;
    let anoTarget = periodo.ano;
    if (mesTarget <= 0) {
      mesTarget += 12;
      anoTarget -= 1;
    }
    mesesValidos.push({ mes: mesTarget, ano: anoTarget, nivel: m });
  }

  // Mapa de precios por CÓDIGO SKU (Columna X)
  // Formato: mapaPrecios[SKU] = { precio, fecha, nivel, mFactura, idx }
  const mapaPreciosPorSku = new Map();

  for (let i = 0; i < datosFacturas.length; i++) {
    const fila = datosFacturas[i];
    
    // 1. Obtener CÓDIGO de Columna X (índice 23)
    let codigoSku = String(fila[FACTURAS_COLS.COL_CODIGO_SKU] || "").trim().toUpperCase();
    if (!codigoSku || codigoSku === "SIN SKU" || codigoSku.includes("#N/A") || codigoSku.includes("ERROR")) {
      continue;
    }

    // 2. Obtener PRECIO UNITARIO INVENTARIO de Columna O (índice 14)
    const precioUnitario = parseNum(fila[FACTURAS_COLS.COL_PRECIO_UN_INV]);
    if (precioUnitario <= 0) continue;

    // 3. Obtener FECHA (Ingreso o Factura)
    const fechaObj = parseFecha(fila[FACTURAS_COLS.COL_FECHA_INGRESO], fila[FACTURAS_COLS.COL_FECHA_FACTURA], periodo.ano);
    if (!fechaObj) continue;

    const mFactura = fechaObj.getMonth() + 1;
    const aFactura = fechaObj.getFullYear();

    const combo = mesesValidos.find(p => p.mes === mFactura && p.ano === aFactura);
    if (!combo) continue;

    // Guardar el más reciente por fecha / posición
    if (!mapaPreciosPorSku.has(codigoSku)) {
      mapaPreciosPorSku.set(codigoSku, {
        precio: precioUnitario,
        fecha: fechaObj,
        nivel: combo.nivel,
        mes: mFactura,
        idx: i
      });
    } else {
      const prev = mapaPreciosPorSku.get(codigoSku);
      if (fechaObj > prev.fecha || (fechaObj.getTime() === prev.fecha.getTime() && i >= prev.idx)) {
        mapaPreciosPorSku.set(codigoSku, {
          precio: precioUnitario,
          fecha: fechaObj,
          nivel: combo.nivel,
          mes: mFactura,
          idx: i
        });
      }
    }
  }

  // 4. Procesar hoja de Inventario local
  const hojaGestion = ssGestion.getSheetByName(CONFIG_PRECIOS.HOJA_GESTION_INV);
  if (!hojaGestion) {
    SpreadsheetApp.getUi().alert(`❌ No encontré la pestaña '${CONFIG_PRECIOS.HOJA_GESTION_INV}'`);
    return;
  }

  const uFilaGestion = hojaGestion.getLastRow();
  if (uFilaGestion < CONFIG_PRECIOS.FILA_INICIO_INV) {
    SpreadsheetApp.getUi().alert("La hoja de Inventario no tiene productos para procesar.");
    return;
  }

  const totalProductos = uFilaGestion - CONFIG_PRECIOS.FILA_INICIO_INV + 1;
  
  const matrizCodigos = hojaGestion.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_CODIGO_INV, totalProductos, 1).getValues();
  const matrizPreciosViejos = hojaGestion.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_PRECIO_VIEJO_INV, totalProductos, 1).getValues();

  const matrizNuevosPrecios = [];
  const matrizAuditoria = [];

  let countMesActual = 0;
  let countHistorico = 0;
  let countMantenidos = 0;

  const nombresMesesTexto = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

  for (let i = 0; i < totalProductos; i++) {
    const codLocal = String(matrizCodigos[i][0] || "").trim().toUpperCase();
    const precioBase = matrizPreciosViejos[i][0];

    if (!codLocal) {
      matrizNuevosPrecios.push([precioBase]);
      matrizAuditoria.push([""]);
      continue;
    }

    if (mapaPreciosPorSku.has(codLocal)) {
      const match = mapaPreciosPorSku.get(codLocal);
      matrizNuevosPrecios.push([match.precio]);
      
      const dia = String(match.fecha.getDate()).padStart(2, "0");
      const mes = String(match.fecha.getMonth() + 1).padStart(2, "0");
      const ano = match.fecha.getFullYear();

      if (match.nivel === 0) {
        matrizAuditoria.push([`Factura: ${dia}/${mes}/${ano}`]);
        countMesActual++;
      } else {
        matrizAuditoria.push([`Histórico: ${nombresMesesTexto[match.mes]} (${dia}/${mes})`]);
        countHistorico++;
      }
    } else {
      matrizNuevosPrecios.push([precioBase]);
      matrizAuditoria.push(["No encontrado - Se mantuvo precio anterior"]);
      countMantenidos++;
    }
  }

  // 5. Inyección en columnas G (Precio) y H (Auditoría)
  hojaGestion.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_PRECIO_NUEVO_INV, totalProductos, 1).setValues(matrizNuevosPrecios);
  hojaGestion.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_ORIGEN_PRECIO_INV, totalProductos, 1).setValues(matrizAuditoria);

  SpreadsheetApp.flush();

  // 6. Recalcular M ($) y N (%)
  calcularValorizadoYPorcentajes(true);

  SpreadsheetApp.getActiveSpreadsheet().toast(
    `✅ Listo: ${countMesActual} de este mes, ${countHistorico} históricos, ${countMantenidos} mantenidos.`,
    "Sincronización SKU",
    6
  );
}

// ──────────────────────────────────────────────────────────────────────────────
// 2. FUNCIÓN DE RECALCULAR COLUMNAS M (VALORIZADO) Y N (% VALORIZADO)
// ──────────────────────────────────────────────────────────────────────────────
function calcularValorizadoYPorcentajes(silencioso = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG_PRECIOS.HOJA_GESTION_INV);
  if (!hoja) return;

  const uFila = hoja.getLastRow();
  if (uFila < CONFIG_PRECIOS.FILA_INICIO_INV) return;

  const totalFilas = uFila - CONFIG_PRECIOS.FILA_INICIO_INV + 1;

  const precios = hoja.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_PRECIO_NUEVO_INV, totalFilas, 1).getValues();
  const finales = hoja.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_FINAL_INV, totalFilas, 1).getValues();

  const valorizados = [];
  let sumaTotalGastado = 0;

  for (let i = 0; i < totalFilas; i++) {
    const p = parseNum(precios[i][0]);
    const cant = parseNum(finales[i][0]);
    const val = (p > 0 && cant > 0) ? (cant * p) : 0;
    
    valorizados.push(val);
    sumaTotalGastado += val;
  }

  const matrizM = [];
  const matrizN = [];

  for (let i = 0; i < totalFilas; i++) {
    const val = valorizados[i];
    if (val > 0) {
      matrizM.push([val]);
      const pct = (sumaTotalGastado > 0) ? (val / sumaTotalGastado) : 0;
      matrizN.push([pct]);
    } else {
      matrizM.push([0]);
      matrizN.push([0]);
    }
  }

  const rangoM = hoja.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_VALORIZADO_INV, totalFilas, 1);
  const rangoN = hoja.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_PORCENTAJE_INV, totalFilas, 1);

  rangoM.setValues(matrizM);
  rangoN.setValues(matrizN);

  rangoM.setNumberFormat("$ #,##0.00");
  rangoN.setNumberFormat("0.00%");

  SpreadsheetApp.flush();

  if (!silencioso) {
    SpreadsheetApp.getUi().alert(
      "✅ RECÁLCULO COMPLETADO",
      `Total valorizado en stock: $ ${sumaTotalGastado.toLocaleString("es-AR", { minimumFractionDigits: 2 })}\n\nColumnas M ($) y N (%) actualizadas.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ──────────────────────────────────────────────────────────────────────────────
// MENÚ DE ACCESO SUPERIOR
// ──────────────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Automatización')
    .addItem('1️⃣ Actualizar Precios por Código SKU (Col X) y Valorizar', 'actualizarUltimoPrecioMensual')
    .addItem('2️⃣ Recalcular solo Valorizado (M) y % (N)', 'calcularValorizadoYPorcentajes')
    .addToUi();
}