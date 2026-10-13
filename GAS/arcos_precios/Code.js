// ==========================================================================
// CONFIGURACIÓN GLOBAL - GESTIÓN ARCOS (v18.0 - CRUCE ESTRICTO SKU X2:X)
// ==========================================================================
const ID_FACTURAS_ARCOS = "1frh93qB8Od5mrtAfMuObKwRW93p7kev6DNgmERqz6ho";

const CONFIG_PRECIOS = {
  HOJA_GESTION_INV: "Inventario",
  FILA_INICIO_INV: 3,
  
  COL_MEDIDA_INV: 3,          // Columna C (Medida)
  COL_CODIGO_INV: 4,          // Columna D (Código / SKU)
  COL_PRODUCTO_INV: 5,        // Columna E (Nombre del producto)
  COL_PRECIO_VIEJO_INV: 6,    // Columna F (Precio unitario antes / Base)
  COL_PRECIO_NUEVO_INV: 7,    // Columna G (Precio unitario actual)
  COL_FINAL_INV: 16,          // Columna P (Conteo FINAL)
  COL_VALORIZADO_INV: 17,     // Columna Q (P * G)
  COL_PORC_VALORIZADO_INV: 18,// Columna R (% del total valorizado)
  COL_ORIGEN_PRECIO_INV: 19   // Columna S (Auditoría / Fecha ultima factura)
};

// Índices base 0 para la hoja Facturas remota
const FACTURAS_COLS = {
  COL_FECHA_FACTURA: 1,       // Columna B (Fecha Factura)
  COL_FECHA_INGRESO: 2,       // Columna C (Fecha Ingreso - Prioridad)
  COL_PRODUCTO: 4,            // Columna E (Nombre en factura)
  COL_PRECIO_UN_INV: 14,      // Columna O (PRECIO UN. INVENTARI)
  COL_CODIGO_SKU: 23          // Columna X (CODIGO SKU)
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
    anoEncontrado = parseInt(matchAno[1], 10);
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

function norm(texto) {
  if (texto === null || texto === undefined) return "";
  return texto.toString().trim().toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9 ]/g, "")
    .replace(/\s+/g, " ");
}

function parseNum(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const clean = val.toString().trim().replace("$", "").replace(/\./g, "").replace(",", ".");
  const num = parseFloat(clean);
  return isNaN(num) ? 0 : num;
}

// ──────────────────────────────────────────────────────────────────────────────
// SCRIPT PRINCIPAL: ACTUALIZAR PRECIOS POR SKU (COLUMNA X) Y RECALCULAR Q Y R
// ──────────────────────────────────────────────────────────────────────────────
function actualizarUltimoPrecioMensualArcos() {
  const ssGestion = SpreadsheetApp.getActiveSpreadsheet();
  const periodo = obtenerPeriodoDesdeNombre();

  // 1. Abrir archivo de Facturas remoto
  let ssFacturas;
  try {
    ssFacturas = SpreadsheetApp.openById(ID_FACTURAS_ARCOS);
  } catch(e) {
    SpreadsheetApp.getUi().alert("❌ No se pudo abrir el archivo remoto de Facturas: " + e.message);
    return;
  }

  const hojaFacturas = ssFacturas.getSheetByName("Facturas") || ssFacturas.getSheetByName("9 Facturas");
  if (!hojaFacturas) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la pestaña 'Facturas' o '9 Facturas' en el archivo remoto.");
    return;
  }

  const uFilaFacturas = hojaFacturas.getLastRow();
  if (uFilaFacturas < 2) {
    SpreadsheetApp.getUi().alert("❌ La hoja Facturas remota no tiene datos.");
    return;
  }

  // Leer columnas de Facturas hasta la Columna X (24 columnas)
  const datosFacturas = hojaFacturas.getRange(2, 1, uFilaFacturas - 1, 24).getValues();

  // Ventana de 4 meses hacia atrás
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

  // Mapa ordenado estrictamente por SKU (Columna X)
  const mapaPreciosPorSku = new Map();

  for (let i = 0; i < datosFacturas.length; i++) {
    const fila = datosFacturas[i];
    
    // 1. SKU directo de Columna X (índice 23)
    let codigoSku = String(fila[FACTURAS_COLS.COL_CODIGO_SKU] || "").trim().toUpperCase();
    if (!codigoSku || codigoSku === "SIN SKU" || codigoSku.includes("#N/A") || codigoSku.includes("ERROR")) {
      continue;
    }

    // 2. Precio unitario inventario de Columna O (índice 14)
    const precioUnitario = parseNum(fila[FACTURAS_COLS.COL_PRECIO_UN_INV]);
    if (precioUnitario <= 0) continue;

    // 3. Fecha (Ingreso o Factura)
    const fechaObj = parseFecha(fila[FACTURAS_COLS.COL_FECHA_INGRESO], fila[FACTURAS_COLS.COL_FECHA_FACTURA], periodo.ano);
    if (!fechaObj) continue;

    const mFactura = fechaObj.getMonth() + 1;
    const aFactura = fechaObj.getFullYear();

    const combo = mesesValidos.find(p => p.mes === mFactura && p.ano === aFactura);
    if (!combo) continue;

    // Guardar el más reciente por fecha o posición
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

  // 2. Procesar e Inyectar en Inventario local
  const hojaGestion = ssGestion.getSheetByName(CONFIG_PRECIOS.HOJA_GESTION_INV);
  if (!hojaGestion) {
    SpreadsheetApp.getUi().alert(`❌ No encontré la pestaña '${CONFIG_PRECIOS.HOJA_GESTION_INV}'`);
    return;
  }

  const uFilaGestion = hojaGestion.getLastRow();
  if (uFilaGestion < CONFIG_PRECIOS.FILA_INICIO_INV) {
    SpreadsheetApp.getUi().alert("La hoja de Inventario no tiene filas para procesar.");
    return;
  }

  const totalProductosAProcesar = uFilaGestion - CONFIG_PRECIOS.FILA_INICIO_INV + 1;
  
  const matrizCodigos = hojaGestion.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_CODIGO_INV, totalProductosAProcesar, 1).getValues();
  const matrizPreciosViejos = hojaGestion.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_PRECIO_VIEJO_INV, totalProductosAProcesar, 1).getValues();
  const matrizFinal = hojaGestion.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_FINAL_INV, totalProductosAProcesar, 1).getValues(); // Columna P

  const matrizNuevosPrecios = [];
  const matrizAuditoria = [];
  const matrizValorizadoQ = [];

  let countMesActual = 0;
  let countHistorico = 0;
  let countMantenidos = 0;
  let granTotalValorizado = 0;

  const nombresMesesTexto = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

  for (let i = 0; i < totalProductosAProcesar; i++) {
    const codInventario = String(matrizCodigos[i][0] || "").trim().toUpperCase();
    const precioBasePlanilla = matrizPreciosViejos[i][0];
    const cantFinal = parseNum(matrizFinal[i][0]);

    if (!codInventario) {
      matrizNuevosPrecios.push([precioBasePlanilla]);
      matrizAuditoria.push([""]);
      matrizValorizadoQ.push([0]);
      continue;
    }

    let precioAsignado = precioBasePlanilla;
    let trackingOrigen = "No encontrado - Se mantuvo precio anterior";

    if (mapaPreciosPorSku.has(codInventario)) {
      const match = mapaPreciosPorSku.get(codInventario);
      precioAsignado = match.precio;
      
      const dia = String(match.fecha.getDate()).padStart(2, "0");
      const mes = String(match.fecha.getMonth() + 1).padStart(2, "0");
      const ano = match.fecha.getFullYear();

      if (match.nivel === 0) {
        trackingOrigen = `Factura: ${dia}/${mes}/${ano}`;
        countMesActual++;
      } else {
        trackingOrigen = `Histórico: ${nombresMesesTexto[match.mes]} (${dia}/${mes})`;
        countHistorico++;
      }
    } else {
      countMantenidos++;
    }

    matrizNuevosPrecios.push([precioAsignado]);
    matrizAuditoria.push([trackingOrigen]);

    // Cálculo automático Q = P * G
    const pUnit = parseNum(precioAsignado);
    const subtotalFila = Math.round(cantFinal * pUnit * 100) / 100;
    matrizValorizadoQ.push([subtotalFila]);
    granTotalValorizado += subtotalFila;
  }

  // Cálculo de porcentajes para Columna R (Q / Total)
  const matrizPorcentajesR = matrizValorizadoQ.map(([val]) => {
    if (granTotalValorizado > 0 && val > 0) {
      return [(val / granTotalValorizado)];
    }
    return [0];
  });

  // Escritura masiva en lote: G, Q, R y S
  hojaGestion.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_PRECIO_NUEVO_INV, totalProductosAProcesar, 1).setValues(matrizNuevosPrecios);
  hojaGestion.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_VALORIZADO_INV, totalProductosAProcesar, 1).setValues(matrizValorizadoQ);
  
  const rangoR = hojaGestion.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_PORC_VALORIZADO_INV, totalProductosAProcesar, 1);
  rangoR.setValues(matrizPorcentajesR);
  rangoR.setNumberFormat("0.00%");

  hojaGestion.getRange(CONFIG_PRECIOS.FILA_INICIO_INV, CONFIG_PRECIOS.COL_ORIGEN_PRECIO_INV, totalProductosAProcesar, 1).setValues(matrizAuditoria);

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(`🚀 Sincronización Arcos Completada (${nombresMesesTexto[periodo.mes]} ${periodo.ano})

• Actualizados de este mes: ${countMesActual}
• Recuperados de meses anteriores: ${countHistorico}
• Sin compras (mantenidos de Col F): ${countMantenidos}
• Total Valorizado Stock: $ ${granTotalValorizado.toLocaleString("es-AR", { minimumFractionDigits: 2 })}

Columnas G (Precio), Q (Valorizado P*G), R (% Valorizado) y S (Auditoría) sincronizadas.`);
}

// ──────────────────────────────────────────────────────────────────────────────
// MENÚ DE LA BARRA SUPERIOR
// ──────────────────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Automatización')
    .addItem('Paso 1: Actualizar Precios por Código SKU (Col X) y Valorizar', 'actualizarUltimoPrecioMensualArcos')
    .addToUi();
}