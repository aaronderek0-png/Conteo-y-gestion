// ==========================================================================
// SISTEMA INTEGRAL: CAFÉ CABELLO (TOTAL2 + CONTROL SEMANAL & GESTIÓN)
// ==========================================================================

const CONFIG = {
  // --- PLANILLA LOCAL (CABELLO) ---
  HOJA_TOTAL2: "Total2",
  HOJA_EQUIVALENCIAS: "Equivalencias",
  HOJA_LOGS: "Control_Semanal",
  FILA_INICIO_TOTAL: 2,

  // --- MAESTRO PRODUCTOS (BASE CENTRAL DE RUBROS) ---
  ID_MAESTRO_PRODUCTOS: "1lOabw9U4ilRwE5r6QXQ2Jc1nqkKUK3Hi096xga0XFnI",
  HOJA_MAESTRO_PRODUCTOS: "Productos",

  // 7 Solapas de Conteo Operativo
  SECTORES: [
    { nombre: "DESPACHO CALIENTE",            colIdx: 0, filaInicio: 2, colCantidad: 5 }, // Col F en Total2
    { nombre: "DELY",                         colIdx: 1, filaInicio: 2, colCantidad: 5 }, // Col G en Total2
    { nombre: "PASTELERIA",                   colIdx: 2, filaInicio: 2, colCantidad: 5 }, // Col H en Total2
    { nombre: "PRODUCCION - CAMARA - ALMACEN", colIdx: 3, filaInicio: 2, colCantidad: 5 }, // Col I en Total2
    { nombre: "BARRA - DEPO",                 colIdx: 4, filaInicio: 2, colCantidad: 5 }, // Col J en Total2
    { nombre: "BARRA",                        colIdx: 5, filaInicio: 2, colCantidad: 5 }, // Col K en Total2
    { nombre: "SALON",                        colIdx: 6, filaInicio: 2, colCantidad: 5 }  // Col L en Total2
  ],

  // --- PLANILLA DESTINO (GESTIÓN CABELLO) ---
  ID_GESTION: "1ayTBVaxzvkKdnb1RpIPDvTnn5dDfR7XSllbVvNF8Rj8", 
  HOJA_DESTINO_GESTION: "Gestion", 
  GESTION_FILA_DATOS: 2,         
  GESTION_COL_DESTINO_INICIO: 8, // Col H (8 a 14: Sectores | 15: FINAL)

  // --- PLANILLA DESTINO EXTERNA PARA HISTORIAL CENTRAL ---
  ID_HISTORIAL_EXTERNO: "1bnJp0suZ8pBoE3AFMnLMdzTy5frC3e3bDZ5XVbBzrLQ",
  HOJA_HISTORIAL_EXTERNO: "Historial"
};

// DICCIONARIO ESTRICTO Y CERRADO: SOLO ESTOS ITEMS SE MULTIPLICAN
const TABLA_CONVERSIONES = [
  { clave: "LEICAJ PORCION",          factor: 0.100 },
  { clave: "BABKA CHOCOLATE PORCION", factor: 0.100 },
  { clave: "PASTRON X 100GR",         factor: 0.100 },
  { clave: "PASTRÓN X100GR",          factor: 0.100 },
  { clave: "PASTRON X100 GR",         factor: 0.100 },
  { clave: "PASTRON X60 GR",          factor: 0.060 },
  { clave: "TRUCHA X 100GR",          factor: 0.100 },
  { clave: "TRUCHA X 70GR",           factor: 0.070 },
  { clave: "FALAFEL PORC",            factor: 0.250 },
  { clave: "ACEITE DE OLIVA 250",     factor: 0.250 },
  { clave: "ACEITE DE OLIVA 500",     factor: 0.500 }
];

const PALABRAS_IGNORAR = [
  "ALMACEN", "ALMACÉN", "LACTEOS", "LÁCTEOS", "GASEOSAS", "BEBIDAS", "DESCARTABLES", 
  "LIMPIEZA", "SUB RUBRO", "RESPONSABLE", "AREA:", "ÁREA:", "UTILES", "ÚTILES", "TAPERS", 
  "HELADERA", "UN.", "UN", "MEDIDA", "PRODUCTO", "PRODUCCION", "PRODUCCIÓN", "QUESOS", 
  "LAMINADOS", "PANES OBRADOR", "HARINERA", "HARINAS", "PRODUCTOS TERCERIZADOS", "CAFETERIA", 
  "PRUEBAS Y OTROS", "OBRADOR", "CARNES", "PESCADOS", "VINOS", "DESTILADOS", "KAMADO", 
  "PRODUCTOS", "BEBIDAS Y DESTILADOS", "SALÓN", "SALON", "BARRA", "DELY", "DESPACHO",
  "CRISTALERIA Y LOZA", "CRISTALERIA", "LOZA", "CONTEO"
];

const ITEMS_CRISTALERIA_LOZA = [
  "VASO YOGURT", "VASO SALON", "COPA VINO BORGONA", "COPA VINO BRUNELLO",
  "CUCHILLO SERVICIO", "CUCHILLO POSTRE", "CUCHILLO CORTE",
  "CUCHARA SERVICIO", "CUCHARA POSTRE", "CUCHARA CAFE CHICO", "CUCHARA CAFE GRANDE",
  "TENEDOR SERVICIO", "TENEDOR POSTRE",
  "PLATO MEZZE", "PLATO PLAYO CHICO", "PLATO PLAYO MEDIANO", "PLATO PLAYO GRANDE",
  "PLATO HONDO HUMMUS", "BOWL", "POCILLO", "JARRITO VIDRIO", "TAZA CAFE DOBLE", "TAZON", 
  "MIX HONGOS Y CEBOLLA ASADA (MANAKISH)"
];

function norm(texto) {
  if (!texto) return "";
  return texto.toString().trim().toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9 ]/g, "")
    .replace(/\s+/g, " ");
}

// Conversión única y exclusivamente por mapeo explícito
function obtenerFactorConversion(nombreProducto) {
  if (!nombreProducto) return 1;
  const pNorm = norm(nombreProducto);
  for (let i = 0; i < TABLA_CONVERSIONES.length; i++) {
    if (pNorm.includes(norm(TABLA_CONVERSIONES[i].clave))) {
      return TABLA_CONVERSIONES[i].factor;
    }
  }
  return 1;
}

function esTituloOSeccion(texto, unVal, medVal) {
  const t = norm(texto);
  if (!t) return true;
  if (t.startsWith("*")) return true;
  if (norm(unVal) === "UN" || norm(medVal) === "MEDIDA") return true;
  return PALABRAS_IGNORAR.some(p => t === norm(p) || t.startsWith(norm(p)));
}

function esCristaleriaOLoza(nombreProducto) {
  const normProd = norm(nombreProducto);
  return ITEMS_CRISTALERIA_LOZA.some(item => norm(item) === normProd);
}

function parseNum(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (typeof val === "number") return isNaN(val) ? 0 : val;
  const num = parseFloat(val.toString().trim().replace(",", "."));
  return isNaN(num) ? 0 : num;
}

function getHoja(libro, nombre) {
  const objetivo = norm(nombre);
  for (const hoja of libro.getSheets()) {
    if (norm(hoja.getName()) === objetivo) return hoja;
  }
  return null;
}

function estaMarcadoExcluir(val) {
  if (!val) return false;
  const v = norm(val);
  return v === "SI" || v === "TRUE" || v === "VERDADERO" || v === "1" || v === "X" || v === "EXCLUIR";
}

function obtenerColumnaExcluir(hoja, colCantidadDefault) {
  const uCol = hoja.getLastColumn();
  if (uCol < 1) return colCantidadDefault + 1;
  const filasCabecera = Math.min(5, hoja.getLastRow());
  const datosCabecera = hoja.getRange(1, 1, filasCabecera, uCol).getValues();
  for (let c = 0; c < uCol; c++) {
    for (let f = 0; f < filasCabecera; f++) {
      if (norm(datosCabecera[f][c]).includes("EXCLUIR")) return c + 1;
    }
  }
  return colCantidadDefault + 1;
}

function cargarMapaRubrosDesdeMaestro() {
  const mapa = new Map();
  try {
    const ssMaestro = SpreadsheetApp.openById(CONFIG.ID_MAESTRO_PRODUCTOS);
    const hojaMaestro = getHoja(ssMaestro, CONFIG.HOJA_MAESTRO_PRODUCTOS);
    if (!hojaMaestro) return mapa;
    const uFila = hojaMaestro.getLastRow();
    if (uFila >= 2) {
      const datos = hojaMaestro.getRange(2, 1, uFila - 1, 4).getValues();
      datos.forEach(f => {
        const skuNorm = norm(f[0]);
        const rubro = String(f[3] || "").trim();
        if (skuNorm && rubro) mapa.set(skuNorm, rubro);
      });
    }
  } catch (err) {
    console.error("Error al cargar rubros desde el Maestro: " + err);
  }
  return mapa;
}

function determinarRubroFinal(skuNorm, mapaRubrosMaestro) {
  if (!skuNorm || skuNorm === "NO ENCONTRADO" || skuNorm.startsWith("NO_EXISTE")) return "OBRADOR";
  if (mapaRubrosMaestro.has(skuNorm)) return mapaRubrosMaestro.get(skuNorm);
  if (skuNorm.startsWith("OBR") || /^P\d+/.test(skuNorm)) return "OBRADOR";
  return "ALMACÉN";
}

function prepararHojaTotal2Limpia(ss) {
  let hojaTotal2 = getHoja(ss, CONFIG.HOJA_TOTAL2);
  if (!hojaTotal2) {
    hojaTotal2 = ss.insertSheet(CONFIG.HOJA_TOTAL2);
  } else {
    hojaTotal2.clearContents();
  }

  const headers = [
    "RUBRO", "UN.", "MEDIDA", "CODIGO", "PRODUCTO",
    "DESPACHO CALIENTE", "DELY", "PASTELERIA", "PRODUCCION - CAMARA - ALMACEN",
    "BARRA - DEPO", "BARRA", "SALON", "FINAL"
  ];
  hojaTotal2.getRange(1, 1, 1, headers.length).setValues([headers]);
  return hojaTotal2;
}

// ==========================================================================
// 1. CONSOLIDAR EN TOTAL2
// ==========================================================================
function consolidarEnTotal2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mapaRubrosMaestro = cargarMapaRubrosDesdeMaestro();

  const mapaEquivalencias = new Map();
  const hojaEq = getHoja(ss, CONFIG.HOJA_EQUIVALENCIAS);
  if (hojaEq && hojaEq.getLastRow() > 1) {
    const datosEq = hojaEq.getRange(2, 1, hojaEq.getLastRow() - 1, 2).getValues();
    datosEq.forEach(r => {
      const alias = norm(r[0]);
      const sku = norm(r[1]);
      if (alias && sku) mapaEquivalencias.set(alias, sku);
    });
  }

  const mapaItemsValidos = new Map();
  const mapaItemsSinSku = new Map();
  const cantSectores = CONFIG.SECTORES.length;

  CONFIG.SECTORES.forEach((sec, sIdx) => {
    const hojaSec = getHoja(ss, sec.nombre);
    if (!hojaSec) return;

    const uFila = hojaSec.getLastRow();
    if (uFila < sec.filaInicio) return;

    const colExcluir = obtenerColumnaExcluir(hojaSec, sec.colCantidad);
    const maxCols = Math.max(6, sec.colCantidad, colExcluir);
    const datosSec = hojaSec.getRange(sec.filaInicio, 1, uFila - sec.filaInicio + 1, maxCols).getValues();

    for (let r = 0; r < datosSec.length; r++) {
      const fila = datosSec[r];
      const unVal = fila[0] || 1;
      const skuRaw = String(fila[1] || "").trim();
      const medVal = fila[2] || "UN";
      const prodRaw = String(fila[3] || "").trim();
      const cantBruta = parseNum(fila[sec.colCantidad - 1]);
      const valorExcluir = fila[colExcluir - 1];

      if (!prodRaw || esTituloOSeccion(prodRaw, unVal, medVal) || esCristaleriaOLoza(prodRaw)) continue;
      if (estaMarcadoExcluir(valorExcluir)) continue;

      let skuNorm = norm(skuRaw);
      const prodNorm = norm(prodRaw);

      if (!skuNorm || skuNorm === "NO ENCONTRADO" || skuNorm === "SKU FALTANTE") {
        if (mapaEquivalencias.has(prodNorm)) {
          skuNorm = mapaEquivalencias.get(prodNorm);
        }
      }

      const factor = obtenerFactorConversion(prodRaw);
      const cantReal = cantBruta * factor;
      const tieneSkuValido = (skuNorm && skuNorm !== "NO ENCONTRADO" && skuNorm !== "SKU FALTANTE");

      if (tieneSkuValido) {
        if (!mapaItemsValidos.has(skuNorm)) {
          const rubroCalculado = determinarRubroFinal(skuNorm, mapaRubrosMaestro);
          mapaItemsValidos.set(skuNorm, {
            rubro: rubroCalculado,
            un: unVal,
            medida: medVal,
            skuRaw: skuRaw,
            producto: prodRaw,
            conteos: new Array(cantSectores).fill(0)
          });
        } else {
          const item = mapaItemsValidos.get(skuNorm);
          if (prodRaw.length < item.producto.length) {
            item.producto = prodRaw;
          }
        }
        mapaItemsValidos.get(skuNorm).conteos[sIdx] += cantReal;

      } else {
        const claveSinSku = prodNorm;
        if (!mapaItemsSinSku.has(claveSinSku)) {
          mapaItemsSinSku.set(claveSinSku, {
            rubro: "REVISAR SKU",
            un: unVal,
            medida: medVal,
            origenHoja: sec.nombre,
            producto: prodRaw,
            conteos: new Array(cantSectores).fill(0)
          });
        } else {
          const item = mapaItemsSinSku.get(claveSinSku);
          if (!item.origenHoja.includes(sec.nombre)) {
            item.origenHoja += ` / ${sec.nombre}`;
          }
          if (prodRaw.length < item.producto.length) {
            item.producto = prodRaw;
          }
        }
        mapaItemsSinSku.get(claveSinSku).conteos[sIdx] += cantReal;
      }
    }
  });

  const filasConSku = [];
  mapaItemsValidos.forEach(item => {
    const cRed = item.conteos.map(c => Math.round(c * 1000) / 1000);
    const finalVal = Math.round(cRed.reduce((a, b) => a + b, 0) * 1000) / 1000;
    filasConSku.push([
      item.rubro, item.un, item.medida, item.skuRaw, item.producto, ...cRed, finalVal
    ]);
  });

  filasConSku.sort((a, b) => {
    if (a[0] === b[0]) return a[4].localeCompare(b[4]);
    return a[0].localeCompare(b[0]);
  });

  const filasSinSku = [];
  mapaItemsSinSku.forEach(item => {
    const cRed = item.conteos.map(c => Math.round(c * 1000) / 1000);
    const finalVal = Math.round(cRed.reduce((a, b) => a + b, 0) * 1000) / 1000;
    const codigoAlerta = `NO_EXISTE_[${item.origenHoja}]`;

    filasSinSku.push([
      item.rubro, item.un, item.medida, codigoAlerta, item.producto, ...cRed, finalVal
    ]);
  });

  const todasLasFilas = [...filasConSku, ...filasSinSku];

  const hojaTotal2 = prepararHojaTotal2Limpia(ss);
  if (todasLasFilas.length > 0) {
    hojaTotal2.getRange(CONFIG.FILA_INICIO_TOTAL, 1, todasLasFilas.length, 13).setValues(todasLasFilas);
  }

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    "Consolidación Lista",
    `• Ítems con SKU: ${filasConSku.length}\n` +
    `• Ítems sin SKU al final: ${filasSinSku.length}\n` +
    `Conversión aplicada únicamente a la tabla de excepciones.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==========================================================================
// 2. MOTOR DE SNAPSHOT (LOCAL EN CONTROL_SEMANAL Y CENTRAL EN HISTORIAL)
// ==========================================================================
function ejecutarRegistroSnapshotInterno_(ssLocal) {
  const hojaTotal2 = getHoja(ssLocal, CONFIG.HOJA_TOTAL2);
  if (!hojaTotal2 || hojaTotal2.getLastRow() < CONFIG.FILA_INICIO_TOTAL) {
    throw new Error("No hay datos consolidados en 'Total2' para generar el snapshot.");
  }

  const uFilaTotal2 = hojaTotal2.getLastRow();
  const cantFilas = uFilaTotal2 - CONFIG.FILA_INICIO_TOTAL + 1;
  // Leemos columnas A a M (13 columnas: Rubro, Un, Med, SKU, Prod, 7 sectores, FINAL)
  const datosTotal2 = hojaTotal2.getRange(CONFIG.FILA_INICIO_TOTAL, 1, cantFilas, 13).getValues();

  // 1. Preparar o abrir solapa local Control_Semanal
  let hojaLogs = getHoja(ssLocal, CONFIG.HOJA_LOGS);
  if (!hojaLogs) {
    hojaLogs = ssLocal.insertSheet(CONFIG.HOJA_LOGS);
    const headers = ["FECHA", "PRODUCTO", "CODIGO", "TOTAL FINAL"];
    hojaLogs.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    hojaLogs.setFrozenRows(1);
  }

  // 2. Preparar o abrir solapa en el Historial externo
  let hojaExterna = null;
  try {
    const ssExterna = SpreadsheetApp.openById(CONFIG.ID_HISTORIAL_EXTERNO);
    hojaExterna = getHoja(ssExterna, CONFIG.HOJA_HISTORIAL_EXTERNO);
    if (!hojaExterna) {
      hojaExterna = ssExterna.insertSheet(CONFIG.HOJA_HISTORIAL_EXTERNO);
      const headersExt = ["FECHA REGISTRO", "PRODUCTO", "CODIGO", "FINAL"];
      hojaExterna.getRange(1, 1, 1, headersExt.length).setValues([headersExt]).setFontWeight("bold");
      hojaExterna.setFrozenRows(1);
    }
  } catch (e) {
    console.warn("No se pudo conectar a la planilla de historial externo: " + e.message);
  }

  const fecha = new Date();
  const registrosSnapshot = [];

  datosTotal2.forEach(fila => {
    const sku = String(fila[3] || "").trim();
    const prod = String(fila[4] || "").trim();
    const finalVal = parseNum(fila[12]); // Col M: Columna 13 (FINAL)

    if (!prod || esTituloOSeccion(prod, fila[1], fila[2])) return;

    if (finalVal > 0) {
      registrosSnapshot.push([
        fecha,
        prod,
        (sku && sku.toUpperCase() !== "NO ENCONTRADO") ? sku : "No encontrado",
        finalVal
      ]);
    }
  });

  if (registrosSnapshot.length > 0) {
    hojaLogs.getRange(hojaLogs.getLastRow() + 1, 1, registrosSnapshot.length, 4).setValues(registrosSnapshot);

    if (hojaExterna) {
      hojaExterna.getRange(hojaExterna.getLastRow() + 1, 1, registrosSnapshot.length, 4).setValues(registrosSnapshot);
    }
  }

  return registrosSnapshot.length;
}

function guardarSnapshotControlSemanal() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "📸 GUARDAR SNAPSHOT",
    "¿Confirmás registrar el snapshot del stock consolidado en 'Control_Semanal' y en el Historial central?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    const cant = ejecutarRegistroSnapshotInterno_(ssLocal);

    ui.alert(
      "✅ SNAPSHOT GUARDADO CON ÉXITO",
      `Se registraron ${cant} productos con stock activo en:\n\n` +
      `1. Local: "${CONFIG.HOJA_LOGS}"\n` +
      `2. Externa: "${CONFIG.HOJA_HISTORIAL_EXTERNO}"\n\n` +
      `Formato: Fecha | Producto | Código | Total Final`,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert("❌ ERROR EN SNAPSHOT", err.message, ui.ButtonSet.OK);
  }
}

// ==========================================================================
// 3. CIERRE SEMANAL: SNAPSHOT AUTOMÁTICO + GESTIÓN CABELLO + LIMPIEZA
// ==========================================================================
function enviarAGestionCabello() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "Confirmación de Cierre Semanal",
    "¿Confirmás guardar Snapshot, transferir los conteos (H:O) a Gestión Cabello y vaciar sectores locales?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const tInicio = new Date().getTime();

  try {
    const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    const hojaTotal2 = getHoja(ssLocal, CONFIG.HOJA_TOTAL2);

    if (!hojaTotal2 || hojaTotal2.getLastRow() < CONFIG.FILA_INICIO_TOTAL) {
      throw new Error("No existe la hoja 'Total2' con datos. Consolidá primero.");
    }

    // --- PASO A: REGISTRO AUTOMÁTICO DE SNAPSHOT ---
    let cantSnapshot = 0;
    try {
      cantSnapshot = ejecutarRegistroSnapshotInterno_(ssLocal);
    } catch (eSnap) {
      console.warn("Aviso en snapshot automático: " + eSnap.message);
    }

    // --- PASO B: TRANSFERENCIA A GESTIÓN ---
    const uFilaTotal2 = hojaTotal2.getLastRow();
    const cantFilasLocal = uFilaTotal2 - CONFIG.FILA_INICIO_TOTAL + 1;
    const datosTotal2 = hojaTotal2.getRange(CONFIG.FILA_INICIO_TOTAL, 1, cantFilasLocal, 13).getValues();

    const mapaLocalPorSku = new Map();
    datosTotal2.forEach(fila => {
      const sku = norm(fila[3]);
      if (sku && sku !== "NO ENCONTRADO" && !sku.startsWith("NO EXISTE")) {
        mapaLocalPorSku.set(sku, {
          rubro: fila[0],
          un: fila[1],
          medida: fila[2],
          skuRaw: fila[3],
          prod: fila[4],
          conteos: fila.slice(5, 12), // F a L (7 sectores)
          final: fila[12]             // M (Total Final)
        });
      }
    });

    const ssRemoto = SpreadsheetApp.openById(CONFIG.ID_GESTION);
    let hojaRemota = ssRemoto.getSheetByName(CONFIG.HOJA_DESTINO_GESTION) || ssRemoto.getSheets()[0];

    const uFilaRemota = hojaRemota.getLastRow();
    const cantFilasRemotas = uFilaRemota >= CONFIG.GESTION_FILA_DATOS ? (uFilaRemota - CONFIG.GESTION_FILA_DATOS + 1) : 0;

    if (cantFilasRemotas === 0) {
      throw new Error("La hoja destino de Gestión no tiene datos desde la fila 2.");
    }

    const clavesGestion = hojaRemota.getRange(CONFIG.GESTION_FILA_DATOS, 4, cantFilasRemotas, 1).getValues();
    const skusProcesados = new Set();
    const matrizConteosNuevos = []; // 8 columnas en Gestión: H a N (7 sectores) + O (FINAL)

    for (let i = 0; i < clavesGestion.length; i++) {
      const skuRemoto = norm(clavesGestion[i][0]);
      let filaConteos = [0, 0, 0, 0, 0, 0, 0, 0];

      if (skuRemoto && mapaLocalPorSku.has(skuRemoto)) {
        const item = mapaLocalPorSku.get(skuRemoto);
        filaConteos = [...item.conteos, item.final];
        skusProcesados.add(skuRemoto);
      }
      matrizConteosNuevos.push(filaConteos);
    }

    // Inyección en H:O (Columna 8, ancho 8)
    hojaRemota.getRange(CONFIG.GESTION_FILA_DATOS, CONFIG.GESTION_COL_DESTINO_INICIO, cantFilasRemotas, 8).setValues(matrizConteosNuevos);

    // Ítems no existentes al fondo
    const filasNuevas = [];
    mapaLocalPorSku.forEach((item, skuKey) => {
      if (!skusProcesados.has(skuKey)) {
        filasNuevas.push([
          item.rubro, item.un, item.medida, item.skuRaw, item.prod, "", "", ...item.conteos, item.final
        ]);
      }
    });

    if (filasNuevas.length > 0) {
      const filaAlta = hojaRemota.getLastRow() + 1;
      hojaRemota.getRange(filaAlta, 1, filasNuevas.length, filasNuevas[0].length).setValues(filasNuevas);
    }

    // --- PASO C: LIMPIEZA DE CONTEOS LOCALES ---
    CONFIG.SECTORES.forEach(sec => {
      const hoja = getHoja(ssLocal, sec.nombre);
      if (!hoja) return;
      const uF = hoja.getLastRow();
      if (uF >= sec.filaInicio) {
        hoja.getRange(sec.filaInicio, sec.colCantidad, uF - sec.filaInicio + 1, 1).clearContent();
      }
    });

    const tTotal = ((new Date().getTime() - tInicio) / 1000).toFixed(1);

    ui.alert(
      "🚀 CIERRE Y GESTIÓN COMPLETADA",
      `Operación finalizada en ${tTotal} segundos.\n\n` +
      `• Snapshot archivado: ${cantSnapshot} ítems registrados en "${CONFIG.HOJA_LOGS}" e Historial.\n` +
      `• Destino Gestión: "${hojaRemota.getName()}" (H:O actualizadas).\n` +
      `• Filas impactadas: ${cantFilasRemotas}\n` +
      `• Nuevos ítems dados de alta: ${filasNuevas.length}\n` +
      `• Sectores reseteados para la próxima semana.`,
      ui.ButtonSet.OK
    );

  } catch (err) {
    ui.alert("Error en el cierre", err.toString(), ui.ButtonSet.OK);
  }
}

// ==========================================================================
// 4. LIMPIEZA MANUAL
// ==========================================================================
function borrarConteosSectores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = SpreadsheetApp.getUi().alert(
    "Atención",
    "¿Confirmás vaciar las cantidades de los 7 sectores y reiniciar Total2 a 0?",
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (resp !== SpreadsheetApp.getUi().Button.YES) return;

  CONFIG.SECTORES.forEach(sec => {
    const hoja = getHoja(ss, sec.nombre);
    if (!hoja) return;

    const uFila = hoja.getLastRow();
    if (uFila >= sec.filaInicio) {
      hoja.getRange(sec.filaInicio, sec.colCantidad, uFila - sec.filaInicio + 1, 1).clearContent();
    }
  });

  const hojaTotal2 = getHoja(ss, CONFIG.HOJA_TOTAL2);
  if (hojaTotal2 && hojaTotal2.getLastRow() >= CONFIG.FILA_INICIO_TOTAL) {
    const cant = hojaTotal2.getLastRow() - CONFIG.FILA_INICIO_TOTAL + 1;
    const ceros = Array.from({ length: cant }, () => new Array(8).fill(0));
    hojaTotal2.getRange(CONFIG.FILA_INICIO_TOTAL, 6, cant, 8).setValues(ceros);
  }

  SpreadsheetApp.getUi().alert("Limpieza Completa", "Se vaciaron los conteos.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================================================
// MENÚ SUPERIOR
// ==========================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ INVENTARIO CABELLO")
    .addItem("1️⃣ Consolidar Conteos en 'Total2'", "consolidarEnTotal2")
    .addItem("📸 Guardar Snapshot en Historial (Local y Externo)", "guardarSnapshotControlSemanal")
    .addSeparator()
    .addItem("🚀 2. Cierre Semanal (Snapshot + Gestión + Limpieza)", "enviarAGestionCabello")
    .addItem("🧹 Resetear Cantidades Manualmente", "borrarConteosSectores")
    .addToUi();
}