// ==========================================================================
// SISTEMA INTEGRAL: ROTISERÍA MISHIGUENE -> CONSOLIDACIÓN & GESTIÓN + SNAPSHOT
// ==========================================================================

const CONFIG = {
  HOJA_TOTAL_ORIGINAL: "Total",
  HOJA_TOTAL2: "Total2",
  HOJA_EQUIVALENCIAS: "Equivalencias",
  HOJA_LOGS: "Control_Semanal",
  FILA_INICIO_TOTAL: 2,

  // --- MAESTRO PRODUCTOS (BASE CENTRAL DE RUBROS) ---
  ID_MAESTRO_PRODUCTOS: "1lOabw9U4ilRwE5r6QXQ2Jc1nqkKUK3Hi096xga0XFnI",
  HOJA_MAESTRO_PRODUCTOS: "Productos",

  // --- GESTIÓN MENSUAL (ROTISERÍA) ---
  ID_GESTION_MENSUAL: "1mzZE5ZZhz3wMpm9SB3cquxC3l3SFaAdp1iUBkMiMBAY",
  NOMBRE_HOJA_INVENTARIO: "Inventario",
  GESTION_FILA_DATOS: 3,
  GESTION_COL_DESTINO_INICIO: 8, // Columna H (8 a 15: Sectores | 16: FINAL)

  // --- PLANILLA DESTINO EXTERNA PARA HISTORIAL CENTRAL ---
  ID_HISTORIAL_EXTERNO: "1bnJp0suZ8pBoE3AFMnLMdzTy5frC3e3bDZ5XVbBzrLQ",
  HOJA_HISTORIAL_EXTERNO: "Historial",

  // Rubros a descartar
  RUBROS_IGNORAR: ["LIMPIEZA", "LIBRERIA", "LIBRERÍA"],

  // SKUs o Nombres puntuales a ignorar completamente
  ITEMS_IGNORAR: ["I0511", "BPT", "I0292", "COSTA Y PAMPA ALBARINO", "BOBINA DE PAPEL", "LIBRO ALEGRIA - TOMAS KALIKA", "LIBRO ALEGRIA"],

  // 8 sectores operativos de conteo
  SECTORES: [
    { nombre: "1 SALON + BARRA",      colNombre: "BARRA / SALON",      filaInicio: 4, colCantidad: 5 },
    { nombre: "2 DESPACHO",            colNombre: "COCINA",             filaInicio: 4, colCantidad: 6 },
    { nombre: "3 DEPOSITO-ALMACEN",    colNombre: "DEPO",               filaInicio: 4, colCantidad: 5 },
    { nombre: "4 ROTISERÍA",           colNombre: "ROTISERIA",          filaInicio: 4, colCantidad: 5 },
    { nombre: "5 DESCARTABLES ROTI",  colNombre: "DESCARTABLES ROTI",  filaInicio: 4, colCantidad: 5 },
    { nombre: "6 DESCARTABLES BARRA", colNombre: "DESCARTABLES BARRA", filaInicio: 4, colCantidad: 5 },
    { nombre: "7 DESCARTABLES DEPO",  colNombre: "DESCARTABLES DEPO",  filaInicio: 4, colCantidad: 6 },
    { nombre: "10 CAMARA",             colNombre: "CAMARA",             filaInicio: 4, colCantidad: 5 }
  ],

  HOJAS_IGNORAR: ["9 BACHA", "BACHA"],

  HOJA_CAJA: {
    nombre: "8 CAJA",
    filaInicio: 4,
    colCantidad: 5,
    skusPermitidos: ["I0021", "BKP4"],
    sectorDestinoIdx: 5 // DESCARTABLES BARRA
  }
};

const TABLA_CONVERSIONES = [
  { clave: "LEICAJ PORCION",          factor: 0.100 },
  { clave: "BABKA CHOCOLATE PORCION", factor: 0.100 },
  { clave: "PASTRON X100 GR",         factor: 0.100 },
  { clave: "PASTRÓN X100GR",         factor: 0.100 },
  { clave: "PASTRON X60 GR",          factor: 0.060 },
  { clave: "TRUCHA X 100GR",          factor: 0.100 },
  { clave: "TRUCHA X 70GR",           factor: 0.070 },
  { clave: "ACEITE DE OLIVA 250",     factor: 0.250 },
  { clave: "ACEITE DE OLIVA 500",     factor: 0.500 }
];

function norm(texto) {
  if (!texto) return "";
  return texto.toString().trim().toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9 ]/g, "")
    .replace(/\s+/g, " ");
}

function obtenerFactorConversion(nombreProducto) {
  if (!nombreProducto) return 1;
  const pNorm = norm(nombreProducto);
  for (const regla of TABLA_CONVERSIONES) {
    if (pNorm.includes(norm(regla.clave))) return regla.factor;
  }
  const matchG = pNorm.match(/X\s*(\d+)\s*(?:GR|G\b)/);
  if (matchG) {
    const gr = parseFloat(matchG[1]);
    if (gr > 0 && gr < 1000) return gr / 1000;
  }
  const matchM = pNorm.match(/(\d+)\s*(?:ML|CC)\b/);
  if (matchM) {
    const ml = parseFloat(matchM[1]);
    if (ml > 0 && ml < 1000) return ml / 1000;
  }
  return 1;
}

function esFilaProductoValida(a, b, c, d) {
  const colA = norm(a);
  const colB = norm(b);
  const colC = norm(c);
  const colD = norm(d);

  if (!colD || colD.length < 2) return false;
  if (colA === "UN" && (colB === "SKU" || colB === "CODIGO" || colB === "COD") && colC === "MEDIDA") return false;
  if (!colA && !colB && !colC) return false;
  if (colD.startsWith("SUB RUBRO") || colD.startsWith("SUBRUBRO")) return false;
  return true;
}

function estaMarcadoExcluir(val) {
  if (!val) return false;
  const v = norm(val);
  return v === "SI" || v === "TRUE" || v === "VERDADERO" || v === "1" || v === "X" || v === "EXCLUIR";
}

function esItemIgnorado(sku, prod, rubro) {
  if (esRubroIgnorado(rubro)) return true;
  const s = norm(sku);
  const p = norm(prod);
  return CONFIG.ITEMS_IGNORAR.some(ign => {
    const ignNorm = norm(ign);
    return s === ignNorm || p === ignNorm || p.includes(ignNorm);
  });
}

function esRubroIgnorado(rubro) {
  const r = norm(rubro);
  return CONFIG.RUBROS_IGNORAR.some(ign => norm(ign) === r);
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
  if (!skuNorm || skuNorm === "NO ENCONTRADO") return "OBRADOR";
  if (mapaRubrosMaestro.has(skuNorm)) return mapaRubrosMaestro.get(skuNorm);
  if (skuNorm.startsWith("OBR") || /^P\d+/.test(skuNorm)) return "OBRADOR";
  return "ALMACÉN";
}

// ==========================================================================
// 1. SINCRONIZAR SKU EN SECTORES
// ==========================================================================
function prepararColumnasSkuEnSectores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTotal = getHoja(ss, CONFIG.HOJA_TOTAL2) || getHoja(ss, CONFIG.HOJA_TOTAL_ORIGINAL);
  if (!hojaTotal) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la hoja base 'Total' o 'Total2'.");
    return;
  }

  const uFilaTotal = hojaTotal.getLastRow();
  const cantFilas = uFilaTotal >= CONFIG.FILA_INICIO_TOTAL ? (uFilaTotal - CONFIG.FILA_INICIO_TOTAL + 1) : 0;
  const mapaCodigoPorNombre = new Map();

  if (cantFilas > 0) {
    const datosTotal = hojaTotal.getRange(CONFIG.FILA_INICIO_TOTAL, 1, cantFilas, 5).getValues();
    datosTotal.forEach(fila => {
      const cod = String(fila[3] || "").trim();
      const prod = norm(fila[4]);
      if (prod && cod && cod.toUpperCase() !== "NO ENCONTRADO") {
        mapaCodigoPorNombre.set(prod, cod);
      }
    });
  }

  const todasLasSolapas = [...CONFIG.SECTORES, CONFIG.HOJA_CAJA];

  todasLasSolapas.forEach(sec => {
    const hoja = getHoja(ss, sec.nombre);
    if (!hoja) return;

    const cabecerasB = hoja.getRange(1, 2, Math.min(5, hoja.getLastRow()), 1).getValues().flat().map(c => norm(c));
    const tieneSku = cabecerasB.some(c => c === "SKU" || c === "CODIGO" || c === "COD");

    if (!tieneSku) {
      hoja.insertColumnBefore(2);
      hoja.getRange(3, 2).setValue("SKU").setFontWeight("bold");
    }

    const uFila = hoja.getLastRow();
    if (uFila >= sec.filaInicio) {
      const prodsSector = hoja.getRange(sec.filaInicio, 4, uFila - sec.filaInicio + 1, 1).getValues();
      const skusGenerados = [];

      for (let i = 0; i < prodsSector.length; i++) {
        const pName = norm(prodsSector[i][0]);
        if (pName && esFilaProductoValida("", "", "", prodsSector[i][0]) && mapaCodigoPorNombre.has(pName)) {
          skusGenerados.push([mapaCodigoPorNombre.get(pName)]);
        } else {
          const skuActual = hoja.getRange(sec.filaInicio + i, 2).getValue();
          skusGenerados.push([skuActual ? String(skuActual).trim() : ""]);
        }
      }
      hoja.getRange(sec.filaInicio, 2, skusGenerados.length, 1).setValues(skusGenerados);
    }
  });

  SpreadsheetApp.getUi().alert("✅ SKU SINCRONIZADOS", "Se actualizaron los códigos en las solapas de conteo.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================================================
// 2. CONSOLIDAR EN TOTAL2 (MANTIENE FILAS FIJAS, NO PISA CATÁLOGO NI FORMATO)
// ==========================================================================
function consolidarEnTotal2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hojaDestino = getHoja(ss, CONFIG.HOJA_TOTAL2);

  if (!hojaDestino) {
    SpreadsheetApp.getUi().alert("❌ No existe la hoja 'Total2'. Por favor renombrá o creala.");
    return;
  }

  const mapaEquivalencias = new Map();
  const mapaRubrosMaestro = cargarMapaRubrosDesdeMaestro();

  const hojaEq = getHoja(ss, CONFIG.HOJA_EQUIVALENCIAS);
  if (hojaEq && hojaEq.getLastRow() > 1) {
    const datosEq = hojaEq.getRange(2, 1, hojaEq.getLastRow() - 1, 2).getValues();
    datosEq.forEach(r => {
      const alias = norm(r[0]);
      const sku = norm(r[1]);
      if (alias && sku) mapaEquivalencias.set(alias, sku);
    });
  }

  const uFilaTotal2 = hojaDestino.getLastRow();
  const cantItemsFijos = uFilaTotal2 >= 2 ? (uFilaTotal2 - 1) : 0;

  if (cantItemsFijos === 0) {
    SpreadsheetApp.getUi().alert("⚠️ La hoja Total2 no tiene catálogo base en filas 2 en adelante.");
    return;
  }

  const baseFija = hojaDestino.getRange(2, 1, cantItemsFijos, 5).getValues();
  const cantSectores = CONFIG.SECTORES.length;

  const matrizConteos = Array.from({ length: cantItemsFijos }, () => new Array(cantSectores + 1).fill(0));

  const idxPorSku = new Map();
  const idxPorNombre = new Map();

  for (let i = 0; i < cantItemsFijos; i++) {
    const skuNorm = norm(baseFija[i][3]);
    const prodNorm = norm(baseFija[i][4]);
    if (skuNorm && skuNorm !== "NO ENCONTRADO") {
      if (!idxPorSku.has(skuNorm)) idxPorSku.set(skuNorm, i);
    }
    if (prodNorm) {
      if (!idxPorNombre.has(prodNorm)) idxPorNombre.set(prodNorm, i);
    }
  }

  const itemsNuevosFinal = [];
  const idxNuevosPorSku = new Map();
  const idxNuevosPorNombre = new Map();

  // --- LECTURA DE SECTORES ---
  CONFIG.SECTORES.forEach((sec, sIdx) => {
    if (CONFIG.HOJAS_IGNORAR.some(h => norm(h) === norm(sec.nombre))) return;

    const hojaSec = getHoja(ss, sec.nombre);
    if (!hojaSec) return;

    const uFila = hojaSec.getLastRow();
    if (uFila < sec.filaInicio) return;

    const colExcluir = obtenerColumnaExcluir(hojaSec, sec.colCantidad);
    const maxCols = Math.max(5, sec.colCantidad, colExcluir);
    const datosSec = hojaSec.getRange(sec.filaInicio, 1, uFila - sec.filaInicio + 1, maxCols).getValues();

    for (let r = 0; r < datosSec.length; r++) {
      const fila = datosSec[r];
      const unVal = fila[0];
      const skuRaw = String(fila[1] || "").trim();
      const medVal = fila[2];
      const prodRaw = String(fila[3] || "").trim();
      const cantBruta = parseNum(fila[sec.colCantidad - 1]);
      const valorExcluir = fila[colExcluir - 1];

      if (!esFilaProductoValida(unVal, skuRaw, medVal, prodRaw)) continue;
      if (estaMarcadoExcluir(valorExcluir)) continue;
      if (esItemIgnorado(skuRaw, prodRaw, "")) continue;
      if (cantBruta === 0) continue;

      const factor = obtenerFactorConversion(prodRaw);
      const cant = cantBruta * factor;

      let skuNorm = norm(skuRaw);
      const prodNorm = norm(prodRaw);

      if (!skuNorm || skuNorm === "NO ENCONTRADO") {
        if (mapaEquivalencias.has(prodNorm)) skuNorm = mapaEquivalencias.get(prodNorm);
      }

      if (skuNorm && skuNorm !== "NO ENCONTRADO" && idxPorSku.has(skuNorm)) {
        const filaIdx = idxPorSku.get(skuNorm);
        matrizConteos[filaIdx][sIdx] += cant;
        continue;
      }

      if (idxPorNombre.has(prodNorm)) {
        const filaIdx = idxPorNombre.get(prodNorm);
        matrizConteos[filaIdx][sIdx] += cant;
        continue;
      }

      if (skuNorm && skuNorm !== "NO ENCONTRADO" && idxNuevosPorSku.has(skuNorm)) {
        const nIdx = idxNuevosPorSku.get(skuNorm);
        itemsNuevosFinal[nIdx].conteos[sIdx] += cant;
      } else if (idxNuevosPorNombre.has(prodNorm)) {
        const nIdx = idxNuevosPorNombre.get(prodNorm);
        itemsNuevosFinal[nIdx].conteos[sIdx] += cant;
      } else {
        const rubroNuevo = skuNorm ? determinarRubroFinal(skuNorm, mapaRubrosMaestro) : "REVISAR";
        const nuevoObj = {
          rubro: rubroNuevo,
          un: unVal || 1,
          medida: medVal || "UN",
          sku: (skuNorm && skuNorm !== "NO ENCONTRADO") ? skuRaw : "",
          producto: prodRaw,
          conteos: new Array(cantSectores).fill(0)
        };
        nuevoObj.conteos[sIdx] = cant;
        const pos = itemsNuevosFinal.length;
        itemsNuevosFinal.push(nuevoObj);
        if (skuNorm && skuNorm !== "NO ENCONTRADO") idxNuevosPorSku.set(skuNorm, pos);
        if (prodNorm) idxNuevosPorNombre.set(prodNorm, pos);
      }
    }
  });

  // --- LECTURA DE CAJA ---
  const hojaCaja = getHoja(ss, CONFIG.HOJA_CAJA.nombre);
  if (hojaCaja) {
    const uFilaCaja = hojaCaja.getLastRow();
    if (uFilaCaja >= CONFIG.HOJA_CAJA.filaInicio) {
      const colExcluirCaja = obtenerColumnaExcluir(hojaCaja, CONFIG.HOJA_CAJA.colCantidad);
      const maxColsCaja = Math.max(CONFIG.HOJA_CAJA.colCantidad, colExcluirCaja);
      const datosCaja = hojaCaja.getRange(CONFIG.HOJA_CAJA.filaInicio, 1, uFilaCaja - CONFIG.HOJA_CAJA.filaInicio + 1, maxColsCaja).getValues();

      datosCaja.forEach(fila => {
        const skuRaw = String(fila[1] || "").trim();
        const cant = parseNum(fila[CONFIG.HOJA_CAJA.colCantidad - 1]);
        const excl = fila[colExcluirCaja - 1];
        const skuKey = norm(skuRaw);

        if (estaMarcadoExcluir(excl)) return;

        if (cant > 0 && CONFIG.HOJA_CAJA.skusPermitidos.some(p => norm(p) === skuKey)) {
          if (idxPorSku.has(skuKey)) {
            const idx = idxPorSku.get(skuKey);
            matrizConteos[idx][CONFIG.HOJA_CAJA.sectorDestinoIdx] += cant;
          }
        }
      });
    }
  }

  // --- CÁLCULO DE TOTALES POR FILA ---
  for (let i = 0; i < cantItemsFijos; i++) {
    let sumaFila = 0;
    for (let s = 0; s < cantSectores; s++) {
      matrizConteos[i][s] = Math.round(matrizConteos[i][s] * 100) / 100;
      sumaFila += matrizConteos[i][s];
    }
    matrizConteos[i][cantSectores] = Math.round(sumaFila * 100) / 100;
  }

  hojaDestino.getRange(2, 6, cantItemsFijos, cantSectores + 1).setValues(matrizConteos);

  const filaFinFijos = cantItemsFijos + 1;
  const filasTotalesHoja = hojaDestino.getLastRow();
  if (filasTotalesHoja > filaFinFijos) {
    hojaDestino.getRange(filaFinFijos + 1, 1, filasTotalesHoja - filaFinFijos, 14).clearContent();
  }

  if (itemsNuevosFinal.length > 0) {
    const filasParaInsertar = itemsNuevosFinal.map(it => {
      const cRound = it.conteos.map(c => Math.round(c * 100) / 100);
      const total = Math.round(cRound.reduce((a, b) => a + b, 0) * 100) / 100;
      return [
        it.rubro,
        it.un,
        it.medida,
        it.sku,
        it.producto,
        ...cRound,
        total
      ];
    });
    hojaDestino.getRange(filaFinFijos + 1, 1, filasParaInsertar.length, 14).setValues(filasParaInsertar);
  }

  SpreadsheetApp.getUi().alert(
    "✅ CONSOLIDACIÓN INMÓVIL COMPLETA",
    `• Catálogo maestro inalterado: ${cantItemsFijos} filas fijas.\n` +
    `• Ítems nuevos agregados al final: ${itemsNuevosFinal.length}.\n` +
    `El diseño, fuentes y columnas A a E se mantuvieron 100% intactos.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ==========================================================================
// 3. MOTOR DEL SNAPSHOT (LOCAL Y EXTERNO)
// ==========================================================================
function ejecutarRegistroSnapshotInterno_(ssLocal) {
  const hojaTotal2 = getHoja(ssLocal, CONFIG.HOJA_TOTAL2);
  if (!hojaTotal2 || hojaTotal2.getLastRow() < 2) {
    throw new Error("No hay datos consolidados en 'Total2' para generar el snapshot.");
  }

  const uFilaTotal2 = hojaTotal2.getLastRow();
  // Leemos columnas A a N (14 columnas: Rubro, Un, Med, SKU, Prod, 8 sectores, FINAL)
  const datosTotal2 = hojaTotal2.getRange(2, 1, uFilaTotal2 - 1, 14).getValues();

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
    const finalVal = parseNum(fila[13]); // Col N: Columna 14 (TOTAL FINAL)

    if (!prod || !esFilaProductoValida(fila[1], sku, fila[2], prod)) return;

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
    "¿Confirmás registrar el snapshot del stock actual de 'Total2' en 'Control_Semanal' y en el Historial externo?",
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
// 4. EXPORTAR A GESTIÓN + DISPARO AUTOMÁTICO DE SNAPSHOT
// ==========================================================================
function enviarAGestionMensual() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "⚠️ CIERRE SEMANAL ROTISERÍA",
    "¿Confirmás transferir los conteos a 'Gestión', generar el Snapshot y reiniciar sectores?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const tInicio = new Date().getTime();

  try {
    const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    const hojaTotal2 = getHoja(ssLocal, CONFIG.HOJA_TOTAL2);

    if (!hojaTotal2 || hojaTotal2.getLastRow() < 2) {
      throw new Error("Primero consolidá en 'Total2'.");
    }

    // --- PASO A: REGISTRAR SNAPSHOT AUTOMÁTICAMENTE ---
    let cantSnapshot = 0;
    try {
      cantSnapshot = ejecutarRegistroSnapshotInterno_(ssLocal);
    } catch (eSnap) {
      console.warn("Aviso en snapshot automático: " + eSnap.message);
    }

    // --- PASO B: TRANSFERENCIA A GESTIÓN ---
    const uFilaTotal2 = hojaTotal2.getLastRow();
    const datosTotal2 = hojaTotal2.getRange(2, 1, uFilaTotal2 - 1, 14).getValues();

    const mapaLocalPorSku = new Map();
    datosTotal2.forEach(fila => {
      const sku = norm(fila[3]);
      if (sku && sku !== "NO ENCONTRADO") {
        mapaLocalPorSku.set(sku, {
          rubro: fila[0],
          un: fila[1],
          medida: fila[2],
          skuRaw: fila[3],
          prod: fila[4],
          conteos: fila.slice(5, 13) // Sectores (Cols F a M)
        });
      }
    });

    const ssRemoto = SpreadsheetApp.openById(CONFIG.ID_GESTION_MENSUAL);
    let hojaRemota = getHoja(ssRemoto, CONFIG.NOMBRE_HOJA_INVENTARIO) || ssRemoto.getSheets()[0];

    const uFilaRemota = hojaRemota.getLastRow();
    const cantFilasRemotas = uFilaRemota >= CONFIG.GESTION_FILA_DATOS 
      ? (uFilaRemota - CONFIG.GESTION_FILA_DATOS + 1) 
      : 0;

    if (cantFilasRemotas === 0) {
      throw new Error("La hoja 'Inventario' de Gestión no tiene datos desde la fila 3.");
    }

    const matrizRemota = hojaRemota.getRange(CONFIG.GESTION_FILA_DATOS, 1, cantFilasRemotas, 18).getValues();

    const skusProcesados = new Set();
    const matrizConteosNuevos = [];  
    const matrizFormulasNuevas = []; 

    for (let i = 0; i < matrizRemota.length; i++) {
      const filaReal = i + CONFIG.GESTION_FILA_DATOS;
      const skuRemoto = norm(matrizRemota[i][3]); 

      let conteosSector = [0, 0, 0, 0, 0, 0, 0, 0];

      if (skuRemoto && mapaLocalPorSku.has(skuRemoto)) {
        conteosSector = mapaLocalPorSku.get(skuRemoto).conteos;
        skusProcesados.add(skuRemoto);
      }

      matrizConteosNuevos.push(conteosSector);

      matrizFormulasNuevas.push([
        `=SUM(H${filaReal}:O${filaReal})`,
        `=P${filaReal}*G${filaReal}`,
        `=IFERROR(Q${filaReal}/SUM(Q$3:Q$${uFilaRemota}),0)`
      ]);
    }

    // Inyección Masiva en Gestión
    hojaRemota.getRange(CONFIG.GESTION_FILA_DATOS, 8, cantFilasRemotas, 8).setValues(matrizConteosNuevos);
    hojaRemota.getRange(CONFIG.GESTION_FILA_DATOS, 16, cantFilasRemotas, 3).setFormulas(matrizFormulasNuevas);

    // Ítems no existentes al final
    const filasNuevas = [];
    mapaLocalPorSku.forEach((item, skuKey) => {
      if (!skusProcesados.has(skuKey)) {
        filasNuevas.push([
          item.rubro, item.un, item.medida, item.skuRaw, item.prod, "", "", ...item.conteos,
          "", "", "", "Nuevo item sin precio"
        ]);
      }
    });

    if (filasNuevas.length > 0) {
      const filaAlta = hojaRemota.getLastRow() + 1;
      hojaRemota.getRange(filaAlta, 1, filasNuevas.length, filasNuevas[0].length).setValues(filasNuevas);

      const formulasAltas = filasNuevas.map((_, idx) => {
        const f = filaAlta + idx;
        return [
          `=SUM(H${f}:O${f})`,
          `=P${f}*G${f}`,
          `=IFERROR(Q${f}/SUM(Q$3:Q$${filaAlta + filasNuevas.length - 1}),0)`
        ];
      });
      hojaRemota.getRange(filaAlta, 16, filasNuevas.length, 3).setFormulas(formulasAltas);
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

    const hojaCaja = getHoja(ssLocal, CONFIG.HOJA_CAJA.nombre);
    if (hojaCaja && hojaCaja.getLastRow() >= CONFIG.HOJA_CAJA.filaInicio) {
      hojaCaja.getRange(CONFIG.HOJA_CAJA.filaInicio, CONFIG.HOJA_CAJA.colCantidad, hojaCaja.getLastRow() - CONFIG.HOJA_CAJA.filaInicio + 1, 1).clearContent();
    }

    const tTotal = ((new Date().getTime() - tInicio) / 1000).toFixed(1);

    ui.alert(
      "🚀 CIERRE Y GESTIÓN COMPLETADA",
      `Operación finalizada en ${tTotal} segundos.\n\n` +
      `• Snapshot registrado: ${cantSnapshot} ítems archivados en "${CONFIG.HOJA_LOGS}" e Historial.\n` +
      `• Filas actualizadas en Gestión: ${cantFilasRemotas}\n` +
      `• Ítems nuevos incorporados: ${filasNuevas.length}\n` +
      `• Hojas de conteo limpiadas y listas para la próxima semana.`,
      ui.ButtonSet.OK
    );

  } catch (err) {
    ui.alert("❌ ERROR EN EL CIERRE", err.toString(), ui.ButtonSet.OK);
  }
}

// ==========================================================================
// 5. LIMPIEZA MANUAL (VUELVE A 0 CANTIDADES SIN BORRAR CATÁLOGO)
// ==========================================================================
function borrarConteosSectores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = SpreadsheetApp.getUi().alert(
    "⚠️ ATENCIÓN: LIMPIAR CONTEOS",
    "¿Estás seguro de que querés vaciar las cantidades de las hojas de conteo?",
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

  const hojaCaja = getHoja(ss, CONFIG.HOJA_CAJA.nombre);
  if (hojaCaja) {
    const uFilaC = hojaCaja.getLastRow();
    if (uFilaC >= CONFIG.HOJA_CAJA.filaInicio) {
      hojaCaja.getRange(CONFIG.HOJA_CAJA.filaInicio, CONFIG.HOJA_CAJA.colCantidad, uFilaC - CONFIG.HOJA_CAJA.filaInicio + 1, 1).clearContent();
    }
  }

  const hojaTotal2 = getHoja(ss, CONFIG.HOJA_TOTAL2);
  if (hojaTotal2 && hojaTotal2.getLastRow() > 1) {
    const filas = hojaTotal2.getLastRow() - 1;
    const ceros = Array.from({ length: filas }, () => new Array(9).fill(0));
    hojaTotal2.getRange(2, 6, filas, 9).setValues(ceros);
  }

  SpreadsheetApp.getUi().alert("🧹 Limpieza Completa", "Se vaciaron las cantidades y se resetearon los conteos a 0 sin tocar el catálogo.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ==========================================================================
// MENÚ DE LA BARRA SUPERIOR
// ==========================================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ CONTROL MISHIGUENE")
    .addItem("1️⃣ Sincronizar SKU en Sectores", "prepararColumnasSkuEnSectores")
    .addItem("2️⃣ Consolidar Conteos en 'Total2'", "consolidarEnTotal2")
    .addItem("📸 Guardar Snapshot en Historial (Local y Externo)", "guardarSnapshotControlSemanal")
    .addSeparator()
    .addItem("🚀 3. Cierre Semanal (Snapshot + Gestión + Limpieza)", "enviarAGestionMensual")
    .addItem("🧹 Limpiar Conteos Manualmente", "borrarConteosSectores")
    .addToUi();
}