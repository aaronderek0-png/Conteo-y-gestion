// ==========================================================================
// SISTEMA INTEGRAL: CONTEO OBRADOR -> GESTIÓN CONGRESO + SNAPSHOT SKU/TOTAL
// ==========================================================================

const CONFIG = {
  // --- CONFIGURACIÓN PLANILLA LOCAL (OBRADOR) ---
  HOJA_TOTAL: "Total",
  HOJA_EQUIVALENCIAS: "Equivalencias",
  HOJA_LOGS: "Control_Semanal",
  FILA_INICIO_TOTAL: 2,

  // --- MAESTRO PRODUCTOS (BASE CENTRAL DE RUBROS) ---
  ID_MAESTRO_PRODUCTOS: "1lOabw9U4ilRwE5r6QXQ2Jc1nqkKUK3Hi096xga0XFnI",
  HOJA_MAESTRO_PRODUCTOS: "Productos",

  // Solapas operativas de Obrador (Cols F, G, H en Total)
  SECTORES: [
    { nombre: "PRODUCCIÓN", colIdx: 0, filaInicio: 5 }, // Col F
    { nombre: "PASTELERÍA", colIdx: 1, filaInicio: 5 }, // Col G
    { nombre: "DEPÓSITO",   colIdx: 2, filaInicio: 5 }  // Col H
  ],

  // --- CONFIGURACIÓN PLANILLA DESTINO (GESTIÓN CONGRESO) ---
  ID_GESTION: "13py03p7PzV1xEHcjSbhrI_I2DTY9M2QKGDLbZ-aKu5c",
  HOJA_DESTINO_GESTION: "Inventario",
  GESTION_FILA_DATOS: 3,         // Datos inician en fila 3
  GESTION_COL_DESTINO_INICIO: 9,  // Columnas I, J, K, L (Producción, Pastelería, Depósito, FINAL)

  // --- CONFIGURACIÓN HISTORIAL EXTERNO ---
  ID_HISTORIAL_EXTERNO: "1bnJp0suZ8pBoE3AFMnLMdzTy5frC3e3bDZ5XVbBzrLQ",
  HOJA_HISTORIAL_EXTERNO: "Historial"
};

const PALABRAS_IGNORAR = [
  "ALMACEN", "ALMACÉN", "LACTEOS", "LÁCTEOS", "GASEOSAS", "BEBIDAS", "DESCARTABLES", 
  "LIMPIEZA", "SUB RUBRO", "RESPONSABLE", "AREA:", "ÁREA:", "UTILES", "ÚTILES", "TAPERS", 
  "HELADERA", "UN.", "UN", "MEDIDA", "PRODUCTO", "PRODUCCION", "PRODUCCIÓN", "QUESOS", 
  "LAMINADOS", "PANES OBRADOR", "HARINERA", "PRODUCTOS TERCERIZADOS", "CAFETERIA", 
  "CAFETERÍA", "PRUEBAS Y OTROS", "OBRADOR", "COCINA PRODUCCIÓN", "COCINA PRODUCCIÓN - PASTELERÍA",
  "DEPÓSITO", "DEPOSITO"
];

function norm(texto) {
  if (!texto) return "";
  return texto.toString().trim().toUpperCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Z0-9 ]/g, "")
    .replace(/\s+/g, " ");
}

function esTituloOSeccion(texto, unVal, medVal) {
  const t = norm(texto);
  if (!t) return true;
  if (norm(unVal) === "UN" || norm(medVal) === "MEDIDA") return true;
  return PALABRAS_IGNORAR.some(p => t === norm(p) || t.startsWith(norm(p)));
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

// --------------------------------------------------------------------------
// CARGA EL DICCIONARIO DE RUBROS DEL MAESTRO DE PRODUCTOS
// --------------------------------------------------------------------------
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
        const skuNorm = norm(f[0]);              // Col A: SKU
        const rubro = String(f[3] || "").trim(); // Col D: Rubro / Categoría
        if (skuNorm && rubro) {
          mapa.set(skuNorm, rubro);
        }
      });
    }
  } catch (err) {
    console.error("Error al cargar rubros desde el Maestro: " + err);
  }
  return mapa;
}

function determinarRubroFinal(skuNorm, mapaRubrosMaestro) {
  if (!skuNorm || skuNorm === "NO ENCONTRADO") {
    return "OBRADOR";
  }
  if (mapaRubrosMaestro.has(skuNorm)) {
    return mapaRubrosMaestro.get(skuNorm);
  }
  if (skuNorm.startsWith("OBR") || /^P\d+/.test(skuNorm)) {
    return "OBRADOR";
  }
  return "OBRADOR";
}

// --------------------------------------------------------------------------
// 1. SINCRONIZAR COLUMNA SKU EN SECTORES (RESPETA CARGAS MANUALES)
// --------------------------------------------------------------------------
function prepararColumnasSkuEnSectores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTotal = getHoja(ss, CONFIG.HOJA_TOTAL);
  if (!hojaTotal) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la hoja 'Total'.");
    return;
  }

  const uFilaTotal = hojaTotal.getLastRow();
  const cantFilas = uFilaTotal - CONFIG.FILA_INICIO_TOTAL + 1;
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

  CONFIG.SECTORES.forEach(sec => {
    const hoja = getHoja(ss, sec.nombre);
    if (!hoja) return;

    const cabecerasB = hoja.getRange(1, 2, Math.min(5, hoja.getLastRow()), 1).getValues().flat().map(c => norm(c));
    const tieneSku = cabecerasB.some(c => c === "SKU" || c === "CODIGO" || c === "COD");

    if (!tieneSku) {
      hoja.insertColumnBefore(2);
      hoja.getRange(4, 2).setValue("SKU").setFontWeight("bold");
    }

    const uFila = hoja.getLastRow();
    if (uFila >= sec.filaInicio) {
      const cantFilasSector = uFila - sec.filaInicio + 1;
      const skusActuales = hoja.getRange(sec.filaInicio, 2, cantFilasSector, 1).getValues();
      const prodsSector = hoja.getRange(sec.filaInicio, 4, cantFilasSector, 1).getValues();
      const skusGenerados = [];

      for (let i = 0; i < prodsSector.length; i++) {
        const skuExistente = String(skusActuales[i][0] || "").trim();
        const pName = norm(prodsSector[i][0]);

        if (skuExistente !== "") {
          skusGenerados.push([skuExistente]); // Preserva SKU manual
        } else if (pName && !esTituloOSeccion(pName) && mapaCodigoPorNombre.has(pName)) {
          skusGenerados.push([mapaCodigoPorNombre.get(pName)]);
        } else {
          skusGenerados.push([""]);
        }
      }
      hoja.getRange(sec.filaInicio, 2, skusGenerados.length, 1).setValues(skusGenerados);
    }
  });

  let hojaEq = getHoja(ss, CONFIG.HOJA_EQUIVALENCIAS);
  if (!hojaEq) {
    hojaEq = ss.insertSheet(CONFIG.HOJA_EQUIVALENCIAS);
    hojaEq.getRange(1, 1, 1, 3).setValues([["TEXTO_EN_CONTEO", "SKU_MAESTRO", "DESCRIPCION_OPCIONAL"]]);
    hojaEq.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert(
    "✅ SKU SINCRONIZADOS",
    "Se actualizaron los códigos en Producción, Pastelería y Depósito respetando cargas manuales.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// --------------------------------------------------------------------------
// 2. CONSOLIDAR EN TOTAL (ACUMULACIÓN POR SKU Y CRUCE CON MAESTRO)
// --------------------------------------------------------------------------
function consolidarConteosEnTotal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaTotal = getHoja(ss, CONFIG.HOJA_TOTAL);
  if (!hojaTotal) {
    SpreadsheetApp.getUi().alert("❌ No se encontró la hoja 'Total'.");
    return;
  }

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

  const uFilaTotal = hojaTotal.getLastRow();
  const cantFilas = uFilaTotal >= CONFIG.FILA_INICIO_TOTAL ? (uFilaTotal - CONFIG.FILA_INICIO_TOTAL + 1) : 0;
  const datosTotal = cantFilas > 0 ? hojaTotal.getRange(CONFIG.FILA_INICIO_TOTAL, 1, cantFilas, 5).getValues() : [];

  const itemsMaestros = [];
  const idxPorSku = new Map();
  const idxPorNombre = new Map();

  for (let i = 0; i < datosTotal.length; i++) {
    let rubro = datosTotal[i][0] || "ALMACEN";
    const un = datosTotal[i][1] || 1;
    const med = datosTotal[i][2] || "UN";
    const sku = String(datosTotal[i][3] || "").trim();
    const prod = String(datosTotal[i][4] || "").trim();

    if (!prod || esTituloOSeccion(prod, un, med)) continue;

    let skuNorm = norm(sku);
    const prodNorm = norm(prod);

    if (!skuNorm || skuNorm === "NO ENCONTRADO") {
      if (mapaEquivalencias.has(prodNorm)) {
        skuNorm = mapaEquivalencias.get(prodNorm);
      }
    }

    if (skuNorm && skuNorm !== "NO ENCONTRADO" && idxPorSku.has(skuNorm)) {
      continue;
    }

    rubro = determinarRubroFinal(skuNorm, mapaRubrosMaestro);

    const itemObj = {
      rubro: rubro,
      un: un,
      medida: med,
      sku: (skuNorm && skuNorm !== "NO ENCONTRADO") ? sku : "No encontrado",
      producto: prod,
      conteos: [0, 0, 0]
    };

    const newIdx = itemsMaestros.length;
    itemsMaestros.push(itemObj);

    if (skuNorm && skuNorm !== "NO ENCONTRADO") idxPorSku.set(skuNorm, newIdx);
    if (prodNorm) idxPorNombre.set(prodNorm, newIdx);
  }

  CONFIG.SECTORES.forEach(sec => {
    const hojaSec = getHoja(ss, sec.nombre);
    if (!hojaSec) return;

    const uFila = hojaSec.getLastRow();
    if (uFila < sec.filaInicio) return;

    const datosSec = hojaSec.getRange(sec.filaInicio, 1, uFila - sec.filaInicio + 1, 5).getValues();

    for (let r = 0; r < datosSec.length; r++) {
      const fila = datosSec[r];
      const unVal = fila[0] || 1;
      const skuSec = norm(fila[1]);
      const medVal = fila[2] || "UN";
      const prodRaw = String(fila[3] || "").trim();
      const cant = parseNum(fila[4]);

      if (!prodRaw || esTituloOSeccion(prodRaw, unVal, medVal) || cant === 0) continue;

      let skuNorm = skuSec;
      const prodNorm = norm(prodRaw);

      if (!skuNorm || skuNorm === "NO ENCONTRADO") {
        if (mapaEquivalencias.has(prodNorm)) {
          skuNorm = mapaEquivalencias.get(prodNorm);
        }
      }

      let idxDestino = -1;

      if (skuNorm && skuNorm !== "NO ENCONTRADO" && idxPorSku.has(skuNorm)) {
        idxDestino = idxPorSku.get(skuNorm);
      } else if (mapaEquivalencias.has(prodNorm)) {
        const skuAlias = mapaEquivalencias.get(prodNorm);
        if (idxPorSku.has(skuAlias)) idxDestino = idxPorSku.get(skuAlias);
      } else if (idxPorNombre.has(prodNorm)) {
        idxDestino = idxPorNombre.get(prodNorm);
      }

      if (idxDestino !== -1) {
        itemsMaestros[idxDestino].conteos[sec.colIdx] += cant;
        itemsMaestros[idxDestino].rubro = determinarRubroFinal(skuNorm, mapaRubrosMaestro);
      } else {
        const rubroNuevo = determinarRubroFinal(skuNorm, mapaRubrosMaestro);

        const nuevoItem = {
          rubro: rubroNuevo,
          un: unVal,
          medida: medVal,
          sku: (skuNorm && skuNorm !== "NO ENCONTRADO") ? fila[1] : "No encontrado",
          producto: prodRaw,
          conteos: [0, 0, 0]
        };
        nuevoItem.conteos[sec.colIdx] = cant;

        const nuevoIdx = itemsMaestros.length;
        itemsMaestros.push(nuevoItem);

        if (skuNorm && skuNorm !== "NO ENCONTRADO") idxPorSku.set(skuNorm, nuevoIdx);
        idxPorNombre.set(prodNorm, nuevoIdx);
      }
    }
  });

  const itemsConSku = [];
  const itemsSinSku = [];

  itemsMaestros.forEach(item => {
    const prod = Math.round(item.conteos[0] * 1000) / 1000;
    const past = Math.round(item.conteos[1] * 1000) / 1000;
    const dep  = Math.round(item.conteos[2] * 1000) / 1000;
    const finalVal = Math.round((prod + past + dep) * 1000) / 1000;

    const filaSalida = [
      item.rubro,
      item.un,
      item.medida,
      item.sku,
      item.producto,
      prod,
      past,
      dep,
      finalVal
    ];

    if (item.sku && item.sku.toUpperCase() !== "NO ENCONTRADO" && item.sku !== "") {
      itemsConSku.push(filaSalida);
    } else {
      itemsSinSku.push(filaSalida);
    }
  });

  const salidaTotal = [...itemsConSku, ...itemsSinSku];

  const filter = hojaTotal.getFilter();
  if (filter) filter.remove();

  if (uFilaTotal >= CONFIG.FILA_INICIO_TOTAL) {
    hojaTotal.getRange(CONFIG.FILA_INICIO_TOTAL, 1, Math.max(uFilaTotal - CONFIG.FILA_INICIO_TOTAL + 1, 1), 9).clearContent();
  }

  if (salidaTotal.length > 0) {
    hojaTotal.getRange(CONFIG.FILA_INICIO_TOTAL, 1, salidaTotal.length, 9).setValues(salidaTotal);
  }

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    "⚡ CONSOLIDACIÓN COMPLETADA",
    `Se consolidaron ${salidaTotal.length} productos en 'Total'.\n\n` +
    `• Ítems con SKU ordenados arriba.\n` +
    `• Ítems sin SKU agrupados al fondo.\n` +
    `• Suma consolidada en Columna I (FINAL).`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// --------------------------------------------------------------------------
// 3. SNAPSHOT EN CONTROL_SEMANAL Y PLANILLA EXTERNA (HISTORIAL)
// --------------------------------------------------------------------------
function guardarSnapshotControlSemanal() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "📸 GUARDAR SNAPSHOT",
    "¿Confirmás guardar el snapshot consolidado en 'Control_Semanal' y en la planilla externa 'Historial'?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaTotal = getHoja(ss, CONFIG.HOJA_TOTAL);
    if (!hojaTotal) throw new Error("No se encontró la hoja 'Total'.");

    const uFilaTotal = hojaTotal.getLastRow();
    const totalFilas = uFilaTotal >= CONFIG.FILA_INICIO_TOTAL ? (uFilaTotal - CONFIG.FILA_INICIO_TOTAL + 1) : 0;
    if (totalFilas <= 0) throw new Error("No hay datos en la hoja 'Total'.");

    const datosTotal = hojaTotal.getRange(CONFIG.FILA_INICIO_TOTAL, 1, totalFilas, 9).getValues();

    // 1. Preparar hoja local Control_Semanal
    let hojaLogs = getHoja(ss, CONFIG.HOJA_LOGS);
    if (!hojaLogs) {
      hojaLogs = ss.insertSheet(CONFIG.HOJA_LOGS);
      const headers = ["FECHA", "PRODUCTO", "CODIGO", "TOTAL FINAL"];
      hojaLogs.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
      hojaLogs.setFrozenRows(1);
    }

    // 2. Preparar hoja destino externa (Historial)
    const ssExterna = SpreadsheetApp.openById(CONFIG.ID_HISTORIAL_EXTERNO);
    let hojaExterna = getHoja(ssExterna, CONFIG.HOJA_HISTORIAL_EXTERNO);
    if (!hojaExterna) {
      hojaExterna = ssExterna.insertSheet(CONFIG.HOJA_HISTORIAL_EXTERNO);
      const headersExt = ["FECHA REGISTRO", "PRODUCTO", "CODIGO", "FINAL"];
      hojaExterna.getRange(1, 1, 1, headersExt.length).setValues([headersExt]).setFontWeight("bold");
      hojaExterna.setFrozenRows(1);
    }

    const fecha = new Date();
    const registrosSnapshot = [];

    datosTotal.forEach(fila => {
      const sku = String(fila[3] || "").trim();
      const prod = String(fila[4] || "").trim();
      const finalVal = parseNum(fila[8]); // Col I: FINAL

      if (!prod || esTituloOSeccion(prod, fila[1], fila[2])) return;

      // Solo guardamos si hubo stock contado
      if (finalVal > 0) {
        // [FECHA (A), PRODUCTO (B), CODIGO (C), TOTAL FINAL (D)]
        registrosSnapshot.push([
          fecha,
          prod,
          (sku && sku.toUpperCase() !== "NO ENCONTRADO") ? sku : "No encontrado",
          finalVal
        ]);
      }
    });

    if (registrosSnapshot.length === 0) {
      throw new Error("No hay productos con conteo mayor a 0 para registrar.");
    }

    // Escribir en hoja local
    const uFilaLogs = hojaLogs.getLastRow();
    hojaLogs.getRange(uFilaLogs + 1, 1, registrosSnapshot.length, 4).setValues(registrosSnapshot);

    // Escribir en hoja externa
    const uFilaExt = hojaExterna.getLastRow();
    hojaExterna.getRange(uFilaExt + 1, 1, registrosSnapshot.length, 4).setValues(registrosSnapshot);

    ui.alert(
      "✅ SNAPSHOT GUARDADO EN AMBAS PLANILLAS",
      `Se registraron ${registrosSnapshot.length} productos en:\n\n` +
      `1. Local: "${CONFIG.HOJA_LOGS}"\n` +
      `2. Externa: "${hojaExterna.getName()}" (ID: ${CONFIG.ID_HISTORIAL_EXTERNO})\n\n` +
      `• Formato: Fecha | Producto | Código | Final`,
      ui.ButtonSet.OK
    );

  } catch (err) {
    ui.alert("❌ ERROR EN SNAPSHOT", err.toString(), ui.ButtonSet.OK);
  }
}

// --------------------------------------------------------------------------
// 4. EXPORTAR A GESTIÓN CONGRESO + REGISTRO AUTOMÁTICO DE SNAPSHOT
// --------------------------------------------------------------------------
function exportarAGestionYLimpiar() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "⚠️ CIERRE SEMANAL",
    "¿Confirmás transferir los conteos a Gestión Congreso (Inventario), registrar snapshot con SKU y reiniciar sectores?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  try {
    const ssLocal = SpreadsheetApp.getActiveSpreadsheet();
    const hojaTotal = getHoja(ssLocal, CONFIG.HOJA_TOTAL);
    if (!hojaTotal) throw new Error("No se encontró la hoja 'Total'.");

    const uFilaTotal = hojaTotal.getLastRow();
    const totalFilas = uFilaTotal >= CONFIG.FILA_INICIO_TOTAL ? (uFilaTotal - CONFIG.FILA_INICIO_TOTAL + 1) : 0;
    if (totalFilas <= 0) throw new Error("No hay datos en la hoja 'Total'.");

    const datosTotal = hojaTotal.getRange(CONFIG.FILA_INICIO_TOTAL, 1, totalFilas, 9).getValues();

    const libroGestion = SpreadsheetApp.openById(CONFIG.ID_GESTION);
    let hojaGestion = getHoja(libroGestion, CONFIG.HOJA_DESTINO_GESTION);
    if (!hojaGestion) {
      hojaGestion = libroGestion.getSheets()[0];
    }

    const uFilaGestion = hojaGestion.getLastRow();
    const cantFilasGestion = uFilaGestion >= CONFIG.GESTION_FILA_DATOS ? (uFilaGestion - CONFIG.GESTION_FILA_DATOS + 1) : 0;

    const mapaGestionPorSku = new Map();
    const mapaGestionPorNombre = new Map();
    const skusConocidos = new Set();
    const filasDuplicadasParaBorrar = [];

    if (cantFilasGestion > 0) {
      const clavesGestion = hojaGestion.getRange(CONFIG.GESTION_FILA_DATOS, 4, cantFilasGestion, 2).getValues();
      for (let i = 0; i < clavesGestion.length; i++) {
        const cod = norm(clavesGestion[i][0]);
        const prod = norm(clavesGestion[i][1]);
        const numFila = i + CONFIG.GESTION_FILA_DATOS;

        if (cod && cod !== "NO ENCONTRADO") {
          if (!mapaGestionPorSku.has(cod)) {
            mapaGestionPorSku.set(cod, { fila: numFila, codRaw: clavesGestion[i][0] });
            skusConocidos.add(cod);
          } else {
            filasDuplicadasParaBorrar.push(numFila);
          }
        }
        if (prod && !mapaGestionPorNombre.has(prod)) {
          mapaGestionPorNombre.set(prod, { fila: numFila, codRaw: clavesGestion[i][0] });
        }
      }
    }

    if (filasDuplicadasParaBorrar.length > 0) {
      for (let idx = filasDuplicadasParaBorrar.length - 1; idx >= 0; idx--) {
        hojaGestion.deleteRow(filasDuplicadasParaBorrar[idx]);
      }
    }

    const updatesConteos = new Map();
    const updatesSkus = new Map();
    const filasParaInsertarAbajo = [];

    datosTotal.forEach(fila => {
      const rubro = fila[0];
      const un = fila[1];
      const med = fila[2];
      const skuRaw = String(fila[3] || "").trim();
      const prodRaw = String(fila[4] || "").trim();
      const conteos4 = fila.slice(5, 9); // [PROD, PAST, DEPO, FINAL]

      if (!prodRaw || esTituloOSeccion(prodRaw, un, med)) return;

      const skuKey = norm(skuRaw);
      const prodKey = norm(prodRaw);

      let filaDestino = null;
      let codActualEnGestion = "";

      if (skuKey && skuKey !== "NO ENCONTRADO" && mapaGestionPorSku.has(skuKey)) {
        filaDestino = mapaGestionPorSku.get(skuKey).fila;
        codActualEnGestion = mapaGestionPorSku.get(skuKey).codRaw;
      } else if (mapaGestionPorNombre.has(prodKey)) {
        filaDestino = mapaGestionPorNombre.get(prodKey).fila;
        codActualEnGestion = mapaGestionPorNombre.get(prodKey).codRaw;
      }

      if (filaDestino) {
        updatesConteos.set(filaDestino, conteos4);

        const codNormGestion = norm(codActualEnGestion);
        if (skuKey && skuKey !== "NO ENCONTRADO" && (!codNormGestion || codNormGestion === "NO ENCONTRADO")) {
          updatesSkus.set(filaDestino, skuRaw);
        }
      } else {
        if (skuKey && skuKey !== "NO ENCONTRADO" && skusConocidos.has(skuKey)) {
          return;
        }

        filasParaInsertarAbajo.push([
          rubro, un, med, skuRaw, prodRaw, "", "", "", ...conteos4
        ]);
      }
    });

    updatesConteos.forEach((valores4, nFila) => {
      hojaGestion.getRange(nFila, CONFIG.GESTION_COL_DESTINO_INICIO, 1, 4).setValues([valores4]);
    });

    updatesSkus.forEach((nuevoSku, nFila) => {
      hojaGestion.getRange(nFila, 4).setValue(nuevoSku);
    });

    if (filasParaInsertarAbajo.length > 0) {
      const filaDestinoInsert = hojaGestion.getLastRow() + 1;
      hojaGestion.getRange(filaDestinoInsert, 1, filasParaInsertarAbajo.length, 12).setValues(filasParaInsertarAbajo);
    }

    // --- GUARDADO DE SNAPSHOT (LOCAL Y EXTERNO) ---
    const fecha = new Date();
    const registrosSnapshot = [];

    datosTotal.forEach(fila => {
      const sku = String(fila[3] || "").trim();
      const prod = String(fila[4] || "").trim();
      const finalVal = parseNum(fila[8]);

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
      // 1. En hoja local
      let hojaLogs = getHoja(ssLocal, CONFIG.HOJA_LOGS);
      if (!hojaLogs) {
        hojaLogs = ssLocal.insertSheet(CONFIG.HOJA_LOGS);
        const headers = ["FECHA", "PRODUCTO", "CODIGO", "TOTAL FINAL"];
        hojaLogs.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
        hojaLogs.setFrozenRows(1);
      }
      hojaLogs.getRange(hojaLogs.getLastRow() + 1, 1, registrosSnapshot.length, 4).setValues(registrosSnapshot);

      // 2. En hoja externa (Historial)
      try {
        const ssExterna = SpreadsheetApp.openById(CONFIG.ID_HISTORIAL_EXTERNO);
        let hojaExterna = getHoja(ssExterna, CONFIG.HOJA_HISTORIAL_EXTERNO);
        if (!hojaExterna) {
          hojaExterna = ssExterna.insertSheet(CONFIG.HOJA_HISTORIAL_EXTERNO);
          const headersExt = ["FECHA REGISTRO", "PRODUCTO", "CODIGO", "FINAL"];
          hojaExterna.getRange(1, 1, 1, headersExt.length).setValues([headersExt]).setFontWeight("bold");
          hojaExterna.setFrozenRows(1);
        }
        hojaExterna.getRange(hojaExterna.getLastRow() + 1, 1, registrosSnapshot.length, 4).setValues(registrosSnapshot);
      } catch (errExt) {
        console.error("No se pudo escribir en el Historial externo durante el cierre: " + errExt);
      }
    }

    // Vaciar cantidades en las 3 solapas de conteo
    CONFIG.SECTORES.forEach(sec => {
      const hojaSector = getHoja(ssLocal, sec.nombre);
      if (!hojaSector) return;
      const uFila = hojaSector.getLastRow();
      if (uFila >= sec.filaInicio) {
        hojaSector.getRange(sec.filaInicio, 5, uFila - sec.filaInicio + 1, 1).clearContent();
      }
    });

    ui.alert(
      "🚀 SINCRONIZACIÓN Y CIERRE EXITOSO",
      `Planilla Gestión: Solapa "${hojaGestion.getName()}"\n\n` +
      `• Actualizados: ${updatesConteos.size} conteos en Cols I, J, K, L.\n` +
      `• SKUs actualizados en Col D: ${updatesSkus.size}.\n` +
      (filasDuplicadasParaBorrar.length > 0 ? `• Eliminadas ${filasDuplicadasParaBorrar.length} filas duplicadas viejas en Gestión.\n` : "") +
      (filasParaInsertarAbajo.length > 0 ? `• Insertados al fondo: ${filasParaInsertarAbajo.length} ítems nuevos.\n` : "• Sin productos nuevos para dar de alta.\n") +
      `• Snapshot guardado en "${CONFIG.HOJA_LOGS}" y en "Historial" externo.\n` +
      `• Conteos reiniciados.`,
      ui.ButtonSet.OK
    );

  } catch (err) {
    ui.alert("❌ ERROR EN EL CIERRE", err.toString(), ui.ButtonSet.OK);
  }
}

// --------------------------------------------------------------------------
// 5. LIMPIEZA MANUAL DE CONTEOS
// --------------------------------------------------------------------------
function borrarConteosSectores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = SpreadsheetApp.getUi().alert(
    "⚠️ ATENCIÓN: LIMPIAR CONTEOS",
    "¿Confirmás vaciar las cantidades de las 3 solapas de conteo?",
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (resp !== SpreadsheetApp.getUi().Button.YES) return;

  CONFIG.SECTORES.forEach(sec => {
    const hoja = getHoja(ss, sec.nombre);
    if (!hoja) return;

    const uFila = hoja.getLastRow();
    if (uFila >= sec.filaInicio) {
      hoja.getRange(sec.filaInicio, 5, uFila - sec.filaInicio + 1, 1).clearContent();
    }
  });

  SpreadsheetApp.getUi().alert("🧹 Limpieza Completa", "Se vaciaron las cantidades en los sectores.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// --------------------------------------------------------------------------
// MENÚ DE INTERFAZ EN GOOGLE SHEETS
// --------------------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ OBRADOR CONTROL")
    .addItem("1️⃣ Sincronizar SKU en Sectores", "prepararColumnasSkuEnSectores")
    .addItem("2️⃣ Consolidar Conteos en 'Total'", "consolidarConteosEnTotal")
    .addItem("📸 Guardar Snapshot en Historial (Local y Externo)", "guardarSnapshotControlSemanal")
    .addSeparator()
    .addItem("🚀 3. Cierre Semanal (Enviar a Gestión y Limpiar)", "exportarAGestionYLimpiar")
    .addItem("🧹 Limpiar Conteos Manualmente", "borrarConteosSectores")
    .addToUi();
}