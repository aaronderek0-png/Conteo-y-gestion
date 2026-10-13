/**
 * SISTEMA INTEGRAL DE CONTROL OPERATIVO, CATALOGACIÓN Y PRECIOS
 * Archivo: Local (Facturas / Desperdicios) - CONGRESO
 */

const CONFIG_SOLICITUDES = {
  EMAIL_DESTINO: "mishiguene.cervino5@gmail.com",
  MAESTRO_ID: "1lOabw9U4ilRwE5r6QXQ2Jc1nqkKUK3Hi096xga0XFnI",
  HOJA_SOLICITUDES: "Solicitudes",
  LOCAL: "CONGRESO", 
  UMBRAL_ALERTA_PORC: 6.0
};

// ============================================================================
// TRIGGERS GLOBALES (ONOPEN / ONEDIT)
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚡ CONGRESO")
    .addItem("⚡ Prorratear Gastos (Llevar a $0)", "distribuirGastosCero")
    .addItem("🧹 Refrescar validaciones de catálogo", "botonRefrescarValidaciones")
    .addSeparator()
    .addItem("⏰ Instalar auditoría diaria (11:00 AM)", "instalarTriggerDiario11AM")
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName().trim();
  const row = e.range.rowStart;
  const col = e.range.columnStart;
  const numCols = e.range.numColumns || 1;

  // CASO 1: HOJA "Desperdicios"
  if (sheetName === "Desperdicios" && col === 4 && row >= 2) {
    const targetCell = sheet.getRange(row, 5);
    const val = (e.value || "").toString().trim().toLowerCase();

    if (!val) {
      targetCell.clearContent().clearDataValidations();
      return;
    }

    const ss = e.source;
    const listaSheet = ss.getSheetByName("Lista");
    if (!listaSheet) return;

    let colIndex = 0;
    if (val === "insumo") colIndex = 1;
    else if (val === "receta") colIndex = 2;
    else if (val === "plato") colIndex = 3;

    if (colIndex === 0) {
      targetCell.clearContent().clearDataValidations();
      return;
    }

    const lastRow = listaSheet.getLastRow();
    if (lastRow < 2) return;

    const data = listaSheet.getRange(2, colIndex, lastRow - 1, 1).getValues();
    const items = [];
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] !== "" && data[i][0] !== null && data[i][0] !== undefined) {
        items.push(data[i][0].toString());
      }
    }

    if (items.length > 0) {
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(items, true)
        .setAllowInvalid(true)
        .build();

      targetCell.clearContent();
      targetCell.setDataValidation(rule);
    } else {
      targetCell.clearContent().clearDataValidations();
    }
    return;
  }

  // CASO 2: HOJA "Facturas"
  if (sheetName === "Facturas" && row >= 2) {
    if ((col === 4 || col === 5) && numCols === 1) {
      registrarItemNuevo(e);
    }

    if (col === 5 || col === 8) {
      procesarProrrateoACero_(sheet, row);
    }
  }
}

// ============================================================================
// NORMALIZADOR Y CONCEPTOS
// ============================================================================

function normalizarTexto(txt) {
  return String(txt || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
}

function esConceptoGasto_(nombre) {
  const txt = normalizarTexto(nombre);
  if (!txt || txt.includes("flete")) return false;
  return txt.includes("envio") || 
         txt.includes("impuesto") || 
         txt.includes("interno") ||
         txt.includes("percepcion") ||
         txt.includes("perc");
}

function registrarItemNuevo(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Facturas") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2 || (col !== 4 && col !== 5)) return;

  const valorIngresado = String(e.value || e.range.getValue() || "").trim();
  if (!valorIngresado) return;

  const valorNorm = normalizarTexto(valorIngresado);
  const ssLocal = e.source;
  const tipo = (col === 4) ? "PROVEEDOR" : "PRODUCTO";

  let listaMaestra = [];
  if (tipo === "PRODUCTO") {
    const hojaProd = ssLocal.getSheetByName("Productos") || ssLocal.getSheetByName("Listas") || ssLocal.getSheetByName("Lista");
    if (hojaProd) listaMaestra = hojaProd.getDataRange().getValues().flat();
  } else {
    const hojaLista = ssLocal.getSheetByName("Listas") || ssLocal.getSheetByName("Lista") || ssLocal.getSheetByName("CONFIGURACION");
    if (hojaLista) listaMaestra = hojaLista.getDataRange().getValues().flat();
  }

  if (listaMaestra.some(item => normalizarTexto(item) === valorNorm && valorNorm !== "")) {
    e.range.setBackground(null);
    e.range.clearNote();
    return;
  }

  try {
    const ssExterno = SpreadsheetApp.openById(CONFIG_SOLICITUDES.MAESTRO_ID);
    const hojaSolicitudes = ssExterno.getSheetByName(CONFIG_SOLICITUDES.HOJA_SOLICITUDES) || ssExterno.getSheets()[0];

    const ultimaFilaSol = hojaSolicitudes.getLastRow();
    if (ultimaFilaSol > 1) {
      const solicitudesExistentes = hojaSolicitudes.getRange(2, 1, ultimaFilaSol - 1, 13).getValues();
      const yaExiste = solicitudesExistentes.some(filaSol => {
        const itemSol = (tipo === "PRODUCTO") ? filaSol[6] : filaSol[10];
        return normalizarTexto(itemSol) === valorNorm;
      });
      if (yaExiste) return;
    }

    const fechaHoy = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy HH:mm");
    const proveedorFila = (tipo === "PROVEEDOR") ? valorIngresado : String(sheet.getRange(row, 4).getValue() || "Sin Proveedor");

    hojaSolicitudes.appendRow([
      fechaHoy, CONFIG_SOLICITUDES.LOCAL, tipo, "", "",
      tipo === "PRODUCTO" ? valorIngresado : "",
      tipo === "PRODUCTO" ? valorIngresado : "",
      tipo === "PRODUCTO" ? 1 : "", tipo === "PRODUCTO" ? "UN" : "",
      tipo === "PRODUCTO" ? 1 : "", proveedorFila, "PENDIENTE", ""
    ]);

    e.range.setBackground("#FFF2CC");
    e.range.setNote("⚠️ Ítem no catalogado. Solicitud enviada a administración.");

    MailApp.sendEmail(
      CONFIG_SOLICITUDES.EMAIL_DESTINO,
      `🔔 Nueva Solicitud de ${tipo}: ${valorIngresado} (${CONFIG_SOLICITUDES.LOCAL})`,
      `Se ingresó un nuevo ${tipo.toLowerCase()} no catalogado en ${CONFIG_SOLICITUDES.LOCAL}.\n\nDetalle: ${valorIngresado}\nProveedor: ${proveedorFila}\nFecha/Hora: ${fechaHoy}`
    );
  } catch(err) {
    Logger.log("Error al procesar solicitud: " + err.message);
  }
}

function botonRefrescarValidaciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Facturas");
  if (!sheet) return SpreadsheetApp.getUi().alert("No se encontró la hoja Facturas.");

  const ultFila = sheet.getLastRow();
  if (ultFila < 2) return SpreadsheetApp.getUi().alert("No hay filas para validar.");

  const numFilas = ultFila - 1;
  const hojaLista = ss.getSheetByName("Lista") || ss.getSheetByName("Listas");
  let catalogoSet = new Set();
  if (hojaLista) {
    catalogoSet = new Set(hojaLista.getDataRange().getValues().flat().map(normalizarTexto).filter(String));
  }

  const valores = sheet.getRange(2, 4, numFilas, 21).getValues();
  const rangoDE = sheet.getRange(2, 4, numFilas, 2);
  const fondos = rangoDE.getBackgrounds();
  const notas = rangoDE.getNotes();
  let limpiadas = 0;

  for (let i = 0; i < numFilas; i++) {
    const provTexto = normalizarTexto(valores[i][0]);
    const prodTexto = normalizarTexto(valores[i][1]);
    const codigoSku = String(valores[i][20] || "").trim();

    const prodValido = (codigoSku && !codigoSku.includes("#N/A") && !codigoSku.includes("ERROR")) || catalogoSet.has(prodTexto);
    const provValido = catalogoSet.has(provTexto) || provTexto !== "";

    if (prodValido && (notas[i][1] !== "" || fondos[i][1] === "#fff2cc")) {
      notas[i][1] = "";
      fondos[i][1] = null;
      limpiadas++;
    }

    if (provValido && (notas[i][0] !== "" || fondos[i][0] === "#fff2cc")) {
      notas[i][0] = "";
      fondos[i][0] = null;
      limpiadas++;
    }
  }

  rangoDE.setNotes(notas);
  rangoDE.setBackgrounds(fondos);
  SpreadsheetApp.getUi().alert(limpiadas ? `🧹 Se limpiaron ${limpiadas} marca(s).` : "✅ Todo limpio.");
}

// ============================================================================
// DICCIONARIO EXPLÍCITO DE ALÍCUOTAS (SKU Y NOMBRE)
// ============================================================================
const ALICUOTAS_IMPUESTOS = {
  // COCA COMÚN: 8.70%
  "coca cola": 0.08696,
  "coca": 0.08696,
  "coca cola zero": 0.08696,
  "cocaz": 0.08696,
  // LOS DEMÁS: 4.17%
  "sprite zero": 0.04167,
  "spritez": 0.04167,
  "sprite comun": 0.04167,
  "sprite": 0.04167,
  "tonica botella": 0.04167,
  "toniba": 0.04167,
  "tonica": 0.04167,
  "i0261": 0.04167
};

function obtenerAlicuotaProducto_(nombreProd) {
  const norm = normalizarTexto(nombreProd);
  const claves = Object.keys(ALICUOTAS_IMPUESTOS).sort((a, b) => b.length - a.length);
  
  for (const clave of claves) {
    if (norm.includes(clave)) {
      return ALICUOTAS_IMPUESTOS[clave];
    }
  }
  return 0;
}

// ============================================================================
// CONCEPTOS SEPARADOS (ENVÍO VS IMPUESTOS)
// ============================================================================
function esConceptoEnvio_(nombre) {
  const txt = normalizarTexto(nombre);
  if (!txt) return false;
  return txt.includes("envio");
}

function esConceptoImpuesto_(nombre) {
  const txt = normalizarTexto(nombre);
  if (!txt) return false;
  return txt.includes("impuesto") || 
         txt.includes("interno") || 
         txt.includes("percepcion") || 
         txt.includes("perc");
}

// ============================================================================
// PRORRATEO DUAL MATEMÁTICO: ENVÍO (PROPORCIONAL) + IMPUESTOS (ALÍCUOTA FIJA)
// ============================================================================
function procesarProrrateoACero_(sheet, filaActual) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1500)) return;

  try {
    const nroFactura = String(sheet.getRange(filaActual, 1).getValue() || "").trim();
    if (!nroFactura) return;

    let filaInicio = filaActual;
    while (filaInicio > 2 && String(sheet.getRange(filaInicio - 1, 1).getValue() || "").trim() === nroFactura) {
      filaInicio--;
    }

    let filaFin = filaActual;
    const maxFilas = sheet.getLastRow();
    while (filaFin < maxFilas && String(sheet.getRange(filaFin + 1, 1).getValue() || "").trim() === nroFactura) {
      filaFin++;
    }

    const numFilas = filaFin - filaInicio + 1;
    const rangoDatos = sheet.getRange(filaInicio, 1, numFilas, 10).getValues();

    const filasEnvios = [];
    const filasImpuestos = [];
    const filasProductos = [];

    for (let i = 0; i < numFilas; i++) {
      const prod = String(rangoDatos[i][4] || "");
      const cant = parseFloat(rangoDatos[i][6]) || 1;
      const precioU = parseFloat(rangoDatos[i][7]) || 0;
      
      let descRaw = parseFloat(rangoDatos[i][8]) || 0;
      const desc = descRaw > 1 ? descRaw / 100 : descRaw; 

      if (esConceptoEnvio_(prod)) {
        if (precioU > 0) filasEnvios.push({ filaReal: filaInicio + i, montoTotal: precioU * cant * (1 - desc) });
      } else if (esConceptoImpuesto_(prod)) {
        if (precioU > 0) filasImpuestos.push({ filaReal: filaInicio + i, montoTotal: precioU * cant * (1 - desc) });
      } else if (precioU > 0 && prod !== "") {
        const subtotalNeto = cant * precioU * (1 - desc);
        filasProductos.push({
          filaReal: filaInicio + i,
          nombre: prod,
          cantidad: cant,
          precioOriginalU: precioU,
          descuento: desc,
          subtotalNeto: subtotalNeto
        });
      }
    }

    if ((filasEnvios.length === 0 && filasImpuestos.length === 0) || filasProductos.length === 0) return;

    const totalSubtotalProductos = filasProductos.reduce((acc, p) => acc + p.subtotalNeto, 0);
    const totalEnvios = filasEnvios.reduce((acc, g) => acc + g.montoTotal, 0);

    filasProductos.forEach(p => {
      let nuevoPrecioU = p.precioOriginalU;
      const factorDesc = 1 - p.descuento;

      if (totalEnvios > 0 && totalSubtotalProductos > 0 && factorDesc > 0) {
        const proporcion = p.subtotalNeto / totalSubtotalProductos; 
        const gastoAsignado = totalEnvios * proporcion; 
        const incrementoBase = gastoAsignado / (p.cantidad * factorDesc);
        nuevoPrecioU += incrementoBase;
      }

      if (filasImpuestos.length > 0 && factorDesc > 0) {
        const alicuota = obtenerAlicuotaProducto_(p.nombre);
        if (alicuota > 0) {
          const extraImpuesto = (p.precioOriginalU * alicuota) / factorDesc;
          nuevoPrecioU += extraImpuesto;
        }
      }

      nuevoPrecioU = Math.round(nuevoPrecioU * 1000) / 1000; 
      if (nuevoPrecioU !== p.precioOriginalU) {
        sheet.getRange(p.filaReal, 8).setValue(nuevoPrecioU);
      }
    });

    filasEnvios.forEach(item => {
      sheet.getRange(item.filaReal, 8).setValue(0);
      sheet.getRange(item.filaReal, 5).setNote(`Envío de $${item.montoTotal} prorrateado (absorbido en precios base) en Fra ${nroFactura}`);
    });

    filasImpuestos.forEach(item => {
      sheet.getRange(item.filaReal, 8).setValue(0);
      sheet.getRange(item.filaReal, 5).setNote(`Impuesto de $${item.montoTotal} asignado por alícuotas fijas en Fra ${nroFactura}`);
    });

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Factura ${nroFactura}: Matemática precisa en Col H (Envío prorrateado: $${totalEnvios}).`,
      "⚡ CONGRESO",
      5
    );
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// AUDITORÍA Y ALERTAS DIARIAS DE PRECIOS (EXCLUSIÓN 100% DE VERDULERÍA)
// ============================================================================

function esArticuloOProveedorVerduleria_(producto, proveedor) {
  const prodNorm = normalizarTexto(producto);
  const provNorm = normalizarTexto(proveedor);

  // Excluye por descripción de producto
  if (prodNorm.includes("verduleria") || prodNorm.includes("verdura") || prodNorm.includes("fruta")) {
    return true;
  }

  // Excluye por proveedores habituales de verdulería
  if (provNorm.includes("raices") || 
      provNorm.includes("open fruit") || 
      provNorm.includes("land fruit") || 
      provNorm.includes("verduleria")) {
    return true;
  }

  return false;
}

function auditarVariacionesPreciosDiario() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Facturas");
  if (!sheet || sheet.getLastRow() < 2) return;

  const numFilas = sheet.getLastRow() - 1;
  const datos = sheet.getRange(2, 1, numFilas, 23).getValues();
  const historialPrecios = {};
  const alertas = [];

  for (let i = 0; i < numFilas; i++) {
    const fila = datos[i];
    const fecha = fila[1] ? Utilities.formatDate(new Date(fila[1]), "GMT-3", "dd/MM/yyyy") : "S/F";
    const proveedor = String(fila[3] || "").trim();
    const producto = String(fila[4] || "").trim();
    const precioActual = parseFloat(fila[14]) || 0;
    const flagCambio = String(fila[21] || "").trim();

    if (!producto || precioActual === 0) continue;

    // FILTRO ESTRICTO: Saltea al 100% ítems o proveedores de Verdulería
    if (esArticuloOProveedorVerduleria_(producto, proveedor)) {
      continue;
    }

    if (flagCambio.includes("CAMBIÓ PRECIO") && historialPrecios[producto]) {
      const precioAnt = historialPrecios[producto];
      const dif = precioActual - precioAnt;
      const pct = (dif / precioAnt) * 100;

      if (Math.abs(pct) >= CONFIG_SOLICITUDES.UMBRAL_ALERTA_PORC) {
        alertas.push({
          fila: i + 2,
          fecha: fecha,
          proveedor: proveedor,
          producto: producto,
          precioAnterior: precioAnt,
          precioActual: precioActual,
          diferencia: dif,
          porcentaje: pct.toFixed(1),
          sentido: dif > 0 ? "🔺 SUBIÓ" : "🔻 BAJÓ"
        });
      }
    }
    historialPrecios[producto] = precioActual;
  }

  if (alertas.length > 0) {
    enviarReporteEmailPrecios(ss.getName(), ss.getUrl(), alertas);
  }
}

function enviarReporteEmailPrecios(nombreLocal, urlPlanilla, alertas) {
  let filasHtml = "";
  alertas.forEach(a => {
    const colorPorc = a.diferencia > 0 ? "#cc0000" : "#274e13";
    const bgFila = a.diferencia > 0 ? "#fce5cd" : "#d9ead3";

    filasHtml += `
      <tr style="background-color: ${bgFila}; border-bottom: 1px solid #ddd; text-align: center;">
        <td style="padding: 8px;">${a.fecha}</td>
        <td style="padding: 8px; text-align: left;"><b>${a.producto}</b></td>
        <td style="padding: 8px;">${a.proveedor}</td>
        <td style="padding: 8px;">$${a.precioAnterior.toLocaleString("es-AR", {minimumFractionDigits: 2})}</td>
        <td style="padding: 8px;"><b>$${a.precioActual.toLocaleString("es-AR", {minimumFractionDigits: 2})}</b></td>
        <td style="padding: 8px; color: ${colorPorc};"><b>${a.sentido} ${Math.abs(a.porcentaje)}%</b></td>
      </tr>`;
  });

  const cuerpoHtml = `
    <div style="font-family: Arial, sans-serif; max-width: 750px; color: #333;">
      <h2 style="color: #b45f06;">📊 Reporte Diario de Variación de Precios (11:00 AM)</h2>
      <p><b>Local:</b> ${CONFIG_SOLICITUDES.LOCAL} (${nombreLocal})</p>
      <p>Se detectaron <b>${alertas.length}</b> insumo(s) con saltos iguales o mayores al <b>${CONFIG_SOLICITUDES.UMBRAL_ALERTA_PORC}%</b> (Verdulería excluida):</p>
      <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
        <thead>
          <tr style="background-color: #434343; color: #fff; text-align: center;">
            <th style="padding: 8px;">Fecha</th>
            <th style="padding: 8px; text-align: left;">Producto</th>
            <th style="padding: 8px;">Proveedor</th>
            <th style="padding: 8px;">Precio Anterior</th>
            <th style="padding: 8px;">Precio Nuevo</th>
            <th style="padding: 8px;">Variación</th>
          </tr>
        </thead>
        <tbody>${filasHtml}</tbody>
      </table>
      <p style="margin-top: 20px; font-size: 12px; color: #666;">
        Reporte generado automáticamente a las 11:00 AM.<br>
        Acceso directo a la planilla: <a href="${urlPlanilla}">${nombreLocal}</a>
      </p>
    </div>`;

  MailApp.sendEmail({
    to: CONFIG_SOLICITUDES.EMAIL_DESTINO,
    subject: `🚨 Alerta Precios (11:00 AM): ${alertas.length} insumo(s) con cambios >= ${CONFIG_SOLICITUDES.UMBRAL_ALERTA_PORC}% en ${CONFIG_SOLICITUDES.LOCAL}`,
    htmlBody: cuerpoHtml
  });
}

function instalarTriggerDiario11AM() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "auditarVariacionesPreciosDiario") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("auditarVariacionesPreciosDiario")
    .timeBased()
    .everyDays(1)
    .atHour(11)
    .create();

  SpreadsheetApp.getUi().alert("⏰ Trigger instalado con éxito: correrá todos los días a las 11:00 AM.");
}