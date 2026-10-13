/**
 * SISTEMA INTEGRAL DE CONTROL OPERATIVO, CATALOGACIÓN Y PRECIOS
 * Archivo: Local (Facturas) - CABELLO
 */

const CONFIG_SOLICITUDES = {
  EMAIL_DESTINO: "mishiguene.cervino5@gmail.com",
  MAESTRO_ID: "1lOabw9U4ilRwE5r6QXQ2Jc1nqkKUK3Hi096xga0XFnI",
  HOJA_SOLICITUDES: "Solicitudes",
  LOCAL: "CABELLO", 
  UMBRAL_ALERTA_PORC: 6.0
};

const CONFIG_REMITOS = {
  HOJA_DESTINO: "Remitos",
  ZONA_HORARIA: "GMT-3",
  CLIENTE: "Cafe mishiguene",
  NOMBRE_INFORME: /^INFORME\d+/i,
  LABEL_GMAIL: "REMITOS_PROCESADOS_CABELLO",
  BUSQUEDA_GMAIL: '-label:REMITOS_PROCESADOS_CABELLO newer_than:2d subject:("Informe: Ventas por artículo") subject:("Cliente:Cafe mishiguene") has:attachment'
};

// ============================================================================
// DICCIONARIO DE ALÍCUOTAS DE IMPUESTOS INTERNOS (NOMBRE Y SKU)
// ============================================================================
const ALICUOTAS_IMPUESTOS = {
  // LÍNEA COLA (8.70%)
  "coca cola zero": 0.08696,
  "cocaz": 0.08696,
  "coca cola": 0.08696,
  "coca": 0.08696,

  // RESTO: SPRITE / TÓNICA (4.17%)
  "sprite zero": 0.04167,
  "spritez": 0.04167,
  "sprite comun": 0.04167,
  "sprite": 0.04167,
  "tonica botella": 0.04167,
  "toniba": 0.04167,
  "tonica lata": 0.04167,
  "tonila": 0.04167,
  "tonica": 0.04167,
  "i0261": 0.04167
};

function obtenerAlicuotaProducto_(nombreProd, codigoSku) {
  const texto = normalizarTexto(nombreProd + " " + (codigoSku || ""));
  const claves = Object.keys(ALICUOTAS_IMPUESTOS).sort((a, b) => b.length - a.length);
  for (const clave of claves) {
    if (texto.includes(clave)) return ALICUOTAS_IMPUESTOS[clave];
  }
  return 0;
}

// ============================================================================
// TRIGGERS GLOBALES (ONOPEN / ONEDIT)
// ============================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("CABELLO")
    .addItem("📦 Sincronizar remitos desde Gmail", "syncRemitos")
    .addItem("🔄 Resetear memoria/caché de remitos", "resetearCacheRemitos")
    .addSeparator()
    .addItem("⚡ Prorratear Gastos (Llevar a $0)", "distribuirGastosCero")
    .addItem("🧹 Refrescar validaciones de catálogo", "botonRefrescarValidaciones")
    .addSeparator()
    .addItem("⏰ Instalar auditoría diaria (11:00 AM)", "instalarTriggerDiario11AM")
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Facturas") return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  if (col === 4 || col === 5) {
    registrarItemNuevo(e);
  }

  // Col 5 = Producto, Col 8 = Precio unitario en H
  if (col === 5 || col === 8) {
    procesarProrrateoACero_(sheet, row);
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

function esConceptoEnvio_(nombre) {
  const txt = normalizarTexto(nombre);
  return txt.includes("envio") || txt.includes("flete") || txt.includes("transporte");
}

function esConceptoImpuesto_(nombre) {
  const txt = normalizarTexto(nombre);
  return txt.includes("impuesto") || 
         txt.includes("interno") || 
         txt.includes("percepcion") || 
         txt.includes("perc");
}

function esConceptoGasto_(nombre) {
  return esConceptoEnvio_(nombre) || esConceptoImpuesto_(nombre);
}

function registrarItemNuevo(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Facturas" || e.range.getRow() < 2) return;
  const col = e.range.getColumn();
  if (col !== 4 && col !== 5) return;

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
    const proveedorFila = (tipo === "PROVEEDOR") ? valorIngresado : String(sheet.getRange(e.range.getRow(), 4).getValue() || "Sin Proveedor");

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
// PRORRATEO QUIRÚRGICO EN COLUMNA H (SOLO TOCA COLUMNA 8 Y PROTEGE LAS FÓRMULAS)
// ============================================================================

function distribuirGastosCero() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Facturas");
  if (!sheet) return;
  const filaActiva = sheet.getActiveCell().getRow();
  if (filaActiva > 1) {
    procesarProrrateoACero_(sheet, filaActiva);
  }
}

function procesarProrrateoACero_(sheet, filaActual) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) return;

  try {
    const nroFactura = String(sheet.getRange(filaActual, 1).getValue() || "").trim();
    if (!nroFactura) return;

    // Detectar el lote de filas de la factura completa
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
    // Leemos hasta Col X (24 columnas)
    const rangoDatos = sheet.getRange(filaInicio, 1, numFilas, 24).getValues();

    const filasEnvio = [];
    const filasImpuesto = [];
    const filasProductos = [];

    let totalMontoEnvio = 0;
    let totalMontoImpuesto = 0;

    for (let i = 0; i < numFilas; i++) {
      const prod = String(rangoDatos[i][4] || "");          // Col E (5)
      const cant = parseFloat(rangoDatos[i][5]) || 1;          // Col F (6)
      const valorH = parseFloat(rangoDatos[i][7]) || 0;        // Col H (8)
      const descPorc = parseFloat(rangoDatos[i][8]) || 0;      // Col I (9): % Descuento
      const sku = String(rangoDatos[i][20] || rangoDatos[i][23] || "");

      let descNorm = descPorc;
      if (descNorm > 1) descNorm = descNorm / 100;
      if (descNorm < 0) descNorm = 0;

      if (esConceptoEnvio_(prod)) {
        if (valorH > 0) {
          filasEnvio.push({ filaReal: filaInicio + i, monto: valorH });
          totalMontoEnvio += valorH;
        }
      } else if (esConceptoImpuesto_(prod)) {
        if (valorH > 0) {
          filasImpuesto.push({ filaReal: filaInicio + i, monto: valorH });
          totalMontoImpuesto += valorH;
        }
      } else if (valorH > 0 && prod !== "") {
        const subtotalNetoFila = valorH * cant * (1 - descNorm);
        filasProductos.push({
          filaReal: filaInicio + i,
          nombre: prod,
          sku: sku,
          cantidad: cant,
          valorH: valorH,
          descNorm: descNorm,
          subtotalNeto: subtotalNetoFila
        });
      }
    }

    // Si no hay gastos o no hay productos, salimos
    if ((filasEnvio.length === 0 && filasImpuesto.length === 0) || filasProductos.length === 0) {
      return;
    }

    const subtotalTotalNeto = filasProductos.reduce((acc, p) => acc + p.subtotalNeto, 0);

    // ========================================================================
    // BLOQUE 1: CÁLCULO DE ENVÍO (Prorrateo 100% independiente)
    // ========================================================================
    const deltaH_Envio = {};
    if (totalMontoEnvio > 0 && subtotalTotalNeto > 0) {
      filasProductos.forEach(p => {
        const porcPeso = p.subtotalNeto / subtotalTotalNeto;
        const envioAsignado = totalMontoEnvio * porcPeso;
        const envioUnitario = envioAsignado / p.cantidad;
        // Compensación de descuento: revertir para que el MAP calcule exacto
        const divisorDesc = (1 - p.descNorm) > 0 ? (1 - p.descNorm) : 1;
        deltaH_Envio[p.filaReal] = envioUnitario / divisorDesc;
      });
    }

    // ========================================================================
    // BLOQUE 2: CÁLCULO DE IMPUESTO INTERNO (Mapeo explícito OR Prorrateo general)
    // ========================================================================
    const deltaH_Impuesto = {};
    let tipoDistribucionImpuesto = "";

    if (totalMontoImpuesto > 0) {
      // Verificar si algún producto tiene alícuota mapeada en el diccionario
      const hayProductosConAlicuota = filasProductos.some(p => obtenerAlicuotaProducto_(p.nombre, p.sku) > 0);

      if (hayProductosConAlicuota) {
        // Modo alícuota específica (Coca, Sprite, etc.)
        tipoDistribucionImpuesto = "alícuota fija mapeada";
        filasProductos.forEach(p => {
          const alicuota = obtenerAlicuotaProducto_(p.nombre, p.sku);
          deltaH_Impuesto[p.filaReal] = p.valorH * alicuota;
        });
      } else {
        // No hay productos mapeados: prorrateo general como el envío
        tipoDistribucionImpuesto = `prorrateo general entre ${filasProductos.length} producto(s)`;
        filasProductos.forEach(p => {
          const porcPeso = p.subtotalNeto / subtotalTotalNeto;
          const impAsignado = totalMontoImpuesto * porcPeso;
          const impUnitario = impAsignado / p.cantidad;
          // Compensación del descuento para que no le coma guita
          const divisorDesc = (1 - p.descNorm) > 0 ? (1 - p.descNorm) : 1;
          deltaH_Impuesto[p.filaReal] = impUnitario / divisorDesc;
        });
      }
    }

    // ========================================================================
    // BLOQUE 3: IMPACTAR NUEVO VALOR EN COLUMNA H (SOLO COLUMNA 8)
    // ========================================================================
    filasProductos.forEach(p => {
      const extraEnvio = deltaH_Envio[p.filaReal] || 0;
      const extraImpuesto = deltaH_Impuesto[p.filaReal] || 0;
      const nuevoH = Math.round((p.valorH + extraEnvio + extraImpuesto) * 100) / 100;
      sheet.getRange(p.filaReal, 8).setValue(nuevoH);
    });

    // ========================================================================
    // BLOQUE 4: LLEVAR GASTOS A $0 Y NOTAS INDEPENDIENTES
    // ========================================================================
    filasEnvio.forEach(item => {
      sheet.getRange(item.filaReal, 8).setValue(0);
      sheet.getRange(item.filaReal, 5).setNote(
        `📦 Envío de $${item.monto.toLocaleString("es-AR")} prorrateado proporcionalmente en Columna H (con compensación de descuento) en Factura ${nroFactura}.`
      );
    });

    filasImpuesto.forEach(item => {
      sheet.getRange(item.filaReal, 8).setValue(0);
      sheet.getRange(item.filaReal, 5).setNote(
        `🏛 Impuesto Interno de $${item.monto.toLocaleString("es-AR")} distribuido por ${tipoDistribucionImpuesto} en Columna H en Factura ${nroFactura}.`
      );
    });

    let msj = `Factura ${nroFactura}:`;
    if (totalMontoEnvio > 0) msj += ` Envío ($${totalMontoEnvio}) impactado.`;
    if (totalMontoImpuesto > 0) msj += ` Impuesto ($${totalMontoImpuesto}) procesado.`;

    SpreadsheetApp.getActiveSpreadsheet().toast(msj, "⚡ CABELLO", 3);

  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// AUDITORÍA Y ALERTAS DE PRECIOS (EXCLUSIÓN 100% DE VERDULERÍA)
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

    // FILTRO ESTRICTO: Si es Verdulería (por producto o proveedor), saltear al 100%
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

// ============================================================================
// REMITOS GMAIL (ÚLTIMAS 48 HS - CLIENTE: CAFE MISHIGUENE)
// ============================================================================

function resetearCacheRemitos() {
  const propiedades = PropertiesService.getDocumentProperties();
  const todasLasKeys = propiedades.getKeys();
  let borradas = 0;

  for (const k of todasLasKeys) {
    if (k.startsWith("CABELLO_REMITO_PROCESADO_")) {
      propiedades.deleteProperty(k);
      borradas++;
    }
  }

  const etiqueta = GmailApp.getUserLabelByName(CONFIG_REMITOS.LABEL_GMAIL);
  if (etiqueta) {
    const hilos = etiqueta.getThreads();
    for (const h of hilos) {
      etiqueta.removeFromThread(h);
    }
  }

  SpreadsheetApp.getUi().alert(`✅ Memoria reseteada (${borradas} claves borradas) y correos desmarcados. Listo para sincronizar.`);
}

function syncRemitos() {
  const ui = SpreadsheetApp.getUi();
  const hojaDestino = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_REMITOS.HOJA_DESTINO);
  if (!hojaDestino) return ui.alert('❌ No existe la hoja "Remitos".');

  const propiedades = PropertiesService.getDocumentProperties();
  const etiqueta = GmailApp.getUserLabelByName(CONFIG_REMITOS.LABEL_GMAIL) || GmailApp.createLabel(CONFIG_REMITOS.LABEL_GMAIL);
  
  const hilos = GmailApp.search(CONFIG_REMITOS.BUSQUEDA_GMAIL);
  
  let mensajesEncontrados = 0;
  let informesProcesados = 0;
  let informesIgnorados = 0;
  let filasImportadas = 0;
  let filasOmitidas = 0;
  const errores = [];

  for (const hilo of hilos) {
    const mensajes = hilo.getMessages();
    let hiloProcesadoExitoso = false;

    for (const mensaje of mensajes) {
      const asuntoRaw = mensaje.getSubject();
      const asuntoNorm = normalizarTexto(asuntoRaw).replace(/\s*:\s*/g, ":");
      const clienteBuscadoNorm = normalizarTexto(CONFIG_REMITOS.CLIENTE);

      if (!asuntoNorm.includes("informe:ventas por articulo") ||
          !asuntoNorm.includes("cliente:" + clienteBuscadoNorm)) {
        continue;
      }

      mensajesEncontrados++;
      const adjuntos = mensaje.getAttachments({ includeInlineImages: false, includeAttachments: true });

      let fechaFallback = mensaje.getDate();
      let matchFechaAsunto = asuntoRaw.match(/al\s+(\d{1,2}\/\d{1,2}\/\d{2,4})/i) || asuntoRaw.match(/Fecha:\s*(\d{1,2}\/\d{1,2}\/\d{2,4})/i);
      if (matchFechaAsunto) {
        let fParsed = convertirFechaRemito_(matchFechaAsunto[1]);
        if (fParsed) fechaFallback = fParsed;
      }

      for (const adjunto of adjuntos) {
        const nombre = adjunto.getName();
        if (!CONFIG_REMITOS.NOMBRE_INFORME.test(nombre) || !/\.(xlsx|xls)$/i.test(nombre)) continue;

        const clave = claveAdjuntoRemito_(mensaje.getId(), nombre, adjunto.getSize());
        if (propiedades.getProperty(clave)) {
          informesIgnorados++;
          hiloProcesadoExitoso = true;
          continue;
        }

        try {
          const resultado = importarAdjuntoRemito_(adjunto, hojaDestino, fechaFallback);
          propiedades.setProperty(clave, JSON.stringify({
            nombre: nombre,
            fechaMensaje: mensaje.getDate().toISOString(),
            filas: resultado.importadas
          }));
          informesProcesados++;
          filasImportadas += resultado.importadas;
          filasOmitidas += resultado.omitidas;
          hiloProcesadoExitoso = true;
        } catch (err) {
          errores.push(`${nombre}: ${err.message}`);
        }
      }
    }

    if (hiloProcesadoExitoso && filasImportadas > 0) {
      etiqueta.addToThread(hilo);
    }
  }

  let mensajeFinal =
    "📦 SYNC REMITOS CABELLO (ÚLTIMAS 48 HS)\n\n" +
    `• Mensajes detectados: ${mensajesEncontrados}\n` +
    `• Archivos procesados y volcados: ${informesProcesados}\n` +
    `• Archivos ya existentes (omitidos): ${informesIgnorados}\n` +
    `• Filas escritas en 'Remitos': ${filasImportadas}\n` +
    `• Filas descartadas (cabeceras/totales): ${filasOmitidas}`;
    
  if (errores.length > 0) mensajeFinal += "\n\n⚠️ ERRORES:\n" + errores.join("\n");
  if (mensajesEncontrados === 0) mensajeFinal += "\n\nℹ️ No se detectaron correos pendientes de 'Cafe mishiguene'.";
  
  ui.alert(mensajeFinal);
}

function importarAdjuntoRemito_(adjunto, hojaDestino, fechaFallback) {
  let archivoConvertido;
  try {
    const resource = {
      title: "TMP_CABELLO_" + adjunto.getName(),
      mimeType: MimeType.GOOGLE_SHEETS
    };

    if (Drive.Files.create) {
      archivoConvertido = Drive.Files.create({ name: resource.title, mimeType: resource.mimeType }, adjunto.copyBlob());
    } else {
      archivoConvertido = Drive.Files.insert(resource, adjunto.copyBlob(), { convert: true });
    }

    const ssOrigen = SpreadsheetApp.openById(archivoConvertido.id);
    const hojaOrigen = ssOrigen.getSheets()[0];
    const ultimaFila = hojaOrigen.getLastRow();
    const ultimaCol = hojaOrigen.getLastColumn();
    
    if (ultimaFila < 2 || ultimaCol < 2) {
      throw new Error("Archivo vacío o con estructura ilegible.");
    }

    const datos = hojaOrigen.getRange(1, 1, ultimaFila, ultimaCol).getValues();

    let filaCabecera = -1;
    let colProd = -1;
    let colCant = -1;
    let colImp = -1;
    let colFecha = -1;

    for (let r = 0; r < Math.min(15, datos.length); r++) {
      for (let c = 0; c < datos[r].length; c++) {
        let valNorm = normalizarTexto(datos[r][c]);
        if (colProd === -1 && (valNorm.includes("articulo") || valNorm.includes("producto") || valNorm.includes("descripcion") || valNorm.includes("nombre") || valNorm === "detalle")) {
          colProd = c;
          filaCabecera = r;
        }
        if (colCant === -1 && (valNorm.includes("cant") || valNorm.includes("unidades"))) {
          colCant = c;
          filaCabecera = r;
        }
        if (colImp === -1 && (valNorm.includes("importe") || valNorm.includes("total") || valNorm.includes("subtotal") || valNorm.includes("monto"))) {
          colImp = c;
        }
        if (colFecha === -1 && valNorm.includes("fecha")) {
          colFecha = c;
        }
      }
      if (colProd !== -1 && colCant !== -1) break;
    }

    if (colProd === -1) colProd = 1;
    if (colCant === -1) colCant = 2;
    if (colImp === -1) colImp = 3;
    let filaInicioDatos = filaCabecera !== -1 ? filaCabecera + 1 : 1;

    const salida = [];
    let omitidas = 0;

    for (let i = filaInicioDatos; i < datos.length; i++) {
      const fila = datos[i];
      const prodRaw = String(fila[colProd] || "").replace(/[\u0000-\u001F\u007F-\u009F]/g, "").trim();
      const cantRaw = fila[colCant];
      const impRaw = colImp !== -1 ? fila[colImp] : 0;
      const fechaRaw = colFecha !== -1 ? fila[colFecha] : null;

      let prodNorm = normalizarTexto(prodRaw);
      if (!prodRaw || prodNorm.length < 2 || 
          prodNorm.startsWith("total") || 
          prodNorm.startsWith("subtotal") || 
          prodNorm.startsWith("cliente") || 
          prodNorm.startsWith("informe") ||
          prodNorm.startsWith("pagina") ||
          prodNorm.includes("articulo")) {
        omitidas++;
        continue;
      }

      const cantidad = convertirNumeroRemito_(cantRaw);
      const importe = convertirNumeroRemito_(impRaw);

      if (typeof cantidad !== "number" || isNaN(cantidad) || cantidad <= 0) {
        omitidas++;
        continue;
      }

      let fecha = convertirFechaRemito_(fechaRaw) || fechaFallback;

      const precioUnitario = (typeof importe === "number" && importe > 0 && cantidad > 0) 
        ? Math.round((importe / cantidad) * 100) / 100 
        : 0;

      salida.push([fecha, prodRaw, cantidad, precioUnitario, importe || 0]);
    }

    if (salida.length > 0) {
      const uFilaReal = encontrarUltimaFilaRealRemitos_(hojaDestino);
      const primeraFila = uFilaReal + 1;

      hojaDestino.getRange(primeraFila, 1, salida.length, 5).setValues(salida);
      hojaDestino.getRange(primeraFila, 1, salida.length, 1).setNumberFormat("dd/MM/yyyy");
      hojaDestino.getRange(primeraFila, 3, salida.length, 1).setNumberFormat("0.##");
      hojaDestino.getRange(primeraFila, 4, salida.length, 2).setNumberFormat("$#,##0.00");
    }

    return { importadas: salida.length, omitidas: omitidas };

  } finally {
    if (archivoConvertido && archivoConvertido.id) {
      try {
        if (Drive.Files.delete) Drive.Files.delete(archivoConvertido.id);
        else Drive.Files.remove(archivoConvertido.id);
      } catch (e) {
        // Silenciar borrado de temporal
      }
    }
  }
}

function encontrarUltimaFilaRealRemitos_(hoja) {
  const uFila = hoja.getLastRow();
  if (uFila < 2) return 1;
  const valores = hoja.getRange(2, 2, uFila - 1, 1).getValues();
  for (let i = valores.length - 1; i >= 0; i--) {
    if (String(valores[i][0]).trim() !== "") {
      return i + 2;
    }
  }
  return 1;
}

function claveAdjuntoRemito_(mensajeId, nombre, tamano) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, mensajeId + "|" + nombre + "|" + tamano);
  return "CABELLO_REMITO_PROCESADO_" + digest.map(byte => {
    const valor = byte < 0 ? byte + 256 : byte;
    return ("0" + valor.toString(16)).slice(-2);
  }).join("");
}

function convertirNumeroRemito_(valor) {
  if (valor === "" || valor === null || valor === undefined) return "";
  if (typeof valor === "number") return isNaN(valor) ? "" : valor;
  let texto = String(valor).trim().replace(/\$/g, "").replace(/\s/g, "");
  if (!texto) return "";
  if (texto.includes(",") && texto.includes(".")) texto = texto.replace(/\./g, "").replace(",", ".");
  else if (texto.includes(",")) texto = texto.replace(",", ".");
  const numero = Number(texto);
  return isNaN(numero) ? "" : numero;
}

function convertirFechaRemito_(valor) {
  if (!valor) return null;
  if (Object.prototype.toString.call(valor) === "[object Date]" && !isNaN(valor.getTime())) return valor;
  
  let str = String(valor).trim();
  let match = str.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (match) {
    let dia = Number(match[1]);
    let mes = Number(match[2]);
    let anio = Number(match[3]);
    if (mes > 12) {
      dia = Number(match[2]);
      mes = Number(match[1]);
    }
    if (anio < 100) anio += 2000;
    return new Date(anio, mes - 1, dia);
  }
  return null;
}