/**
 * Web App para registrar ventas en la hoja "movimientos" y hacer el registro en caja.
 * Recibe POST (form-data o JSON payload) con múltiples ítems y agrega filas en lote.
 * También registra cada movimiento en la pestaña "caja" con tipo 1 (venta).
 */
const SHEET_NAME = 'movimientos';
const CAJA_SHEET_NAME = 'caja';
const LOTES_SHEET_NAME = 'lotes_facturacion';
const LOTES_ITEMS_SHEET_NAME = 'lotes_items';
const LOTES_HEADERS = [
  'lote_id', 'fecha_creacion', 'fecha_desde', 'fecha_hasta', 'estado', 'fecha_actualizacion',
  'total_items', 'total_monto', 'archivo_nombre', 'archivo_id', 'url', 'nota',
];
const FACTURACION_FOLDER_ID = '1AK9CZlEeh6oLlnf6PcV_cp5uZMPH2Qqa';
const FACTURACION_TEMPLATE_ID = '16gxGia3t367ImiAW_2t0dhJ2Zy7D5E5HGooSWAemQ9s';

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  const callback = e && e.parameter && e.parameter.callback;
  try {
    if (action === 'facturacion_generar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = generateFacturacion(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'facturacion_confirmar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = confirmFacturacion(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'facturacion_denegar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = denyFacturacion(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'facturacion_lotes') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = listFacturacionLotes(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'movimiento_eliminar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = deleteMovimiento(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'movimiento_actualizar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = updateMovimiento(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'facturacion_descargar') {
      const loteId = e && e.parameter ? e.parameter.lote_id : null;
      return downloadFacturacion(loteId);
    }
    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    const out = { ok: false, error: err.message || String(err) };
    return callback ? jsonpResponse(out, callback) : jsonResponse(out, 400);
  }
}

// Preflight CORS
function doOptions() {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const payload = parsePayload(e);
    const action = (e && e.parameter && e.parameter.action) || payload.action;
    if (action === 'facturacion_generar') {
      return jsonResponse(generateFacturacion(payload));
    }
    if (action === 'facturacion_confirmar') {
      return jsonResponse(confirmFacturacion(payload));
    }
    if (action === 'movimiento_eliminar') {
      return jsonResponse(deleteMovimiento(payload));
    }
    if (action === 'movimiento_actualizar') {
      return jsonResponse(updateMovimiento(payload));
    }
    const items = payload.items || [];
    if (!items.length) throw new Error('No hay items para registrar.');
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}".`);

    let rows = [];
    let nextNum = 1;
    const lock = LockService.getDocumentLock();
    lock.waitLock(15000);
    try {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const lastValue = sheet.getRange(lastRow, 1).getValue();
        nextNum = (typeof lastValue === 'number' && !isNaN(lastValue)) ? lastValue + 1 : lastRow;
      }

      rows = items.map((item, idx) => {
        const fecha = item.fecha ? new Date(item.fecha) : (payload.fecha ? new Date(payload.fecha) : new Date());
        const codigo = item.codigo || '';
        const tipo = item.tipo === undefined ? 1 : (Number(item.tipo) || 1);
        const cantidad = Number(item.cantidad || 0);
        const valorUnitario = Number(item.valor_unitario || 0);
        const descuento = Number(item.descuento || 0);
        const comentario = item.comentario && String(item.comentario).trim() ? item.comentario : 'Venta mostrador';
        const facturado = item.facturado !== undefined ? item.facturado : (payload.facturado || 0);
        const credito = item.credito !== undefined ? item.credito : (payload.credito || 0);
        const costoTotal = (Math.abs(cantidad) * valorUnitario) - descuento;
        return [
          nextNum + idx,
          fecha,
          codigo,
          tipo,
          cantidad,
          valorUnitario,
          descuento,
          costoTotal,
          comentario,
          facturado,
          credito,
        ];
      });

      sheet.getRange(lastRow + 1, 1, rows.length, 11).setValues(rows);
    } finally {
      lock.releaseLock();
    }

    insertCaja(payload, rows);

    return jsonResponse({ ok: true, inserted: rows.length, first_numero: nextNum });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message || String(err) }, 400);
  }
}

function jsonResponse(obj, status) {
  const out = ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  return out;
}

function jsonpResponse(obj, callback) {
  const safeCallback = String(callback || '').replace(/[^\w$.]/g, '');
  const payload = `${safeCallback}(${JSON.stringify(obj)});`;
  return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function insertCaja(payload, rows) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CAJA_SHEET_NAME);
  if (!sheet || !rows.length) return;
  // row: [numero, fecha, codigo, tipo, cantidad, valorUnitario, descuento, costo_total, ...]
  const lastRow = sheet.getLastRow();
  const nextId = Math.max(1, lastRow);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  let idxMonto = headers.indexOf('monto');
  if (idxMonto < 0 && headers.length >= 4) idxMonto = 3;
  let idxConcepto = headers.indexOf('concepto');
  if (idxConcepto < 0 && headers.length >= 5) idxConcepto = 4;
  let idxReferencia = headers.indexOf('referencia');
  if (idxReferencia < 0 && headers.length >= 6) idxReferencia = 5;
  let idxTipoDescripcion = headers.indexOf('tipo_descripcion');
  const cajaRows = rows.map((row, idx) => {
    const numero = row[0];
    const fecha = row[1];
    const codigo = row[2] || '';
    const cantidad = Number(row[4]) || 0;
    const valor = Number(row[5]) || 0;
    const descuento = Number(row[6]) || 0;
    const monto = (Math.abs(cantidad) * valor) - descuento;
    const producto = row[8] || payload.concepto || 'Producto';
    const concepto = `Por venta de ${Math.abs(cantidad)} * ${producto} (${codigo})`;
    const id = nextId + idx;
    const rowData = [];
    rowData[0] = id;
    rowData[1] = fecha;
    rowData[2] = 1;
    if (idxMonto >= 0) rowData[idxMonto] = monto;
    if (idxConcepto >= 0) rowData[idxConcepto] = concepto;
    if (idxReferencia >= 0) rowData[idxReferencia] = numero;
    if (idxTipoDescripcion >= 0) rowData[idxTipoDescripcion] = 'venta';
    return rowData;
  });
  sheet.getRange(lastRow + 1, 1, cajaRows.length, sheet.getLastColumn()).setValues(cajaRows.map(row => {
    const filled = new Array(sheet.getLastColumn()).fill('');
    row.forEach((val, index) => { filled[index] = val; });
    return filled;
  }));
}

function generateFacturacion(payload) {
  const dateRaw = payload.fecha;
  if (!dateRaw) throw new Error('Fecha requerida.');
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const dateTarget = parseDateValue(dateRaw);
  if (!dateTarget) throw new Error('Fecha inválida.');
  const dateKey = Utilities.formatDate(dateTarget, tz, 'yyyy-MM-dd');

  const movimientosSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!movimientosSheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}".`);
  const data = movimientosSheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('No hay movimientos.');

  const headers = data[0].map(h => String(h).toLowerCase());
  const idxNumero = headers.indexOf('numero');
  const idxFecha = headers.indexOf('fecha');
  const idxCodigo = headers.indexOf('codigo_producto');
  const idxTipo = headers.indexOf('tipo');
  const idxCantidad = headers.indexOf('cantidad');
  const idxValor = headers.indexOf('valor_unitario');
  const idxDescuento = headers.indexOf('descuento');
  const idxTotal = headers.indexOf('costo_total');
  const idxFacturado = headers.indexOf('facturado');

  const requiredIdx = [idxNumero, idxFecha, idxCodigo, idxTipo, idxCantidad, idxValor, idxDescuento, idxTotal, idxFacturado];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en movimientos.');
  }

  const catalogMap = loadCatalogMap();
  const summary = {};
  const loteItems = [];
  let totalMonto = 0;

  data.slice(1).forEach(row => {
    const tipo = Number(row[idxTipo] || 0);
    const facturado = Number(row[idxFacturado] || 0);
    const codigo = String(row[idxCodigo] || '').trim();
    const numero = row[idxNumero];
    if (!codigo || tipo !== 1 || facturado !== 0) return;
    const rowDate = parseDateValue(row[idxFecha]);
    if (!rowDate) return;
    const rowKey = Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd');
    if (rowKey !== dateKey) return;

    const cantidad = Math.abs(Number(row[idxCantidad] || 0));
    const valorUnitario = Number(row[idxValor] || 0);
    const descuento = Number(row[idxDescuento] || 0);
    const total = Number(row[idxTotal] || ((cantidad * valorUnitario) - descuento));
    totalMonto += total;

    if (!summary[codigo]) {
      const productName = catalogMap[codigo] || 'Producto';
      summary[codigo] = {
        codigo,
        nombre: productName,
        cantidad: 0,
        descuento: 0,
        valorUnitario: valorUnitario || 0,
      };
    }
    summary[codigo].cantidad += cantidad;
    summary[codigo].descuento += descuento;
    if (!summary[codigo].valorUnitario && valorUnitario) {
      summary[codigo].valorUnitario = valorUnitario;
    }

    loteItems.push([
      null,
      numero,
      row[idxFecha],
      codigo,
      cantidad,
      valorUnitario,
      descuento,
      total,
    ]);
  });

  const summaryRows = Object.keys(summary).map(code => summary[code]);
  if (!summaryRows.length) throw new Error('No hay movimientos para facturar en esa fecha.');

  const { loteId, loteRowIndex } = createLoteRecord(dateTarget, summaryRows.length, totalMonto);
  for (let i = 0; i < loteItems.length; i += 1) {
    loteItems[i][0] = loteId;
  }
  appendLoteItems(loteItems);

  const fileName = `facturacion_${dateKey}_lote${loteId}.xlsx`;
  const fileId = buildFacturacionFile(summaryRows, fileName);
  updateLoteFile(loteRowIndex, fileName, fileId);

  const link = `https://drive.google.com/file/d/${fileId}/view`;
  return { ok: true, lote_id: loteId, file_id: fileId, file_name: fileName, link: link };
}

function confirmFacturacion(payload) {
  const loteId = Number(payload.lote_id || 0);
  if (!loteId) throw new Error('Lote inválido.');
  const loteMeta = getLoteMetaById(loteId);
  if (!loteMeta) throw new Error('Lote no encontrado.');
  if (String(loteMeta.estado || '').toLowerCase() !== 'pendiente') {
    throw new Error('Solo se pueden confirmar lotes pendientes.');
  }
  const lotesItemsSheet = getOrCreateSheet(LOTES_ITEMS_SHEET_NAME, [
    'lote_id', 'numero_movimiento', 'fecha', 'codigo_producto', 'cantidad', 'valor_unitario', 'descuento', 'costo_total',
  ]);
  const itemsData = lotesItemsSheet.getDataRange().getValues();
  const rows = itemsData.slice(1).filter(row => Number(row[0] || 0) === loteId);
  if (!rows.length) throw new Error('No se encontraron items para el lote.');

  const numeros = rows.map(row => row[1]).filter(val => val !== '' && val !== null);
  const alreadyFacturados = findMovimientosFacturados(numeros);
  if (alreadyFacturados.length) {
    const note = `Rechazado: movimientos ya facturados (${alreadyFacturados.join(', ')}).`;
    updateLoteRechazado(loteId, note);
    return { ok: true, lote_id: loteId, estado: 'rechazado', already_facturados: alreadyFacturados };
  }
  markMovimientosFacturados(numeros);
  updateLoteConfirmado(loteId);
  return { ok: true, lote_id: loteId, items: rows.length };
}

function denyFacturacion(payload) {
  const loteId = Number(payload.lote_id || 0);
  if (!loteId) throw new Error('Lote inválido.');
  const loteMeta = getLoteMetaById(loteId);
  if (!loteMeta) throw new Error('Lote no encontrado.');
  if (String(loteMeta.estado || '').toLowerCase() !== 'pendiente') {
    throw new Error('Solo se pueden denegar lotes pendientes.');
  }
  updateLoteDenegado(loteId);
  return { ok: true, lote_id: loteId, estado: 'denegado' };
}

function deleteMovimiento(payload) {
  const numeroRaw = payload && payload.numero ? payload.numero : null;
  const numero = String(numeroRaw || '').trim();
  if (!numero) throw new Error('Numero requerido.');
  const movimientosSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!movimientosSheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}".`);
  const data = movimientosSheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('No hay movimientos.');
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxNumero = headers.indexOf('numero');
  const idxFacturado = headers.indexOf('facturado');
  if (idxNumero < 0 || idxFacturado < 0) {
    throw new Error('Columnas requeridas faltantes en movimientos.');
  }
  let rowIndex = -1;
  let facturado = 0;
  for (let i = 1; i < data.length; i += 1) {
    if (String(data[i][idxNumero] || '') === numero) {
      rowIndex = i + 1;
      facturado = Number(data[i][idxFacturado] || 0);
      break;
    }
  }
  if (rowIndex < 0) throw new Error('Movimiento no encontrado.');
  if (facturado === 1) throw new Error('No se puede eliminar un movimiento facturado.');
  movimientosSheet.deleteRow(rowIndex);
  deleteCajaByNumero(numero);
  return { ok: true, numero: numero };
}

function updateMovimiento(payload) {
  const numeroRaw = payload && payload.numero ? payload.numero : null;
  const numero = String(numeroRaw || '').trim();
  if (!numero) throw new Error('Numero requerido.');
  const cantidadRaw = payload && payload.cantidad !== undefined ? payload.cantidad : null;
  const descuentoRaw = payload && payload.descuento !== undefined ? payload.descuento : null;
  const facturadoRaw = payload && payload.facturado !== undefined ? payload.facturado : undefined;
  const cantidadInput = Number(cantidadRaw);
  const descuentoInput = Number(descuentoRaw);
  if (!isFinite(cantidadInput) || !isFinite(descuentoInput)) {
    throw new Error('Cantidad o descuento inválidos.');
  }
  const movimientosSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!movimientosSheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}".`);
  const data = movimientosSheet.getDataRange().getValues();
  if (data.length < 2) throw new Error('No hay movimientos.');
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxNumero = headers.indexOf('numero');
  const idxCantidad = headers.indexOf('cantidad');
  const idxValor = headers.indexOf('valor_unitario');
  const idxDescuento = headers.indexOf('descuento');
  const idxTotal = headers.indexOf('costo_total');
  const idxFacturado = headers.indexOf('facturado');
  const idxTipo = headers.indexOf('tipo');
  if (idxNumero < 0 || idxCantidad < 0 || idxValor < 0 || idxDescuento < 0 || idxTotal < 0 || idxFacturado < 0 || idxTipo < 0) {
    throw new Error('Columnas requeridas faltantes en movimientos.');
  }
  let rowIndex = -1;
  let facturado = 0;
  let valorUnitario = 0;
  let tipo = 0;
  for (let i = 1; i < data.length; i += 1) {
    if (String(data[i][idxNumero] || '') === numero) {
      rowIndex = i + 1;
      facturado = Number(data[i][idxFacturado] || 0);
      valorUnitario = Number(data[i][idxValor] || 0);
      tipo = Number(data[i][idxTipo] || 0);
      break;
    }
  }
  if (rowIndex < 0) throw new Error('Movimiento no encontrado.');
  if (facturado === 1) throw new Error('No se puede editar un movimiento facturado.');
  let nextFacturado = facturado;
  if (facturadoRaw !== undefined) {
    nextFacturado = Number(facturadoRaw) === 1 ? 1 : 0;
    if (nextFacturado === 1 && isNumeroInLoteConfirmado(numero)) {
      throw new Error('Movimiento pertenece a lote confirmado.');
    }
  }
  const cantidad = tipo === 1 ? -Math.abs(cantidadInput) : cantidadInput;
  const descuento = Math.max(0, descuentoInput);
  const costoTotal = (Math.abs(cantidad) * valorUnitario) - descuento;
  movimientosSheet.getRange(rowIndex, idxCantidad + 1, 1, 3).setValues([[cantidad, valorUnitario, descuento]]);
  movimientosSheet.getRange(rowIndex, idxTotal + 1).setValue(costoTotal);
  if (facturadoRaw !== undefined) {
    movimientosSheet.getRange(rowIndex, idxFacturado + 1).setValue(nextFacturado);
  }
  updateCajaMontoByNumero(numero, costoTotal);
  return { ok: true, numero: numero, total: costoTotal, facturado: nextFacturado };
}

function listFacturacionLotes(payload) {
  const sheet = getLotesSheet();
  if (!sheet) return { ok: true, items: [] };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, items: [] };
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxId = headers.indexOf('lote_id');
  const idxCreacion = headers.indexOf('fecha_creacion');
  const idxDesde = headers.indexOf('fecha_desde');
  const idxHasta = headers.indexOf('fecha_hasta');
  const idxEstado = headers.indexOf('estado');
  const idxActualizacion = headers.indexOf('fecha_actualizacion');
  const idxItems = headers.indexOf('total_items');
  const idxMonto = headers.indexOf('total_monto');
  const idxArchivo = headers.indexOf('archivo_nombre');
  const idxUrl = headers.indexOf('url');
  const idxNota = headers.indexOf('nota');
  const requiredIdx = [idxId, idxCreacion, idxDesde, idxHasta, idxEstado, idxActualizacion, idxItems, idxMonto, idxArchivo, idxUrl, idxNota];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en lotes_facturacion.');
  }

  const filterDate = payload && payload.fecha ? parseDateValue(payload.fecha) : null;
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const filterKey = filterDate ? Utilities.formatDate(filterDate, tz, 'yyyy-MM-dd') : null;

  const items = data.slice(1).map(row => ({
    lote_id: row[idxId],
    fecha_creacion: row[idxCreacion],
    fecha_desde: row[idxDesde],
    fecha_hasta: row[idxHasta],
    estado: row[idxEstado],
    fecha_actualizacion: row[idxActualizacion],
    total_items: row[idxItems],
    total_monto: row[idxMonto],
    archivo_nombre: row[idxArchivo],
    url: row[idxUrl],
    nota: row[idxNota],
  })).filter(item => {
    if (!filterKey) return true;
    const dateValue = parseDateValue(item.fecha_desde || item.fecha_creacion);
    if (!dateValue) return false;
    const dateKey = Utilities.formatDate(dateValue, tz, 'yyyy-MM-dd');
    return dateKey === filterKey;
  });

  items.sort((a, b) => {
    const dateA = parseDateValue(a.fecha_creacion) || new Date(0);
    const dateB = parseDateValue(b.fecha_creacion) || new Date(0);
    if (dateB.getTime() !== dateA.getTime()) {
      return dateB.getTime() - dateA.getTime();
    }
    return Number(b.lote_id || 0) - Number(a.lote_id || 0);
  });

  return { ok: true, items: items };
}

function getLoteMetaById(loteId) {
  const sheet = getLotesSheet();
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxId = headers.indexOf('lote_id');
  const idxEstado = headers.indexOf('estado');
  const idxActualizacion = headers.indexOf('fecha_actualizacion');
  const idxNota = headers.indexOf('nota');
  if (idxId < 0 || idxEstado < 0 || idxActualizacion < 0) return null;
  for (let i = 1; i < data.length; i += 1) {
    if (Number(data[i][idxId] || 0) === loteId) {
      return {
        rowIndex: i + 1,
        estado: data[i][idxEstado],
        fecha_actualizacion: data[i][idxActualizacion],
        nota: idxNota >= 0 ? data[i][idxNota] : '',
      };
    }
  }
  return null;
}

function downloadFacturacion(loteIdRaw) {
  return jsonResponse({ ok: false, error: 'Descarga directa deshabilitada.' });
}

function createLoteRecord(dateTarget, totalItems, totalMonto) {
  const lotesSheet = getLotesSheet();
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const lastRow = lotesSheet.getLastRow();
    let nextId = 1;
    if (lastRow >= 2) {
      const lastValue = lotesSheet.getRange(lastRow, 1).getValue();
      nextId = (typeof lastValue === 'number' && !isNaN(lastValue)) ? lastValue + 1 : lastRow;
    }
    const now = new Date();
    lotesSheet.appendRow([nextId, now, dateTarget, dateTarget, 'pendiente', now, totalItems, totalMonto, '', '', '', '']);
    return { loteId: nextId, loteRowIndex: lotesSheet.getLastRow() };
  } finally {
    lock.releaseLock();
  }
}

function appendLoteItems(rows) {
  const lotesItemsSheet = getOrCreateSheet(LOTES_ITEMS_SHEET_NAME, [
    'lote_id', 'numero_movimiento', 'fecha', 'codigo_producto', 'cantidad', 'valor_unitario', 'descuento', 'costo_total',
  ]);
  if (!rows.length) return;
  const startRow = lotesItemsSheet.getLastRow() + 1;
  lotesItemsSheet.getRange(startRow, 1, rows.length, 8).setValues(rows);
}

function updateLoteFile(loteRowIndex, fileName, fileId) {
  const lotesSheet = getLotesSheet();
  const link = `https://drive.google.com/file/d/${fileId}/view`;
  lotesSheet.getRange(loteRowIndex, 9, 1, 3).setValues([[fileName, fileId, link]]);
  lotesSheet.getRange(loteRowIndex, 6).setValue(new Date());
}

function updateLoteConfirmado(loteId) {
  const lotesSheet = getLotesSheet();
  const loteMeta = getLoteMetaById(loteId);
  if (!loteMeta) return;
  lotesSheet.getRange(loteMeta.rowIndex, 5, 1, 2).setValues([['confirmado', new Date()]]);
}

function updateLoteDenegado(loteId) {
  const lotesSheet = getLotesSheet();
  const loteMeta = getLoteMetaById(loteId);
  if (!loteMeta) return;
  lotesSheet.getRange(loteMeta.rowIndex, 5, 1, 2).setValues([['denegado', new Date()]]);
}

function updateLoteRechazado(loteId, nota) {
  const lotesSheet = getLotesSheet();
  const loteMeta = getLoteMetaById(loteId);
  if (!loteMeta) return;
  lotesSheet.getRange(loteMeta.rowIndex, 5, 1, 2).setValues([['rechazado', new Date()]]);
  const noteValue = nota || '';
  if (!noteValue) return;
  const headers = lotesSheet.getRange(1, 1, 1, lotesSheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  const idxNota = headers.indexOf('nota');
  if (idxNota >= 0) {
    lotesSheet.getRange(loteMeta.rowIndex, idxNota + 1, 1, 1).setValue(noteValue);
  }
}

function findMovimientosFacturados(numeros) {
  if (!numeros.length) return [];
  const movimientosSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!movimientosSheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}".`);
  const data = movimientosSheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxNumero = headers.indexOf('numero');
  const idxFacturado = headers.indexOf('facturado');
  if (idxNumero < 0 || idxFacturado < 0) {
    throw new Error('Columnas requeridas faltantes en movimientos.');
  }
  const numeroSet = {};
  numeros.forEach(val => {
    if (val !== '' && val !== null) numeroSet[String(val)] = true;
  });
  const found = {};
  for (let i = 1; i < data.length; i += 1) {
    const numero = String(data[i][idxNumero] || '');
    if (!numeroSet[numero]) continue;
    const facturado = Number(data[i][idxFacturado] || 0);
    if (facturado === 1) found[numero] = true;
  }
  return Object.keys(found);
}

function isNumeroInLoteConfirmado(numero) {
  const itemsSheet = SpreadsheetApp.getActive().getSheetByName(LOTES_ITEMS_SHEET_NAME);
  if (!itemsSheet) return false;
  const itemsData = itemsSheet.getDataRange().getValues();
  if (itemsData.length < 2) return false;
  const itemsHeaders = itemsData[0].map(h => String(h).toLowerCase());
  const idxLote = itemsHeaders.indexOf('lote_id');
  const idxNumero = itemsHeaders.indexOf('numero_movimiento');
  if (idxLote < 0 || idxNumero < 0) return false;
  const loteIds = {};
  for (let i = 1; i < itemsData.length; i += 1) {
    if (String(itemsData[i][idxNumero] || '') === String(numero)) {
      const loteId = Number(itemsData[i][idxLote] || 0);
      if (loteId) loteIds[loteId] = true;
    }
  }
  const loteKeys = Object.keys(loteIds);
  if (!loteKeys.length) return false;
  const lotesSheet = getLotesSheet();
  if (!lotesSheet) return false;
  const lotesData = lotesSheet.getDataRange().getValues();
  if (lotesData.length < 2) return false;
  const lotesHeaders = lotesData[0].map(h => String(h).toLowerCase());
  const idxLoteId = lotesHeaders.indexOf('lote_id');
  const idxEstado = lotesHeaders.indexOf('estado');
  if (idxLoteId < 0 || idxEstado < 0) return false;
  for (let i = 1; i < lotesData.length; i += 1) {
    const loteId = Number(lotesData[i][idxLoteId] || 0);
    if (!loteIds[loteId]) continue;
    const estado = String(lotesData[i][idxEstado] || '').toLowerCase();
    if (estado === 'confirmado') return true;
  }
  return false;
}

function getLotesSheet() {
  const sheet = getOrCreateSheet(LOTES_SHEET_NAME, LOTES_HEADERS);
  ensureLotesHeaders(sheet, LOTES_HEADERS);
  return sheet;
}

function ensureLotesHeaders(sheet, headers) {
  if (!sheet) return;
  const lastCol = Math.max(sheet.getLastColumn(), headers.length);
  const row = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const normalized = row.map(h => String(h).toLowerCase());
  let changed = false;
  headers.forEach((header) => {
    if (!normalized.includes(header)) {
      row.push(header);
      normalized.push(header);
      changed = true;
    }
  });
  if (changed) {
    sheet.getRange(1, 1, 1, row.length).setValues([row]);
  }
}

function markMovimientosFacturados(numeros) {
  if (!numeros.length) return;
  const movimientosSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!movimientosSheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}".`);
  const lastRow = movimientosSheet.getLastRow();
  if (lastRow < 2) return;
  const numerosRange = movimientosSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const facturadoRange = movimientosSheet.getRange(2, 10, lastRow - 1, 1).getValues();
  const numeroSet = {};
  numeros.forEach(val => { numeroSet[String(val)] = true; });
  let updated = false;
  for (let i = 0; i < numerosRange.length; i += 1) {
    const numero = String(numerosRange[i][0] || '');
    if (numeroSet[numero]) {
      facturadoRange[i][0] = 1;
      updated = true;
    }
  }
  if (updated) {
    movimientosSheet.getRange(2, 10, facturadoRange.length, 1).setValues(facturadoRange);
  }
}

function deleteCajaByNumero(numero) {
  const cajaSheet = SpreadsheetApp.getActive().getSheetByName(CAJA_SHEET_NAME);
  if (!cajaSheet) return;
  const data = cajaSheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0].map(h => String(h).toLowerCase());
  let idxNumero = headers.indexOf('numero_movimiento');
  if (idxNumero < 0) idxNumero = headers.indexOf('numero');
  if (idxNumero < 0 && data[0].length >= 6) idxNumero = 5;
  if (idxNumero < 0) return;
  for (let i = data.length - 1; i >= 1; i -= 1) {
    if (String(data[i][idxNumero] || '') === String(numero)) {
      cajaSheet.deleteRow(i + 1);
    }
  }
}

function updateCajaMontoByNumero(numero, monto) {
  const cajaSheet = SpreadsheetApp.getActive().getSheetByName(CAJA_SHEET_NAME);
  if (!cajaSheet) return;
  const data = cajaSheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0].map(h => String(h).toLowerCase());
  let idxNumero = headers.indexOf('numero_movimiento');
  if (idxNumero < 0) idxNumero = headers.indexOf('numero');
  if (idxNumero < 0 && data[0].length >= 6) idxNumero = 5;
  let idxMonto = headers.indexOf('monto');
  if (idxMonto < 0 && data[0].length >= 4) idxMonto = 3;
  if (idxNumero < 0 || idxMonto < 0) return;
  for (let i = 1; i < data.length; i += 1) {
    if (String(data[i][idxNumero] || '') === String(numero)) {
      cajaSheet.getRange(i + 1, idxMonto + 1).setValue(monto);
    }
  }
}

function buildFacturacionFile(summaryRows, fileName) {
  const templateFile = DriveApp.getFileById(FACTURACION_TEMPLATE_ID);
  const tempFile = templateFile.makeCopy(fileName);
  const sheet = SpreadsheetApp.openById(tempFile.getId()).getSheets()[0];
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
  }
  const rows = summaryRows.map(item => ([
    'CF',
    'servicio',
    'FACT',
    item.cantidad,
    `Por venta de ${item.cantidad} * ${item.nombre} (${item.codigo})`,
    item.valorUnitario || 0,
    item.descuento || 0,
  ]));
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  SpreadsheetApp.flush();
  const blob = exportSheetAsXlsx_(tempFile.getId(), fileName);
  const folder = DriveApp.getFolderById(FACTURACION_FOLDER_ID);
  const file = folder.createFile(blob);
  tempFile.setTrashed(true);
  return file.getId();
}

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0 && headers && headers.length) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function parseDateValue(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]') return value;
  if (typeof value === 'number') return new Date(value);
  if (typeof value === 'string') {
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) {
      const parts = value.split('-').map(Number);
      return new Date(parts[0], parts[1] - 1, parts[2]);
    }
    if (value.includes('/')) {
      const parts = value.split(' ');
      const dateParts = parts[0].split('/');
      if (dateParts.length === 3) {
        const day = Number(dateParts[0]);
        const month = Number(dateParts[1]) - 1;
        const year = Number(dateParts[2]);
        const time = (parts[1] || '00:00:00').split(':').map(Number);
        return new Date(year, month, day, time[0] || 0, time[1] || 0, time[2] || 0);
      }
    }
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function loadCatalogMap() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i += 1) {
    const sheet = sheets[i];
    if (sheet.getLastRow() < 2) continue;
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headers = headerRow.map(h => String(h).toLowerCase());
    const idxCodigo = headers.indexOf('codigo');
    const idxNombre = headers.indexOf('nombre');
    if (idxCodigo >= 0 && idxNombre >= 0) {
      const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      const map = {};
      data.forEach(row => {
        const codigo = String(row[idxCodigo] || '').trim();
        const nombre = String(row[idxNombre] || '').trim();
        if (codigo) map[codigo] = nombre;
      });
      return map;
    }
  }
  return {};
}

function parsePayload(e) {
  if (e && e.parameter && e.parameter.payload) {
    try {
      return JSON.parse(e.parameter.payload);
    } catch (err) {
      throw new Error('JSON inválido en payload');
    }
  }
  if (e && e.postData && e.postData.type === 'application/json') {
    try {
      return JSON.parse(e.postData.contents || '{}') || {};
    } catch (err) {
      throw new Error('JSON inválido');
    }
  }
  // Fallback a formulario simple (mantiene compat).
  const data = e.parameter || {};
  const item = {
    codigo: data.codigo || '',
    tipo: data.tipo === undefined ? 1 : (Number(data.tipo) || 1),
    cantidad: Number(data.cantidad || 0),
    valor_unitario: Number(data.valor_unitario || 0),
    comentario: data.comentario,
    facturado: data.facturado === 'on' || data.facturado === 'true' ? 1 : data.facturado || 0,
    credito: data.credito === 'on' || data.credito === 'true' ? 1 : data.credito || 0,
  };
  return {
    fecha: data.fecha,
    facturado: item.facturado,
    credito: item.credito,
    items: [item],
  };
}

function authorizeDrive() {
  const templateId = '15LPPOmsS9HrpOw6SP3LYmi9Xm5hGR70n';
  DriveApp.getFileById(templateId).getName();
}

function testDriveAndSheets() {
  const templateFile = DriveApp.getFileById(FACTURACION_TEMPLATE_ID);
  const tempFile = templateFile.makeCopy('test_facturacion_permissions');
  const sheet = SpreadsheetApp.openById(tempFile.getId()).getSheets()[0];
  sheet.getRange(2, 1).setValue('ok');
  exportSheetAsXlsx_(tempFile.getId(), 'test_facturacion_permissions.xlsx');
  tempFile.setTrashed(true);
}

function exportSheetAsXlsx_(sheetId, fileName) {
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
    muteHttpExceptions: true,
  });
  const status = response.getResponseCode();
  if (status >= 400) {
    throw new Error(`Error exportando XLSX (${status})`);
  }
  return response.getBlob().setName(fileName);
}
