/**
 * Web App para registrar ventas en la hoja "movimientos" y hacer el registro en caja.
 * Recibe POST (form-data o JSON payload) con múltiples ítems y agrega filas en lote.
 * También registra cada movimiento en la pestaña "caja" con tipo 1 (venta).
 */
const SHEET_NAME = 'movimientos';
const CAJA_SHEET_NAME = 'caja';
const TRANSACCIONES_SHEET_NAME = 'transacciones';
const LOTES_SHEET_NAME = 'lotes_facturacion';
const LOTES_ITEMS_SHEET_NAME = 'lotes_items';
const CIERRES_SHEET_NAME = 'cierres_caja';
const PRODUCTOS_SHEET_NAME = 'productos';
const TRANSACCIONES_HEADERS = [
  'id', 'tipo', 'fecha', 'total', 'estado', 'usuario', 'timestamp', 'referencia_externa', 'nota',
];
const LOTES_HEADERS = [
  'lote_id', 'fecha_creacion', 'fecha_desde', 'fecha_hasta', 'estado', 'fecha_actualizacion',
  'total_items', 'total_monto', 'archivo_nombre', 'archivo_id', 'url', 'nota',
];
const CIERRES_HEADERS = [
  'fecha', 'saldo_inicial', 'total_entradas', 'total_salidas', 'saldo_final', 'timestamp', 'usuario',
];
const PRODUCTOS_HEADERS = [
  'codigo', 'nombre', 'afecto', 'proveedor', 'costo_unitario', 'precio_venta', 'activo',
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
    if (action === 'caja_listar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = listCaja(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'productos_listar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = listProductos(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'cierres_listar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = listCierres(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'transacciones_listar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = listTransacciones(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'caja_crear') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = createCaja(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'caja_actualizar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = updateCaja(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'caja_eliminar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = deleteCaja(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'productos_crear') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = createProducto(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'productos_actualizar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = updateProducto(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'productos_eliminar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = deleteProducto(payload);
      return callback ? jsonpResponse(result, callback) : jsonResponse(result);
    }
    if (action === 'caja_cerrar') {
      const payload = e && e.parameter ? e.parameter : {};
      const result = closeCaja(payload);
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
    if (action === 'caja_crear') {
      return jsonResponse(createCaja(payload));
    }
    if (action === 'productos_listar') {
      return jsonResponse(listProductos(payload));
    }
    if (action === 'cierres_listar') {
      return jsonResponse(listCierres(payload));
    }
    if (action === 'transacciones_listar') {
      return jsonResponse(listTransacciones(payload));
    }
    if (action === 'caja_actualizar') {
      return jsonResponse(updateCaja(payload));
    }
    if (action === 'caja_eliminar') {
      return jsonResponse(deleteCaja(payload));
    }
    if (action === 'productos_crear') {
      return jsonResponse(createProducto(payload));
    }
    if (action === 'productos_actualizar') {
      return jsonResponse(updateProducto(payload));
    }
    if (action === 'productos_eliminar') {
      return jsonResponse(deleteProducto(payload));
    }
    if (action === 'caja_cerrar') {
      return jsonResponse(closeCaja(payload));
    }
    const items = payload.items || [];
    if (!items.length) throw new Error('No hay items para registrar.');
    items.forEach((item) => {
      const fechaItem = item.fecha ? new Date(item.fecha) : (payload.fecha ? new Date(payload.fecha) : new Date());
      assertDateNotClosed(fechaItem);
    });
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}".`);
    let headers = ensureHeader(sheet, 'transaccion_id');
    headers = ensureHeader(sheet, 'timestamp');
    const idxNumero = headers.indexOf('numero') >= 0 ? headers.indexOf('numero') : 0;
    const idxFecha = headers.indexOf('fecha') >= 0 ? headers.indexOf('fecha') : 1;
    const idxCodigo = headers.indexOf('codigo_producto') >= 0 ? headers.indexOf('codigo_producto') : 2;
    const idxTipo = headers.indexOf('tipo') >= 0 ? headers.indexOf('tipo') : 3;
    const idxCantidad = headers.indexOf('cantidad') >= 0 ? headers.indexOf('cantidad') : 4;
    const idxValor = headers.indexOf('valor_unitario') >= 0 ? headers.indexOf('valor_unitario') : 5;
    const idxDescuento = headers.indexOf('descuento') >= 0 ? headers.indexOf('descuento') : 6;
    const idxTotal = headers.indexOf('costo_total') >= 0 ? headers.indexOf('costo_total') : 7;
    const idxComentario = headers.indexOf('comentario') >= 0 ? headers.indexOf('comentario') : 8;
    const idxConfirmado = headers.indexOf('confirmado') >= 0 ? headers.indexOf('confirmado') : 9;
    const idxCredito = headers.indexOf('credito') >= 0 ? headers.indexOf('credito') : 10;
    const idxTransaccionId = headers.indexOf('transaccion_id');
    const idxTimestamp = headers.indexOf('timestamp');

    let rows = [];
    let nextNum = 1;
    const timestamp = new Date();
    const tipoNumerico = getTipoNumericoFromItems(items);
    const tipoTextoFromItems = normalizeTransaccionTipo(tipoNumerico);
    const tipoTextoFromPayload = normalizeTransaccionTipo(payload.tipo);
    const creditoFlag = getCreditoFlag(items, payload);
    const referenciaExterna = payload.referencia_externa || payload.referencia || '';
    const nota = payload.nota || '';
    const usuario = payload.usuario || '';
    let transaccionId = Number(payload.transaccion_id || 0);
    const transaccionIds = {};
    items.forEach((item) => {
      if (item.transaccion_id) transaccionIds[String(item.transaccion_id)] = true;
    });
    const transaccionIdKeys = Object.keys(transaccionIds);
    if (!transaccionId && transaccionIdKeys.length) {
      if (transaccionIdKeys.length > 1) {
        throw new Error('Todos los items deben pertenecer a la misma transacción.');
      }
      transaccionId = Number(transaccionIdKeys[0]);
    }
    let transaccionMeta = transaccionId ? getTransaccionMetaById(transaccionId) : null;
    if (transaccionId && !transaccionMeta) throw new Error('Transacción no encontrada.');
    if (transaccionMeta && String(transaccionMeta.estado || '').toLowerCase() === 'anulada') {
      throw new Error('Transacción anulada.');
    }
    const tipoTexto = transaccionMeta
      ? normalizeTransaccionTipo(transaccionMeta.tipo)
      : (tipoTextoFromPayload || tipoTextoFromItems);
    if (!tipoTexto) throw new Error('Tipo de transacción inválido.');
    if (tipoTextoFromItems && tipoTextoFromItems !== tipoTexto) {
      throw new Error('Los items no coinciden con el tipo de transacción.');
    }
    if (tipoTextoFromPayload && tipoTextoFromPayload !== tipoTexto) {
      throw new Error('El tipo del payload no coincide con la transacción.');
    }
    if (transaccionIdKeys.length && transaccionId) {
      if (transaccionIdKeys.some(id => Number(id) !== transaccionId)) {
        throw new Error('Todos los items deben pertenecer a la misma transacción.');
      }
    }
    if (transaccionMeta) {
      const summary = getTransaccionSummary(transaccionId);
      if (summary && summary.count > 0 && summary.credito !== creditoFlag) {
        throw new Error('La transacción ya tiene un crédito distinto.');
      }
    } else {
      const fechaTransaccion = payload.fecha ? new Date(payload.fecha) : (items[0].fecha ? new Date(items[0].fecha) : new Date());
      const created = createTransaccionRecord({
        tipo: tipoTexto,
        fecha: fechaTransaccion,
        usuario: usuario,
        referencia_externa: referenciaExterna,
        nota: nota,
      });
      transaccionId = created.id;
      transaccionMeta = getTransaccionMetaById(transaccionId);
    }
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
        const confirmado = item.confirmado !== undefined
          ? item.confirmado
          : (item.facturado !== undefined
            ? item.facturado
            : (payload.confirmado !== undefined
              ? payload.confirmado
              : (payload.facturado || 0)));
        const credito = item.credito !== undefined ? item.credito : (payload.credito || 0);
        const costoTotal = (Math.abs(cantidad) * valorUnitario) - descuento;
        const row = new Array(headers.length).fill('');
        row[idxNumero] = nextNum + idx;
        row[idxFecha] = fecha;
        row[idxCodigo] = codigo;
        row[idxTipo] = tipo;
        row[idxCantidad] = cantidad;
        row[idxValor] = valorUnitario;
        row[idxDescuento] = descuento;
        row[idxTotal] = costoTotal;
        row[idxComentario] = comentario;
        row[idxConfirmado] = confirmado;
        row[idxCredito] = credito;
        if (idxTransaccionId >= 0) row[idxTransaccionId] = transaccionId;
        if (idxTimestamp >= 0) row[idxTimestamp] = timestamp;
        return row;
      });

      sheet.getRange(lastRow + 1, 1, rows.length, headers.length).setValues(rows);
    } finally {
      lock.releaseLock();
    }

    syncTransaccionAndCaja(transaccionId);

    return jsonResponse({ ok: true, inserted: rows.length, first_numero: nextNum, transaccion_id: transaccionId });
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

function normalizeTransaccionTipo(value) {
  if (value === undefined || value === null) return '';
  const raw = String(value).toLowerCase().trim();
  if (raw === 'venta' || raw === 'compra') return raw;
  const num = Number(value);
  if (num === 1) return 'venta';
  if (num === 2) return 'compra';
  return '';
}

function getTipoNumericoFromItems(items) {
  let tipo = null;
  items.forEach((item) => {
    const current = item.tipo === undefined ? 1 : (Number(item.tipo) || 1);
    if (tipo === null) {
      tipo = current;
      return;
    }
    if (tipo !== current) {
      throw new Error('Todos los items deben tener el mismo tipo.');
    }
  });
  return tipo === null ? 1 : tipo;
}

function getCreditoFlag(items, payload) {
  let credito = payload.credito !== undefined ? Number(payload.credito) : null;
  if (credito === null && items.length) {
    credito = items[0].credito !== undefined ? Number(items[0].credito) : 0;
  }
  credito = Number(credito) === 1 ? 1 : 0;
  items.forEach((item) => {
    if (item.credito === undefined) return;
    const itemCredito = Number(item.credito) === 1 ? 1 : 0;
    if (itemCredito !== credito) {
      throw new Error('Todos los items deben tener el mismo crédito.');
    }
  });
  return credito;
}

function getTransaccionesSheet() {
  const sheet = getOrCreateSheet(TRANSACCIONES_SHEET_NAME, TRANSACCIONES_HEADERS);
  ensureTransaccionesHeaders(sheet, TRANSACCIONES_HEADERS);
  return sheet;
}

function ensureTransaccionesHeaders(sheet, headers) {
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

function createTransaccionRecord(data) {
  const sheet = getTransaccionesSheet();
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const lastRow = sheet.getLastRow();
    let nextId = 1;
    if (lastRow >= 2) {
      const lastValue = sheet.getRange(lastRow, 1).getValue();
      nextId = (typeof lastValue === 'number' && !isNaN(lastValue)) ? lastValue + 1 : lastRow;
    }
    const row = new Array(sheet.getLastColumn()).fill('');
    row[0] = nextId;
    row[1] = data.tipo;
    row[2] = data.fecha || new Date();
    row[3] = 0;
    row[4] = 'pendiente';
    row[5] = data.usuario || '';
    row[6] = new Date();
    row[7] = data.referencia_externa || '';
    row[8] = data.nota || '';
    sheet.appendRow(row);
    return { id: nextId, rowIndex: sheet.getLastRow() };
  } finally {
    lock.releaseLock();
  }
}

function getTransaccionMetaById(transaccionId) {
  const sheet = getTransaccionesSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxId = headers.indexOf('id');
  const idxTipo = headers.indexOf('tipo');
  const idxFecha = headers.indexOf('fecha');
  const idxTotal = headers.indexOf('total');
  const idxEstado = headers.indexOf('estado');
  const idxUsuario = headers.indexOf('usuario');
  const idxTimestamp = headers.indexOf('timestamp');
  const idxReferencia = headers.indexOf('referencia_externa');
  const idxNota = headers.indexOf('nota');
  if (idxId < 0) return null;
  for (let i = 1; i < data.length; i += 1) {
    if (Number(data[i][idxId] || 0) === Number(transaccionId)) {
      return {
        rowIndex: i + 1,
        id: data[i][idxId],
        tipo: idxTipo >= 0 ? data[i][idxTipo] : '',
        fecha: idxFecha >= 0 ? data[i][idxFecha] : '',
        total: idxTotal >= 0 ? data[i][idxTotal] : 0,
        estado: idxEstado >= 0 ? data[i][idxEstado] : '',
        usuario: idxUsuario >= 0 ? data[i][idxUsuario] : '',
        timestamp: idxTimestamp >= 0 ? data[i][idxTimestamp] : '',
        referencia_externa: idxReferencia >= 0 ? data[i][idxReferencia] : '',
        nota: idxNota >= 0 ? data[i][idxNota] : '',
      };
    }
  }
  return null;
}

function updateTransaccionTotal(transaccionId, total) {
  const sheet = getTransaccionesSheet();
  const meta = getTransaccionMetaById(transaccionId);
  if (!meta) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  const idxTotal = headers.indexOf('total');
  const idxTimestamp = headers.indexOf('timestamp');
  if (idxTotal >= 0) sheet.getRange(meta.rowIndex, idxTotal + 1).setValue(total);
  if (idxTimestamp >= 0) sheet.getRange(meta.rowIndex, idxTimestamp + 1).setValue(new Date());
}

function updateTransaccionEstado(transaccionId, estado) {
  const sheet = getTransaccionesSheet();
  const meta = getTransaccionMetaById(transaccionId);
  if (!meta) return;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  const idxEstado = headers.indexOf('estado');
  const idxTimestamp = headers.indexOf('timestamp');
  if (idxEstado >= 0) sheet.getRange(meta.rowIndex, idxEstado + 1).setValue(estado);
  if (idxTimestamp >= 0) sheet.getRange(meta.rowIndex, idxTimestamp + 1).setValue(new Date());
}

function deleteTransaccionRecord(transaccionId) {
  const sheet = getTransaccionesSheet();
  const meta = getTransaccionMetaById(transaccionId);
  if (!meta) return;
  sheet.deleteRow(meta.rowIndex);
}

function getTransaccionSummary(transaccionId) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { total: 0, count: 0, hasConfirmado: false, allConfirmados: false, credito: 0, tipo: 0, fecha: null };
  }
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxTransaccionId = headers.indexOf('transaccion_id');
  const idxTotal = headers.indexOf('costo_total');
  const idxConfirmado = headers.indexOf('confirmado');
  const idxCredito = headers.indexOf('credito');
  const idxTipo = headers.indexOf('tipo');
  const idxFecha = headers.indexOf('fecha');
  if (idxTransaccionId < 0) return null;
  let total = 0;
  let count = 0;
  let hasConfirmado = false;
  let credito = 0;
  let tipo = 0;
  let fecha = null;
  let allConfirmados = true;
  for (let i = 1; i < data.length; i += 1) {
    if (Number(data[i][idxTransaccionId] || 0) !== Number(transaccionId)) continue;
    count += 1;
    total += Number(data[i][idxTotal] || 0);
    if (Number(data[i][idxConfirmado] || 0) === 1) {
      hasConfirmado = true;
    } else {
      allConfirmados = false;
    }
    if (Number(data[i][idxCredito] || 0) === 1) credito = 1;
    if (!tipo) tipo = Number(data[i][idxTipo] || 0);
    if (!fecha) fecha = data[i][idxFecha];
  }
  if (!count) allConfirmados = false;
  return {
    total: total,
    count: count,
    hasConfirmado: hasConfirmado,
    allConfirmados: allConfirmados,
    credito: credito,
    tipo: tipo,
    fecha: fecha,
  };
}

function confirmTransaccionesFromMovimientos(numeros) {
  if (!numeros.length) return;
  const movimientosSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!movimientosSheet) return;
  const data = movimientosSheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxNumero = headers.indexOf('numero');
  const idxTransaccionId = headers.indexOf('transaccion_id');
  if (idxNumero < 0 || idxTransaccionId < 0) return;
  const numeroSet = {};
  numeros.forEach(val => { if (val !== '' && val !== null) numeroSet[String(val)] = true; });
  const transaccionIds = {};
  for (let i = 1; i < data.length; i += 1) {
    const numero = String(data[i][idxNumero] || '');
    if (!numeroSet[numero]) continue;
    const transaccionId = Number(data[i][idxTransaccionId] || 0);
    if (transaccionId) transaccionIds[transaccionId] = true;
  }
  Object.keys(transaccionIds).forEach((id) => {
    const transaccionId = Number(id);
    const summary = getTransaccionSummary(transaccionId);
    if (!summary || !summary.count) return;
    const meta = getTransaccionMetaById(transaccionId);
    if (meta && String(meta.estado || '').toLowerCase() === 'anulada') return;
    if (!summary.allConfirmados) {
      return;
    }
    updateTransaccionEstado(transaccionId, 'confirmada');
  });
}

function syncTransaccionAndCaja(transaccionId) {
  const summary = getTransaccionSummary(transaccionId);
  if (!summary) return;
  if (!summary.count) {
    deleteTransaccionRecord(transaccionId);
    deleteCajaByTransaccionId(transaccionId);
    return;
  }
  updateTransaccionTotal(transaccionId, summary.total);
  if (summary.credito === 1) {
    deleteCajaByTransaccionId(transaccionId);
    return;
  }
  const meta = getTransaccionMetaById(transaccionId);
  const tipoTexto = normalizeTransaccionTipo(meta && meta.tipo ? meta.tipo : summary.tipo);
  if (!tipoTexto) return;
  const fecha = (meta && meta.fecha) ? meta.fecha : (summary.fecha || new Date());
  const conceptoBase = tipoTexto === 'compra' ? 'Compra' : 'Venta';
  upsertCajaByTransaccionId({
    transaccionId: transaccionId,
    tipoTexto: tipoTexto,
    fecha: fecha,
    monto: summary.total,
    concepto: `${conceptoBase} #${transaccionId}`,
  });
}

function upsertCajaByTransaccionId(data) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CAJA_SHEET_NAME);
  if (!sheet) return;
  let headers = ensureHeader(sheet, 'referencia');
  headers = ensureHeader(sheet, 'tipo_descripcion');
  headers = ensureHeader(sheet, 'timestamp');
  const idxId = headers.indexOf('id');
  const idxFecha = headers.indexOf('fecha');
  const idxTipo = headers.indexOf('tipo');
  let idxMonto = headers.indexOf('monto');
  if (idxMonto < 0 && headers.length >= 4) idxMonto = 3;
  let idxConcepto = headers.indexOf('concepto');
  if (idxConcepto < 0 && headers.length >= 5) idxConcepto = 4;
  let idxReferencia = headers.indexOf('referencia');
  if (idxReferencia < 0 && headers.length >= 6) idxReferencia = 5;
  let idxTipoDescripcion = headers.indexOf('tipo_descripcion');
  const idxTimestamp = headers.indexOf('timestamp');
  const requiredIdx = [idxId, idxFecha, idxTipo, idxMonto, idxConcepto, idxReferencia, idxTipoDescripcion];
  if (requiredIdx.some(idx => idx < 0)) return;
  const dataRange = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < dataRange.length; i += 1) {
    if (String(dataRange[i][idxReferencia] || '') === String(data.transaccionId)) {
      rowIndex = i + 1;
      break;
    }
  }
  const tipoNum = data.tipoTexto === 'compra' ? 2 : 1;
  if (rowIndex < 0) {
    const lastRow = sheet.getLastRow();
    let nextId = 1;
    if (lastRow >= 2) {
      const lastValue = sheet.getRange(lastRow, 1).getValue();
      nextId = (typeof lastValue === 'number' && !isNaN(lastValue)) ? lastValue + 1 : lastRow;
    }
    const row = new Array(sheet.getLastColumn()).fill('');
    row[idxId] = nextId;
    row[idxFecha] = data.fecha;
    row[idxTipo] = tipoNum;
    row[idxMonto] = data.monto;
    row[idxConcepto] = data.concepto;
    row[idxReferencia] = data.transaccionId;
    row[idxTipoDescripcion] = data.tipoTexto;
    if (idxTimestamp >= 0) row[idxTimestamp] = new Date();
    sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
    return;
  }
  sheet.getRange(rowIndex, idxFecha + 1).setValue(data.fecha);
  sheet.getRange(rowIndex, idxTipo + 1).setValue(tipoNum);
  sheet.getRange(rowIndex, idxMonto + 1).setValue(data.monto);
  sheet.getRange(rowIndex, idxConcepto + 1).setValue(data.concepto);
  sheet.getRange(rowIndex, idxReferencia + 1).setValue(data.transaccionId);
  sheet.getRange(rowIndex, idxTipoDescripcion + 1).setValue(data.tipoTexto);
  if (idxTimestamp >= 0) sheet.getRange(rowIndex, idxTimestamp + 1).setValue(new Date());
}

function deleteCajaByTransaccionId(transaccionId) {
  const cajaSheet = SpreadsheetApp.getActive().getSheetByName(CAJA_SHEET_NAME);
  if (!cajaSheet) return;
  const data = cajaSheet.getDataRange().getValues();
  if (data.length < 2) return;
  const headers = data[0].map(h => String(h).toLowerCase());
  let idxReferencia = headers.indexOf('referencia');
  if (idxReferencia < 0 && data[0].length >= 6) idxReferencia = 5;
  if (idxReferencia < 0) return;
  for (let i = data.length - 1; i >= 1; i -= 1) {
    if (String(data[i][idxReferencia] || '') === String(transaccionId)) {
      cajaSheet.deleteRow(i + 1);
    }
  }
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
  const idxConfirmado = headers.indexOf('confirmado');

  const requiredIdx = [idxNumero, idxFecha, idxCodigo, idxTipo, idxCantidad, idxValor, idxDescuento, idxTotal, idxConfirmado];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en movimientos.');
  }

  const catalogMap = loadCatalogMap();
  const summary = {};
  const loteItems = [];
  let totalMonto = 0;

  data.slice(1).forEach(row => {
    const tipo = Number(row[idxTipo] || 0);
    const confirmado = Number(row[idxConfirmado] || 0);
    const codigo = String(row[idxCodigo] || '').trim();
    const numero = row[idxNumero];
    if (!codigo || tipo !== 1 || confirmado !== 0) return;
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
  const alreadyConfirmados = findMovimientosConfirmados(numeros);
  if (alreadyConfirmados.length) {
    const note = `Rechazado: movimientos ya confirmados (${alreadyConfirmados.join(', ')}).`;
    updateLoteRechazado(loteId, note);
    return { ok: true, lote_id: loteId, estado: 'rechazado', already_confirmados: alreadyConfirmados };
  }
  markMovimientosConfirmados(numeros);
  confirmTransaccionesFromMovimientos(numeros);
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
  const idxFecha = headers.indexOf('fecha');
  const idxConfirmado = headers.indexOf('confirmado');
  const idxCredito = headers.indexOf('credito');
  const idxTransaccionId = headers.indexOf('transaccion_id');
  if (idxNumero < 0 || idxFecha < 0 || idxConfirmado < 0 || idxCredito < 0) {
    throw new Error('Columnas requeridas faltantes en movimientos.');
  }
  let rowIndex = -1;
  let confirmado = 0;
  let credito = 0;
  let rowDate = null;
  let transaccionId = 0;
  for (let i = 1; i < data.length; i += 1) {
    if (String(data[i][idxNumero] || '') === numero) {
      rowIndex = i + 1;
      confirmado = Number(data[i][idxConfirmado] || 0);
      credito = Number(data[i][idxCredito] || 0);
      rowDate = data[i][idxFecha];
      if (idxTransaccionId >= 0) transaccionId = Number(data[i][idxTransaccionId] || 0);
      break;
    }
  }
  if (rowIndex < 0) throw new Error('Movimiento no encontrado.');
  if (isDateClosed(rowDate)) throw new Error('Día ya conciliado y cuadrado.');
  if (transaccionId) {
    const summary = getTransaccionSummary(transaccionId);
    if (summary && summary.hasConfirmado) {
      throw new Error('No se puede eliminar una transacción con movimientos confirmados.');
    }
  }
  if (confirmado === 1) throw new Error('No se puede eliminar un movimiento confirmado.');
  movimientosSheet.deleteRow(rowIndex);
  if (transaccionId) {
    syncTransaccionAndCaja(transaccionId);
  } else if (credito !== 1) {
    deleteCajaByNumero(numero);
  }
  return { ok: true, numero: numero, transaccion_id: transaccionId || '' };
}

function updateMovimiento(payload) {
  const numeroRaw = payload && payload.numero ? payload.numero : null;
  const numero = String(numeroRaw || '').trim();
  if (!numero) throw new Error('Numero requerido.');
  const cantidadRaw = payload && payload.cantidad !== undefined ? payload.cantidad : null;
  const descuentoRaw = payload && payload.descuento !== undefined ? payload.descuento : null;
  const confirmadoRaw = payload && payload.confirmado !== undefined
    ? payload.confirmado
    : (payload && payload.facturado !== undefined ? payload.facturado : undefined);
  const valorUnitarioRaw = payload && payload.valor_unitario !== undefined ? payload.valor_unitario : undefined;
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
  const idxFecha = headers.indexOf('fecha');
  const idxCantidad = headers.indexOf('cantidad');
  const idxValor = headers.indexOf('valor_unitario');
  const idxDescuento = headers.indexOf('descuento');
  const idxTotal = headers.indexOf('costo_total');
  const idxConfirmado = headers.indexOf('confirmado');
  const idxTipo = headers.indexOf('tipo');
  const idxCredito = headers.indexOf('credito');
  const idxTransaccionId = headers.indexOf('transaccion_id');
  if (idxNumero < 0 || idxFecha < 0 || idxCantidad < 0 || idxValor < 0 || idxDescuento < 0 || idxTotal < 0 || idxConfirmado < 0 || idxTipo < 0 || idxCredito < 0) {
    throw new Error('Columnas requeridas faltantes en movimientos.');
  }
  let rowIndex = -1;
  let confirmado = 0;
  let valorUnitario = 0;
  let tipo = 0;
  let credito = 0;
  let rowDate = null;
  let transaccionId = 0;
  for (let i = 1; i < data.length; i += 1) {
    if (String(data[i][idxNumero] || '') === numero) {
      rowIndex = i + 1;
      confirmado = Number(data[i][idxConfirmado] || 0);
      valorUnitario = Number(data[i][idxValor] || 0);
      tipo = Number(data[i][idxTipo] || 0);
      credito = Number(data[i][idxCredito] || 0);
      rowDate = data[i][idxFecha];
      if (idxTransaccionId >= 0) transaccionId = Number(data[i][idxTransaccionId] || 0);
      break;
    }
  }
  if (rowIndex < 0) throw new Error('Movimiento no encontrado.');
  if (isDateClosed(rowDate)) throw new Error('Día ya conciliado y cuadrado.');
  if (transaccionId) {
    const summary = getTransaccionSummary(transaccionId);
    if (summary && summary.hasConfirmado) {
      throw new Error('No se puede editar una transacción con movimientos confirmados.');
    }
  }
  if (confirmado === 1) throw new Error('No se puede editar un movimiento confirmado.');
  let nextConfirmado = confirmado;
  if (confirmadoRaw !== undefined) {
    nextConfirmado = Number(confirmadoRaw) === 1 ? 1 : 0;
    if (nextConfirmado === 1 && isNumeroInLoteConfirmado(numero)) {
      throw new Error('Movimiento pertenece a lote confirmado.');
    }
  }
  let nextValorUnitario = valorUnitario;
  if (valorUnitarioRaw !== undefined && tipo === 2) {
    const valorInput = Number(valorUnitarioRaw);
    if (!isFinite(valorInput)) throw new Error('Valor unitario inválido.');
    nextValorUnitario = valorInput;
  }
  const cantidad = tipo === 1 ? -Math.abs(cantidadInput) : cantidadInput;
  const descuento = Math.max(0, descuentoInput);
  const costoTotal = (Math.abs(cantidad) * nextValorUnitario) - descuento;
  movimientosSheet.getRange(rowIndex, idxCantidad + 1, 1, 3).setValues([[cantidad, nextValorUnitario, descuento]]);
  movimientosSheet.getRange(rowIndex, idxTotal + 1).setValue(costoTotal);
  if (confirmadoRaw !== undefined) {
    movimientosSheet.getRange(rowIndex, idxConfirmado + 1).setValue(nextConfirmado);
  }
  if (transaccionId) {
    syncTransaccionAndCaja(transaccionId);
  } else if (credito !== 1) {
    updateCajaMontoByNumero(numero, costoTotal);
  }
  return { ok: true, numero: numero, total: costoTotal, confirmado: nextConfirmado, transaccion_id: transaccionId || '' };
}

function assertDateNotClosed(dateValue) {
  if (!dateValue) return;
  if (isDateClosed(dateValue)) throw new Error('Día ya conciliado y cuadrado.');
}

function isDateClosed(dateValue) {
  const dateTarget = parseDateValue(dateValue);
  if (!dateTarget) return false;
  return !!getCierreByDate(dateTarget);
}

function getSaldoInicialForDate(dateValue) {
  const dateTarget = parseDateValue(dateValue);
  if (!dateTarget) return 0;
  const previous = getLastCierreBefore(dateTarget);
  return previous ? Number(previous.saldo_final || 0) : 0;
}

function getCierreByDate(dateValue) {
  const sheet = getCierresSheet();
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxFecha = headers.indexOf('fecha');
  const idxSaldoInicial = headers.indexOf('saldo_inicial');
  const idxEntradas = headers.indexOf('total_entradas');
  const idxSalidas = headers.indexOf('total_salidas');
  const idxSaldoFinal = headers.indexOf('saldo_final');
  const idxTimestamp = headers.indexOf('timestamp');
  const idxUsuario = headers.indexOf('usuario');
  if (idxFecha < 0) return null;
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const targetKey = Utilities.formatDate(parseDateValue(dateValue), tz, 'yyyy-MM-dd');
  for (let i = 1; i < data.length; i += 1) {
    const rowDate = parseDateValue(data[i][idxFecha]);
    if (!rowDate) continue;
    const rowKey = Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd');
    if (rowKey === targetKey) {
      return {
        rowIndex: i + 1,
        fecha: data[i][idxFecha],
        saldo_inicial: data[i][idxSaldoInicial],
        total_entradas: data[i][idxEntradas],
        total_salidas: data[i][idxSalidas],
        saldo_final: data[i][idxSaldoFinal],
        timestamp: data[i][idxTimestamp],
        usuario: data[i][idxUsuario],
      };
    }
  }
  return null;
}

function getLastCierreBefore(dateValue) {
  const sheet = getCierresSheet();
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxFecha = headers.indexOf('fecha');
  const idxSaldoFinal = headers.indexOf('saldo_final');
  if (idxFecha < 0 || idxSaldoFinal < 0) return null;
  const targetDate = parseDateValue(dateValue);
  if (!targetDate) return null;
  let latest = null;
  for (let i = 1; i < data.length; i += 1) {
    const rowDate = parseDateValue(data[i][idxFecha]);
    if (!rowDate) continue;
    if (rowDate.getTime() >= targetDate.getTime()) continue;
    if (!latest || rowDate.getTime() > latest.date.getTime()) {
      latest = { date: rowDate, saldo_final: data[i][idxSaldoFinal] };
    }
  }
  return latest;
}

function confirmComprasForDate(dateTarget) {
  const movimientosSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!movimientosSheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}".`);
  const data = movimientosSheet.getDataRange().getValues();
  if (data.length < 2) return { numerosConfirmados: [], pendientes: false };
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxNumero = headers.indexOf('numero');
  const idxFecha = headers.indexOf('fecha');
  const idxTipo = headers.indexOf('tipo');
  const idxConfirmado = headers.indexOf('confirmado');
  if (idxNumero < 0 || idxFecha < 0 || idxTipo < 0 || idxConfirmado < 0) {
    throw new Error('Columnas requeridas faltantes en movimientos.');
  }
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const targetKey = Utilities.formatDate(dateTarget, tz, 'yyyy-MM-dd');
  const confirmadoRange = movimientosSheet.getRange(2, idxConfirmado + 1, data.length - 1, 1).getValues();
  let updated = false;
  let pendientes = false;
  const numerosConfirmados = [];
  for (let i = 1; i < data.length; i += 1) {
    const rowDate = parseDateValue(data[i][idxFecha]);
    if (!rowDate) continue;
    const rowKey = Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd');
    if (rowKey !== targetKey) continue;
    const tipo = Number(data[i][idxTipo] || 0);
    let confirmado = Number(data[i][idxConfirmado] || 0);
    if (tipo === 2 && confirmado !== 1) {
      confirmado = 1;
      data[i][idxConfirmado] = 1;
      confirmadoRange[i - 1][0] = 1;
      updated = true;
      const numero = data[i][idxNumero];
      if (numero !== '' && numero !== null) numerosConfirmados.push(numero);
    }
    if (confirmado !== 1) pendientes = true;
  }
  if (updated) {
    movimientosSheet.getRange(2, idxConfirmado + 1, confirmadoRange.length, 1).setValues(confirmadoRange);
  }
  return { numerosConfirmados: numerosConfirmados, pendientes: pendientes };
}

function closeCaja(payload) {
  const dateRaw = payload && payload.fecha ? payload.fecha : null;
  if (!dateRaw) throw new Error('Fecha requerida.');
  const dateTarget = parseDateValue(dateRaw);
  if (!dateTarget) throw new Error('Fecha inválida.');
  if (isDateClosed(dateTarget)) throw new Error('Día ya conciliado y cuadrado.');
  const previousDate = new Date(dateTarget.getTime());
  previousDate.setDate(previousDate.getDate() - 1);
  if (!getCierreByDate(previousDate)) {
    throw new Error('Debe cerrar el día anterior antes de cerrar este día.');
  }

  const cierreConfirm = confirmComprasForDate(dateTarget);
  if (cierreConfirm.numerosConfirmados.length) {
    confirmTransaccionesFromMovimientos(cierreConfirm.numerosConfirmados);
  }
  if (cierreConfirm.pendientes) {
    throw new Error('No se puede cerrar el día: hay movimientos sin confirmar.');
  }

  const cajaSheet = SpreadsheetApp.getActive().getSheetByName(CAJA_SHEET_NAME);
  if (!cajaSheet) throw new Error('No se encontró la hoja "caja".');
  const data = cajaSheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxFecha = headers.indexOf('fecha');
  const idxTipo = headers.indexOf('tipo');
  const idxMonto = headers.indexOf('monto');
  const idxConciliado = headers.indexOf('conciliado');
  const requiredIdx = [idxFecha, idxTipo, idxMonto, idxConciliado];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en caja.');
  }
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const targetKey = Utilities.formatDate(dateTarget, tz, 'yyyy-MM-dd');
  let totalEntradas = 0;
  let totalSalidas = 0;
  const conciliadoRange = data.length > 1
    ? cajaSheet.getRange(2, idxConciliado + 1, data.length - 1, 1).getValues()
    : [];
  let updateConciliado = false;
  for (let i = 1; i < data.length; i += 1) {
    const rowDate = parseDateValue(data[i][idxFecha]);
    if (!rowDate) continue;
    const rowKey = Utilities.formatDate(rowDate, tz, 'yyyy-MM-dd');
    if (rowKey !== targetKey) continue;
    const tipo = Number(data[i][idxTipo] || 0);
    const monto = Number(data[i][idxMonto] || 0);
    if (tipo === 1) totalEntradas += monto;
    if (tipo === 2) totalSalidas += monto;
    if (conciliadoRange[i - 1] && conciliadoRange[i - 1][0] !== 1) {
      conciliadoRange[i - 1][0] = 1;
      updateConciliado = true;
    }
  }
  if (updateConciliado && conciliadoRange.length) {
    cajaSheet.getRange(2, idxConciliado + 1, conciliadoRange.length, 1).setValues(conciliadoRange);
  }

  const saldoInicial = getSaldoInicialForDate(dateTarget);
  const saldoFinal = saldoInicial + totalEntradas - totalSalidas;
  const cierresSheet = getCierresSheet();
  const row = [dateTarget, saldoInicial, totalEntradas, totalSalidas, saldoFinal, new Date(), ''];
  cierresSheet.appendRow(row);
  return {
    ok: true,
    fecha: dateTarget,
    saldo_inicial: saldoInicial,
    total_entradas: totalEntradas,
    total_salidas: totalSalidas,
    saldo_final: saldoFinal,
  };
}

function listCaja(payload) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CAJA_SHEET_NAME);
  if (!sheet) return { ok: true, items: [] };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, items: [] };
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxId = headers.indexOf('id');
  const idxFecha = headers.indexOf('fecha');
  const idxTipo = headers.indexOf('tipo');
  const idxMonto = headers.indexOf('monto');
  const idxConcepto = headers.indexOf('concepto');
  const idxReferencia = headers.indexOf('referencia');
  const idxConciliado = headers.indexOf('conciliado');
  const idxTipoDescripcion = headers.indexOf('tipo_descripcion');
  const requiredIdx = [idxId, idxFecha, idxTipo, idxMonto, idxConcepto, idxReferencia, idxConciliado, idxTipoDescripcion];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en caja.');
  }
  const filterDate = payload && payload.fecha ? parseDateValue(payload.fecha) : null;
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const filterKey = filterDate ? Utilities.formatDate(filterDate, tz, 'yyyy-MM-dd') : null;
  const items = data.slice(1).map(row => ({
    id: row[idxId],
    fecha: row[idxFecha],
    tipo: row[idxTipo],
    monto: row[idxMonto],
    concepto: row[idxConcepto],
    referencia: row[idxReferencia],
    conciliado: row[idxConciliado],
    tipo_descripcion: row[idxTipoDescripcion],
  })).filter(item => {
    if (!filterKey) return true;
    const dateValue = parseDateValue(item.fecha);
    if (!dateValue) return false;
    const dateKey = Utilities.formatDate(dateValue, tz, 'yyyy-MM-dd');
    return dateKey === filterKey;
  });
  items.sort((a, b) => {
    const dateA = parseDateValue(a.fecha) || new Date(0);
    const dateB = parseDateValue(b.fecha) || new Date(0);
    return dateB.getTime() - dateA.getTime();
  });
  const summaryDate = filterDate || new Date();
  const saldoInicial = getSaldoInicialForDate(summaryDate);
  const cerrado = !!getCierreByDate(summaryDate);
  return { ok: true, items: items, saldo_inicial: saldoInicial, cerrado: cerrado };
}

function listCierres(payload) {
  const sheet = getCierresSheet();
  if (!sheet) return { ok: true, items: [] };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, items: [] };
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxFecha = headers.indexOf('fecha');
  const idxSaldoInicial = headers.indexOf('saldo_inicial');
  const idxEntradas = headers.indexOf('total_entradas');
  const idxSalidas = headers.indexOf('total_salidas');
  const idxSaldoFinal = headers.indexOf('saldo_final');
  const idxTimestamp = headers.indexOf('timestamp');
  const idxUsuario = headers.indexOf('usuario');
  const requiredIdx = [idxFecha, idxSaldoInicial, idxEntradas, idxSalidas, idxSaldoFinal];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en cierres_caja.');
  }
  const fromRaw = payload && (payload.fecha_desde || payload.desde);
  const toRaw = payload && (payload.fecha_hasta || payload.hasta);
  const fromDate = parseDateValue(fromRaw);
  const toDate = parseDateValue(toRaw);
  const fromTime = fromDate ? new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate()).getTime() : null;
  const toTime = toDate ? new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate()).getTime() : null;
  const items = data.slice(1).map(row => ({
    fecha: row[idxFecha],
    saldo_inicial: row[idxSaldoInicial],
    total_entradas: row[idxEntradas],
    total_salidas: row[idxSalidas],
    saldo_final: row[idxSaldoFinal],
    timestamp: idxTimestamp >= 0 ? row[idxTimestamp] : '',
    usuario: idxUsuario >= 0 ? row[idxUsuario] : '',
  })).filter(item => {
    const dateValue = parseDateValue(item.fecha);
    if (!dateValue) return false;
    const dayTime = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate()).getTime();
    if (fromTime && dayTime < fromTime) return false;
    if (toTime && dayTime > toTime) return false;
    return true;
  });
  items.sort((a, b) => {
    const dateA = parseDateValue(a.fecha) || new Date(0);
    const dateB = parseDateValue(b.fecha) || new Date(0);
    return dateB.getTime() - dateA.getTime();
  });
  return { ok: true, items: items };
}

function listTransacciones(payload) {
  const sheet = getTransaccionesSheet();
  if (!sheet) return { ok: true, items: [] };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, items: [] };
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxId = headers.indexOf('id');
  const idxTipo = headers.indexOf('tipo');
  const idxFecha = headers.indexOf('fecha');
  const idxTotal = headers.indexOf('total');
  const idxEstado = headers.indexOf('estado');
  const idxUsuario = headers.indexOf('usuario');
  const idxReferencia = headers.indexOf('referencia_externa');
  const idxNota = headers.indexOf('nota');
  const requiredIdx = [idxId, idxTipo, idxFecha, idxTotal, idxEstado, idxUsuario];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en transacciones.');
  }
  const rawTipo = payload && payload.tipo !== undefined ? payload.tipo : '';
  const tipoFilter = rawTipo ? normalizeTransaccionTipo(rawTipo) : '';
  if (rawTipo && !tipoFilter) throw new Error('Tipo de transacción inválido.');
  let fromRaw = payload && (payload.fecha_desde || payload.desde);
  let toRaw = payload && (payload.fecha_hasta || payload.hasta);
  if (payload && payload.fecha) {
    fromRaw = payload.fecha;
    toRaw = payload.fecha;
  }
  const fromDate = parseDateValue(fromRaw);
  const toDate = parseDateValue(toRaw);
  const fromTime = fromDate ? new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate()).getTime() : null;
  const toTime = toDate ? new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate()).getTime() : null;
  const items = data.slice(1).map(row => ({
    id: row[idxId],
    tipo: row[idxTipo],
    fecha: row[idxFecha],
    total: row[idxTotal],
    estado: row[idxEstado],
    usuario: row[idxUsuario],
    referencia_externa: idxReferencia >= 0 ? row[idxReferencia] : '',
    nota: idxNota >= 0 ? row[idxNota] : '',
  })).filter(item => {
    if (tipoFilter && normalizeTransaccionTipo(item.tipo) !== tipoFilter) return false;
    if (!fromTime && !toTime) return true;
    const dateValue = parseDateValue(item.fecha);
    if (!dateValue) return false;
    const dayTime = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate()).getTime();
    if (fromTime && dayTime < fromTime) return false;
    if (toTime && dayTime > toTime) return false;
    return true;
  });
  items.sort((a, b) => {
    const dateA = parseDateValue(a.fecha) || new Date(0);
    const dateB = parseDateValue(b.fecha) || new Date(0);
    if (dateB.getTime() !== dateA.getTime()) {
      return dateB.getTime() - dateA.getTime();
    }
    return Number(b.id || 0) - Number(a.id || 0);
  });
  return { ok: true, items: items };
}

function listProductos(payload) {
  const sheet = getProductosSheet();
  if (!sheet) return { ok: true, items: [] };
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, items: [] };
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxCodigo = headers.indexOf('codigo');
  const idxNombre = headers.indexOf('nombre');
  const idxAfecto = headers.indexOf('afecto');
  const idxProveedor = headers.indexOf('proveedor');
  const idxCosto = headers.indexOf('costo_unitario');
  const idxPrecio = headers.indexOf('precio_venta');
  const idxActivo = headers.indexOf('activo');
  const requiredIdx = [idxCodigo, idxNombre, idxAfecto, idxProveedor, idxCosto, idxPrecio, idxActivo];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en productos.');
  }
  const q = payload && payload.q ? String(payload.q).toLowerCase().trim() : '';
  const items = data.slice(1).map(row => ({
    codigo: row[idxCodigo],
    nombre: row[idxNombre],
    afecto: row[idxAfecto],
    proveedor: row[idxProveedor],
    costo_unitario: row[idxCosto],
    precio_venta: row[idxPrecio],
    activo: row[idxActivo],
  })).filter(item => {
    if (!q) return true;
    const codigo = String(item.codigo || '').toLowerCase();
    const nombre = String(item.nombre || '').toLowerCase();
    const proveedor = String(item.proveedor || '').toLowerCase();
    return codigo.includes(q) || nombre.includes(q) || proveedor.includes(q);
  });
  items.sort((a, b) => Number(a.codigo || 0) - Number(b.codigo || 0));
  return { ok: true, items: items };
}

function createProducto(payload) {
  const sheet = getProductosSheet();
  if (!sheet) throw new Error('No se encontró la hoja "productos".');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  const idxCodigo = headers.indexOf('codigo');
  const idxNombre = headers.indexOf('nombre');
  const idxAfecto = headers.indexOf('afecto');
  const idxProveedor = headers.indexOf('proveedor');
  const idxCosto = headers.indexOf('costo_unitario');
  const idxPrecio = headers.indexOf('precio_venta');
  const idxActivo = headers.indexOf('activo');
  const requiredIdx = [idxCodigo, idxNombre, idxAfecto, idxProveedor, idxCosto, idxPrecio, idxActivo];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en productos.');
  }
  const codigo = String(payload.codigo || '').trim();
  const nombre = String(payload.nombre || '').trim();
  if (!codigo) throw new Error('Código requerido.');
  if (!nombre) throw new Error('Nombre requerido.');
  const afecto = payload.afecto !== undefined ? Number(payload.afecto || 0) : 0;
  const proveedor = String(payload.proveedor || '').trim();
  const costoUnitario = payload.costo_unitario !== undefined ? Number(payload.costo_unitario || 0) : 0;
  const precioVenta = payload.precio_venta !== undefined ? Number(payload.precio_venta || 0) : 0;
  const activo = payload.activo !== undefined ? Number(payload.activo || 0) : 1;
  if (!isFinite(afecto) || !isFinite(costoUnitario) || !isFinite(precioVenta) || !isFinite(activo)) {
    throw new Error('Campos numéricos inválidos.');
  }
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i += 1) {
      if (String(data[i][idxCodigo] || '').trim() === codigo) {
        throw new Error('Código ya existe.');
      }
    }
    const row = new Array(sheet.getLastColumn()).fill('');
    row[idxCodigo] = codigo;
    row[idxNombre] = nombre;
    row[idxAfecto] = afecto;
    row[idxProveedor] = proveedor;
    row[idxCosto] = costoUnitario;
    row[idxPrecio] = precioVenta;
    row[idxActivo] = activo;
    sheet.appendRow(row);
    return { ok: true, codigo: codigo };
  } finally {
    lock.releaseLock();
  }
}

function updateProducto(payload) {
  const sheet = getProductosSheet();
  if (!sheet) throw new Error('No se encontró la hoja "productos".');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  const idxCodigo = headers.indexOf('codigo');
  const idxNombre = headers.indexOf('nombre');
  const idxAfecto = headers.indexOf('afecto');
  const idxProveedor = headers.indexOf('proveedor');
  const idxCosto = headers.indexOf('costo_unitario');
  const idxPrecio = headers.indexOf('precio_venta');
  const idxActivo = headers.indexOf('activo');
  const requiredIdx = [idxCodigo, idxNombre, idxAfecto, idxProveedor, idxCosto, idxPrecio, idxActivo];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en productos.');
  }
  const codigo = String(payload.codigo || '').trim();
  if (!codigo) throw new Error('Código requerido.');
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i += 1) {
      if (String(data[i][idxCodigo] || '').trim() === codigo) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex < 0) throw new Error('Producto no encontrado.');
    const current = data[rowIndex - 1];
    const nombre = payload.nombre !== undefined ? String(payload.nombre || '').trim() : current[idxNombre];
    const afecto = payload.afecto !== undefined ? Number(payload.afecto || 0) : current[idxAfecto];
    const proveedor = payload.proveedor !== undefined ? String(payload.proveedor || '').trim() : current[idxProveedor];
    const costoUnitario = payload.costo_unitario !== undefined ? Number(payload.costo_unitario || 0) : current[idxCosto];
    const precioVenta = payload.precio_venta !== undefined ? Number(payload.precio_venta || 0) : current[idxPrecio];
    const activo = payload.activo !== undefined ? Number(payload.activo || 0) : current[idxActivo];
    if (!isFinite(afecto) || !isFinite(costoUnitario) || !isFinite(precioVenta) || !isFinite(activo)) {
      throw new Error('Campos numéricos inválidos.');
    }
    sheet.getRange(rowIndex, idxNombre + 1).setValue(nombre);
    sheet.getRange(rowIndex, idxAfecto + 1).setValue(afecto);
    sheet.getRange(rowIndex, idxProveedor + 1).setValue(proveedor);
    sheet.getRange(rowIndex, idxCosto + 1).setValue(costoUnitario);
    sheet.getRange(rowIndex, idxPrecio + 1).setValue(precioVenta);
    sheet.getRange(rowIndex, idxActivo + 1).setValue(activo);
    return { ok: true, codigo: codigo };
  } finally {
    lock.releaseLock();
  }
}

function deleteProducto(payload) {
  const sheet = getProductosSheet();
  if (!sheet) throw new Error('No se encontró la hoja "productos".');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  const idxCodigo = headers.indexOf('codigo');
  if (idxCodigo < 0) throw new Error('Columnas requeridas faltantes en productos.');
  const codigo = String(payload.codigo || '').trim();
  if (!codigo) throw new Error('Código requerido.');
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i += 1) {
      if (String(data[i][idxCodigo] || '').trim() === codigo) {
        sheet.deleteRow(i + 1);
        return { ok: true, codigo: codigo };
      }
    }
    throw new Error('Producto no encontrado.');
  } finally {
    lock.releaseLock();
  }
}

function createCaja(payload) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CAJA_SHEET_NAME);
  if (!sheet) throw new Error('No se encontró la hoja "caja".');
  const headers = ensureHeader(sheet, 'timestamp');
  const idxId = headers.indexOf('id');
  const idxFecha = headers.indexOf('fecha');
  const idxTipo = headers.indexOf('tipo');
  const idxMonto = headers.indexOf('monto');
  const idxConcepto = headers.indexOf('concepto');
  const idxReferencia = headers.indexOf('referencia');
  const idxConciliado = headers.indexOf('conciliado');
  const idxTipoDescripcion = headers.indexOf('tipo_descripcion');
  const idxTimestamp = headers.indexOf('timestamp');
  const requiredIdx = [idxId, idxFecha, idxTipo, idxMonto, idxConcepto, idxReferencia, idxConciliado, idxTipoDescripcion];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en caja.');
  }
  const tipo = Number(payload.tipo || 0);
  const monto = Number(payload.monto || 0);
  const concepto = String(payload.concepto || '').trim();
  const referencia = String(payload.referencia || '').trim();
  const tipoDescripcion = String(payload.tipo_descripcion || '').trim();
  if (![1, 2].includes(tipo)) throw new Error('Tipo inválido.');
  if (!concepto) throw new Error('Concepto requerido.');
  if (!isFinite(monto)) throw new Error('Monto inválido.');
  const dateValue = payload.fecha ? parseDateValue(payload.fecha) : new Date();
  const fecha = dateValue || new Date();
  assertDateNotClosed(fecha);
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const lastRow = sheet.getLastRow();
    let nextId = 1;
    if (lastRow >= 2) {
      const lastValue = sheet.getRange(lastRow, 1).getValue();
      nextId = (typeof lastValue === 'number' && !isNaN(lastValue)) ? lastValue + 1 : lastRow;
    }
    const row = new Array(sheet.getLastColumn()).fill('');
    row[idxId] = nextId;
    row[idxFecha] = fecha;
    row[idxTipo] = tipo;
    row[idxMonto] = monto;
    row[idxConcepto] = concepto;
    row[idxReferencia] = referencia;
    row[idxConciliado] = '';
    row[idxTipoDescripcion] = tipoDescripcion || 'otros';
    if (idxTimestamp >= 0) row[idxTimestamp] = new Date();
    sheet.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
    return { ok: true, id: nextId };
  } finally {
    lock.releaseLock();
  }
}

function updateCaja(payload) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CAJA_SHEET_NAME);
  if (!sheet) throw new Error('No se encontró la hoja "caja".');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  const idxId = headers.indexOf('id');
  const idxFecha = headers.indexOf('fecha');
  const idxTipo = headers.indexOf('tipo');
  const idxMonto = headers.indexOf('monto');
  const idxConcepto = headers.indexOf('concepto');
  const idxReferencia = headers.indexOf('referencia');
  const idxConciliado = headers.indexOf('conciliado');
  const idxTipoDescripcion = headers.indexOf('tipo_descripcion');
  const requiredIdx = [idxId, idxFecha, idxTipo, idxMonto, idxConcepto, idxReferencia, idxConciliado, idxTipoDescripcion];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en caja.');
  }
  const id = String(payload.id || '').trim();
  if (!id) throw new Error('ID requerido.');
  const tipo = Number(payload.tipo || 0);
  const monto = Number(payload.monto || 0);
  const concepto = String(payload.concepto || '').trim();
  const referencia = String(payload.referencia || '').trim();
  const tipoDescripcion = String(payload.tipo_descripcion || '').trim();
  if (![1, 2].includes(tipo)) throw new Error('Tipo inválido.');
  if (!concepto) throw new Error('Concepto requerido.');
  if (!isFinite(monto)) throw new Error('Monto inválido.');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let conciliado = 0;
  let currentTipoDescripcion = '';
  let rowDate = null;
  for (let i = 1; i < data.length; i += 1) {
    if (String(data[i][idxId] || '') === id) {
      rowIndex = i + 1;
      conciliado = Number(data[i][idxConciliado] || 0);
      currentTipoDescripcion = String(data[i][idxTipoDescripcion] || '');
      rowDate = data[i][idxFecha];
      break;
    }
  }
  if (rowIndex < 0) throw new Error('Movimiento no encontrado.');
  if (isDateClosed(rowDate)) throw new Error('Día ya conciliado y cuadrado.');
  if (conciliado === 1) throw new Error('Movimiento conciliado.');
  const tipoDescripcionLower = currentTipoDescripcion.toLowerCase();
  if (tipoDescripcionLower === 'venta' || tipoDescripcionLower === 'compra') {
    throw new Error('No se puede editar un movimiento de venta.');
  }
  sheet.getRange(rowIndex, idxTipo + 1).setValue(tipo);
  sheet.getRange(rowIndex, idxMonto + 1).setValue(monto);
  sheet.getRange(rowIndex, idxConcepto + 1).setValue(concepto);
  sheet.getRange(rowIndex, idxReferencia + 1).setValue(referencia);
  sheet.getRange(rowIndex, idxTipoDescripcion + 1).setValue(tipoDescripcion || 'otros');
  return { ok: true, id: id };
}

function deleteCaja(payload) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CAJA_SHEET_NAME);
  if (!sheet) throw new Error('No se encontró la hoja "caja".');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  const idxId = headers.indexOf('id');
  const idxFecha = headers.indexOf('fecha');
  const idxConciliado = headers.indexOf('conciliado');
  const idxTipoDescripcion = headers.indexOf('tipo_descripcion');
  const requiredIdx = [idxId, idxFecha, idxConciliado, idxTipoDescripcion];
  if (requiredIdx.some(idx => idx < 0)) {
    throw new Error('Columnas requeridas faltantes en caja.');
  }
  const id = String(payload.id || '').trim();
  if (!id) throw new Error('ID requerido.');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let conciliado = 0;
  let currentTipoDescripcion = '';
  let rowDate = null;
  for (let i = 1; i < data.length; i += 1) {
    if (String(data[i][idxId] || '') === id) {
      rowIndex = i + 1;
      conciliado = Number(data[i][idxConciliado] || 0);
      currentTipoDescripcion = String(data[i][idxTipoDescripcion] || '');
      rowDate = data[i][idxFecha];
      break;
    }
  }
  if (rowIndex < 0) throw new Error('Movimiento no encontrado.');
  if (isDateClosed(rowDate)) throw new Error('Día ya conciliado y cuadrado.');
  if (conciliado === 1) throw new Error('Movimiento conciliado.');
  const tipoDescripcionLower = currentTipoDescripcion.toLowerCase();
  if (tipoDescripcionLower === 'venta' || tipoDescripcionLower === 'compra') {
    throw new Error('No se puede eliminar un movimiento de venta.');
  }
  sheet.deleteRow(rowIndex);
  return { ok: true, id: id };
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

function findMovimientosConfirmados(numeros) {
  if (!numeros.length) return [];
  const movimientosSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!movimientosSheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}".`);
  const data = movimientosSheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(h => String(h).toLowerCase());
  const idxNumero = headers.indexOf('numero');
  const idxConfirmado = headers.indexOf('confirmado');
  if (idxNumero < 0 || idxConfirmado < 0) {
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
    const confirmado = Number(data[i][idxConfirmado] || 0);
    if (confirmado === 1) found[numero] = true;
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

function getCierresSheet() {
  const sheet = getOrCreateSheet(CIERRES_SHEET_NAME, CIERRES_HEADERS);
  ensureCierresHeaders(sheet, CIERRES_HEADERS);
  return sheet;
}

function getProductosSheet() {
  const sheet = getOrCreateSheet(PRODUCTOS_SHEET_NAME, PRODUCTOS_HEADERS);
  ensureProductosHeaders(sheet, PRODUCTOS_HEADERS);
  return sheet;
}

function ensureCierresHeaders(sheet, headers) {
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

function ensureProductosHeaders(sheet, headers) {
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

function ensureHeader(sheet, headerName) {
  if (!sheet) return [];
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const row = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const normalized = row.map(h => String(h).toLowerCase());
  const target = String(headerName || '').toLowerCase();
  if (target && !normalized.includes(target)) {
    row.push(target);
    normalized.push(target);
    sheet.getRange(1, 1, 1, row.length).setValues([row]);
  }
  return normalized;
}

function markMovimientosConfirmados(numeros) {
  if (!numeros.length) return;
  const movimientosSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!movimientosSheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}".`);
  const lastRow = movimientosSheet.getLastRow();
  if (lastRow < 2) return;
  const headers = movimientosSheet.getRange(1, 1, 1, movimientosSheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase());
  const idxNumero = headers.indexOf('numero');
  const idxConfirmado = headers.indexOf('confirmado');
  if (idxNumero < 0 || idxConfirmado < 0) {
    throw new Error('Columnas requeridas faltantes en movimientos.');
  }
  const numerosRange = movimientosSheet.getRange(2, idxNumero + 1, lastRow - 1, 1).getValues();
  const confirmadoRange = movimientosSheet.getRange(2, idxConfirmado + 1, lastRow - 1, 1).getValues();
  const numeroSet = {};
  numeros.forEach(val => { numeroSet[String(val)] = true; });
  let updated = false;
  for (let i = 0; i < numerosRange.length; i += 1) {
    const numero = String(numerosRange[i][0] || '');
    if (numeroSet[numero]) {
      confirmadoRange[i][0] = 1;
      updated = true;
    }
  }
  if (updated) {
    movimientosSheet.getRange(2, idxConfirmado + 1, confirmadoRange.length, 1).setValues(confirmadoRange);
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
  const confirmadoInput = data.confirmado !== undefined ? data.confirmado : data.facturado;
  const item = {
    codigo: data.codigo || '',
    tipo: data.tipo === undefined ? 1 : (Number(data.tipo) || 1),
    cantidad: Number(data.cantidad || 0),
    valor_unitario: Number(data.valor_unitario || 0),
    comentario: data.comentario,
    confirmado: confirmadoInput === 'on' || confirmadoInput === 'true' ? 1 : (confirmadoInput || 0),
    credito: data.credito === 'on' || data.credito === 'true' ? 1 : data.credito || 0,
    transaccion_id: data.transaccion_id || '',
  };
  return {
    fecha: data.fecha,
    tipo: data.tipo,
    confirmado: item.confirmado,
    credito: item.credito,
    transaccion_id: data.transaccion_id || '',
    referencia_externa: data.referencia_externa || '',
    nota: data.nota || '',
    usuario: data.usuario || '',
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
