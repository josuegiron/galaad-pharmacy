/**
 * Web App para registrar ventas en la hoja "movimientos" y hacer el registro en caja.
 * Recibe POST (form-data o JSON payload) con múltiples ítems y agrega filas en lote.
 * También registra cada movimiento en la pestaña "caja" con tipo 1 (venta).
 */
const SHEET_NAME = 'movimientos';
const CAJA_SHEET_NAME = 'caja';

function doGet() {
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

// Preflight CORS
function doOptions() {
  return ContentService.createTextOutput('').setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  try {
    const payload = parsePayload(e);
    const items = payload.items || [];
    if (!items.length) throw new Error('No hay items para registrar.');
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}".`);

    const lock = LockService.getDocumentLock();
    lock.waitLock(15000);
    const lastRow = sheet.getLastRow();
    let nextNum = 1;
    if (lastRow >= 2) {
      const lastValue = sheet.getRange(lastRow, 1).getValue();
      nextNum = (typeof lastValue === 'number' && !isNaN(lastValue)) ? lastValue + 1 : lastRow;
    }

    const rows = items.map((item, idx) => {
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
    lock.releaseLock();

    insertCaja(payload, rows);

    return jsonResponse({ ok: true, inserted: rows.length, first_numero: nextNum });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message || String(err) }, 400);
  }
}

function jsonResponse(obj, status) {
  const out = ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
  if (status) out.setResponseCode(status);
  return out;
}

function insertCaja(payload, rows) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(CAJA_SHEET_NAME);
  if (!sheet || !rows.length) return;
  // row: [numero, fecha, codigo, tipo, cantidad, valorUnitario, descuento, costo_total, ...]
  const lastRow = sheet.getLastRow();
  const nextId = Math.max(1, lastRow);
  const cajaRows = rows.map((row, idx) => {
    const numero = row[0];
    const fecha = row[1];
    const cantidad = Number(row[4]) || 0;
    const valor = Number(row[5]) || 0;
    const descuento = Number(row[6]) || 0;
    const monto = (Math.abs(cantidad) * valor) - descuento;
    const concepto = row[8] || payload.concepto || 'Venta';
    const id = nextId + idx;
    return [id, fecha, 1, monto, concepto, numero];
  });
  sheet.getRange(lastRow + 1, 1, cajaRows.length, 6).setValues(cajaRows);
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
