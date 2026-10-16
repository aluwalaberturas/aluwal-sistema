/**
 * ALUWAL HUB — backend único para presupuestos / ventas / obra.
 *
 * SISTEMA ALUWAL (producción). Va en su propia Google Sheet, aparte de las que usan las apps viejas.
 *
 * Qué hace:
 *  - Lleva UN contador de presupuestos (B-0001, B-0002...) compartido por todas las apps.
 *  - Guarda cada presupuesto con sus ítems detallados (JSON), no como texto cortado.
 *  - Permite marcar un presupuesto como vendido y consultarlo por número.
 *
 * Uso (desde las apps):
 *  GET  ?action=nuevo&origen=online|interno   -> reserva el próximo número y devuelve {ok,num}
 *  GET  ?action=get&num=B-0002                -> devuelve el presupuesto completo
 *  GET  ?action=list[&estado=vendido]         -> lista de presupuestos
 *  POST {action:'guardar', num, ...campos}    -> completa/actualiza un presupuesto reservado
 *  POST {action:'vendido', num, vendedor?}    -> lo marca vendido (con fecha) y lo devuelve
 *  (El POST se manda como text/plain con el JSON adentro, para evitar problemas de CORS.)
 */

// ===== Configuración =====
const PREFIJO = 'B-';
const DIGITOS = 4;          // B-0001
const NUM_INICIAL = 0;      // el primer número que se entrega es NUM_INICIAL + 1  ->  B-0001
const HOJA = 'Presupuestos';
const TZ = 'America/Argentina/Buenos_Aires';
const COLS = ['num', 'fecha', 'origen', 'cliente', 'tel', 'mail', 'dir', 'localidad',
              'modo', 'total', 'obs', 'items', 'estado', 'fechaVenta', 'vendedor'];

// Correr UNA vez a mano desde el editor, después de pegar el código en la Sheet nueva.
function setup() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) sh = ss.insertSheet(HOJA);
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight('bold');
  sh.setFrozenRows(1);
  // Todo como texto (menos el total) para que Sheets no convierta teléfonos ni fechas
  COLS.forEach((c, i) => { if (c !== 'total') sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@'); });
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ULTIMO') === null) props.setProperty('ULTIMO', String(NUM_INICIAL));
}

// ===== Utilidades =====
function hoja_() { return SpreadsheetApp.getActive().getSheetByName(HOJA); }
function fmtNum_(n) { return PREFIJO + String(n).padStart(DIGITOS, '0'); }
function ahora_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function filaDe_(num) {
  const sh = hoja_();
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const nums = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < nums.length; i++) if (String(nums[i][0]) === num) return i + 2;
  return 0;
}
function registro_(fila) {
  const v = hoja_().getRange(fila, 1, 1, COLS.length).getValues()[0];
  const r = {};
  COLS.forEach((c, i) => { r[c] = v[i]; });
  try { r.items = r.items ? JSON.parse(r.items) : []; } catch (e) { r.items = []; }
  return r;
}

// ===== Contador (con candado: dos personas guardando a la vez nunca reciben el mismo número) =====
function reservarNumero_(origen) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const props = PropertiesService.getScriptProperties();
    const n = Number(props.getProperty('ULTIMO') || NUM_INICIAL) + 1;
    props.setProperty('ULTIMO', String(n));
    const num = fmtNum_(n);
    const fila = COLS.map(() => '');
    fila[COLS.indexOf('num')] = num;
    fila[COLS.indexOf('fecha')] = ahora_();
    fila[COLS.indexOf('origen')] = origen || '';
    fila[COLS.indexOf('estado')] = 'reservado';
    hoja_().appendRow(fila);
    return num;
  } finally {
    lock.releaseLock();
  }
}

// ===== Guardar / vender =====
function guardar_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const fila = filaDe_(d.num);
    if (!fila) return { ok: false, error: 'No existe el número ' + d.num + ' (pedí uno con action=nuevo)' };
    const sh = hoja_();
    COLS.forEach((c, i) => {
      if (c === 'num' || c === 'fecha' || c === 'estado' || c === 'fechaVenta') return;
      if (d[c] === undefined) return;
      const val = c === 'items' ? JSON.stringify(d.items) : (c === 'total' ? d[c] : String(d[c]));
      sh.getRange(fila, i + 1).setValue(val);
    });
    // al guardar datos reales, deja de estar "reservado"
    const iEstado = COLS.indexOf('estado') + 1;
    if (sh.getRange(fila, iEstado).getValue() === 'reservado') sh.getRange(fila, iEstado).setValue('presupuesto');
    return { ok: true, presupuesto: registro_(fila) };
  } finally {
    lock.releaseLock();
  }
}

function vendido_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const fila = filaDe_(d.num);
    if (!fila) return { ok: false, error: 'No existe el número ' + d.num };
    const sh = hoja_();
    sh.getRange(fila, COLS.indexOf('estado') + 1).setValue('vendido');
    sh.getRange(fila, COLS.indexOf('fechaVenta') + 1).setValue(ahora_());
    if (d.vendedor) sh.getRange(fila, COLS.indexOf('vendedor') + 1).setValue(d.vendedor);
    return { ok: true, presupuesto: registro_(fila) };
  } finally {
    lock.releaseLock();
  }
}

// ===== Endpoints =====
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    if (p.action === 'nuevo') return out_({ ok: true, num: reservarNumero_(p.origen) });
    if (p.action === 'get') {
      const fila = filaDe_(p.num);
      return out_(fila ? { ok: true, presupuesto: registro_(fila) } : { ok: false, error: 'No existe ' + p.num });
    }
    if (p.action === 'list') {
      const sh = hoja_();
      const n = sh.getLastRow();
      const lista = [];
      for (let f = 2; f <= n; f++) {
        const r = registro_(f);
        if (p.estado && r.estado !== p.estado) continue;
        lista.push(r);
      }
      return out_({ ok: true, presupuestos: lista });
    }
    return out_({ ok: false, error: 'action inválida' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.action === 'guardar') return out_(guardar_(d));
    if (d.action === 'vendido') return out_(vendido_(d));
    return out_({ ok: false, error: 'action inválida' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}
