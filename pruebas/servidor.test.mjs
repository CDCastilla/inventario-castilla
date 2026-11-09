// Simulador mínimo de Apps Script para probar servidor/Codigo.gs sin tocar la hoja real.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

function crearEntorno() {
  const hojas = new Map();
  class Hoja {
    constructor(nombre) { this.nombre = nombre; this.d = []; }
    getLastRow() { return this.d.length; }
    getMaxRows() { return 1000; }
    setFrozenRows() {}
    getRange(fila, col, nf = 1, nc = 1) {
      const h = this;
      const r = {
        getValues: () => Array.from({ length: nf }, (_, i) => Array.from({ length: nc }, (_, j) => (h.d[fila - 1 + i] || [])[col - 1 + j] ?? '')),
        setValues: (v) => { v.forEach((f, i) => { h.d[fila - 1 + i] = h.d[fila - 1 + i] || []; f.forEach((x, j) => { h.d[fila - 1 + i][col - 1 + j] = x; }); }); return r; },
        setNumberFormat: () => r, setFontWeight: () => r,
        clearContent: () => { for (let i = 0; i < nf; i++) if (h.d[fila - 1 + i]) h.d[fila - 1 + i] = []; h.d = h.d.filter((f, k) => k === 0 || f.length); return r; }
      };
      return r;
    }
  }
  const ss = {
    getSheetByName: (n) => hojas.get(n) || null,
    insertSheet: (n) => { const h = new Hoja(n); hojas.set(n, h); return h; },
    getSheets: () => [...hojas.values()],
    deleteSheet: (h) => hojas.delete(h.nombre),
    getId: () => 'ID'
  };
  const props = {};
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, getUi: () => { throw new Error('sin UI'); } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: { getUuid: () => 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee' },
    Logger: { log() {} },
    ScriptApp: { getProjectTriggers: () => [], newTrigger: () => { const t = { timeBased: () => t, onWeekDay: () => t, atHour: () => t, inTimezone: () => t, create: () => t }; return t; }, WeekDay: { MONDAY: 1 } },
    UrlFetchApp: { fetch: (url, opt) => ctx.__fetch(url, opt) },
    ContentService: { createTextOutput: (s) => ({ setMimeType: () => s }), MimeType: { JSON: 'json' } },
    console
  };
  vm.createContext(ctx);
  vm.runInContext(readFileSync(new URL('../servidor/Codigo.gs', import.meta.url), 'utf8'), ctx);
  ctx.configurar();
  const llamar = (pet) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ clave: props.CLAVE, ...pet }) } }));
  return { ctx, hojas, props, llamar };
}

const guantes = { id: 'P-GUANTES01', nombre: 'Guantes S', categoria: 'Fungible', uds_envase: 1, modo: 'abrir', lote: false, minimo: 10 };
const implante = { id: 'P-IMPL0001', nombre: 'Implante 4.0×10', categoria: 'Implantes y cirugía', uds_envase: 1, modo: 'unidades', lote: true, minimo: 2 };
let n = 0;
const mov = (m) => ({ tipo: 'movimiento', op_id: 'o' + (++n), movimiento: { uuid: 'u' + n, fecha: '2026-11-09T10:00:00Z', ...m } });

test('configurar crea las hojas, la clave y la configuración inicial', () => {
  const { hojas, props } = crearEntorno();
  for (const n of ['PRODUCTOS', 'CODIGOS', 'MOVIMIENTOS', 'FACTURAS', 'DOTACION', 'CONFIG', 'STOCK', 'RECHAZADOS']) assert.ok(hojas.has(n), n);
  assert.equal(props.CLAVE.length, 12);
  assert.ok(hojas.get('CONFIG').getLastRow() > 10);
});

test('clave incorrecta', () => {
  const { ctx } = crearEntorno();
  const r = JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ clave: 'mala', accion: 'estado' }) } }));
  assert.equal(r.ok, false);
});

test('flujo completo: alta, entrada, salida a gabinete, duplicado, rechazos, anulación', () => {
  const { llamar, hojas } = crearEntorno();
  let r = llamar({ accion: 'enviar', ops: [
    { tipo: 'producto', op_id: 'p1', producto: guantes, codigos: ['08437014559484'] },
    { tipo: 'producto', op_id: 'p2', producto: implante, codigos: ['08435451622884'] }
  ] });
  assert.ok(r.ok);
  assert.equal(r.productos.length, 2);
  assert.equal(r.codigos.length, 2);

  const e1 = mov({ tipo: 'ENTRADA', producto_id: guantes.id, cantidad: 12 });
  const e2 = mov({ tipo: 'ENTRADA', producto_id: implante.id, lote: '75958-2447', caducidad: '2031-03-16', cantidad: 3 });
  const eSinLote = mov({ tipo: 'ENTRADA', producto_id: implante.id, cantidad: 1 });
  const s1 = mov({ tipo: 'SALIDA', producto_id: guantes.id, cantidad: 2, destino: 'Gabinete 3' });
  const sMucho = mov({ tipo: 'SALIDA', producto_id: implante.id, lote: '75958-2447', cantidad: 5, destino: 'Consumido' });
  const sMalDestino = mov({ tipo: 'SALIDA', producto_id: guantes.id, cantidad: 1, destino: 'Pasillo' });
  r = llamar({ accion: 'enviar', ops: [e1, e2, eSinLote, s1, sMucho, sMalDestino] });
  const est = Object.fromEntries(r.resultados.map((x) => [x.op, x.estado]));
  assert.equal(est[e1.op_id], 'ok');
  assert.equal(est[e2.op_id], 'ok');
  assert.equal(est[eSinLote.op_id], 'rechazado');
  assert.equal(est[s1.op_id], 'ok');
  assert.equal(est[sMucho.op_id], 'rechazado');
  assert.equal(est[sMalDestino.op_id], 'rechazado');
  assert.equal(hojas.get('RECHAZADOS').getLastRow() - 1, 3);

  const stock = Object.fromEntries(r.stock.map((s) => [s.producto_id + '|' + s.lote, s.unidades]));
  assert.equal(stock[guantes.id + '|'], 10);
  assert.equal(stock[implante.id + '|75958-2447'], 3);
  assert.equal(r.reposiciones.length, 1);
  assert.equal(r.reposiciones[0].gabinete, 'Gabinete 3');

  // Reenvío del mismo movimiento (corte de red): se ignora.
  r = llamar({ accion: 'enviar', ops: [s1] });
  assert.equal(r.resultados[0].estado, 'duplicado');
  assert.equal(r.stock.find((s) => s.producto_id === guantes.id).unidades, 10);

  // Anular la salida devuelve el stock; anular dos veces se rechaza.
  const a1 = mov({ tipo: 'ANULACION', anula_uuid: s1.movimiento.uuid, cantidad: 0 });
  const a2 = mov({ tipo: 'ANULACION', anula_uuid: s1.movimiento.uuid, cantidad: 0 });
  r = llamar({ accion: 'enviar', ops: [a1, a2] });
  assert.deepEqual(r.resultados.map((x) => x.estado), ['ok', 'rechazado']);
  assert.equal(r.stock.find((s) => s.producto_id === guantes.id).unidades, 12);
  assert.equal(r.reposiciones.length, 0);

  // Ajuste por recuento.
  r = llamar({ accion: 'enviar', ops: [mov({ tipo: 'AJUSTE', producto_id: guantes.id, cantidad: -1, nota: 'Recuento' })] });
  assert.equal(r.stock.find((s) => s.producto_id === guantes.id).unidades, 11);

  // Editar la ficha no mueve stock; un código de otro producto da aviso.
  r = llamar({ accion: 'enviar', ops: [{ tipo: 'producto', op_id: 'p3', producto: { ...guantes, minimo: 15 }, codigos: ['08435451622884'] }] });
  assert.equal(r.resultados[0].avisos.length, 1);
  assert.equal(r.productos.find((p) => p.id === guantes.id).minimo, 15);
  assert.equal(r.stock.find((s) => s.producto_id === guantes.id).unidades, 11);
  assert.equal(hojas.get('MOVIMIENTOS').getLastRow() - 1, 5); // e1, e2, s1, anulación y ajuste

  // Hoja STOCK para personas.
  const filasStock = hojas.get('STOCK').d.slice(1);
  assert.ok(filasStock.some((f) => f[0] === guantes.id && f[10] === 'SÍ'));
});

test('identificar: pide a Gemini y devuelve una propuesta saneada', () => {
  const { llamar, props, ctx } = crearEntorno();
  let r = llamar({ accion: 'identificar', imagenes: ['x'.repeat(200)] });
  assert.equal(r.ok, false);
  assert.match(r.error, /GEMINI_KEY/);
  props.GEMINI_KEY = 'k';
  let peticion;
  ctx.__fetch = (url, opt) => {
    peticion = { url, cuerpo: JSON.parse(opt.payload), clave: opt.headers['x-goog-api-key'] };
    const texto = JSON.stringify({ nombre: 'Guantes látex Santex talla S', marca: 'Santex', ref_proveedor: 'GD05NB', uds_envase: 100, categoria: 'Fungible', lote: '', caducidad: '2031-04', dudas: '' });
    return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ candidates: [{ content: { parts: [{ text: texto }] } }] }) };
  };
  r = llamar({ accion: 'identificar', imagenes: ['x'.repeat(200)], lectura: { gtin: '08437014559484' } });
  assert.ok(r.ok, r.error);
  assert.equal(peticion.clave, 'k');
  assert.match(peticion.url, /gemini-flash-latest:generateContent$/);
  assert.match(peticion.cuerpo.contents[0].parts[0].text, /08437014559484/);
  assert.equal(peticion.cuerpo.contents[0].parts[1].inline_data.mime_type, 'image/jpeg');
  assert.deepEqual(r.propuesta, { nombre: 'Guantes látex Santex talla S', marca: 'Santex', ref_proveedor: 'GD05NB', uds_envase: 100, categoria: 'Fungible', lote: '', caducidad: '2031-04', dudas: '' });
  ctx.__fetch = () => ({ getResponseCode: () => 200, getContentText: () => JSON.stringify({ candidates: [{ content: { parts: [{ text: '{"nombre":"X","categoria":"Inventada","uds_envase":"abc"}' }] } }] }) });
  r = llamar({ accion: 'identificar', imagenes: ['x'.repeat(200)] });
  assert.equal(r.propuesta.categoria, '');
  assert.equal(r.propuesta.uds_envase, 1);
});

test('identificar: si Gemini está saturado reintenta y cambia de modelo', () => {
  const { llamar, props, ctx } = crearEntorno();
  props.GEMINI_KEY = 'k';
  ctx.Utilities.sleep = () => {};
  const urls = [];
  ctx.__fetch = (url) => {
    urls.push(url);
    if (url.includes('gemini-flash-latest')) return { getResponseCode: () => 503, getContentText: () => '{"error":"high demand"}' };
    return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ candidates: [{ content: { parts: [{ text: '{"nombre":"Eugenol 20 ml","categoria":"Material clínico","uds_envase":1}' }] } }] }) };
  };
  const r = llamar({ accion: 'identificar', imagenes: ['x'.repeat(200)] });
  assert.ok(r.ok, r.error);
  assert.equal(r.propuesta.nombre, 'Eugenol 20 ml');
  assert.equal(urls.filter((u) => u.includes('gemini-flash-latest')).length, 2);
  assert.match(urls[2], /gemini-flash-lite-latest/);
});
