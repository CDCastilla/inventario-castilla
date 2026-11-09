// Simulador mínimo de Apps Script para probar servidor/Codigo.gs sin tocar la hoja real.
import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { createHash } from 'node:crypto';

function crearEntorno() {
  const hojas = new Map();
  class Hoja {
    constructor(nombre) { this.nombre = nombre; this.d = []; }
    getLastRow() { return this.d.length; }
    getDataRange() { const h = this; return { getValues: () => h.d.map((f) => [...f]) }; }
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
  const disparadores = [];
  const ficheros = new Map();
  const carpeta = { getId: () => 'CARP', createFile: (b) => { const id = 'F' + (ficheros.size + 1); ficheros.set(id, b.bytes); return { getId: () => id }; } };
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, getUi: () => { throw new Error('sin UI'); } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => props[k] ?? null, setProperty: (k, v) => { props[k] = v; } }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    Utilities: {
      getUuid: () => 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee', sleep() {},
      base64Decode: (b) => Buffer.from(b, 'base64'), base64Encode: (b) => Buffer.from(b).toString('base64'),
      newBlob: (bytes, tipo, nombre) => ({ bytes, tipo, nombre }),
      computeDigest: (_, txt) => [...createHash('md5').update(txt).digest()].map((b) => (b > 127 ? b - 256 : b)),
      DigestAlgorithm: { MD5: 'md5' }, formatDate: (d) => d.toISOString().slice(0, 10)
    },
    Logger: { log() {} },
    ScriptApp: {
      getProjectTriggers: () => disparadores.map((d) => ({ getHandlerFunction: () => d })),
      newTrigger: (f) => { const t = { timeBased: () => t, onWeekDay: () => t, atHour: () => t, inTimezone: () => t, after: () => t, everyMinutes: () => t, create: () => { disparadores.push(f); return t; } }; return t; },
      deleteTrigger: (t) => { const i = disparadores.indexOf(t.getHandlerFunction()); if (i >= 0) disparadores.splice(i, 1); },
      WeekDay: { MONDAY: 1 }
    },
    DriveApp: {
      getFolderById: () => carpeta, createFolder: () => carpeta,
      getFileById: (id) => ({ getBlob: () => ({ getBytes: () => ficheros.get(id), getContentType: () => 'image/jpeg' }) })
    },
    UrlFetchApp: { fetch: (url, opt) => ctx.__fetch(url, opt) },
    ContentService: { createTextOutput: (s) => ({ setMimeType: () => s }), MimeType: { JSON: 'json' } },
    console
  };
  vm.createContext(ctx);
  vm.runInContext(readFileSync(new URL('../servidor/Codigo.gs', import.meta.url), 'utf8'), ctx);
  ctx.configurar();
  const llamar = (pet) => JSON.parse(ctx.doPost({ postData: { contents: JSON.stringify({ clave: props.CLAVE, ...pet }) } }));
  return { ctx, hojas, props, llamar, disparadores, ficheros };
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
    const texto = JSON.stringify({ nombre: 'Guantes látex Santex talla S', marca: 'Santex', ref_proveedor: 'GD05NB', uds_envase: 100, categoria: 'Fungible', lote: '', caducidad: '2031-04', fuente: '', dudas: '' });
    return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ candidates: [{ content: { parts: [{ text: texto }] } }] }) };
  };
  r = llamar({ accion: 'identificar', imagenes: ['x'.repeat(200)], lectura: { gtin: '08437014559484' } });
  assert.ok(r.ok, r.error);
  assert.equal(peticion.clave, 'k');
  assert.match(peticion.url, /gemini-flash-latest:generateContent$/);
  assert.match(peticion.cuerpo.contents[0].parts[0].text, /08437014559484/);
  assert.equal(peticion.cuerpo.contents[0].parts[1].inline_data.mime_type, 'image/jpeg');
  assert.deepEqual(r.propuesta, { nombre: 'Guantes látex Santex talla S', marca: 'Santex', ref_proveedor: 'GD05NB', uds_envase: 100, categoria: 'Fungible', lote: '', caducidad: '2031-04', fuente: '', dudas: '' });
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

test('alta sin esperar: foto a Drive, pendiente, identificación en segundo plano y revisión', () => {
  const { llamar, props, ctx, disparadores, ficheros } = crearEntorno();
  props.GEMINI_KEY = 'k';
  const foto = Buffer.from('x'.repeat(300)).toString('base64');
  const prov = { id: 'P-EUGENOL001', nombre: 'Pendiente de identificar · …240128', uds_envase: 1, modo: 'unidades', lote: true, estado_ia: 'pendiente' };
  let r = llamar({ accion: 'enviar', ops: [
    { tipo: 'producto', op_id: 'a', producto: prov, codigos: ['08436042240128'], fotos: [foto], identificar: true },
    { tipo: 'movimiento', op_id: 'b', movimiento: { uuid: 'ue1', tipo: 'ENTRADA', producto_id: prov.id, lote: '020226', caducidad: '2028-02-28', cantidad: 2 } }
  ] });
  assert.ok(r.ok, r.error);
  assert.equal(r.productos[0].estado_ia, 'pendiente');
  assert.equal(ficheros.size, 1);
  assert.ok(disparadores.includes('identificarAhora'));
  assert.ok(disparadores.includes('identificarPendientes'));

  let cuerpo;
  ctx.__fetch = (url, opt) => {
    cuerpo = JSON.parse(opt.payload);
    const texto = 'Según la búsqueda: ```json\n' + JSON.stringify({ nombre: 'Eugenol 20 ml', marca: 'Proclinic', ref_proveedor: 'EUG', uds_envase: 1, categoria: 'Material clínico', fuente: 'foto e internet', dudas: '' }) + '\n```';
    return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ candidates: [{ content: { parts: [{ text: texto }] } }] }) };
  };
  ctx.identificarAhora();
  assert.ok(!disparadores.includes('identificarAhora'));
  assert.ok(cuerpo.tools && cuerpo.tools[0].google_search, 'usa la búsqueda en internet');
  assert.match(cuerpo.contents[0].parts[0].text, /08436042240128/);
  r = llamar({ accion: 'estado' });
  const p = r.productos[0];
  assert.equal(p.nombre, 'Eugenol 20 ml');
  assert.equal(p.estado_ia, 'revisar');
  assert.equal(p.categoria, 'Material clínico');
  assert.equal(p.ref_proveedor, 'EUG');
  assert.equal(r.stock[0].unidades, 2);

  // El usuario confirma: deja de estar por revisar y no se pierde la foto.
  r = llamar({ accion: 'enviar', ops: [{ tipo: 'producto', op_id: 'c', producto: { ...p, estado_ia: '' }, codigos: [] }] });
  assert.equal(r.productos[0].estado_ia, '');
  assert.equal(r.productos[0].con_foto, true);
});

test('identificación en segundo plano: tras 4 fallos queda en error y no se reintenta', () => {
  const { llamar, props, ctx } = crearEntorno();
  props.GEMINI_KEY = 'k';
  const foto = Buffer.from('y'.repeat(300)).toString('base64');
  llamar({ accion: 'enviar', ops: [{ tipo: 'producto', op_id: 'a', producto: { id: 'P-XXXXXXXX', nombre: 'Pendiente', uds_envase: 1, modo: 'unidades' }, codigos: ['08400000000000'], fotos: [foto], identificar: true }] });
  ctx.__fetch = () => ({ getResponseCode: () => 503, getContentText: () => 'busy' });
  for (let i = 0; i < 5; i++) ctx.identificarPendientes();
  const p = llamar({ accion: 'estado' }).productos[0];
  assert.equal(p.estado_ia, 'error');
  assert.match(p.notas, /IA:/);
});

test('foto del producto: añadir otra foto la pone la primera y se puede leer', () => {
  const { llamar } = crearEntorno();
  const f1 = Buffer.from('a'.repeat(300)).toString('base64');
  const f2 = Buffer.from('b'.repeat(300)).toString('base64');
  const prod = { id: 'P-FOTO0001', nombre: 'Guantes', uds_envase: 1, modo: 'abrir' };
  llamar({ accion: 'enviar', ops: [{ tipo: 'producto', op_id: '1', producto: prod, codigos: [], fotos: [f1] }] });
  let r = llamar({ accion: 'foto', producto_id: prod.id });
  assert.equal(r.base64, f1);
  r = llamar({ accion: 'enviar', ops: [{ tipo: 'producto', op_id: '2', producto: prod, codigos: [], fotos: [f2], identificar: false }] });
  assert.equal(r.productos[0].estado_ia, '');
  assert.equal(llamar({ accion: 'foto', producto_id: prod.id }).base64, f2);
  assert.equal(llamar({ accion: 'foto', producto_id: 'P-NOEXISTE' }).base64, '');
});

test('combinar productos: códigos, stock, catálogo y movimientos antiguos van al que se conserva', () => {
  const { llamar, hojas } = crearEntorno();
  const a = { id: 'P-AAAA0001', nombre: 'Guantes látex S Santex', uds_envase: 1, modo: 'abrir' };
  const b = { id: 'P-BBBB0001', nombre: 'Guantes latex talla S', uds_envase: 1, modo: 'abrir', minimo: 10 };
  llamar({ accion: 'enviar', ops: [
    { tipo: 'producto', op_id: '1', producto: a, codigos: ['111'] },
    { tipo: 'producto', op_id: '2', producto: b, codigos: ['222'] },
    { tipo: 'movimiento', op_id: '3', movimiento: { uuid: 'm1', tipo: 'ENTRADA', producto_id: a.id, cantidad: 3 } },
    { tipo: 'movimiento', op_id: '4', movimiento: { uuid: 'm2', tipo: 'ENTRADA', producto_id: b.id, cantidad: 5 } }
  ] });
  let r = llamar({ accion: 'enviar', ops: [{ tipo: 'fusionar', op_id: 'f', de: b.id, en: a.id }] });
  assert.equal(r.resultados[0].estado, 'ok');
  assert.equal(r.stock.length, 1);
  assert.equal(r.stock[0].producto_id, a.id);
  assert.equal(r.stock[0].unidades, 8);
  assert.ok(r.codigos.every((c) => c.producto_id === a.id));
  const pb = r.productos.find((p) => p.id === b.id);
  assert.equal(pb.activo, false);
  assert.equal(pb.fusionado_en, a.id);
  assert.equal(r.productos.find((p) => p.id === a.id).minimo, 10);
  // Una salida enviada desde un móvil antiguo con el id viejo se aplica al que se conserva.
  r = llamar({ accion: 'enviar', ops: [{ tipo: 'movimiento', op_id: '5', movimiento: { uuid: 'm3', tipo: 'SALIDA', producto_id: b.id, cantidad: 6, destino: 'Gabinete 1' } }] });
  assert.equal(r.resultados[0].estado, 'ok');
  assert.equal(r.stock[0].unidades, 2);
  // Anular una entrada antigua del producto combinado.
  r = llamar({ accion: 'enviar', ops: [{ tipo: 'movimiento', op_id: '6', movimiento: { uuid: 'm4', tipo: 'ANULACION', anula_uuid: 'm1' } }] });
  assert.equal(r.resultados[0].estado, 'rechazado', 'dejaría stock negativo');
  r = llamar({ accion: 'enviar', ops: [{ tipo: 'no_duplicado', op_id: '7', a: 'P-X', b: 'P-Y' }] });
  assert.deepEqual(r.noDuplicados, [['P-X', 'P-Y']]);
  assert.ok(hojas.has('NO_DUPLICADOS'));
});

test('cargar catálogo: crea productos por artículo, enlaza referencias y no duplica al repetir', () => {
  const { ctx, hojas, llamar } = crearEntorno();
  const imp = hojas.get('PRODUCTOS').constructor ? null : null;
  const ss = ctx.SpreadsheetApp.getActiveSpreadsheet();
  const h = ss.insertSheet('CATALOGO_IMPORTAR');
  h.d = [
    ['articulo', 'nombre', 'categoria', 'marca', 'proveedor', 'ref', 'descripcion', 'precio', 'fecha', 'factura'],
    ['A1', 'Dentasept Tri Enzymatic 5 L', 'Limpieza, desinfección e higiene', 'Anios', 'Proclinic', '49889', 'DENTASEPT TRI ENZYMATIC 5L', '45.10', '2026-09-10', '0040570298'],
    ['A1', 'Dentasept Tri Enzymatic 5 L', 'Limpieza, desinfección e higiene', 'Anios', 'Henry Schein', '900-1234', 'DENTASEPT TRI ENZ 5L ANIOS', '47.00', '2026-03-02', 'M948822'],
    ['A2', 'Guantes látex S', 'Fungible', '', 'Dentared', 'DTX-GLSP100S', 'Guantes de Látex sin Polvo 100 uds S', '2.80', '2026-01-10', 'DR2026_05177']
  ];
  ctx.cargarCatalogo();
  let r = llamar({ accion: 'estado' });
  assert.equal(r.productos.length, 2);
  assert.equal(r.catalogo.length, 3);
  const dent = r.productos.find((p) => p.nombre.startsWith('Dentasept'));
  assert.equal(dent.proveedor, 'Proclinic');
  assert.equal(dent.modo, 'abrir');
  assert.equal(r.catalogo.filter((c) => c.producto_id === dent.id).length, 2);
  h.d.push(['A1', 'Dentasept Tri Enzymatic 5 L', '', 'Anios', 'Proclinic', '49889', 'DENTASEPT TRI ENZYMATIC 5L', '46.20', '2026-10-20', '0040638705']);
  ctx.cargarCatalogo();
  r = llamar({ accion: 'estado' });
  assert.equal(r.productos.length, 2);
  assert.equal(r.catalogo.length, 3);
  assert.equal(r.catalogo.find((c) => c.ref === '49889').precio, 46.2);
});

test('imágenes de internet: busca, comprueba que es imagen y la guarda como foto', () => {
  const { ctx, props, llamar, ficheros } = crearEntorno();
  props.GEMINI_KEY = 'k';
  llamar({ accion: 'enviar', ops: [
    { tipo: 'producto', op_id: '1', producto: { id: 'P-IMG00001', nombre: 'MTA Angelus 1 g', uds_envase: 1, modo: 'unidades' }, codigos: [] },
    { tipo: 'producto', op_id: '2', producto: { id: 'P-IMG00002', nombre: 'Producto raro', uds_envase: 1, modo: 'unidades' }, codigos: [] }
  ] });
  const jpg = Buffer.concat([Buffer.from([0xff, 0xd8, 0xff]), Buffer.alloc(5000, 1)]);
  ctx.__fetch = (url, opt) => {
    if (url.includes('generativelanguage')) {
      const raro = JSON.parse(opt.payload).contents[0].parts[0].text.includes('Producto raro');
      const t = raro ? '{"imagen":"https://x.com/falsa.jpg","pagina":""}' : '{"imagen":"https://angelus.com/mta.jpg","pagina":""}';
      return { getResponseCode: () => 200, getContentText: () => JSON.stringify({ candidates: [{ content: { parts: [{ text: t }] } }] }) };
    }
    if (url === 'https://angelus.com/mta.jpg') return { getResponseCode: () => 200, getHeaders: () => ({ 'Content-Type': 'image/jpeg' }), getBlob: () => ({ getBytes: () => [...jpg].map((b) => (b > 127 ? b - 256 : b)) }) };
    return { getResponseCode: () => 404, getHeaders: () => ({}), getBlob: () => ({ getBytes: () => [] }) };
  };
  ctx.identificarPendientes();
  let r = llamar({ accion: 'estado' });
  assert.equal(r.productos.find((p) => p.id === 'P-IMG00001').con_foto, true);
  assert.equal(r.productos.find((p) => p.id === 'P-IMG00002').con_foto, false);
  assert.ok(ficheros.size >= 1);
  ctx.identificarPendientes(); // segundo intento fallido → ya no se busca más
  ctx.identificarPendientes();
  const fila = ctx.filas_('PRODUCTOS').find((p) => p.id === 'P-IMG00002');
  assert.equal(fila.img_estado, 'no');
});
