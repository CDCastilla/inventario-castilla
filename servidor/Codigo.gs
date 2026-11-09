/**
 * Inventario · Clínica Dental Castilla — servidor (Google Apps Script)
 *
 * Va pegado en el editor de Apps Script de la hoja de cálculo del inventario
 * (Extensiones → Apps Script) y se publica como aplicación web.
 * Instrucciones completas: README.md del repositorio.
 *
 * Reglas:
 *  - MOVIMIENTOS solo se añade. Cada movimiento lleva un uuid; si llega dos veces, se ignora.
 *  - La cantidad se guarda con signo (entrada +, salida −) en unidades de control del producto.
 *  - El stock se calcula sumando movimientos por producto y lote.
 *  - Todo se valida aquí, dentro de un bloqueo. Lo que no se acepta va a RECHAZADOS.
 *  - Editar la ficha de un producto nunca mueve stock.
 */

var VERSION_SERVIDOR = '1.0.1';

var HOJAS = {
  PRODUCTOS: ['id', 'nombre', 'categoria', 'proveedor', 'ref_proveedor', 'uds_envase', 'modo', 'lote', 'armario', 'minimo', 'activo', 'notas', 'actualizado'],
  CODIGOS: ['codigo', 'producto_id', 'alta'],
  MOVIMIENTOS: ['uuid', 'fecha', 'tipo', 'producto_id', 'lote', 'caducidad', 'cantidad', 'destino', 'factura', 'origen', 'nota', 'anula_uuid', 'recibido'],
  FACTURAS: ['id', 'proveedor', 'numero', 'fecha', 'archivo', 'estado', 'alta'],
  DOTACION: ['gabinete', 'producto_id', 'cantidad', 'dias_revision'],
  CONFIG: ['tipo', 'nombre', 'modo_defecto', 'lote_defecto'],
  STOCK: ['producto_id', 'nombre', 'categoria', 'armario', 'lote', 'caducidad', 'unidades', 'uds_envase', 'minimo_producto', 'total_producto', 'bajo_minimo'],
  RECHAZADOS: ['recibido', 'uuid', 'motivo', 'datos']
};

var CONFIG_INICIAL = [
  ['categoria', 'Implantes y cirugía', 'unidades', 'SI'],
  ['categoria', 'Medicamentos', 'unidades', 'SI'],
  ['categoria', 'Material clínico', 'unidades', 'SI'],
  ['categoria', 'Fungible', 'abrir', 'NO'],
  ['categoria', 'Limpieza, desinfección e higiene', 'abrir', 'NO'],
  ['categoria', 'Equipos e instrumental', 'unidades', 'NO'],
  ['armario', 'Armario Gabinete 1', '', ''],
  ['armario', 'Armario Gabinete 3', '', ''],
  ['armario', 'Laboratorio', '', ''],
  ['armario', 'Armario cirugía', '', ''],
  ['armario', 'Armario aditamentos', '', ''],
  ['armario', 'Armario Rayos X', '', ''],
  ['armario', 'Cuarto de limpieza', '', ''],
  ['destino', 'Consumido', '', ''],
  ['destino', 'Gabinete 1', '', ''],
  ['destino', 'Gabinete 2', '', ''],
  ['destino', 'Gabinete 3', '', ''],
  ['destino', 'Caja quirúrgica', '', '']
];

var TIPOS = ['ENTRADA', 'SALIDA', 'AJUSTE', 'ANULACION'];
var MAX_OPS = 200;

/* ───────────────────────── Menú y configuración ───────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventario')
    .addItem('Configurar (primera vez)', 'configurar')
    .addItem('Ver la clave de la app', 'verClave')
    .addItem('Cambiar la clave de la app', 'cambiarClave')
    .addItem('Recalcular hoja STOCK', 'recalcularStock')
    .addItem('Hacer copia de seguridad ahora', 'copiaSeguridad')
    .addToUi();
}

/** Crea las pestañas que falten, carga la configuración inicial, genera la clave y programa la copia semanal. */
function configurar() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HOJAS).forEach(function (nombre) {
    var hoja = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
    var cab = HOJAS[nombre];
    if (hoja.getLastRow() === 0) {
      hoja.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
      hoja.setFrozenRows(1);
    }
    // Texto plano: así Sheets no convierte códigos (ceros a la izquierda) ni fechas.
    if (nombre !== 'STOCK') hoja.getRange(1, 1, hoja.getMaxRows(), cab.length).setNumberFormat('@');
  });
  var config = ss.getSheetByName('CONFIG');
  if (config.getLastRow() === 1) {
    config.getRange(2, 1, CONFIG_INICIAL.length, 4).setValues(CONFIG_INICIAL);
  }
  var hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);

  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('CLAVE')) props.setProperty('CLAVE', nuevaClave_());
  instalarCopiaSemanal();
  // Sin ventana: si se ejecuta desde el editor, una ventana se queda esperando en la hoja y agota el tiempo.
  var texto = 'Inventario configurado. Clave de la app: ' + props.getProperty('CLAVE') +
    ' (también en el menú Inventario → Ver la clave de la app).';
  Logger.log(texto);
  try { ss.toast(texto, 'Inventario', 30); } catch (e) { /* sin hoja abierta */ }
}

function verClave() {
  aviso_('Clave de la app: ' + (PropertiesService.getScriptProperties().getProperty('CLAVE') || '(sin configurar)'));
}

function cambiarClave() {
  var clave = nuevaClave_();
  PropertiesService.getScriptProperties().setProperty('CLAVE', clave);
  aviso_('Clave nueva: ' + clave + '\n\nHay que escribirla otra vez en todos los móviles.');
}

function nuevaClave_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function aviso_(texto) {
  try { SpreadsheetApp.getUi().alert(texto); } catch (e) { Logger.log(texto); }
}

/* ───────────────────────── Aplicación web ───────────────────────── */

function doGet() {
  return json_({ ok: true, app: 'Inventario Castilla', version: VERSION_SERVIDOR });
}

/** La app envía POST con text/plain (evita el preflight CORS): {clave, accion, ops?} */
function doPost(e) {
  var pet;
  try { pet = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Petición no válida' }); }
  var clave = PropertiesService.getScriptProperties().getProperty('CLAVE');
  if (!clave || pet.clave !== clave) return json_({ ok: false, error: 'Clave incorrecta' });
  try {
    if (pet.accion === 'estado') return json_(conVersion_(leerEstado_()));
    if (pet.accion === 'enviar') return json_(conVersion_(procesarOps_(pet.ops || [])));
    if (pet.accion === 'exportar') return json_(conVersion_(exportarXlsx_()));
    return json_({ ok: false, error: 'Acción desconocida' });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function conVersion_(r) { r.version = VERSION_SERVIDOR; return r; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────────── Lectura ───────────────────────── */

function filas_(nombre) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!hoja) throw new Error('Falta la hoja ' + nombre + ': ejecuta Inventario → Configurar');
  var cab = HOJAS[nombre];
  var n = hoja.getLastRow() - 1;
  if (n < 1) return [];
  return hoja.getRange(2, 1, n, cab.length).getValues().map(function (v) {
    var o = {};
    cab.forEach(function (c, i) { o[c] = v[i] === null || v[i] === undefined ? '' : String(v[i]); });
    return o;
  });
}

function producto_(o) {
  return {
    id: o.id, nombre: o.nombre, categoria: o.categoria, proveedor: o.proveedor, ref_proveedor: o.ref_proveedor,
    uds_envase: Number(o.uds_envase) || 1, modo: o.modo === 'abrir' ? 'abrir' : 'unidades',
    lote: o.lote === 'SI', armario: o.armario, minimo: Number(o.minimo) || 0,
    activo: o.activo !== 'NO', notas: o.notas
  };
}

/** Suma de movimientos por producto|lote. Devuelve {clave: {producto_id, lote, caducidad, unidades}} */
function calcularStock_(movs) {
  var st = {};
  movs.forEach(function (m) {
    var k = m.producto_id + '|' + (m.lote || '');
    if (!st[k]) st[k] = { producto_id: m.producto_id, lote: m.lote || '', caducidad: '', unidades: 0 };
    st[k].unidades += Number(m.cantidad) || 0;
    if (m.caducidad && !st[k].caducidad) st[k].caducidad = m.caducidad;
  });
  return st;
}

function leerEstado_() {
  var movs = filas_('MOVIMIENTOS');
  var anulados = {};
  movs.forEach(function (m) { if (m.tipo === 'ANULACION' && m.anula_uuid) anulados[m.anula_uuid] = true; });

  var stock = [];
  var st = calcularStock_(movs);
  Object.keys(st).forEach(function (k) { if (Math.abs(st[k].unidades) > 1e-9) stock.push(st[k]); });

  // Reposiciones de gabinete: última salida a cada gabinete por producto.
  var repos = {};
  movs.forEach(function (m) {
    if (m.tipo !== 'SALIDA' || anulados[m.uuid] || String(m.destino).indexOf('Gabinete') !== 0) return;
    var k = m.destino + '|' + m.producto_id;
    if (!repos[k] || m.fecha > repos[k].fecha) repos[k] = { gabinete: m.destino, producto_id: m.producto_id, fecha: m.fecha, unidades: -Number(m.cantidad) };
  });

  var config = { categoria: [], armario: [], destino: [], proveedor: [] };
  filas_('CONFIG').forEach(function (c) {
    if (!config[c.tipo]) config[c.tipo] = [];
    config[c.tipo].push({ nombre: c.nombre, modo_defecto: c.modo_defecto, lote_defecto: c.lote_defecto });
  });

  var ultimos = movs.slice(-300).map(function (m) {
    return { uuid: m.uuid, fecha: m.fecha, tipo: m.tipo, producto_id: m.producto_id, lote: m.lote, caducidad: m.caducidad,
      cantidad: Number(m.cantidad), destino: m.destino, factura: m.factura, origen: m.origen, nota: m.nota,
      anula_uuid: m.anula_uuid, anulado: !!anulados[m.uuid] };
  });

  return {
    ok: true,
    productos: filas_('PRODUCTOS').map(producto_),
    codigos: filas_('CODIGOS').map(function (c) { return { codigo: c.codigo, producto_id: c.producto_id }; }),
    config: config,
    dotacion: filas_('DOTACION').map(function (d) { return { gabinete: d.gabinete, producto_id: d.producto_id, cantidad: Number(d.cantidad) || 0, dias_revision: Number(d.dias_revision) || 0 }; }),
    stock: stock,
    reposiciones: Object.keys(repos).map(function (k) { return repos[k]; }),
    movimientos: ultimos,
    ultimaCopia: PropertiesService.getScriptProperties().getProperty('ULTIMA_COPIA') || '',
    hora: new Date().toISOString()
  };
}

/* ───────────────────────── Escritura ───────────────────────── */

function procesarOps_(ops) {
  if (!Array.isArray(ops)) throw new Error('ops no es una lista');
  if (ops.length > MAX_OPS) throw new Error('Demasiadas operaciones en un envío (máx. ' + MAX_OPS + ')');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ahora = new Date().toISOString();

    var hojaProd = ss.getSheetByName('PRODUCTOS');
    var prodFilas = filas_('PRODUCTOS');
    var prodIdx = {};
    prodFilas.forEach(function (p, i) { prodIdx[p.id] = i; });

    var codigos = {};
    filas_('CODIGOS').forEach(function (c) { codigos[c.codigo] = c.producto_id; });
    var nuevosCodigos = [];

    var movs = filas_('MOVIMIENTOS');
    var uuids = {}, anulados = {}, porUuid = {};
    movs.forEach(function (m) {
      uuids[m.uuid] = true; porUuid[m.uuid] = m;
      if (m.tipo === 'ANULACION') anulados[m.anula_uuid] = true;
    });
    var stock = calcularStock_(movs);
    var destinos = {};
    filas_('CONFIG').forEach(function (c) { if (c.tipo === 'destino') destinos[c.nombre] = true; });

    var nuevosMovs = [], rechazados = [], resultados = [];
    var cambiosProd = false;

    ops.forEach(function (op) {
      try {
        if (op.tipo === 'producto') {
          var p = validarProducto_(op.producto);
          var fila = [p.id, p.nombre, p.categoria, p.proveedor, p.ref_proveedor, p.uds_envase, p.modo, p.lote, p.armario, p.minimo, p.activo, p.notas, ahora];
          if (prodIdx[p.id] !== undefined) {
            prodFilas[prodIdx[p.id]] = filaAObjeto_('PRODUCTOS', fila);
          } else {
            prodIdx[p.id] = prodFilas.length;
            prodFilas.push(filaAObjeto_('PRODUCTOS', fila));
          }
          cambiosProd = true;
          var avisos = [];
          (op.codigos || []).forEach(function (cod) {
            cod = String(cod || '').trim();
            if (!cod) return;
            if (codigos[cod] && codigos[cod] !== p.id) { avisos.push('El código ' + cod + ' ya es de otro producto'); return; }
            if (!codigos[cod]) { codigos[cod] = p.id; nuevosCodigos.push([cod, p.id, ahora]); }
          });
          resultados.push({ op: op.op_id, estado: 'ok', id: p.id, avisos: avisos });
          return;
        }

        if (op.tipo === 'movimiento') {
          var m = op.movimiento || {};
          if (!m.uuid) throw new Error('Movimiento sin uuid');
          if (uuids[m.uuid]) { resultados.push({ op: op.op_id, uuid: m.uuid, estado: 'duplicado' }); return; }
          var r = validarMovimiento_(m, prodFilas, prodIdx, stock, destinos, porUuid, anulados);
          var filaM = [m.uuid, String(m.fecha || ahora), r.tipo, r.producto_id, r.lote, r.caducidad, r.cantidad, r.destino,
            String(m.factura || ''), String(m.origen || ''), String(m.nota || '').slice(0, 300), r.anula_uuid, ahora];
          nuevosMovs.push(filaM);
          var obj = filaAObjeto_('MOVIMIENTOS', filaM);
          uuids[m.uuid] = true; porUuid[m.uuid] = obj;
          if (r.tipo === 'ANULACION') anulados[r.anula_uuid] = true;
          var k = r.producto_id + '|' + r.lote;
          if (!stock[k]) stock[k] = { producto_id: r.producto_id, lote: r.lote, caducidad: r.caducidad, unidades: 0 };
          stock[k].unidades += r.cantidad;
          if (r.caducidad && !stock[k].caducidad) stock[k].caducidad = r.caducidad;
          resultados.push({ op: op.op_id, uuid: m.uuid, estado: 'ok' });
          return;
        }
        throw new Error('Operación desconocida');
      } catch (err) {
        var motivo = String(err && err.message || err);
        var uuid = op.movimiento && op.movimiento.uuid || (op.producto && op.producto.id) || '';
        rechazados.push([ahora, uuid, motivo, JSON.stringify(op).slice(0, 4000)]);
        resultados.push({ op: op.op_id, uuid: uuid, estado: 'rechazado', motivo: motivo });
      }
    });

    if (cambiosProd) {
      var cab = HOJAS.PRODUCTOS;
      var valores = prodFilas.map(function (p) { return cab.map(function (c) { return p[c]; }); });
      hojaProd.getRange(2, 1, valores.length, cab.length).setNumberFormat('@').setValues(valores);
    }
    anadir_('CODIGOS', nuevosCodigos);
    anadir_('MOVIMIENTOS', nuevosMovs);
    anadir_('RECHAZADOS', rechazados);
    if (cambiosProd || nuevosMovs.length) recalcularStock();

    var estado = leerEstado_();
    estado.resultados = resultados;
    return estado;
  } finally {
    lock.releaseLock();
  }
}

function filaAObjeto_(nombre, fila) {
  var o = {};
  HOJAS[nombre].forEach(function (c, i) { o[c] = fila[i] === undefined ? '' : String(fila[i]); });
  return o;
}

function anadir_(nombre, filas) {
  if (!filas.length) return;
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  var r = hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, filas[0].length);
  r.setNumberFormat('@');
  r.setValues(filas.map(function (f) { return f.map(function (v) { return v === null || v === undefined ? '' : String(v); }); }));
}

function texto_(v, max) { return String(v === null || v === undefined ? '' : v).trim().slice(0, max || 200); }

function validarProducto_(p) {
  if (!p) throw new Error('Producto vacío');
  var id = texto_(p.id, 40);
  if (!/^P-[A-Za-z0-9-]{4,36}$/.test(id)) throw new Error('Id de producto no válido');
  var nombre = texto_(p.nombre, 120);
  if (!nombre) throw new Error('El producto necesita nombre');
  var uds = Number(p.uds_envase);
  if (!(uds >= 1 && uds <= 100000 && Math.floor(uds) === uds)) throw new Error('Unidades por envase no válidas');
  var minimo = Number(p.minimo || 0);
  if (!(minimo >= 0 && minimo <= 1000000)) throw new Error('Mínimo no válido');
  return {
    id: id, nombre: nombre, categoria: texto_(p.categoria, 80), proveedor: texto_(p.proveedor, 80),
    ref_proveedor: texto_(p.ref_proveedor, 60), uds_envase: uds,
    modo: p.modo === 'abrir' ? 'abrir' : 'unidades', lote: p.lote ? 'SI' : 'NO',
    armario: texto_(p.armario, 80), minimo: minimo, activo: p.activo === false ? 'NO' : 'SI', notas: texto_(p.notas, 300)
  };
}

function fechaCad_(v) {
  v = texto_(v, 10);
  if (!v) return '';
  if (!/^\d{4}-\d{2}(-\d{2})?$/.test(v)) throw new Error('Caducidad no válida (AAAA-MM o AAAA-MM-DD)');
  return v;
}

function validarMovimiento_(m, prodFilas, prodIdx, stock, destinos, porUuid, anulados) {
  var tipo = texto_(m.tipo, 12);
  if (TIPOS.indexOf(tipo) < 0) throw new Error('Tipo de movimiento no válido');

  if (tipo === 'ANULACION') {
    var orig = porUuid[texto_(m.anula_uuid, 60)];
    if (!orig) throw new Error('El movimiento a anular no existe');
    if (orig.tipo === 'ANULACION') throw new Error('No se puede anular una anulación');
    if (anulados[orig.uuid]) throw new Error('Ese movimiento ya está anulado');
    var cant = -Number(orig.cantidad);
    var kk = orig.producto_id + '|' + (orig.lote || '');
    var disp = stock[kk] ? stock[kk].unidades : 0;
    if (disp + cant < -1e-9) throw new Error('Anularlo dejaría el stock en negativo');
    return { tipo: tipo, producto_id: orig.producto_id, lote: orig.lote || '', caducidad: orig.caducidad || '', cantidad: cant, destino: orig.destino || '', anula_uuid: orig.uuid };
  }

  var pid = texto_(m.producto_id, 40);
  if (prodIdx[pid] === undefined) throw new Error('Producto desconocido');
  var prod = producto_(prodFilas[prodIdx[pid]]);
  var lote = texto_(m.lote, 40);
  var caducidad = fechaCad_(m.caducidad);
  var n = Number(m.cantidad);
  if (!isFinite(n) || n === 0 || Math.abs(n) > 1000000) throw new Error('Cantidad no válida');
  if (Math.round(n * 1000) !== n * 1000) throw new Error('Cantidad con demasiados decimales');
  var k = pid + '|' + lote;
  var disponible = stock[k] ? stock[k].unidades : 0;

  if (tipo === 'ENTRADA') {
    if (n < 0) throw new Error('Una entrada no puede ser negativa');
    if (prod.lote && !lote) throw new Error('Este producto lleva control de lote: falta el lote');
    return { tipo: tipo, producto_id: pid, lote: lote, caducidad: caducidad, cantidad: n, destino: '', anula_uuid: '' };
  }
  if (tipo === 'SALIDA') {
    if (n < 0) throw new Error('Una salida se envía en positivo');
    var destino = texto_(m.destino, 60) || 'Consumido';
    if (!destinos[destino]) throw new Error('Destino no válido: ' + destino);
    if (disponible - n < -1e-9) throw new Error('Stock insuficiente (hay ' + disponible + ')');
    return { tipo: tipo, producto_id: pid, lote: lote, caducidad: caducidad || (stock[k] && stock[k].caducidad) || '', cantidad: -n, destino: destino, anula_uuid: '' };
  }
  // AJUSTE: cantidad con signo (diferencia del recuento)
  if (disponible + n < -1e-9) throw new Error('El ajuste dejaría el stock en negativo');
  return { tipo: tipo, producto_id: pid, lote: lote, caducidad: caducidad || (stock[k] && stock[k].caducidad) || '', cantidad: n, destino: '', anula_uuid: '' };
}

/** Reescribe la hoja STOCK (vista para personas). */
function recalcularStock() {
  var prods = {};
  filas_('PRODUCTOS').forEach(function (p) { prods[p.id] = producto_(p); });
  var st = calcularStock_(filas_('MOVIMIENTOS'));
  var totales = {};
  Object.keys(st).forEach(function (k) { totales[st[k].producto_id] = (totales[st[k].producto_id] || 0) + st[k].unidades; });
  var filas = [];
  Object.keys(st).sort().forEach(function (k) {
    var s = st[k];
    if (Math.abs(s.unidades) < 1e-9) return;
    var p = prods[s.producto_id] || { nombre: '?', categoria: '', armario: '', uds_envase: 1, minimo: 0 };
    var total = totales[s.producto_id];
    filas.push([s.producto_id, p.nombre, p.categoria, p.armario, s.lote, s.caducidad, s.unidades, p.uds_envase, p.minimo, total, p.minimo && total < p.minimo ? 'SÍ' : '']);
  });
  // Productos con mínimo y sin stock también aparecen (para pedir).
  Object.keys(prods).forEach(function (id) {
    var p = prods[id];
    if (p.activo && p.minimo > 0 && !totales[id]) filas.push([id, p.nombre, p.categoria, p.armario, '', '', 0, p.uds_envase, p.minimo, 0, 'SÍ']);
  });
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('STOCK');
  var cab = HOJAS.STOCK;
  if (hoja.getLastRow() > 1) hoja.getRange(2, 1, hoja.getLastRow() - 1, cab.length).clearContent();
  if (filas.length) hoja.getRange(2, 1, filas.length, cab.length).setValues(filas);
}

/* ───────────────────────── Copias y exportación ───────────────────────── */

function xlsxBlob_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } });
  var hoy = Utilities.formatDate(new Date(), 'Europe/Madrid', 'yyyy-MM-dd');
  return resp.getBlob().setName('Inventario Castilla ' + hoy + '.xlsx');
}

/** Copia .xlsx en Drive (carpeta «Inventario Castilla - Copias»); se guardan las 4 últimas. */
function copiaSeguridad() {
  var props = PropertiesService.getScriptProperties();
  var carpeta;
  var id = props.getProperty('CARPETA_COPIAS');
  try { carpeta = id && DriveApp.getFolderById(id); } catch (e) { carpeta = null; }
  if (!carpeta) {
    carpeta = DriveApp.createFolder('Inventario Castilla - Copias');
    props.setProperty('CARPETA_COPIAS', carpeta.getId());
  }
  carpeta.createFile(xlsxBlob_());
  var ficheros = [];
  var it = carpeta.getFiles();
  while (it.hasNext()) ficheros.push(it.next());
  ficheros.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  ficheros.slice(4).forEach(function (f) { f.setTrashed(true); });
  props.setProperty('ULTIMA_COPIA', new Date().toISOString());
}

/** Programa la copia cada lunes a las 7:00 (hora de Madrid). */
function instalarCopiaSemanal() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'copiaSeguridad') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('copiaSeguridad').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).inTimezone('Europe/Madrid').create();
}

function exportarXlsx_() {
  var blob = xlsxBlob_();
  return { ok: true, nombre: blob.getName(), base64: Utilities.base64Encode(blob.getBytes()) };
}
