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

var VERSION_SERVIDOR = '1.6.0';

var HOJAS = {
  PRODUCTOS: ['id', 'nombre', 'categoria', 'proveedor', 'ref_proveedor', 'uds_envase', 'modo', 'lote', 'armario', 'minimo', 'activo', 'notas', 'actualizado', 'foto', 'estado_ia', 'intentos_ia', 'img_estado', 'fusionado_en'],
  CODIGOS: ['codigo', 'producto_id', 'alta'],
  MOVIMIENTOS: ['uuid', 'fecha', 'tipo', 'producto_id', 'lote', 'caducidad', 'cantidad', 'destino', 'factura', 'origen', 'nota', 'anula_uuid', 'recibido'],
  FACTURAS: ['id', 'proveedor', 'numero', 'fecha', 'archivo', 'estado', 'alta'],
  DOTACION: ['gabinete', 'producto_id', 'cantidad', 'dias_revision'],
  CONFIG: ['tipo', 'nombre', 'modo_defecto', 'lote_defecto'],
  STOCK: ['producto_id', 'nombre', 'categoria', 'armario', 'lote', 'caducidad', 'unidades', 'uds_envase', 'minimo_producto', 'total_producto', 'bajo_minimo'],
  RECHAZADOS: ['recibido', 'uuid', 'motivo', 'datos'],
  CATALOGO: ['proveedor', 'ref', 'descripcion', 'precio', 'fecha', 'factura', 'producto_id', 'alta'],
  NO_DUPLICADOS: ['producto_a', 'producto_b', 'fecha'],
  PEDIDOS: ['pedido', 'linea', 'proveedor', 'numero', 'fecha', 'producto_id', 'ref', 'descripcion', 'cantidad', 'precio', 'iva', 'estado', 'factura', 'alta']
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
    .addItem('Cargar catálogo (pestaña CATALOGO_IMPORTAR)', 'cargarCatalogo')
    .addItem('Buscar imágenes ahora (5 min)', 'buscarImagenesAhora')
    .addItem('Volver a buscar imágenes no encontradas', 'reintentarImagenes')
    .addItem('Probar conexión con Gemini', 'probarGemini')
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
    } else asegurarCabeceras_(nombre);
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
  try { programarIdentificacion_(); } catch (e) { Logger.log(e); }
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

/** Sin ventanas que esperen: se escribe en el registro de ejecución y en un aviso de la hoja. */
function aviso_(texto) {
  Logger.log(texto);
  try { SpreadsheetApp.getActiveSpreadsheet().toast(texto, 'Inventario', 60); } catch (e) { /* sin hoja abierta */ }
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
    if (pet.accion === 'estado') {
      // Autocuración: si falta el disparador periódico (p. ej. no se ejecutó configurar), se crea.
      try {
        if (PropertiesService.getScriptProperties().getProperty('GEMINI_KEY') &&
          !ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'identificarPeriodico'; })) programarIdentificacion_(true);
      } catch (e) { Logger.log('Disparador: ' + e); }
      return json_(conVersion_(leerEstado_()));
    }
    if (pet.accion === 'enviar') return json_(conVersion_(procesarOps_(pet.ops || [])));
    if (pet.accion === 'exportar') return json_(conVersion_(exportarXlsx_()));
    if (pet.accion === 'identificar') return json_(conVersion_(identificar_(pet)));
    if (pet.accion === 'foto') return json_(conVersion_(leerFoto_(pet.producto_id)));
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

/** Como filas_, pero devuelve [] si la pestaña aún no existe (versiones antiguas). */
function filasOpc_(nombre) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre) ? filas_(nombre) : [];
}

function producto_(o) {
  return {
    id: o.id, nombre: o.nombre, categoria: o.categoria, proveedor: o.proveedor, ref_proveedor: o.ref_proveedor,
    uds_envase: Number(o.uds_envase) || 1, modo: o.modo === 'abrir' ? 'abrir' : 'unidades',
    lote: o.lote === 'SI', armario: o.armario, minimo: Number(o.minimo) || 0,
    activo: o.activo !== 'NO' && !o.fusionado_en, notas: o.notas, estado_ia: o.estado_ia || '', con_foto: !!o.foto,
    fusionado_en: o.fusionado_en || '', foto_v: o.foto ? String(o.foto).split(',')[0].slice(-10) : ''
  };
}

/** Productos combinados: id antiguo → id que se conserva (sigue cadenas). */
function alias_(prodFilas) {
  var directo = {}, a = {};
  prodFilas.forEach(function (p) { if (p.fusionado_en) directo[p.id] = p.fusionado_en; });
  Object.keys(directo).forEach(function (id) {
    var x = id, n = 0;
    while (directo[x] && n++ < 20) x = directo[x];
    a[id] = x;
  });
  return a;
}
function real_(alias, pid) { return alias[pid] || pid; }

/** Suma de movimientos por producto|lote. Devuelve {clave: {producto_id, lote, caducidad, unidades}} */
function calcularStock_(movs, alias) {
  var st = {};
  alias = alias || {};
  movs.forEach(function (m) {
    var pid = real_(alias, m.producto_id);
    var k = pid + '|' + (m.lote || '');
    if (!st[k]) st[k] = { producto_id: pid, lote: m.lote || '', caducidad: '', unidades: 0 };
    st[k].unidades += Number(m.cantidad) || 0;
    if (m.caducidad && !st[k].caducidad) st[k].caducidad = m.caducidad;
  });
  return st;
}

/** Pedidos con lo recibido de cada línea (suma de las entradas con origen «pedido:ID:línea»). Sin los anulados. */
function pedidos_(movs, anulados, alias) {
  var recibido = {};
  movs.forEach(function (m) {
    if (m.tipo !== 'ENTRADA' || anulados[m.uuid] || String(m.origen).indexOf('pedido:') !== 0) return;
    var k = String(m.origen).slice(7);
    recibido[k] = (recibido[k] || 0) + (Number(m.cantidad) || 0);
  });
  var porId = {}, orden = [];
  filasOpc_('PEDIDOS').forEach(function (x) {
    if (x.estado === 'anulado') return;
    if (!porId[x.pedido]) { porId[x.pedido] = { id: x.pedido, proveedor: x.proveedor, numero: x.numero, fecha: x.fecha, estado: x.estado, factura: x.factura, alta: x.alta, lineas: [] }; orden.push(x.pedido); }
    porId[x.pedido].lineas.push({ linea: Number(x.linea), producto_id: x.producto_id ? real_(alias, x.producto_id) : '', ref: x.ref, descripcion: x.descripcion,
      cantidad: Number(x.cantidad) || 0, precio: Number(x.precio) || 0, iva: x.iva, recibido: recibido[x.pedido + ':' + x.linea] || 0 });
  });
  // Abiertos todos; cerrados, los de los últimos 90 días.
  var limite = new Date(Date.now() - 90 * 864e5).toISOString().slice(0, 10);
  return orden.map(function (id) { return porId[id]; }).filter(function (p) { return !p.estado || p.fecha >= limite; });
}

function leerEstado_() {
  var movs = filas_('MOVIMIENTOS');
  var prodFilas = filas_('PRODUCTOS');
  var alias = alias_(prodFilas);
  var anulados = {};
  movs.forEach(function (m) { if (m.tipo === 'ANULACION' && m.anula_uuid) anulados[m.anula_uuid] = true; });

  var stock = [];
  var st = calcularStock_(movs, alias);
  Object.keys(st).forEach(function (k) { if (Math.abs(st[k].unidades) > 1e-9) stock.push(st[k]); });

  // Reposiciones de gabinete: última salida a cada gabinete por producto.
  var repos = {};
  movs.forEach(function (m) {
    if (m.tipo !== 'SALIDA' || anulados[m.uuid] || String(m.destino).indexOf('Gabinete') !== 0) return;
    var pid = real_(alias, m.producto_id);
    var k = m.destino + '|' + pid;
    if (!repos[k] || m.fecha > repos[k].fecha) repos[k] = { gabinete: m.destino, producto_id: pid, fecha: m.fecha, unidades: -Number(m.cantidad) };
  });

  var config = { categoria: [], armario: [], destino: [], proveedor: [] };
  filas_('CONFIG').forEach(function (c) {
    if (!config[c.tipo]) config[c.tipo] = [];
    config[c.tipo].push({ nombre: c.nombre, modo_defecto: c.modo_defecto, lote_defecto: c.lote_defecto });
  });

  var ultimos = movs.slice(-300).map(function (m) {
    return { uuid: m.uuid, fecha: m.fecha, tipo: m.tipo, producto_id: real_(alias, m.producto_id), lote: m.lote, caducidad: m.caducidad,
      cantidad: Number(m.cantidad), destino: m.destino, factura: m.factura, origen: m.origen, nota: m.nota,
      anula_uuid: m.anula_uuid, anulado: !!anulados[m.uuid] };
  });

  return {
    ok: true,
    productos: prodFilas.map(producto_),
    codigos: filas_('CODIGOS').map(function (c) { return { codigo: c.codigo, producto_id: real_(alias, c.producto_id) }; }),
    catalogo: filasOpc_('CATALOGO').map(function (c) {
      return { proveedor: c.proveedor, ref: c.ref, descripcion: c.descripcion, precio: Number(c.precio) || 0, fecha: c.fecha, producto_id: real_(alias, c.producto_id) };
    }),
    noDuplicados: filasOpc_('NO_DUPLICADOS').map(function (x) { return [x.producto_a, x.producto_b]; }),
    pedidos: pedidos_(movs, anulados, alias),
    config: config,
    dotacion: filas_('DOTACION').map(function (d) { return { gabinete: d.gabinete, producto_id: d.producto_id, cantidad: Number(d.cantidad) || 0, dias_revision: Number(d.dias_revision) || 0 }; }),
    stock: stock,
    reposiciones: Object.keys(repos).map(function (k) { return repos[k]; }),
    movimientos: ultimos,
    ultimaCopia: PropertiesService.getScriptProperties().getProperty('ULTIMA_COPIA') || '',
    imagenes: diagnosticoImagenes_(),
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
    var alias = alias_(prodFilas);
    var stock = calcularStock_(movs, alias);
    var destinos = {};
    filas_('CONFIG').forEach(function (c) { if (c.tipo === 'destino') destinos[c.nombre] = true; });

    var nuevosMovs = [], rechazados = [], resultados = [];
    var cambiosProd = false, hayPendientes = false, cambiosCodigos = false, noDup = [], fusionesHechas = [];
    var pedFilas = null, cambiosPed = false, refsPedido = [];
    asegurarCabeceras_('PRODUCTOS');

    ops.forEach(function (op) {
      try {
        if (op.tipo === 'producto') {
          var p = validarProducto_(op.producto);
          var prev = prodIdx[p.id] !== undefined ? prodFilas[prodIdx[p.id]] : null;
          var foto = prev ? prev.foto : '', intentos = prev ? prev.intentos_ia : '';
          var estadoIa = op.producto.estado_ia !== undefined ? String(op.producto.estado_ia) : (prev ? prev.estado_ia : '');
          if (['', 'pendiente', 'revisar', 'error'].indexOf(estadoIa) < 0) estadoIa = '';
          if (Array.isArray(op.fotos) && op.fotos.length) {
            var nuevas = guardarFotos_(p.id, op.fotos);
            if (nuevas) foto = (nuevas + (foto ? ',' + foto : '')).split(',').slice(0, 3).join(','); // la última primero
            if (op.identificar) { estadoIa = 'pendiente'; intentos = '0'; hayPendientes = true; }
          }
          if (prev && prev.fusionado_en) throw new Error('Este producto se combinó con otro: actualiza la app');
          var obj = {};
          HOJAS.PRODUCTOS.forEach(function (c) { obj[c] = prev ? prev[c] : ''; });
          [['id', p.id], ['nombre', p.nombre], ['categoria', p.categoria], ['proveedor', p.proveedor], ['ref_proveedor', p.ref_proveedor],
            ['uds_envase', p.uds_envase], ['modo', p.modo], ['lote', p.lote], ['armario', p.armario], ['minimo', p.minimo], ['activo', p.activo],
            ['notas', p.notas], ['actualizado', ahora], ['foto', foto], ['estado_ia', estadoIa], ['intentos_ia', intentos]]
            .forEach(function (kv) { obj[kv[0]] = String(kv[1]); });
          if (prodIdx[p.id] !== undefined) {
            prodFilas[prodIdx[p.id]] = obj;
          } else {
            prodIdx[p.id] = prodFilas.length;
            prodFilas.push(obj);
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

        if (op.tipo === 'fusionar') {
          var de = texto_(op.de, 40), en = texto_(op.en, 40);
          if (prodIdx[de] === undefined || prodIdx[en] === undefined) throw new Error('Producto desconocido');
          de = real_(alias, de); en = real_(alias, en);
          if (de === en) { resultados.push({ op: op.op_id, estado: 'ok' }); return; }
          var pde = prodFilas[prodIdx[de]], pen = prodFilas[prodIdx[en]];
          pde.fusionado_en = en; pde.activo = 'NO'; pde.actualizado = ahora;
          if (!pen.foto && pde.foto) { pen.foto = pde.foto; pen.img_estado = pde.img_estado; }
          if (!pen.proveedor && pde.proveedor) pen.proveedor = pde.proveedor;
          if (!pen.ref_proveedor && pde.ref_proveedor) pen.ref_proveedor = pde.ref_proveedor;
          if ((Number(pen.minimo) || 0) < (Number(pde.minimo) || 0)) pen.minimo = pde.minimo;
          if (pen.lote !== 'SI' && pde.lote === 'SI') pen.lote = 'SI';
          if (pen.estado_ia === 'pendiente' || pen.estado_ia === 'revisar') pen.estado_ia = '';
          pen.actualizado = ahora;
          alias = alias_(prodFilas);
          Object.keys(codigos).forEach(function (c) { if (codigos[c] === de) { codigos[c] = en; cambiosCodigos = true; } });
          stock = calcularStock_(movs, alias);
          cambiosProd = true;
          fusionesHechas.push([de, en]);
          resultados.push({ op: op.op_id, estado: 'ok' });
          return;
        }

        if (op.tipo === 'pedido') {
          var pd = op.pedido || {};
          var pidPed = texto_(pd.id, 40);
          if (!/^PD-[A-Za-z0-9-]+$/.test(pidPed)) throw new Error('Pedido sin identificador');
          var lineas = Array.isArray(pd.lineas) ? pd.lineas : [];
          if (!lineas.length || lineas.length > 300) throw new Error('El pedido no tiene líneas');
          if (!pedFilas) pedFilas = filasOpc_('PEDIDOS');
          var previas = pedFilas.filter(function (x) { return x.pedido === pidPed; });
          var base = { proveedor: texto_(pd.proveedor, 80), numero: texto_(pd.numero, 40), fecha: fechaCad_(pd.fecha) || ahora.slice(0, 10),
            estado: previas.length ? previas[0].estado : '', factura: previas.length ? previas[0].factura : '' };
          var nuevasPed = lineas.map(function (l, i) {
            var prodL = texto_(l.producto_id, 40);
            if (prodL) { if (prodIdx[prodL] === undefined) throw new Error('Producto desconocido en la línea ' + (i + 1)); prodL = real_(alias, prodL); }
            var cant = Number(l.cantidad);
            if (!isFinite(cant) || cant <= 0 || cant > 100000) throw new Error('Cantidad no válida en la línea ' + (i + 1));
            return { pedido: pidPed, linea: String(Number(l.linea) || i + 1), proveedor: base.proveedor, numero: base.numero, fecha: base.fecha,
              producto_id: prodL, ref: texto_(l.ref, 60), descripcion: texto_(l.descripcion, 200), cantidad: String(cant),
              precio: String(Number(l.precio) || 0), iva: texto_(l.iva, 4), estado: base.estado, factura: base.factura, alta: previas.length ? previas[0].alta : ahora };
          });
          pedFilas = pedFilas.filter(function (x) { return x.pedido !== pidPed; }).concat(nuevasPed);
          cambiosPed = true;
          nuevasPed.forEach(function (l) { if (l.ref && l.producto_id) refsPedido.push(l); });
          resultados.push({ op: op.op_id, estado: 'ok' });
          return;
        }

        if (op.tipo === 'pedido_estado') {
          var idPe = texto_(op.id, 40);
          if (!pedFilas) pedFilas = filasOpc_('PEDIDOS');
          var suyas = pedFilas.filter(function (x) { return x.pedido === idPe; });
          if (!suyas.length) throw new Error('Pedido desconocido');
          var est = texto_(op.estado, 12);
          if (op.estado !== undefined && ['', 'cerrado', 'anulado'].indexOf(est) < 0) throw new Error('Estado de pedido no válido');
          suyas.forEach(function (x) {
            if (op.estado !== undefined) x.estado = est;
            if (op.factura !== undefined) x.factura = texto_(op.factura, 60);
          });
          cambiosPed = true;
          resultados.push({ op: op.op_id, estado: 'ok' });
          return;
        }

        if (op.tipo === 'no_duplicado') {
          noDup.push([texto_(op.a, 40), texto_(op.b, 40), ahora]);
          resultados.push({ op: op.op_id, estado: 'ok' });
          return;
        }

        if (op.tipo === 'movimiento') {
          var m = op.movimiento || {};
          if (!m.uuid) throw new Error('Movimiento sin uuid');
          if (uuids[m.uuid]) { resultados.push({ op: op.op_id, uuid: m.uuid, estado: 'duplicado' }); return; }
          var r = validarMovimiento_(m, prodFilas, prodIdx, stock, destinos, porUuid, anulados, alias);
          var filaM = [m.uuid, String(m.fecha || ahora), r.tipo, r.producto_id, r.lote, r.caducidad, r.cantidad, r.destino,
            String(m.factura || ''), String(m.origen || ''), String(m.nota || '').slice(0, 300), r.anula_uuid, ahora];
          nuevosMovs.push(filaM);
          var objM = filaAObjeto_('MOVIMIENTOS', filaM);
          movs.push(objM);
          uuids[m.uuid] = true; porUuid[m.uuid] = objM;
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
    if (cambiosCodigos) {
      var hc = ss.getSheetByName('CODIGOS');
      var cods = filas_('CODIGOS').map(function (c) { return [c.codigo, codigos[c.codigo] || c.producto_id, c.alta]; });
      if (cods.length) hc.getRange(2, 1, cods.length, 3).setNumberFormat('@').setValues(cods);
    }
    anadir_('CODIGOS', nuevosCodigos);
    anadir_('MOVIMIENTOS', nuevosMovs);
    anadir_('RECHAZADOS', rechazados);
    if (noDup.length) { asegurarHoja_('NO_DUPLICADOS'); anadir_('NO_DUPLICADOS', noDup); }
    if (fusionesHechas.length) repuntarFusiones_(fusionesHechas);
    if (cambiosPed) {
      var hp = asegurarHoja_('PEDIDOS');
      var cabP = HOJAS.PEDIDOS;
      if (hp.getLastRow() > 1) hp.getRange(2, 1, hp.getLastRow() - 1, cabP.length).clearContent();
      if (pedFilas.length) hp.getRange(2, 1, pedFilas.length, cabP.length).setNumberFormat('@').setValues(pedFilas.map(function (x) { return cabP.map(function (c) { return x[c]; }); }));
    }
    if (refsPedido.length) catalogoDesdePedido_(refsPedido, ahora);
    if (cambiosProd || nuevosMovs.length) recalcularStock();
    if (hayPendientes) { try { programarIdentificacion_(); } catch (e) { Logger.log('No se pudo programar la identificación: ' + e); } }

    var estado = leerEstado_();
    estado.resultados = resultados;
    return estado;
  } finally {
    lock.releaseLock();
  }
}

function asegurarHoja_(nombre) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.getRange(1, 1, 1, HOJAS[nombre].length).setValues([HOJAS[nombre]]).setFontWeight('bold');
    h.setFrozenRows(1);
    h.getRange(1, 1, h.getMaxRows(), HOJAS[nombre].length).setNumberFormat('@');
  }
  return h;
}

/** Tras combinar productos: catálogo y dotación apuntan al que se conserva. */
function repuntarFusiones_(pares) {
  var mapa = {};
  pares.forEach(function (x) { mapa[x[0]] = x[1]; });
  [['CATALOGO', 'producto_id'], ['DOTACION', 'producto_id']].forEach(function (hc) {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hc[0]);
    if (!hoja || hoja.getLastRow() < 2) return;
    var cab = HOJAS[hc[0]], col = cab.indexOf(hc[1]);
    var r = hoja.getRange(2, col + 1, hoja.getLastRow() - 1, 1);
    var v = r.getValues(), cambio = false;
    v.forEach(function (f) { if (mapa[f[0]]) { f[0] = mapa[f[0]]; cambio = true; } });
    if (cambio) r.setValues(v);
  });
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
    armario: texto_(p.armario, 80), minimo: minimo, activo: p.activo === false ? 'NO' : 'SI', notas: texto_(p.notas, 500)
  };
}

function fechaCad_(v) {
  v = texto_(v, 10);
  if (!v) return '';
  if (!/^\d{4}-\d{2}(-\d{2})?$/.test(v)) throw new Error('Caducidad no válida (AAAA-MM o AAAA-MM-DD)');
  return v;
}

function validarMovimiento_(m, prodFilas, prodIdx, stock, destinos, porUuid, anulados, alias) {
  alias = alias || {};
  var tipo = texto_(m.tipo, 12);
  if (TIPOS.indexOf(tipo) < 0) throw new Error('Tipo de movimiento no válido');

  if (tipo === 'ANULACION') {
    var orig = porUuid[texto_(m.anula_uuid, 60)];
    if (!orig) throw new Error('El movimiento a anular no existe');
    if (orig.tipo === 'ANULACION') throw new Error('No se puede anular una anulación');
    if (anulados[orig.uuid]) throw new Error('Ese movimiento ya está anulado');
    var cant = -Number(orig.cantidad);
    var kk = real_(alias, orig.producto_id) + '|' + (orig.lote || '');
    var disp = stock[kk] ? stock[kk].unidades : 0;
    if (disp + cant < -1e-9) throw new Error('Anularlo dejaría el stock en negativo');
    return { tipo: tipo, producto_id: real_(alias, orig.producto_id), lote: orig.lote || '', caducidad: orig.caducidad || '', cantidad: cant, destino: orig.destino || '', anula_uuid: orig.uuid };
  }

  var pid = texto_(m.producto_id, 40);
  if (prodIdx[pid] === undefined) throw new Error('Producto desconocido');
  pid = real_(alias, pid);
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
  var pf = filas_('PRODUCTOS');
  pf.forEach(function (p) { prods[p.id] = producto_(p); });
  var st = calcularStock_(filas_('MOVIMIENTOS'), alias_(pf));
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

/* ───────────────────────── Identificar productos con Gemini ───────────────────────── */
/*
 * Necesita la propiedad del script GEMINI_KEY (clave de Google AI Studio).
 * Opcional: GEMINI_MODELO (por defecto gemini-flash-latest).
 *
 * Flujo normal (sin esperas en el móvil): la app da de alta el producto con un nombre provisional
 * y sus fotos; las fotos se guardan en Drive («Inventario Castilla - Fotos») y el producto queda
 * «pendiente». Un disparador en segundo plano pide a Gemini que lo identifique con la foto y, si hace
 * falta, buscando el código (GTIN) en internet, y deja el producto «por revisar».
 */

/** Añade a la cabecera las columnas nuevas (actualizaciones de versión). */
function asegurarCabeceras_(nombre) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!hoja) return;
  var cab = HOJAS[nombre];
  var actual = hoja.getRange(1, 1, 1, cab.length).getValues()[0].map(String);
  if (actual.join('|') !== cab.join('|')) hoja.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
}

function carpeta_(prop, nombre) {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty(prop), c = null;
  try { c = id && DriveApp.getFolderById(id); } catch (e) { c = null; }
  if (!c) { c = DriveApp.createFolder(nombre); props.setProperty(prop, c.getId()); }
  return c;
}

function guardarFotos_(pid, fotos) {
  var c = carpeta_('CARPETA_FOTOS', 'Inventario Castilla - Fotos');
  var ids = [];
  fotos.filter(function (b) { return typeof b === 'string' && b.length > 100; }).slice(0, 3).forEach(function (b, i) {
    var bytes = Utilities.base64Decode(b);
    var png = bytes[0] === -119 && bytes[1] === 80, webp = bytes[8] === 87 && bytes[9] === 69;
    var tipo = png ? 'image/png' : webp ? 'image/webp' : 'image/jpeg';
    var blob = Utilities.newBlob(bytes, tipo, pid + '-' + (i + 1) + (png ? '.png' : webp ? '.webp' : '.jpg'));
    ids.push(c.createFile(blob).getId());
  });
  return ids.join(',');
}

/**
 * Disparadores: uno inmediato (al dar de alta productos) y uno periódico cada hora.
 * Cada hora y con tiempo limitado por vuelta: las cuentas gratuitas de Google solo permiten
 * 90 minutos al día de ejecución por disparadores; si se pasa, Google los para hasta el día siguiente.
 */
function programarIdentificacion_(soloPeriodico) {
  var props = PropertiesService.getScriptProperties();
  var ts = ScriptApp.getProjectTriggers();
  if (props.getProperty('DISPARADOR_V') !== '2') {
    ts.forEach(function (t) { if (t.getHandlerFunction() === 'identificarPendientes') ScriptApp.deleteTrigger(t); });
    ts = ScriptApp.getProjectTriggers();
    props.setProperty('DISPARADOR_V', '2');
  }
  var ahora = ts.some(function (t) { return t.getHandlerFunction() === 'identificarAhora'; });
  var periodico = ts.some(function (t) { return t.getHandlerFunction() === 'identificarPeriodico'; });
  if (!ahora && !soloPeriodico) ScriptApp.newTrigger('identificarAhora').timeBased().after(10 * 1000).create();
  if (!periodico) ScriptApp.newTrigger('identificarPeriodico').timeBased().everyHours(1).create();
}

function identificarPeriodico() { identificarPendientes(150000); }

/** Desde el menú: sin límite de disparadores, hasta 5 minutos. */
function buscarImagenesAhora() {
  identificarPendientes(300000);
  var d = diagnosticoImagenes_();
  aviso_('Imágenes: ' + d.conFoto + ' con foto, ' + d.pendientes + ' por buscar, ' + d.sinImagen + ' sin encontrar.' +
    (d.ultimoError ? ' Último error: ' + d.ultimoError : ''));
}

/** Vuelve a buscar las que no se encontraron (por ejemplo, tras un fallo de cuota). */
function reintentarImagenes() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var filas = filas_('PRODUCTOS'), n = 0;
    filas.forEach(function (p) { if (!p.foto && (p.img_estado === 'no' || p.img_estado === 'no1')) { p.img_estado = ''; n++; } });
    escribirProductos_(filas);
    aviso_(n + ' productos volverán a buscar imagen.');
  } finally { lock.releaseLock(); }
  try { programarIdentificacion_(); } catch (e) { Logger.log(e); }
}

function escribirProductos_(filas) {
  var cab = HOJAS.PRODUCTOS;
  if (!filas.length) return;
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PRODUCTOS').getRange(2, 1, filas.length, cab.length)
    .setNumberFormat('@').setValues(filas.map(function (p) { return cab.map(function (c) { return p[c]; }); }));
}

function diagnosticoImagenes_() {
  var props = PropertiesService.getScriptProperties();
  var d = { conFoto: 0, pendientes: 0, sinImagen: 0, ultimoError: props.getProperty('ULTIMO_ERROR_IMG') || '',
    ultimaBusqueda: props.getProperty('ULTIMA_BUSQUEDA_IMG') || '', gemini: !!props.getProperty('GEMINI_KEY') };
  filasOpc_('PRODUCTOS').forEach(function (p) {
    if (p.activo === 'NO' || p.fusionado_en) return;
    if (p.foto) d.conFoto++;
    else if (p.img_estado === 'no') d.sinImagen++;
    else d.pendientes++;
  });
  try { d.disparador = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'identificarPeriodico'; }); } catch (e) { d.disparador = null; }
  return d;
}

function identificarAhora() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'identificarAhora') ScriptApp.deleteTrigger(t); });
  identificarPendientes();
}

function catsConfig_() {
  return filas_('CONFIG').filter(function (c) { return c.tipo === 'categoria'; });
}

/** Pide a Gemini los datos del producto. buscar = permitir búsqueda en Google por el GTIN. */
function pedirGemini_(imagenes, lec, cats, buscar) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_KEY');
  if (!key) throw new Error('Falta la clave de Gemini: Apps Script → Configuración del proyecto → Propiedades del script → GEMINI_KEY');
  var modelo = props.getProperty('GEMINI_MODELO') || 'gemini-flash-latest';
  lec = lec || {};
  var prompt = [
    'Eres el asistente de inventario de una clínica dental en España. Hay que identificar un producto (material dental, medicamento, fungible o limpieza).',
    'Datos leídos del código de barras: GTIN ' + (lec.gtin || '—') + ', lote ' + (lec.lote || '—') + ', caducidad ' + (lec.caducidad || '—') + '.',
    imagenes.length ? 'Te adjunto foto(s) del envase.' : 'No hay fotos.',
    buscar && lec.gtin ? 'Si la foto no basta para saber qué producto es, busca el GTIN en internet (también sin el cero inicial, como EAN-13).' : '',
    'Responde SOLO con un objeto JSON con estas claves:',
    '- nombre: en español, corto (máx. 60 caracteres), como lo llamaría la clínica: tipo de producto + marca + modelo, medida, talla, color o concentración. Ejemplos: "Guantes látex Santex talla S", "Implante Klockner Essential Cone 4.0×10", "Meganest 40 mg/ml 1:100.000", "Dentasept Tri Enzymatic 5 L".',
    '- marca: marca o fabricante.',
    '- ref_proveedor: la referencia del fabricante (texto tras REF o Ref.), o "" si no aparece.',
    '- uds_envase: número entero de unidades separables que contiene el envase (100 guantes, 50 carpules, 5 fresas); 1 si no consta.',
    '- categoria: exactamente una de: ' + cats.join(' | ') + '.',
    '- lote y caducidad: solo si se leen en la foto y no venían en el código (caducidad AAAA-MM o AAAA-MM-DD), si no "".',
    '- fuente: "foto", "internet" o "foto e internet".',
    '- dudas: frase corta si algo no está claro; "" si todo está claro.',
    'No inventes datos: si no lo sabes, deja el campo vacío.'
  ].filter(String).join('\n');
  var partes = [{ text: prompt }].concat(imagenes.map(function (b) { return { inline_data: { mime_type: 'image/jpeg', data: b } }; }));
  var cuerpoJson = { contents: [{ parts: partes }], generationConfig: { responseMimeType: 'application/json', temperature: 0.1 } };
  var cuerpoBusqueda = { contents: [{ parts: partes }], tools: [{ google_search: {} }], generationConfig: { temperature: 0.1 } };
  var cuerpos = buscar && lec.gtin ? [cuerpoBusqueda, cuerpoJson] : [cuerpoJson];

  var modelos = modelosGemini_();
  var ultimoError = '';
  for (var c = 0; c < cuerpos.length; c++) {
    for (var i = 0; i < modelos.length; i++) {
      for (var intento = 0; intento < 2; intento++) {
        var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelos[i]) + ':generateContent', {
          method: 'post', contentType: 'application/json', muteHttpExceptions: true,
          headers: { 'x-goog-api-key': key }, payload: JSON.stringify(cuerpos[c])
        });
        var code = r.getResponseCode();
        if (code === 200) {
          var d = leerJsonGemini_(r.getContentText());
          if (d) return sanearPropuesta_(d, cats);
          ultimoError = 'respuesta sin JSON';
          break;
        }
        ultimoError = modelos[i] + ' respondió ' + code + ': ' + r.getContentText().slice(0, 200);
        if (code === 401 || code === 403) throw new Error('Gemini: ' + ultimoError);
        if (code === 400 || code === 404) break; // modelo o herramienta no disponible: probar lo siguiente
        Utilities.sleep(1500 * (intento + 1));
      }
    }
  }
  throw new Error('Gemini no ha podido identificarlo ahora (' + ultimoError + ')');
}

function leerJsonGemini_(texto) {
  try {
    var j = JSON.parse(texto);
    var t = ((j.candidates && j.candidates[0] && j.candidates[0].content && j.candidates[0].content.parts) || [])
      .map(function (p) { return p.text || ''; }).join('');
    var a = t.indexOf('{'), b = t.lastIndexOf('}');
    if (a < 0 || b < a) return null;
    return JSON.parse(t.slice(a, b + 1));
  } catch (e) { return null; }
}

function sanearPropuesta_(d, cats) {
  var cad = String(d.caducidad || '');
  return {
    nombre: texto_(d.nombre, 120), marca: texto_(d.marca, 80), ref_proveedor: texto_(d.ref_proveedor, 60),
    uds_envase: Math.max(1, Math.floor(Number(d.uds_envase) || 1)),
    categoria: cats.indexOf(String(d.categoria || '')) >= 0 ? String(d.categoria) : '',
    lote: texto_(d.lote, 40), caducidad: /^\d{4}-\d{2}(-\d{2})?$/.test(cad) ? cad : '',
    fuente: texto_(d.fuente, 30), dudas: texto_(d.dudas, 200)
  };
}

/** Identificación inmediata (la app espera la respuesta). */
function identificar_(pet) {
  var imagenes = (pet.imagenes || []).filter(function (b) { return typeof b === 'string' && b.length > 100; }).slice(-3);
  if (!imagenes.length) throw new Error('No ha llegado ninguna foto');
  var cats = catsConfig_().map(function (c) { return c.nombre; });
  return { ok: true, propuesta: pedirGemini_(imagenes, pet.lectura || {}, cats, false) };
}

/** Identifica en segundo plano los productos «pendiente» (máx. 5 por vuelta). */
function identificarPendientes(limiteMs) {
  if (!PropertiesService.getScriptProperties().getProperty('GEMINI_KEY')) return;
  var fin = Date.now() + (Number(limiteMs) || 150000);
  asegurarCabeceras_('PRODUCTOS');
  identificarFotos_();
  if (Date.now() < fin - 20000) buscarImagenes_(40, fin);
}

function identificarFotos_() {
  var pendientes = filas_('PRODUCTOS').filter(function (p) { return p.estado_ia === 'pendiente'; }).slice(0, 5);
  if (!pendientes.length) return;
  var codigos = filas_('CODIGOS');
  var cfgCats = catsConfig_();
  var cats = cfgCats.map(function (c) { return c.nombre; });
  var resultados = {};
  pendientes.forEach(function (p) {
    var gtin = (codigos.filter(function (c) { return c.producto_id === p.id && /^\d{14}$/.test(c.codigo); })[0] || {}).codigo || '';
    var imgs = String(p.foto || '').split(',').filter(String).map(function (id) {
      try { return Utilities.base64Encode(DriveApp.getFileById(id).getBlob().getBytes()); } catch (e) { return null; }
    }).filter(Boolean);
    try { resultados[p.id] = { ok: true, d: pedirGemini_(imgs, { gtin: gtin }, cats, true) }; }
    catch (e) { resultados[p.id] = { ok: false, error: String(e.message || e) }; }
  });

  // Escribir con bloqueo, sobre la versión actual de la hoja.
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var filas = filas_('PRODUCTOS');
    var conMovs = {};
    filas_('MOVIMIENTOS').forEach(function (m) { conMovs[m.producto_id] = true; });
    var cambios = false;
    filas.forEach(function (p) {
      var r = resultados[p.id];
      if (!r || p.estado_ia !== 'pendiente') return; // si alguien la ha editado mientras tanto, no se toca
      cambios = true;
      if (!r.ok) {
        var n = (Number(p.intentos_ia) || 0) + 1;
        p.intentos_ia = String(n);
        if (n >= 4) { p.estado_ia = 'error'; p.notas = (p.notas ? p.notas + ' · ' : '') + 'IA: ' + r.error.slice(0, 120); }
        return;
      }
      aplicarPropuesta_(p, r.d, cfgCats, !!conMovs[p.id]);
      p.estado_ia = 'revisar';
    });
    if (cambios) {
      var cab = HOJAS.PRODUCTOS;
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName('PRODUCTOS').getRange(2, 1, filas.length, cab.length)
        .setNumberFormat('@').setValues(filas.map(function (p) { return cab.map(function (c) { return p[c]; }); }));
      recalcularStock();
    }
  } finally {
    lock.releaseLock();
  }
}

function aplicarPropuesta_(p, d, cfgCats, tieneMovimientos) {
  if (d.nombre) p.nombre = d.nombre;
  if (!p.proveedor && d.marca) p.proveedor = d.marca;
  if (!p.ref_proveedor && d.ref_proveedor) p.ref_proveedor = d.ref_proveedor;
  if (!p.categoria && d.categoria) {
    p.categoria = d.categoria;
    var c = cfgCats.filter(function (x) { return x.nombre === d.categoria; })[0];
    if (c && !tieneMovimientos) p.modo = c.modo_defecto === 'abrir' ? 'abrir' : 'unidades';
    if (c && c.lote_defecto === 'SI') p.lote = 'SI';
  }
  var notas = [];
  if (d.uds_envase > 1) {
    // Cambiar las unidades por envase con stock ya registrado alteraría el stock: entonces solo se anota.
    if (p.modo !== 'abrir' && !tieneMovimientos && Number(p.uds_envase) <= 1) p.uds_envase = String(d.uds_envase);
    else notas.push('envase de ' + d.uds_envase + ' u.');
  }
  if (d.fuente) notas.push('identificado por IA (' + d.fuente + ')');
  if (d.dudas) notas.push(d.dudas);
  if (notas.length) p.notas = ((p.notas ? p.notas + ' · ' : '') + notas.join(' · ')).slice(0, 500);
}

/** Devuelve la foto principal de un producto (la más reciente) en base64. */
function leerFoto_(pid) {
  var p = filas_('PRODUCTOS').filter(function (x) { return x.id === String(pid || ''); })[0];
  if (!p || !p.foto) return { ok: true, base64: '' };
  var id = String(p.foto).split(',')[0];
  try { var b = DriveApp.getFileById(id).getBlob(); return { ok: true, base64: Utilities.base64Encode(b.getBytes()), tipo: b.getContentType() || 'image/jpeg' }; }
  catch (e) { return { ok: true, base64: '' }; }
}

/* ───────────────────────── Catálogo desde facturas ───────────────────────── */
/*
 * Se importa el Excel «Catalogo inventario.xlsx» como pestaña nueva llamada CATALOGO_IMPORTAR
 * (Archivo → Importar → Subir → Insertar hojas nuevas; luego renombrarla) y se ejecuta
 * Inventario → Cargar catálogo. Columnas: articulo, nombre, categoria, marca, proveedor, ref,
 * descripcion, precio, fecha, factura. Se puede repetir: no duplica (clave proveedor + referencia).
 */
function claveCat_(prov, ref) { return String(prov).trim().toLowerCase() + '|' + String(ref).trim().toUpperCase().replace(/\s+/g, ''); }

/** Las líneas de un pedido enlazadas a un producto pasan al catálogo (depósito + referencia + último precio). */
function catalogoDesdePedido_(lineas, ahora) {
  var hCat = asegurarHoja_('CATALOGO');
  var cat = filas_('CATALOGO');
  var porClave = {};
  cat.forEach(function (c, i) { porClave[claveCat_(c.proveedor, c.ref)] = i; });
  var cambio = false, nuevas = [];
  lineas.forEach(function (l) {
    if (!l.proveedor) return;
    var k = claveCat_(l.proveedor, l.ref);
    var i = porClave[k];
    var fac = 'Pedido ' + l.numero;
    if (i === undefined) {
      porClave[k] = cat.length + nuevas.length;
      nuevas.push([l.proveedor, l.ref, l.descripcion, l.precio, l.fecha, fac, l.producto_id, ahora]);
      return;
    }
    var e = cat[i];
    if (!e.producto_id) { e.producto_id = l.producto_id; cambio = true; }
    if (String(l.fecha) >= String(e.fecha) && Number(l.precio) > 0) { e.precio = l.precio; e.fecha = l.fecha; e.factura = fac; e.descripcion = l.descripcion || e.descripcion; cambio = true; }
  });
  var cab = HOJAS.CATALOGO;
  if (cambio && cat.length) hCat.getRange(2, 1, cat.length, cab.length).setNumberFormat('@').setValues(cat.map(function (c) { return cab.map(function (k) { return c[k]; }); }));
  anadir_('CATALOGO', nuevas);
}

function cargarCatalogo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var imp = ss.getSheetByName('CATALOGO_IMPORTAR');
  if (!imp) { aviso_('No encuentro la pestaña CATALOGO_IMPORTAR. Importa el Excel del catálogo como hoja nueva y renómbrala así.'); return; }
  var datos = imp.getDataRange().getValues();
  var cab = datos.shift().map(function (c) { return String(c).trim().toLowerCase(); });
  var col = function (n) { return cab.indexOf(n); };
  ['articulo', 'nombre', 'proveedor', 'ref'].forEach(function (n) { if (col(n) < 0) throw new Error('Falta la columna ' + n + ' en CATALOGO_IMPORTAR'); });
  var filas = datos.filter(function (r) { return String(r[col('articulo')]).trim() && String(r[col('ref')]).trim(); }).map(function (r) {
    var o = {};
    cab.forEach(function (c, i) { o[c] = r[i] instanceof Date ? Utilities.formatDate(r[i], 'Europe/Madrid', 'yyyy-MM-dd') : String(r[i] === null ? '' : r[i]).trim(); });
    return o;
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    asegurarCabeceras_('PRODUCTOS');
    var hCat = asegurarHoja_('CATALOGO');
    var cat = filas_('CATALOGO');
    var porClave = {};
    cat.forEach(function (c, i) { porClave[claveCat_(c.proveedor, c.ref)] = i; });
    var prods = filas_('PRODUCTOS');
    var prodIdx = {};
    prods.forEach(function (p, i) { prodIdx[p.id] = i; });
    var alias = alias_(prods);
    var cfgCats = catsConfig_();
    var ahora = new Date().toISOString();

    // Artículo → producto (si alguna de sus referencias ya estaba enlazada, se reutiliza).
    var artProd = {};
    filas.forEach(function (f) {
      var i = porClave[claveCat_(f.proveedor, f.ref)];
      if (i !== undefined && cat[i].producto_id) artProd[f.articulo] = real_(alias, cat[i].producto_id);
    });
    var nuevosProd = 0, nuevasRefs = 0, actualizadas = 0;
    var porArticulo = {};
    filas.forEach(function (f) { (porArticulo[f.articulo] = porArticulo[f.articulo] || []).push(f); });
    Object.keys(porArticulo).forEach(function (art) {
      var grupo = porArticulo[art].sort(function (a, b) { return String(b.fecha).localeCompare(String(a.fecha)); });
      var reciente = grupo[0];
      var pid = artProd[art];
      if (!pid) {
        pid = 'P-C' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, art).map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, '0'); }).join('').slice(0, 10).toUpperCase();
        if (prodIdx[pid] === undefined) {
          var c = cfgCats.filter(function (x) { return x.nombre === reciente.categoria; })[0];
          var o = {};
          HOJAS.PRODUCTOS.forEach(function (k) { o[k] = ''; });
          o.id = pid; o.nombre = String(reciente.nombre || reciente.descripcion).slice(0, 120); o.categoria = c ? c.nombre : '';
          o.proveedor = reciente.proveedor; o.ref_proveedor = reciente.ref; o.uds_envase = '1';
          o.modo = c && c.modo_defecto === 'abrir' ? 'abrir' : 'unidades'; o.lote = c && c.lote_defecto === 'SI' ? 'SI' : 'NO';
          o.minimo = '0'; o.activo = 'SI'; o.notas = 'Del catálogo de facturas' + (reciente.marca ? ' · marca ' + reciente.marca : '');
          o.actualizado = ahora;
          prodIdx[pid] = prods.length; prods.push(o); nuevosProd++;
        }
      }
      grupo.forEach(function (f) {
        var k = claveCat_(f.proveedor, f.ref);
        var i = porClave[k];
        if (i === undefined) {
          porClave[k] = cat.length;
          cat.push({ proveedor: f.proveedor, ref: f.ref, descripcion: f.descripcion, precio: f.precio, fecha: f.fecha, factura: f.factura, producto_id: pid, alta: ahora });
          nuevasRefs++;
        } else {
          var e = cat[i];
          if (String(f.fecha) > String(e.fecha)) { e.precio = f.precio; e.fecha = f.fecha; e.factura = f.factura; e.descripcion = f.descripcion; actualizadas++; }
          if (!e.producto_id) e.producto_id = pid;
        }
      });
    });
    var cp = HOJAS.PRODUCTOS;
    ss.getSheetByName('PRODUCTOS').getRange(2, 1, prods.length, cp.length).setNumberFormat('@').setValues(prods.map(function (p) { return cp.map(function (k) { return p[k]; }); }));
    var cc = HOJAS.CATALOGO;
    if (cat.length) hCat.getRange(2, 1, cat.length, cc.length).setNumberFormat('@').setValues(cat.map(function (c) { return cc.map(function (k) { return c[k] === undefined ? '' : String(c[k]); }); }));
    recalcularStock();
  } finally {
    lock.releaseLock();
  }
  try { programarIdentificacion_(); } catch (e) { Logger.log(e); }
  aviso_('Catálogo cargado: ' + nuevosProd + ' productos nuevos, ' + nuevasRefs + ' referencias nuevas, ' + actualizadas + ' precios actualizados. Las imágenes se irán buscando solas en segundo plano.');
}

/* ───────────────────────── Imágenes de internet ───────────────────────── */

/**
 * Busca imagen para productos sin foto, hasta `n` o hasta la hora `fin`.
 * Si Gemini falla (cuota, saturación) se para y no marca nada: se reintentará en la siguiente vuelta.
 */
function buscarImagenes_(n, fin) {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('GEMINI_KEY')) return 0;
  fin = fin || Date.now() + 120000;
  var candidatos = filas_('PRODUCTOS').filter(function (p) {
    return p.activo !== 'NO' && !p.fusionado_en && !p.foto && p.estado_ia !== 'pendiente' && (p.img_estado === '' || p.img_estado === 'no1') && p.nombre.indexOf('Pendiente') !== 0;
  }).slice(0, n || 8);
  if (!candidatos.length) return 0;
  var res = {}, hechos = 0;
  for (var i = 0; i < candidatos.length; i++) {
    if (Date.now() > fin - 25000) break;
    var p = candidatos[i], r;
    try { r = buscarUrlImagen_(p); } catch (e) { r = { estado: 'error', error: String(e.message || e) }; }
    if (r.estado === 'error') { props.setProperty('ULTIMO_ERROR_IMG', new Date().toISOString().slice(0, 16) + ' ' + r.error.slice(0, 180)); break; }
    var blob = r.url ? descargarImagen_(r.url) : null;
    res[p.id] = blob ? guardarFotos_(p.id, [Utilities.base64Encode(blob.getBytes())]) : '';
    hechos++;
  }
  props.setProperty('ULTIMA_BUSQUEDA_IMG', new Date().toISOString());
  if (!hechos) return 0;
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var filas = filas_('PRODUCTOS');
    filas.forEach(function (p) {
      if (!(p.id in res) || p.foto) return;
      if (res[p.id]) { p.foto = res[p.id]; p.img_estado = 'internet'; }
      else p.img_estado = p.img_estado === 'no1' ? 'no' : 'no1';
    });
    escribirProductos_(filas);
  } finally {
    lock.releaseLock();
  }
  return hechos;
}

/** {estado: 'ok'|'sin'|'error', url, error}. Usa las páginas que encuentra la búsqueda de Google (og:image). */
function buscarUrlImagen_(p) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('GEMINI_KEY');
  var modelo = props.getProperty('GEMINI_MODELO') || 'gemini-flash-latest';
  var prompt = 'Busca en internet la página de este producto dental en la web del fabricante o de un distribuidor dental: "' + p.nombre + '"' +
    (p.ref_proveedor ? ', referencia ' + p.ref_proveedor : '') + (p.proveedor ? ', vendido por ' + p.proveedor : '') +
    '. Responde SOLO con JSON: {"pagina": "URL de la página del producto o vacío", "imagen": "URL directa de la imagen del producto o vacío"}. No inventes URLs.';
  var modelos = modelosGemini_();
  var errores = [];
  for (var i = 0; i < modelos.length; i++) {
    var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(modelos[i]) + ':generateContent', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': key },
      payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }], tools: [{ google_search: {} }], generationConfig: { temperature: 0.1 } })
    });
    var code = r.getResponseCode();
    if (code !== 200) {
      errores.push(modelos[i] + ' ' + code + ': ' + mensajeGemini_(r.getContentText()));
      if (code === 401 || code === 403) break;
      continue;
    }
    var texto = r.getContentText();
    var d = leerJsonGemini_(texto) || {};
    var paginas = [];
    if (d.pagina && /^https?:\/\//.test(d.pagina)) paginas.push(d.pagina);
    try {
      var j = JSON.parse(texto);
      var chunks = (j.candidates && j.candidates[0] && j.candidates[0].groundingMetadata && j.candidates[0].groundingMetadata.groundingChunks) || [];
      chunks.forEach(function (c) { if (c.web && c.web.uri) paginas.push(c.web.uri); });
    } catch (e) { /* sin metadatos */ }
    if (d.imagen && /^https?:\/\//.test(d.imagen) && descargarImagen_(d.imagen, true)) return { estado: 'ok', url: d.imagen };
    for (var k = 0; k < paginas.length && k < 4; k++) {
      var og = imagenDePagina_(paginas[k]);
      if (og) return { estado: 'ok', url: og };
    }
    return { estado: 'sin', url: '' };
  }
  return { estado: 'error', error: errores.join(' | ') || 'Gemini no responde' };
}

/** Lee la imagen principal (og:image) de una página de producto. */
function imagenDePagina_(url) {
  try {
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 (inventario clinica)' } });
    if (r.getResponseCode() !== 200) return '';
    var html = r.getContentText().slice(0, 300000);
    var m = html.match(/<meta[^>]+(?:property|name)=["'](?:og:image(?::secure_url)?|twitter:image)["'][^>]+content=["']([^"']+)["']/i) ||
      html.match(/<meta[^>]+content=["']([^"']+)["'][^>]+(?:property|name)=["'](?:og:image|twitter:image)["']/i);
    if (!m) return '';
    var img = m[1].replace(/&amp;/g, '&');
    if (img.indexOf('//') === 0) img = 'https:' + img;
    if (!/^https?:\/\//.test(img)) return ''; // relativa tras una redirección: no se puede resolver con seguridad
    return descargarImagen_(img, true) ? img : '';
  } catch (e) { return ''; }
}

/** Descarga y comprueba que es una imagen razonable (3 KB – 4 MB). */
function descargarImagen_(url, soloComprobar) {
  try {
    var r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0 (inventario clinica)' } });
    if (r.getResponseCode() !== 200) return null;
    var tipo = String(r.getHeaders()['Content-Type'] || r.getHeaders()['content-type'] || '');
    var b = r.getBlob();
    var n = b.getBytes().length;
    if (!/^image\/(jpeg|jpg|png|webp)/i.test(tipo) || n < 3000 || n > 4000000) return null;
    return soloComprobar ? true : b;
  } catch (e) { return null; }
}

/* ───────────────────────── Modelos de Gemini ───────────────────────── */

function mensajeGemini_(texto) {
  try { var j = JSON.parse(texto); return String(j.error && j.error.message || texto).slice(0, 160); } catch (e) { return String(texto).slice(0, 160); }
}

/**
 * Lista de modelos a probar, en orden. Google retira modelos con el tiempo: se piden los disponibles
 * para esta clave (se guarda la lista 12 horas). GEMINI_MODELO, si existe, va primero.
 */
function modelosGemini_() {
  var props = PropertiesService.getScriptProperties();
  var fijo = props.getProperty('GEMINI_MODELO');
  var guardado = props.getProperty('GEMINI_MODELOS');
  var lista = null;
  if (guardado) {
    try { var g = JSON.parse(guardado); if (Date.now() - g.t < 12 * 3600 * 1000 && g.m.length) lista = g.m; } catch (e) { lista = null; }
  }
  if (!lista) {
    lista = [];
    try {
      var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', {
        muteHttpExceptions: true, headers: { 'x-goog-api-key': props.getProperty('GEMINI_KEY') }
      });
      if (r.getResponseCode() === 200) {
        lista = (JSON.parse(r.getContentText()).models || [])
          .filter(function (m) {
            var n = String(m.name || '');
            return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0 && /flash/.test(n) &&
              !/image|tts|audio|live|embedding|thinking|exp|preview/.test(n);
          })
          .map(function (m) { return String(m.name).replace(/^models\//, ''); })
          .sort(function (a, b) {
            var la = /-latest$/.test(a) ? 0 : 1, lb = /-latest$/.test(b) ? 0 : 1;
            if (la !== lb) return la - lb;
            var lite = (/lite/.test(a) ? 1 : 0) - (/lite/.test(b) ? 1 : 0);
            if (lite) return lite;
            return b.localeCompare(a, 'en', { numeric: true });
          });
        if (lista.length) props.setProperty('GEMINI_MODELOS', JSON.stringify({ t: Date.now(), m: lista.slice(0, 6) }));
      }
    } catch (e) { Logger.log('Lista de modelos: ' + e); }
    if (!lista.length) lista = ['gemini-flash-latest', 'gemini-flash-lite-latest'];
  }
  if (fijo) lista = [fijo].concat(lista.filter(function (m) { return m !== fijo; }));
  return lista.slice(0, 4);
}

/** Diagnóstico desde el editor: qué modelos responden, con y sin búsqueda en Google. */
function probarGemini() {
  PropertiesService.getScriptProperties().deleteProperty('GEMINI_MODELOS');
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
  if (!key) { aviso_('Falta GEMINI_KEY en las propiedades del script'); return; }
  var modelos = modelosGemini_();
  var lineas = ['Modelos: ' + modelos.join(', ')];
  modelos.slice(0, 3).forEach(function (m) {
    [false, true].forEach(function (buscar) {
      var cuerpo = { contents: [{ parts: [{ text: 'Responde solo: OK' }] }] };
      if (buscar) cuerpo.tools = [{ google_search: {} }];
      var r = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(m) + ':generateContent', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true, headers: { 'x-goog-api-key': key }, payload: JSON.stringify(cuerpo)
      });
      var c = r.getResponseCode();
      lineas.push(m + (buscar ? ' + búsqueda' : '') + ': ' + (c === 200 ? 'OK' : c + ' ' + mensajeGemini_(r.getContentText())));
    });
  });
  aviso_(lineas.join('\n'));
}
