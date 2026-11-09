// Inventario · Clínica Dental Castilla — app (PWA)
import { leerCodigo } from './gs1.js';
import {
  stockActual, lotesDe, totalDe, formatoCantidad, avisos, nuevoIdProducto, nuevoUuid,
  unidadesSalida, normaliza, diasEntre, finCaducidad
} from './logica.js';

const VERSION_APP = '1.1.0';
const $ = (s) => document.querySelector(s);
const vista = $('#vista');

/* ───────────── Estado persistente ───────────── */
const leer = (k, def) => { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : def; } catch { return def; } };
const guardar = (k, v) => { try { localStorage.setItem(k, JSON.stringify(v)); } catch { aviso('No se pudo guardar en el móvil'); } };

let cfg = leer('cfg', { url: '', clave: '' });
let datos = leer('datos', { productos: [], codigos: [], config: { categoria: [], armario: [], destino: [] }, dotacion: [], stock: [], reposiciones: [], movimientos: [] });
let cola = leer('cola', []);
let rechazos = leer('rechazos', []);
let ultimaSync = leer('ultimaSync', '');
let estadoSync = 'ok';

const prods = () => new Map(datos.productos.map((p) => [p.id, p]));
const codigos = () => new Map(datos.codigos.map((c) => [c.codigo, c.producto_id]));
const st = () => stockActual(datos.stock, cola, datos.movimientos);
const hoy = () => new Date().toISOString().slice(0, 10);
const destinos = () => (datos.config.destino || []).map((d) => d.nombre);

/* ───────────── Utilidades de interfaz ───────────── */
function h(v) {
  return String(v ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}
let toastT;
function aviso(t) {
  const el = $('#toast'); el.textContent = t; el.hidden = false;
  clearTimeout(toastT); toastT = setTimeout(() => { el.hidden = true; }, 3200);
}
function fechaCorta(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d)) return iso;
  return d.toLocaleDateString('es-ES', { day: '2-digit', month: '2-digit', year: '2-digit' }) + (iso.length > 10 ? ' ' + d.toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' }) : '');
}
function cad(c) {
  if (!c) return '';
  const [a, m, d] = c.split('-');
  return d ? `${d}/${m}/${a}` : `${m}/${a}`;
}
const icoEscanear = '<svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" aria-hidden="true"><path d="M4 8V5a1 1 0 0 1 1-1h3M16 4h3a1 1 0 0 1 1 1v3M20 16v3a1 1 0 0 1-1 1h-3M8 20H5a1 1 0 0 1-1-1v-3"/><path d="M7 9v6M10 9v6M13 9v6M16 9v6"/></svg>';

/* ───────────── Navegación ───────────── */
const pila = [];
function ir(nombre, params = {}, reemplazar = false) {
  if (reemplazar && pila.length) pila.pop();
  else if (pila.length) history.pushState({ n: pila.length }, '');
  pila.push({ nombre, params });
  pintar();
  window.scrollTo(0, 0);
}
function atras() {
  if (pila.length > 1) history.back();
}
$('#volver').addEventListener('click', atras);
window.addEventListener('popstate', () => {
  if (!$('#escaner').hidden) cerrarEscaner();
  if (pila.length > 1) pila.pop();
  pintar();
});

function pintar() {
  const actual = pila[pila.length - 1] || { nombre: 'inicio', params: {} };
  $('#volver').hidden = pila.length <= 1;
  $('#logo').hidden = pila.length > 1;
  pintarSync();
  const fn = VISTAS[actual.nombre] || VISTAS.inicio;
  fn(actual.params);
}

function titulo(t) { $('#titulo').textContent = t; document.title = t + ' · Inventario'; }

/* ───────────── Sincronización ───────────── */
async function api(accion, extra = {}) {
  if (!cfg.url || !cfg.clave) throw new Error('Falta configurar la URL y la clave en Ajustes');
  const r = await fetch(cfg.url, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ clave: cfg.clave, accion, ...extra })
  });
  if (!r.ok) throw new Error('El servidor respondió ' + r.status);
  const j = await r.json();
  if (!j.ok) throw new Error(j.error || 'Error del servidor');
  return j;
}

function aplicarEstado(j) {
  datos = {
    productos: j.productos, codigos: j.codigos, config: j.config, dotacion: j.dotacion,
    stock: j.stock, reposiciones: j.reposiciones, movimientos: j.movimientos, ultimaCopia: j.ultimaCopia, version: j.version
  };
  guardar('datos', datos);
  ultimaSync = new Date().toISOString();
  guardar('ultimaSync', ultimaSync);
}

let sincronizando = false;
async function sincronizar(silencioso = true) {
  if (sincronizando || !cfg.url) return;
  sincronizando = true; estadoSync = 'enviando'; pintarSync();
  try {
    while (cola.length) {
      const lote = cola.slice(0, 100);
      const j = await api('enviar', { ops: lote });
      const hechos = new Set();
      for (const r of j.resultados || []) {
        hechos.add(r.op);
        if (r.estado === 'rechazado') {
          const op = lote.find((o) => o.op_id === r.op);
          rechazos.unshift({ fecha: new Date().toISOString(), motivo: r.motivo, resumen: resumenOp(op) });
        }
        if (r.avisos && r.avisos.length) aviso(r.avisos.join('. '));
      }
      cola = cola.filter((o) => !hechos.has(o.op_id));
      guardar('cola', cola); guardar('rechazos', rechazos.slice(0, 50));
      aplicarEstado(j);
      if (!hechos.size) break;
    }
    if (!cola.length) aplicarEstado(await api('estado'));
    estadoSync = 'ok';
  } catch (e) {
    estadoSync = 'error';
    if (!silencioso) aviso(e.message);
  } finally {
    sincronizando = false;
    pintar();
  }
}

function resumenOp(op) {
  if (!op) return '';
  if (op.tipo === 'producto') return 'Ficha: ' + op.producto.nombre;
  const m = op.movimiento, p = prods().get(m.producto_id);
  return `${m.tipo} · ${p ? p.nombre : m.producto_id} · ${m.cantidad}${m.lote ? ' · lote ' + m.lote : ''}`;
}

function pintarSync() {
  const b = $('#sync'), t = $('#syncTexto');
  b.classList.remove('pendiente', 'error');
  if (!cfg.url) { t.textContent = 'Sin configurar'; b.classList.add('error'); return; }
  if (estadoSync === 'enviando') { t.textContent = 'Enviando…'; return; }
  if (cola.length) { t.textContent = `${cola.length} sin enviar`; b.classList.add('pendiente'); return; }
  if (estadoSync === 'error') { t.textContent = 'Sin conexión'; b.classList.add('error'); return; }
  t.textContent = 'Al día';
}
$('#sync').addEventListener('click', () => sincronizar(false));

function encolar(op) {
  op.op_id = nuevoUuid();
  cola.push(op);
  guardar('cola', cola);
  sincronizar();
}

/* ───────────── Escáner ───────────── */
let flujo = null, detector = null, alLeer = null, bucle = null;
const FORMATOS = ['data_matrix', 'ean_13', 'ean_8', 'upc_a', 'upc_e', 'code_128', 'qr_code'];

async function abrirEscaner(callback, msg = 'Apunta al código de la caja') {
  alLeer = callback;
  $('#escanerMsg').textContent = msg;
  $('#manualTxt').value = '';
  $('#escaner').hidden = false;
  if (!('BarcodeDetector' in window)) {
    $('#escanerMsg').textContent = 'Este navegador no puede leer códigos: teclea el código o usa Chrome en Android.';
    return;
  }
  try {
    const soportados = await window.BarcodeDetector.getSupportedFormats();
    detector = new window.BarcodeDetector({ formats: FORMATOS.filter((f) => soportados.includes(f)) });
    flujo = await navigator.mediaDevices.getUserMedia({ video: { facingMode: 'environment', width: { ideal: 1920 }, height: { ideal: 1080 } }, audio: false });
    const v = $('#video'); v.srcObject = flujo; await v.play();
    let ultimo = 0;
    const paso = async (t) => {
      if (!flujo) return;
      if (t - ultimo > 150) {
        ultimo = t;
        try {
          const res = await detector.detect(v);
          if (res.length) { const r = res[0]; const foto = fotograma(v); cerrarEscaner(); navigator.vibrate?.(80); alLeer?.(r.rawValue, r.format, foto); return; }
        } catch { /* fotograma no disponible */ }
      }
      bucle = requestAnimationFrame(paso);
    };
    bucle = requestAnimationFrame(paso);
  } catch (e) {
    $('#escanerMsg').textContent = 'No se pudo abrir la cámara (' + e.message + '). Puedes teclear el código.';
  }
}
/** JPEG en base64 (sin cabecera) de un vídeo, imagen o fichero, reducido a `max` px. */
function fotograma(fuente, max = 1280) {
  try {
    const w = fuente.videoWidth || fuente.naturalWidth || fuente.width, hh = fuente.videoHeight || fuente.naturalHeight || fuente.height;
    if (!w || !hh) return '';
    const k = Math.min(1, max / Math.max(w, hh));
    const c = document.createElement('canvas'); c.width = Math.round(w * k); c.height = Math.round(hh * k);
    c.getContext('2d').drawImage(fuente, 0, 0, c.width, c.height);
    return c.toDataURL('image/jpeg', 0.82).split(',')[1];
  } catch { return ''; }
}
function fotoDeFichero(file) {
  return new Promise((ok) => {
    const img = new Image();
    img.onload = () => { ok(fotograma(img)); URL.revokeObjectURL(img.src); };
    img.onerror = () => ok('');
    img.src = URL.createObjectURL(file);
  });
}

function cerrarEscaner() {
  cancelAnimationFrame(bucle);
  if (flujo) flujo.getTracks().forEach((t) => t.stop());
  flujo = null;
  $('#video').srcObject = null;
  $('#escaner').hidden = true;
}
$('#escanerCerrar').addEventListener('click', cerrarEscaner);
$('#manual').addEventListener('submit', (e) => {
  e.preventDefault();
  const v = $('#manualTxt').value.trim();
  if (!v) return;
  cerrarEscaner();
  alLeer?.(v, 'manual');
});
let buscarDesdeEscaner = null;
$('#escanerBuscar').addEventListener('click', () => { cerrarEscaner(); buscarDesdeEscaner?.(); });

/** Escanear para encontrar un producto y abrir la vista indicada (ver / entrada / salida). */
function escanearPara(intencion) {
  buscarDesdeEscaner = () => ir('stock', { elegir: intencion });
  abrirEscaner((raw, formato, foto) => {
    const lec = leerCodigo(raw, formato);
    const pid = codigos().get(lec.clave);
    if (pid && prods().has(pid)) abrirProducto(pid, intencion, lec);
    else ir('codigoNuevo', { lec, intencion, fotos: foto ? [foto] : [] });
  });
}

function volverAProducto(pid) {
  const prev = pila[pila.length - 2];
  if (prev && prev.nombre === 'producto' && prev.params.pid === pid) { prev.params.lec = null; atras(); }
  else ir('producto', { pid }, true);
}

function abrirProducto(pid, intencion, lec = null, reemplazar = false) {
  if (intencion === 'entrada') ir('entrada', { pid, lec }, reemplazar);
  else if (intencion === 'salida') ir('salida', { pid, lec }, reemplazar);
  else ir('producto', { pid, lec }, reemplazar);
}

/* ───────────── Vistas ───────────── */
const VISTAS = {};

VISTAS.inicio = () => {
  titulo('Inventario');
  const s = st();
  const a = avisos(datos, s, hoy());
  const filas = [];
  if (!cfg.url) filas.push(`<button class="aviso-fila" data-ir="ajustes"><span>Configura la conexión con la hoja</span><span class="chip error">Ajustes</span></button>`);
  if (a.bajoMinimo.length) filas.push(`<button class="aviso-fila" data-ir="pedido"><span>Bajo mínimo · pedir</span><span class="chip aviso">${a.bajoMinimo.length}</span></button>`);
  if (a.caducan.length) filas.push(`<button class="aviso-fila" data-filtro="caduca"><span>Caducan en 60 días</span><span class="chip aviso">${a.caducan.length}</span></button>`);
  if (a.gabinetes.length) filas.push(`<button class="aviso-fila" data-ir="gabinetes"><span>Gabinetes · revisar dotación</span><span class="chip">${a.gabinetes.length}</span></button>`);
  if (cola.length) filas.push(`<button class="aviso-fila" data-sync><span>Movimientos sin enviar</span><span class="chip aviso">${cola.length}</span></button>`);
  if (rechazos.length) filas.push(`<button class="aviso-fila" data-ir="ajustes"><span>Rechazados por el servidor</span><span class="chip error">${rechazos.length}</span></button>`);
  if (cfg.url && datos.version && a.copiaVieja) filas.push(`<div class="aviso-fila"><span>Copia de seguridad ${a.copiaDias === null ? 'sin hacer' : 'de hace ' + a.copiaDias + ' días'}</span><span class="chip aviso">!</span></div>`);

  vista.innerHTML = `
    <button class="btn-escanear" id="escanear">${icoEscanear} Escanear</button>
    <nav class="rejilla" aria-label="Acciones">
      <button class="mosaico" data-intencion="entrada"><b>Entrada</b>Escanear o buscar</button>
      <button class="mosaico" data-intencion="salida"><b>Salida</b>Consumo, gabinete o caja</button>
      <button class="mosaico" data-ir="stock"><b>Stock</b>${datos.productos.filter((p) => p.activo).length} productos</button>
      <button class="mosaico" data-ir="gabinetes"><b>Gabinetes</b>Dotación y reposición</button>
    </nav>
    <section class="lista" aria-label="Avisos">
      <h2>Avisos</h2>
      ${filas.length ? filas.join('') : '<p class="peq">Todo en orden.</p>'}
    </section>
    <div class="fila">
      <button class="btn claro" data-ir="ficha" style="flex:1">Nuevo producto</button>
      <button class="btn claro" data-ir="ajustes" style="flex:1">Ajustes</button>
    </div>
    <p class="peq" style="text-align:center">${ultimaSync ? 'Última sincronización: ' + fechaCorta(ultimaSync) : 'Aún sin sincronizar'}</p>`;
  $('#escanear').onclick = () => escanearPara('ver');
  vista.querySelectorAll('[data-intencion]').forEach((b) => { b.onclick = () => escanearPara(b.dataset.intencion); });
  vista.querySelectorAll('[data-ir]').forEach((b) => { b.onclick = () => ir(b.dataset.ir); });
  vista.querySelectorAll('[data-filtro]').forEach((b) => { b.onclick = () => ir('stock', { filtro: b.dataset.filtro }); });
  vista.querySelectorAll('[data-sync]').forEach((b) => { b.onclick = () => sincronizar(false); });
};

VISTAS.stock = (p) => {
  titulo(p.elegir ? 'Elegir producto' : 'Stock');
  const s = st();
  const estado = { texto: p.texto || '', filtro: p.filtro || 'todo', armario: p.armario || '' };
  const armarios = (datos.config.armario || []).map((a) => a.nombre);
  vista.innerHTML = `
    <label>Buscar<input type="search" id="buscar" placeholder="Nombre, referencia o proveedor" value="${h(estado.texto)}" autocomplete="off"></label>
    <div class="segmentos" role="group" aria-label="Filtro">
      ${[['todo', 'Todo'], ['bajo', 'Bajo mínimo'], ['caduca', 'Caduca pronto'], ['sin', 'Sin stock']].map(([k, t]) => `<button data-f="${k}" aria-pressed="${estado.filtro === k}">${t}</button>`).join('')}
    </div>
    <label>Armario<select id="armario"><option value="">Todos</option>${armarios.map((a) => `<option ${a === estado.armario ? 'selected' : ''}>${h(a)}</option>`).join('')}</select></label>
    <div class="lista" id="res"></div>`;
  const pintarLista = () => {
    const q = normaliza(estado.texto);
    const a = avisos(datos, s, hoy());
    const caducan = new Set(a.caducan.map((c) => c.producto_id));
    const lista = datos.productos.filter((pr) => {
      if (!pr.activo) return false;
      if (estado.armario && pr.armario !== estado.armario) return false;
      if (q && !normaliza(`${pr.nombre} ${pr.ref_proveedor} ${pr.proveedor} ${pr.categoria}`).includes(q)) return false;
      const t = totalDe(s, pr.id);
      if (estado.filtro === 'bajo') return pr.minimo > 0 && t < pr.minimo;
      if (estado.filtro === 'caduca') return caducan.has(pr.id);
      if (estado.filtro === 'sin') return t <= 0;
      return true;
    }).sort((x, y) => x.nombre.localeCompare(y.nombre, 'es'));
    $('#res').innerHTML = lista.length ? lista.map((pr) => {
      const t = totalDe(s, pr.id);
      const lotes = lotesDe(s, pr.id);
      const bajo = pr.minimo > 0 && t < pr.minimo;
      const prox = lotes.find((l) => l.caducidad);
      return `<button class="tarjeta item" data-pid="${h(pr.id)}">
        <span class="entre"><span class="fuerte">${h(pr.nombre)}</span><span class="cifra">${h(formatoCantidad(t, pr))}</span></span>
        <span class="entre peq"><span>${h([pr.categoria, pr.armario].filter(Boolean).join(' · '))}${prox ? ' · cad. ' + cad(prox.caducidad) : ''}</span>${bajo ? `<span class="chip aviso">mín. ${pr.minimo}</span>` : ''}</span>
      </button>`;
    }).join('') : `<p class="vacio">${datos.productos.length ? 'Nada con ese filtro.' : 'Aún no hay productos. Escanea una caja o crea uno nuevo.'}</p>`;
    $('#res').querySelectorAll('[data-pid]').forEach((b) => { b.onclick = () => abrirProducto(b.dataset.pid, p.elegir || 'ver'); });
  };
  $('#buscar').oninput = (e) => { estado.texto = e.target.value; p.texto = estado.texto; pintarLista(); };
  $('#armario').onchange = (e) => { estado.armario = e.target.value; p.armario = estado.armario; pintarLista(); };
  vista.querySelectorAll('[data-f]').forEach((b) => {
    b.onclick = () => {
      estado.filtro = b.dataset.f; p.filtro = estado.filtro;
      vista.querySelectorAll('[data-f]').forEach((x) => x.setAttribute('aria-pressed', x === b));
      pintarLista();
    };
  });
  pintarLista();
};

VISTAS.producto = ({ pid, lec }) => {
  const pr = prods().get(pid);
  if (!pr) { vista.innerHTML = '<p class="vacio">Producto no encontrado.</p>'; return; }
  titulo(pr.nombre);
  const s = st();
  const t = totalDe(s, pid);
  const lotes = lotesDe(s, pid);
  const movs = datos.movimientos.filter((m) => m.producto_id === pid).slice(-25).reverse();
  const pend = cola.filter((o) => o.tipo === 'movimiento' && o.movimiento.producto_id === pid);
  const cods = datos.codigos.filter((c) => c.producto_id === pid).map((c) => c.codigo);
  vista.innerHTML = `
    ${lec ? `<p class="nota info">Escaneado${lec.lote ? ' · lote ' + h(lec.lote) : ''}${lec.caducidad ? ' · cad. ' + cad(lec.caducidad) : ''}</p>` : ''}
    <section class="tarjeta">
      <span class="entre"><h2>${h(pr.nombre)}</h2><span class="cifra">${h(formatoCantidad(t, pr))}</span></span>
      <span class="peq">${h([pr.categoria, pr.armario, pr.proveedor, pr.ref_proveedor && 'REF ' + pr.ref_proveedor].filter(Boolean).join(' · '))}</span>
      <span class="fila"><span class="chip gris">${pr.modo === 'abrir' ? 'Consumido al abrir' : 'Por unidades'}</span>${pr.uds_envase > 1 ? `<span class="chip gris">${pr.uds_envase} ud./envase</span>` : ''}<span class="chip gris">${pr.lote ? 'Con lote' : 'Sin lote'}</span>${pr.minimo ? `<span class="chip ${t < pr.minimo ? 'aviso' : 'gris'}">mín. ${pr.minimo}</span>` : ''}</span>
    </section>
    <div class="rejilla">
      <button class="btn" id="bEntrada">Entrada</button>
      <button class="btn" id="bSalida" ${t <= 0 ? 'disabled' : ''}>Salida</button>
    </div>
    ${pr.lote || lotes.length > 1 ? `<section class="tarjeta"><h2>Lotes</h2>${lotes.length ? lotes.map((l) => `<div class="hist"><span>${h(l.lote || 'sin lote')}${l.caducidad ? ' · cad. ' + cad(l.caducidad) : ''}</span><span class="cifra">${h(formatoCantidad(l.unidades, pr))}</span></div>`).join('') : '<p class="peq">Sin stock.</p>'}</section>` : ''}
    <section class="tarjeta"><h2>Movimientos</h2>
      ${pend.map((o) => `<div class="hist"><span>${h(o.movimiento.tipo)} · ${h(o.movimiento.cantidad)} ${h(o.movimiento.destino || '')}</span><span class="chip aviso">sin enviar</span></div>`).join('')}
      ${movs.length ? movs.map((m) => `<div class="hist ${m.anulado ? 'tachado' : ''}"><span>${fechaCorta(m.fecha)} · ${h(m.tipo)} ${m.cantidad > 0 ? '+' : ''}${h(m.cantidad)}${m.destino ? ' · ' + h(m.destino) : ''}${m.lote ? ' · ' + h(m.lote) : ''}</span>${!m.anulado && m.tipo !== 'ANULACION' ? `<button class="btn peligro" style="min-height:36px;padding:0 10px;font-size:13px" data-anular="${h(m.uuid)}">Anular</button>` : ''}</div>`).join('') : (pend.length ? '' : '<p class="peq">Sin movimientos todavía.</p>')}
    </section>
    <div class="rejilla">
      <button class="btn claro" id="bAjuste">Ajustar stock</button>
      <button class="btn claro" id="bFicha">Editar ficha</button>
    </div>
    <p class="peq">Códigos: ${cods.length ? cods.map(h).join(', ') : 'ninguno'}</p>`;
  $('#bEntrada').onclick = () => ir('entrada', { pid, lec });
  $('#bSalida').onclick = () => ir('salida', { pid, lec });
  $('#bAjuste').onclick = () => ir('ajuste', { pid });
  $('#bFicha').onclick = () => ir('ficha', { pid });
  vista.querySelectorAll('[data-anular]').forEach((b) => {
    b.onclick = () => {
      if (!confirm('¿Anular este movimiento? Quedará registrado como anulado.')) return;
      encolar({ tipo: 'movimiento', movimiento: { uuid: nuevoUuid(), fecha: new Date().toISOString(), tipo: 'ANULACION', anula_uuid: b.dataset.anular, cantidad: 0, origen: 'app' } });
      aviso('Anulación registrada');
      pintar();
    };
  });
};

VISTAS.codigoNuevo = (p) => {
  const { lec, intencion } = p;
  p.fotos = p.fotos || [];
  titulo('Producto nuevo');
  const prop = p.propuesta;
  vista.innerHTML = `
    <section class="tarjeta">
      <h2>Código nuevo</h2>
      <span class="peq">${h(lec.clave)}${lec.lote ? ' · lote ' + h(lec.lote) : ''}${lec.caducidad ? ' · cad. ' + cad(lec.caducidad) : ''}</span>
    </section>
    <section class="tarjeta" id="ia">
      ${p.cargando ? '<h2>Identificando el producto…</h2><span class="peq">Leyendo la foto de la caja.</span>'
        : prop ? `<span class="peq">Propuesta a partir de la foto</span><h2>${h(prop.nombre || '¿?')}</h2>
          <span class="peq">${h([prop.marca, prop.ref_proveedor && 'REF ' + prop.ref_proveedor, prop.uds_envase > 1 && prop.uds_envase + ' ud./envase', prop.categoria].filter(Boolean).join(' · '))}</span>
          ${prop.dudas ? `<span class="nota">${h(prop.dudas)}</span>` : ''}`
        : p.error ? `<h2>No se ha podido identificar</h2><span class="peq">${h(p.error)}</span>`
        : '<h2>Haz una foto de la cara con el nombre</h2><span class="peq">La app lee el nombre, la marca, la referencia y las unidades.</span>'}
    </section>
    ${prop ? '<button class="btn grande" id="usar">Usar estos datos</button>' : ''}
    <label class="btn ${prop ? 'claro' : 'grande'}" style="display:flex;align-items:center;justify-content:center;flex-direction:row;color:${prop ? 'var(--verde)' : 'var(--beige)'}">
      ${p.fotos.length ? 'Añadir foto de la parte delantera' : 'Hacer foto de la caja'}
      <input type="file" accept="image/*" capture="environment" id="foto" class="oculto">
    </label>
    <div class="fila"><button class="btn claro" id="aMano" style="flex:1">Escribir a mano</button><button class="btn claro" id="existe" style="flex:1">Ya existe en la lista</button></div>
    <div class="lista" id="res" hidden></div>`;
  const identificar = async () => {
    p.cargando = true; p.error = ''; pintar();
    try {
      const j = await api('identificar', { imagenes: p.fotos.slice(-3), lectura: { gtin: lec.gtin, lote: lec.lote, caducidad: lec.caducidad, bruto: lec.clave } });
      p.propuesta = j.propuesta;
    } catch (e) { p.error = e.message; }
    p.cargando = false;
    if (pila[pila.length - 1].params === p) pintar();
  };
  if (p.fotos.length && !prop && !p.cargando && !p.error && !p.intentado && cfg.url) { p.intentado = true; identificar(); return; }
  $('#foto').onchange = async (e) => {
    const f = e.target.files && e.target.files[0];
    if (!f) return;
    const b64 = await fotoDeFichero(f);
    if (!b64) { aviso('No se pudo leer la foto'); return; }
    p.fotos.push(b64); p.propuesta = null; identificar();
  };
  if ($('#usar')) $('#usar').onclick = () => ir('ficha', { codigo: lec.clave, lec, intencion, propuesta: prop }, true);
  $('#aMano').onclick = () => ir('ficha', { codigo: lec.clave, lec, intencion }, true);
  $('#existe').onclick = () => {
    const res = $('#res'); res.hidden = false;
    res.innerHTML = '<label>Buscar<input type="search" id="buscar" placeholder="Nombre o referencia" autocomplete="off"></label><div class="lista" id="res2"></div>';
    const pintarLista = (q) => {
      q = normaliza(q);
      const l = datos.productos.filter((x) => x.activo && (!q || normaliza(x.nombre + ' ' + x.ref_proveedor).includes(q))).slice(0, 30);
      $('#res2').innerHTML = l.map((x) => `<button class="tarjeta item" data-pid="${h(x.id)}"><span class="fuerte">${h(x.nombre)}</span><span class="peq">${h(x.ref_proveedor || x.categoria)}</span></button>`).join('');
      $('#res2').querySelectorAll('[data-pid]').forEach((b) => {
        b.onclick = () => {
          const pr = prods().get(b.dataset.pid);
          if (!confirm(`¿Añadir este código a «${pr.nombre}»?`)) return;
          encolar({ tipo: 'producto', producto: { ...pr }, codigos: [lec.clave] });
          datos.codigos.push({ codigo: lec.clave, producto_id: pr.id });
          guardar('datos', datos);
          aviso('Código enlazado');
          abrirProducto(pr.id, intencion, lec, true);
        };
      });
    };
    $('#buscar').oninput = (e) => pintarLista(e.target.value);
    pintarLista('');
    $('#buscar').focus();
  };
};

VISTAS.ficha = ({ pid, codigo, lec, intencion, propuesta }) => {
  const existente = pid ? prods().get(pid) : null;
  titulo(existente ? 'Editar ficha' : 'Nuevo producto');
  const cats = datos.config.categoria || [];
  const arms = (datos.config.armario || []).map((a) => a.nombre);
  const provs = [...new Set(datos.productos.map((p) => p.proveedor).filter(Boolean).concat((datos.config.proveedor || []).map((p) => p.nombre)))];
  const p = existente || { nombre: '', categoria: '', proveedor: '', ref_proveedor: '', uds_envase: 1, modo: 'unidades', lote: false, armario: leer('ultimoArmario', ''), minimo: 0, notas: '', activo: true };
  if (!existente && propuesta) {
    const c = cats.find((x) => x.nombre === propuesta.categoria);
    p.nombre = propuesta.nombre || '';
    p.categoria = c ? c.nombre : '';
    p.proveedor = propuesta.marca || '';
    p.ref_proveedor = propuesta.ref_proveedor || '';
    if (c) { p.modo = c.modo_defecto === 'abrir' ? 'abrir' : 'unidades'; p.lote = c.lote_defecto === 'SI'; }
    if (lec && (lec.lote || lec.caducidad)) p.lote = true;
    const uds = Math.max(1, parseInt(propuesta.uds_envase, 10) || 1);
    if (p.modo === 'abrir') { p.uds_envase = 1; if (uds > 1) p.notas = `Envase de ${uds} u.`; } else p.uds_envase = uds;
  }
  const cods = existente ? datos.codigos.filter((c) => c.producto_id === pid).map((c) => c.codigo) : [];
  const nuevos = codigo ? [codigo] : [];
  vista.innerHTML = `
    ${propuesta ? '<p class="nota info">Datos leídos de la foto: revísalos, elige el armario y guarda.</p>' : ''}
    <form class="campos" id="f">
      <label>Nombre *<input name="nombre" required maxlength="120" value="${h(p.nombre)}" placeholder="Ej.: Guantes látex Santex talla S"></label>
      <label>Categoría<select name="categoria"><option value=""></option>${cats.map((c) => `<option ${c.nombre === p.categoria ? 'selected' : ''}>${h(c.nombre)}</option>`).join('')}</select></label>
      <label>Armario habitual<select name="armario"><option value=""></option>${arms.map((a) => `<option ${a === p.armario ? 'selected' : ''}>${h(a)}</option>`).join('')}</select></label>
      <label>Proveedor<input name="proveedor" list="provs" value="${h(p.proveedor)}"><datalist id="provs">${provs.map((x) => `<option value="${h(x)}">`).join('')}</datalist></label>
      <label>Referencia del proveedor<input name="ref_proveedor" value="${h(p.ref_proveedor)}"></label>
      <fieldset class="tarjeta suave" style="border:0">
        <legend class="oculto">Consumo</legend>
        <label class="check"><input type="radio" name="modo" value="abrir" ${p.modo === 'abrir' ? 'checked' : ''}> Consumido al abrir (el envase entero)</label>
        <label class="check"><input type="radio" name="modo" value="unidades" ${p.modo !== 'abrir' ? 'checked' : ''}> Por unidades (se descuentan las que se usan)</label>
        <label>Unidades por envase<input name="uds_envase" type="number" min="1" step="1" value="${h(p.uds_envase)}"></label>
        <label class="check"><input type="checkbox" name="lote" ${p.lote ? 'checked' : ''}> Controlar lote y caducidad</label>
      </fieldset>
      <label>Mínimo para pedir (en unidades)<input name="minimo" type="number" min="0" step="1" value="${h(p.minimo || '')}" placeholder="0 = sin aviso"></label>
      <label>Notas<textarea name="notas" maxlength="300">${h(p.notas)}</textarea></label>
      ${existente ? `<label class="check"><input type="checkbox" name="activo" ${p.activo ? 'checked' : ''}> Producto activo</label>` : ''}
      <p class="peq" id="cods">Códigos: ${[...cods, ...nuevos].map(h).join(', ') || 'ninguno'}</p>
      <button type="button" class="btn claro" id="addCod">Añadir código escaneando</button>
      <p class="peq">Editar la ficha no mueve stock.</p>
      <button class="btn grande">Guardar</button>
    </form>`;
  const f = $('#f');
  f.categoria.onchange = () => {
    if (existente) return;
    const c = cats.find((x) => x.nombre === f.categoria.value);
    if (!c) return;
    if (c.modo_defecto) f.querySelector(`[name=modo][value=${c.modo_defecto === 'abrir' ? 'abrir' : 'unidades'}]`).checked = true;
    f.lote.checked = c.lote_defecto === 'SI';
  };
  $('#addCod').onclick = () => abrirEscaner((raw, formato) => {
    const l = leerCodigo(raw, formato);
    const otro = codigos().get(l.clave);
    if (otro && otro !== pid) { aviso('Ese código ya es de otro producto'); return; }
    if (![...cods, ...nuevos].includes(l.clave)) nuevos.push(l.clave);
    $('#cods').textContent = 'Códigos: ' + [...cods, ...nuevos].join(', ');
  });
  f.onsubmit = (e) => {
    e.preventDefault();
    const fd = new FormData(f);
    const prod = {
      id: existente ? existente.id : nuevoIdProducto(),
      nombre: String(fd.get('nombre')).trim(),
      categoria: fd.get('categoria') || '', armario: fd.get('armario') || '',
      proveedor: String(fd.get('proveedor') || '').trim(), ref_proveedor: String(fd.get('ref_proveedor') || '').trim(),
      modo: fd.get('modo') === 'abrir' ? 'abrir' : 'unidades',
      uds_envase: Math.max(1, parseInt(fd.get('uds_envase'), 10) || 1),
      lote: !!fd.get('lote'), minimo: Math.max(0, Number(fd.get('minimo')) || 0),
      notas: String(fd.get('notas') || '').trim(), activo: existente ? !!fd.get('activo') : true
    };
    if (!prod.nombre) { aviso('Falta el nombre'); return; }
    if (prod.armario) guardar('ultimoArmario', prod.armario);
    encolar({ tipo: 'producto', producto: prod, codigos: nuevos });
    // Copia local inmediata (funciona sin conexión).
    const i = datos.productos.findIndex((x) => x.id === prod.id);
    if (i >= 0) datos.productos[i] = prod; else datos.productos.push(prod);
    nuevos.forEach((c) => datos.codigos.push({ codigo: c, producto_id: prod.id }));
    guardar('datos', datos);
    aviso('Ficha guardada');
    if (!existente) abrirProducto(prod.id, intencion === 'salida' ? 'ver' : 'entrada', lec, true);
    else atras();
  };
};

function campoCantidad(id, valor, etiqueta) {
  return `<label for="${id}">${etiqueta}</label>
    <div class="cantidad"><button type="button" data-menos="${id}" aria-label="Menos">−</button>
    <input id="${id}" name="${id}" type="number" inputmode="decimal" min="0" step="any" value="${valor}">
    <button type="button" data-mas="${id}" aria-label="Más">+</button></div>`;
}
function activarCantidad(onCambio) {
  vista.querySelectorAll('[data-menos]').forEach((b) => { b.onclick = () => { const i = $('#' + b.dataset.menos); i.value = Math.max(0, (Number(i.value) || 0) - 1); onCambio?.(); }; });
  vista.querySelectorAll('[data-mas]').forEach((b) => { b.onclick = () => { const i = $('#' + b.dataset.mas); i.value = (Number(i.value) || 0) + 1; onCambio?.(); }; });
}

VISTAS.entrada = ({ pid, lec }) => {
  const pr = prods().get(pid);
  if (!pr) { atras(); return; }
  titulo('Entrada');
  vista.innerHTML = `
    <section class="tarjeta"><h2>${h(pr.nombre)}</h2><span class="peq">${h(pr.armario || '')}${pr.uds_envase > 1 ? ' · ' + pr.uds_envase + ' ud. por envase' : ''}</span></section>
    <form class="campos" id="f">
      ${campoCantidad('envases', 1, 'Envases recibidos')}
      <p class="peq" id="eq"></p>
      ${pr.lote ? `<label>Lote *<input name="lote" value="${h(lec?.lote || '')}" required autocomplete="off"></label>
      <label>Caducidad (AAAA-MM o AAAA-MM-DD)<input name="caducidad" value="${h(lec?.caducidad || '')}" placeholder="2028-02" pattern="\\d{4}-\\d{2}(-\\d{2})?" autocomplete="off"></label>
      <button type="button" class="btn claro" id="leerCaja">Leer lote y caducidad de la caja</button>` : ''}
      <label>Nº de factura o albarán<input name="factura" placeholder="Vacío = pendiente de factura" autocomplete="off"></label>
      <label>Nota<input name="nota" maxlength="300" autocomplete="off"></label>
      <button class="btn grande">Registrar entrada</button>
    </form>`;
  const f = $('#f');
  const eq = () => { $('#eq').textContent = pr.uds_envase > 1 ? `= ${(Number(f.envases.value) || 0) * pr.uds_envase} unidades` : ''; };
  activarCantidad(eq); f.envases.oninput = eq; eq();
  if (pr.lote) $('#leerCaja').onclick = () => abrirEscaner((raw, formato) => {
    const l = leerCodigo(raw, formato);
    const otro = codigos().get(l.clave);
    if (otro && otro !== pid) aviso('Ojo: ese código es de otro producto');
    if (l.lote) f.lote.value = l.lote;
    if (l.caducidad) f.caducidad.value = l.caducidad;
    if (!l.lote && !l.caducidad) aviso('El código no trae lote ni caducidad: escríbelos a mano');
  });
  f.onsubmit = (e) => {
    e.preventDefault();
    const env = Number(f.envases.value);
    if (!(env > 0)) { aviso('Indica cuántos envases'); return; }
    const factura = f.factura.value.trim();
    encolar({ tipo: 'movimiento', movimiento: {
      uuid: nuevoUuid(), fecha: new Date().toISOString(), tipo: 'ENTRADA', producto_id: pid,
      lote: pr.lote ? f.lote.value.trim() : '', caducidad: pr.lote ? f.caducidad.value.trim() : '',
      cantidad: env * pr.uds_envase, factura, origen: factura ? 'app' : 'pendiente de factura', nota: f.nota.value.trim()
    } });
    aviso('Entrada registrada');
    volverAProducto(pid);
  };
};

VISTAS.salida = ({ pid, lec }) => {
  const pr = prods().get(pid);
  if (!pr) { atras(); return; }
  titulo('Salida');
  const s = st();
  const lotes = lotesDe(s, pid);
  const total = totalDe(s, pid);
  let lote = lotes.length ? lotes[0].lote : '';
  if (lec?.lote && lotes.some((l) => l.lote === lec.lote)) lote = lec.lote;
  let destino = '';
  const dests = destinos();
  vista.innerHTML = `
    <section class="tarjeta"><h2>${h(pr.nombre)}</h2>
      <span class="peq">${h(pr.armario || '')} · ${pr.modo === 'abrir' ? 'consumido al abrir' : 'por unidades'}</span>
      <span>En stock: <b>${h(formatoCantidad(total, pr))}</b>${pr.minimo ? ' · mínimo ' + pr.minimo : ''}</span>
    </section>
    <form class="campos" id="f">
      ${lotes.length > 1 || pr.lote ? `<label>Lote<select name="lote">${lotes.map((l) => `<option value="${h(l.lote)}" ${l.lote === lote ? 'selected' : ''}>${h(l.lote || 'sin lote')}${l.caducidad ? ' · cad. ' + cad(l.caducidad) : ''} · ${h(formatoCantidad(l.unidades, pr))}</option>`).join('')}</select></label>` : ''}
      ${campoCantidad('cant', 1, pr.modo === 'abrir' ? 'Envases abiertos' : 'Unidades usadas')}
      <p class="peq" id="eq"></p>
      <h2>¿A dónde va?</h2>
      <div class="destinos" role="group" aria-label="Destino">${dests.map((d) => `<button type="button" data-d="${h(d)}" aria-pressed="false">${h(d)}</button>`).join('')}</div>
      <p class="peq" id="dmsg"></p>
      <label>Nota<input name="nota" maxlength="300" autocomplete="off"></label>
      <p class="nota" id="alerta" hidden></p>
      <button class="btn grande" id="ok" disabled>Elige destino</button>
    </form>`;
  const f = $('#f');
  const calc = () => {
    const u = unidadesSalida(pr, f.cant.value);
    const selLote = f.lote ? f.lote.value : lote;
    const disp = (lotes.find((l) => l.lote === selLote) || { unidades: 0 }).unidades;
    $('#eq').textContent = pr.modo === 'abrir' && pr.uds_envase > 1 ? `= ${u} unidades` : '';
    const al = $('#alerta');
    if (u > disp + 1e-9) { al.hidden = false; al.textContent = `Solo hay ${formatoCantidad(disp, pr)} en ese lote.`; }
    else if (pr.minimo && total - u < pr.minimo) { al.hidden = false; al.textContent = `Tras esta salida quedan ${formatoCantidad(total - u, pr)}: pasa a la lista de pedido.`; }
    else al.hidden = true;
    $('#ok').disabled = !destino || !(u > 0) || u > disp + 1e-9;
    $('#ok').textContent = destino ? 'Registrar salida' : 'Elige destino';
  };
  activarCantidad(calc); f.cant.oninput = calc; if (f.lote) f.lote.onchange = calc;
  vista.querySelectorAll('[data-d]').forEach((b) => {
    b.onclick = () => {
      destino = b.dataset.d;
      vista.querySelectorAll('[data-d]').forEach((x) => x.setAttribute('aria-pressed', x === b));
      $('#dmsg').textContent = destino.startsWith('Gabinete') ? `Sale del stock para el pedido y queda anotado en la reposición del ${destino}.` : destino === 'Caja quirúrgica' ? 'Se da por usado al ponerlo en la caja.' : '';
      calc();
    };
  });
  calc();
  f.onsubmit = (e) => {
    e.preventDefault();
    const u = unidadesSalida(pr, f.cant.value);
    const selLote = f.lote ? f.lote.value : lote;
    const l = lotes.find((x) => x.lote === selLote);
    encolar({ tipo: 'movimiento', movimiento: {
      uuid: nuevoUuid(), fecha: new Date().toISOString(), tipo: 'SALIDA', producto_id: pid,
      lote: selLote, caducidad: l?.caducidad || '', cantidad: u, destino, origen: 'app', nota: f.nota.value.trim()
    } });
    aviso('Salida registrada');
    volverAProducto(pid);
  };
};

VISTAS.ajuste = ({ pid }) => {
  const pr = prods().get(pid);
  if (!pr) { atras(); return; }
  titulo('Ajustar stock');
  const s = st();
  const lotes = lotesDe(s, pid);
  vista.innerHTML = `
    <section class="tarjeta"><h2>${h(pr.nombre)}</h2><span class="peq">Escribe lo que hay de verdad. La diferencia queda registrada como ajuste.</span></section>
    <form class="campos" id="f">
      <label>Lote<select name="lote">${lotes.map((l) => `<option value="${h(l.lote)}">${h(l.lote || 'sin lote')} · ahora ${h(formatoCantidad(l.unidades, pr))}</option>`).join('')}<option value="__nuevo">Otro lote / sin stock registrado</option></select></label>
      <label id="nl" hidden>Lote<input name="nuevoLote" autocomplete="off"></label>
      <label>Hay realmente (unidades)<input name="real" type="number" inputmode="decimal" min="0" step="any" required></label>
      <label>Motivo<input name="nota" value="Recuento" maxlength="300"></label>
      <button class="btn grande">Guardar ajuste</button>
    </form>`;
  const f = $('#f');
  f.lote.onchange = () => { $('#nl').hidden = f.lote.value !== '__nuevo'; };
  if (!lotes.length) { f.lote.value = '__nuevo'; $('#nl').hidden = false; }
  f.onsubmit = (e) => {
    e.preventDefault();
    const nuevo = f.lote.value === '__nuevo';
    const lote = nuevo ? f.nuevoLote.value.trim() : f.lote.value;
    if (nuevo && pr.lote && !lote) { aviso('Falta el lote'); return; }
    const actual = (lotes.find((l) => l.lote === lote) || { unidades: 0 }).unidades;
    const dif = Math.round((Number(f.real.value) - actual) * 1000) / 1000;
    if (!dif) { aviso('No hay diferencia: nada que ajustar'); return; }
    encolar({ tipo: 'movimiento', movimiento: { uuid: nuevoUuid(), fecha: new Date().toISOString(), tipo: 'AJUSTE', producto_id: pid, lote, cantidad: dif, origen: 'recuento', nota: f.nota.value.trim() } });
    aviso(`Ajuste ${dif > 0 ? '+' : ''}${dif} registrado`);
    atras();
  };
};

VISTAS.pedido = () => {
  titulo('Lista de pedido');
  const s = st();
  const a = avisos(datos, s, hoy());
  const grupos = new Map();
  for (const b of a.bajoMinimo) {
    const k = b.producto.proveedor || 'Sin proveedor';
    if (!grupos.has(k)) grupos.set(k, []);
    grupos.get(k).push(b);
  }
  vista.innerHTML = grupos.size ? [...grupos.entries()].sort().map(([prov, l]) => `
    <section class="tarjeta"><h2>${h(prov)}</h2>
      ${l.map((b) => `<div class="hist"><span>${h(b.producto.nombre)}${b.producto.ref_proveedor ? ' · REF ' + h(b.producto.ref_proveedor) : ''}</span><span class="cifra">${h(formatoCantidad(b.total, b.producto))} / mín. ${b.producto.minimo}</span></div>`).join('')}
    </section>`).join('') : '<p class="vacio">Nada por debajo del mínimo.</p>';
};

VISTAS.gabinetes = (p) => {
  titulo('Gabinetes');
  const gabs = destinos().filter((d) => d.startsWith('Gabinete'));
  const g = p.g || gabs[0];
  const pm = prods();
  const dot = datos.dotacion.filter((d) => d.gabinete === g);
  const repos = datos.reposiciones.filter((r) => r.gabinete === g).sort((x, y) => y.fecha.localeCompare(x.fecha));
  const reposDe = new Map(repos.map((r) => [r.producto_id, r]));
  vista.innerHTML = `
    <div class="segmentos" role="group" aria-label="Gabinete">${gabs.map((x) => `<button data-g="${h(x)}" aria-pressed="${x === g}">${h(x)}</button>`).join('')}</div>
    <p class="peq">Lo que está en el gabinete ya cuenta como consumido para el pedido. Aquí se ve para reponer.</p>
    <section class="tarjeta"><h2>Dotación</h2>
      ${dot.length ? dot.map((d) => { const r = reposDe.get(d.producto_id); const dias = r ? diasEntre(r.fecha.slice(0, 10), hoy()) : null; const revisar = d.dias_revision > 0 && (dias === null || dias >= d.dias_revision); return `<div class="hist"><span>${h(pm.get(d.producto_id)?.nombre || d.producto_id)} · debe tener ${d.cantidad}<br><span class="peq">${r ? 'repuesto ' + (dias === 0 ? 'hoy' : 'hace ' + dias + ' días') : 'nunca repuesto'}</span></span><span class="chip ${revisar ? 'aviso' : ''}">${revisar ? 'Revisar' : 'Al día'}</span></div>`; }).join('') : '<p class="peq">Dotación por definir (pestaña DOTACION de la hoja; se podrá editar desde la app más adelante).</p>'}
    </section>
    <section class="tarjeta"><h2>Últimas reposiciones</h2>
      ${repos.length ? repos.slice(0, 30).map((r) => `<div class="hist"><span>${h(pm.get(r.producto_id)?.nombre || r.producto_id)}</span><span class="peq">${fechaCorta(r.fecha)}</span></div>`).join('') : '<p class="peq">Aún no se ha subido nada a este gabinete.</p>'}
    </section>
    <button class="btn grande" id="reponer">Reponer escaneando</button>`;
  vista.querySelectorAll('[data-g]').forEach((b) => { b.onclick = () => { p.g = b.dataset.g; pintar(); }; });
  $('#reponer').onclick = () => escanearPara('salida');
};

VISTAS.ajustes = () => {
  titulo('Ajustes');
  vista.innerHTML = `
    <form class="campos" id="f">
      <label>URL de la aplicación web (Apps Script)<input name="url" value="${h(cfg.url)}" placeholder="https://script.google.com/macros/s/…/exec" autocomplete="off" inputmode="url"></label>
      <label>Clave<input name="clave" value="${h(cfg.clave)}" autocomplete="off"></label>
      <button class="btn grande">Guardar y probar</button>
    </form>
    <section class="tarjeta">
      <h2>Estado</h2>
      <span class="peq">App ${VERSION_APP} · servidor ${h(datos.version || '—')}</span>
      <span class="peq">Sin enviar: ${cola.length} · última sincronización: ${ultimaSync ? fechaCorta(ultimaSync) : '—'}</span>
      <span class="peq">Última copia de seguridad: ${datos.ultimaCopia ? fechaCorta(datos.ultimaCopia) : '—'}</span>
    </section>
    <button class="btn borde" id="exportar">Exportar a Excel</button>
    ${rechazos.length ? `<section class="tarjeta"><h2>Rechazados por el servidor</h2>${rechazos.map((r) => `<div class="hist"><span>${h(r.resumen)}<br><span class="peq">${h(r.motivo)} · ${fechaCorta(r.fecha)}</span></span></div>`).join('')}<button class="btn claro" id="limpiar">Borrar esta lista</button></section>` : ''}`;
  $('#f').onsubmit = async (e) => {
    e.preventDefault();
    const fd = new FormData(e.target);
    cfg = { url: String(fd.get('url')).trim(), clave: String(fd.get('clave')).trim() };
    guardar('cfg', cfg);
    try { aplicarEstado(await api('estado')); estadoSync = 'ok'; aviso('Conectado'); sincronizar(); } catch (err) { estadoSync = 'error'; aviso(err.message); }
    pintar();
  };
  $('#exportar').onclick = async () => {
    try {
      aviso('Preparando el Excel…');
      const j = await api('exportar');
      const bin = Uint8Array.from(atob(j.base64), (c) => c.charCodeAt(0));
      const url = URL.createObjectURL(new Blob([bin], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      const a = document.createElement('a'); a.href = url; a.download = j.nombre; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
    } catch (err) { aviso(err.message); }
  };
  if ($('#limpiar')) $('#limpiar').onclick = () => { rechazos = []; guardar('rechazos', rechazos); pintar(); };
};

/* ───────────── Arranque ───────────── */
ir('inicio');
sincronizar();
setInterval(() => { if (navigator.onLine) sincronizar(); }, 60000);
window.addEventListener('online', () => sincronizar());
document.addEventListener('visibilitychange', () => { if (!document.hidden) sincronizar(); });
if ('serviceWorker' in navigator) navigator.serviceWorker.register('sw.js').catch(() => {});

// Para pruebas
window.__inventario = { leerCodigo, finCaducidad };
