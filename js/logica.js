// Lógica pura (sin DOM): stock, formatos y avisos. Se prueba con node --test.

export function clave(pid, lote) { return pid + '|' + (lote || ''); }

/** Stock del servidor + movimientos pendientes en la cola del móvil. */
export function stockActual(stockServidor, cola, movimientosServidor = []) {
  const st = new Map();
  for (const s of stockServidor || []) st.set(clave(s.producto_id, s.lote), { ...s, unidades: Number(s.unidades) || 0 });
  const porUuid = new Map(movimientosServidor.map((m) => [m.uuid, m]));
  for (const op of cola || []) {
    if (op.tipo !== 'movimiento') continue;
    const m = op.movimiento;
    let pid = m.producto_id, lote = m.lote || '', cad = m.caducidad || '', n = 0;
    if (m.tipo === 'ENTRADA') n = Number(m.cantidad);
    else if (m.tipo === 'SALIDA') n = -Number(m.cantidad);
    else if (m.tipo === 'AJUSTE') n = Number(m.cantidad);
    else if (m.tipo === 'ANULACION') {
      const o = porUuid.get(m.anula_uuid);
      if (!o) continue;
      pid = o.producto_id; lote = o.lote || ''; cad = o.caducidad || ''; n = -Number(o.cantidad);
    }
    const k = clave(pid, lote);
    const prev = st.get(k) || { producto_id: pid, lote, caducidad: cad, unidades: 0 };
    prev.unidades += n;
    if (!prev.caducidad && cad) prev.caducidad = cad;
    st.set(k, prev);
  }
  return st;
}

export function lotesDe(st, pid) {
  return [...st.values()]
    .filter((s) => s.producto_id === pid && s.unidades > 1e-9)
    .sort((a, b) => (a.caducidad || '9999').localeCompare(b.caducidad || '9999') || a.lote.localeCompare(b.lote));
}

export function totalDe(st, pid) {
  let t = 0;
  for (const s of st.values()) if (s.producto_id === pid) t += s.unidades;
  return Math.round(t * 1000) / 1000;
}

/** «2 env. + 34 ud.», «8 envases», «3 ud.» */
export function formatoCantidad(unidades, prod) {
  const u = Math.round((Number(unidades) || 0) * 1000) / 1000;
  const uds = Math.max(1, Number(prod?.uds_envase) || 1);
  if (prod?.modo === 'abrir' && uds === 1) return `${u} ${u === 1 ? 'envase' : 'envases'}`;
  if (uds === 1) return `${u} ud.`;
  const env = Math.floor(u / uds);
  const resto = Math.round((u - env * uds) * 1000) / 1000;
  if (!env) return `${resto} ud.`;
  if (!resto) return `${env} env. (${u} ud.)`;
  return `${env} env. + ${resto} ud.`;
}

export function diasEntre(a, b) {
  return Math.round((new Date(b) - new Date(a)) / 86400000);
}

/** Fecha de caducidad comparable: «2028-02» → último día del mes. */
export function finCaducidad(cad) {
  if (!cad) return '';
  if (/^\d{4}-\d{2}$/.test(cad)) {
    const [a, m] = cad.split('-').map(Number);
    return `${cad}-${String(new Date(Date.UTC(a, m, 0)).getUTCDate()).padStart(2, '0')}`;
  }
  return cad;
}

export function avisos(datos, st, hoyIso, diasCad = 60) {
  const prods = new Map((datos.productos || []).map((p) => [p.id, p]));
  const bajoMinimo = [];
  for (const p of prods.values()) {
    if (!p.activo || !(p.minimo > 0)) continue;
    const t = totalDe(st, p.id);
    if (t < p.minimo) bajoMinimo.push({ producto: p, total: t });
  }
  const caducan = [];
  for (const s of st.values()) {
    if (s.unidades <= 1e-9 || !s.caducidad) continue;
    const d = diasEntre(hoyIso, finCaducidad(s.caducidad));
    if (d <= diasCad) caducan.push({ ...s, producto: prods.get(s.producto_id), dias: d });
  }
  caducan.sort((a, b) => a.dias - b.dias);
  const repos = new Map((datos.reposiciones || []).map((r) => [r.gabinete + '|' + r.producto_id, r]));
  const gabinetes = [];
  for (const d of datos.dotacion || []) {
    const r = repos.get(d.gabinete + '|' + d.producto_id);
    const dias = r ? diasEntre(r.fecha.slice(0, 10), hoyIso) : null;
    if (d.dias_revision > 0 && (dias === null || dias >= d.dias_revision)) gabinetes.push({ ...d, producto: prods.get(d.producto_id), dias });
  }
  const copiaDias = datos.ultimaCopia ? diasEntre(datos.ultimaCopia.slice(0, 10), hoyIso) : null;
  return { bajoMinimo, caducan, gabinetes, copiaDias, copiaVieja: copiaDias === null || copiaDias > 8 };
}

export function nuevoIdProducto() {
  const r = (globalThis.crypto?.randomUUID?.() || String(Math.random()).slice(2) + Date.now()).replace(/-/g, '');
  return 'P-' + r.slice(0, 12).toUpperCase();
}

export function nuevoUuid() {
  if (globalThis.crypto?.randomUUID) return globalThis.crypto.randomUUID();
  return 'u' + Date.now().toString(36) + Math.random().toString(36).slice(2, 12);
}

/** Unidades que descuenta una salida según el modo del producto. */
export function unidadesSalida(prod, cantidadIntroducida) {
  const n = Number(cantidadIntroducida) || 0;
  return prod.modo === 'abrir' ? n * Math.max(1, Number(prod.uds_envase) || 1) : n;
}

export function normaliza(t) {
  return String(t || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

/* ───────────── Posibles duplicados (como los contactos del móvil) ───────────── */
const VACIAS = new Set(['de', 'la', 'el', 'los', 'las', 'y', 'con', 'para', 'sin', 'en', 'por', 'a', 'u', 'ud', 'uds', 'un', 'una', 'x', 'c', 'talla', 'del', 'al', 'caja', 'pack', 'envase']);

const TALLAS = new Set(['xs', 's', 'm', 'l', 'xl', 'xxl', 'pequeno', 'mediano', 'grande', 'corta', 'larga', 'fino', 'extrafino', 'grueso', 'blanco', 'blanca', 'amarillo', 'amarilla', 'verde', 'azul', 'rojo', 'roja', 'negro', 'negra', 'rosa', 'gris', 'naranja', 'transparente', 'universal']);
const esVariante = (t) => /\d/.test(t) || TALLAS.has(t) || /^[a-d][1-4](\.5)?$/.test(t);
/** Talla, medida, color o número distintos en cada lado → variantes, no duplicados. */
export function sonVariantes(a, b) {
  const soloA = [...a].filter((t) => !b.has(t) && esVariante(t));
  const soloB = [...b].filter((t) => !a.has(t) && esVariante(t));
  return soloA.length > 0 && soloB.length > 0;
}

export function fichas(texto) {
  return new Set(normaliza(texto).replace(/[^a-z0-9.,]+/g, ' ')
    .replace(/(\d)\s+(l|lt|ml|g|gr|mm|cm|kg|m)\b/g, '$1$2').replace(/(\d)lt\b/g, '$1l').replace(/(\d)gr\b/g, '$1g')
    .split(' ')
    .map((t) => t.replace(/^[.,]+|[.,]+$/g, ''))
    .filter((t) => t && !VACIAS.has(t) && (t.length > 1 || /\d/.test(t) || TALLAS.has(t))));
}

export function parecido(a, b) {
  let comun = 0;
  for (const t of a) if (b.has(t)) comun++;
  const union = a.size + b.size - comun;
  return { indice: union ? comun / union : 0, comun };
}

/**
 * Parejas de productos activos que podrían ser el mismo. Se descartan las marcadas «no son el mismo».
 * Devuelve [{a, b, indice, motivo}] ordenado de más a menos parecido.
 */
export function posiblesDuplicados(productos, catalogo = [], noDuplicados = [], max = 60) {
  const activos = productos.filter((p) => p.activo);
  const descartes = new Set(noDuplicados.map(([x, y]) => [x, y].sort().join('|')));
  const refs = new Map();
  for (const c of catalogo) {
    if (!c.producto_id) continue;
    if (!refs.has(c.producto_id)) refs.set(c.producto_id, []);
    refs.get(c.producto_id).push(c.descripcion || '');
  }
  const tok = new Map(activos.map((p) => [p.id, fichas([p.nombre, ...(refs.get(p.id) || [])].join(' '))]));
  const indice = new Map();
  for (const p of activos) for (const t of tok.get(p.id)) { if (!indice.has(t)) indice.set(t, []); indice.get(t).push(p); }
  const vistos = new Set(), res = [];
  const porId = new Map(activos.map((p) => [p.id, p]));
  for (const p of activos) {
    const candidatos = new Set();
    for (const t of tok.get(p.id)) { const l = indice.get(t); if (l && l.length < 80) l.forEach((q) => candidatos.add(q.id)); }
    for (const qid of candidatos) {
      if (qid === p.id) continue;
      const clave = [p.id, qid].sort().join('|');
      if (vistos.has(clave) || descartes.has(clave)) continue;
      vistos.add(clave);
      const q = porId.get(qid);
      const mismaRef = p.ref_proveedor && q.ref_proveedor && normaliza(p.ref_proveedor).replace(/\s/g, '') === normaliza(q.ref_proveedor).replace(/\s/g, '');
      const { indice: s, comun } = parecido(tok.get(p.id), tok.get(qid));
      if (!mismaRef && sonVariantes(tok.get(p.id), tok.get(qid))) continue;
      if (mismaRef || (s >= 0.5 && comun >= 2)) res.push({ a: p, b: q, indice: mismaRef ? 1 : s, motivo: mismaRef ? 'misma referencia' : 'nombre parecido' });
    }
  }
  return res.sort((x, y) => y.indice - x.indice).slice(0, max);
}

/** Productos más parecidos a uno dado (para «¿es alguno de estos?»). */
export function parecidos(producto, productos, n = 3) {
  const a = fichas(producto.nombre);
  return productos.filter((p) => p.activo && p.id !== producto.id)
    .map((p) => ({ p, s: parecido(a, fichas(p.nombre)) }))
    .filter((x) => x.s.comun >= 2 && x.s.indice >= 0.3 && !sonVariantes(a, fichas(x.p.nombre)))
    .sort((x, y) => y.s.indice - x.s.indice).slice(0, n).map((x) => x.p);
}

/* ───────────── Buscar por el texto leído en la caja (sin código de barras) ───────────── */
const soloAlnum = (t) => normaliza(t).replace(/[^a-z0-9]/g, '');

/**
 * Productos que encajan con lo leído en la foto de la caja ({nombre, marca, ref_proveedor}).
 * Cuenta palabras en común con el nombre, el proveedor y las descripciones del catálogo;
 * la referencia igual pesa mucho. Devuelve [{p, puntos, porRef}] de más a menos parecido.
 */
export function coincidenciasTexto(leido, productos, catalogo = [], n = 8) {
  const q = fichas(`${leido.nombre || ''} ${leido.marca || ''}`);
  const ref = soloAlnum(leido.ref_proveedor);
  const textoDe = new Map(), refsDe = new Map();
  for (const c of catalogo) {
    textoDe.set(c.producto_id, (textoDe.get(c.producto_id) || '') + ' ' + (c.descripcion || '') + ' ' + (c.proveedor || ''));
    if (!refsDe.has(c.producto_id)) refsDe.set(c.producto_id, []);
    refsDe.get(c.producto_id).push(soloAlnum(c.ref));
  }
  const peso = (t) => (/\d/.test(t) ? 1.5 : t.length <= 2 ? 0.5 : 1);
  const total = [...q].reduce((s, t) => s + peso(t), 0) || 1;
  const res = [];
  for (const p of productos) {
    if (!p.activo) continue;
    const f = fichas(`${p.nombre} ${p.proveedor || ''} ${textoDe.get(p.id) || ''}`);
    let s = 0;
    for (const t of q) if (f.has(t)) s += peso(t);
    let puntos = s / total;
    if (sonVariantes(q, fichas(p.nombre))) puntos *= 0.6;
    const refs = [soloAlnum(p.ref_proveedor), ...(refsDe.get(p.id) || [])].filter((r) => r.length >= 3);
    const porRef = ref.length >= 3 && refs.some((r) => r === ref || (ref.length >= 5 && (r.endsWith(ref) || ref.endsWith(r))));
    if (porRef) puntos += 2;
    if (porRef || puntos >= 0.3) res.push({ p, puntos, porRef });
  }
  return res.sort((a, b) => b.puntos - a.puntos || a.p.nombre.localeCompare(b.p.nombre, 'es')).slice(0, n);
}

/* ───────────── Pedido pegado desde la web del depósito ───────────── */
const CABECERAS_PEDIDO = /^(precio( unitario)?|cantidad|iva|total|subtotal|repetir pedido|en preparaci[oó]n:?|enviado:?|entregado:?|pendiente:?|producto|art[ií]culo|descripci[oó]n)$/i;
const numeroEs = (t) => Number(String(t).replace(/\./g, '').replace(',', '.'));

/**
 * Lee el texto copiado de un pedido (formato tipo Dentaltix: descripción, «Referencia:», precio, «2u.», IVA).
 * Ignora «Otros servicios» (portes, catálogos). Devuelve {numero, fecha, lineas:[{linea, descripcion, ref, precio, cantidad, iva}]}.
 */
export function leerPedido(texto) {
  const filas = String(texto || '').replace(/\r/g, '').split(/\n|\t/).map((x) => x.replace(/\s+/g, ' ').trim()).filter(Boolean);
  let numero = '', fecha = '';
  const lineas = [];
  let desc = [];
  for (let i = 0; i < filas.length; i++) {
    const f = filas[i];
    if (/^otros servicios/i.test(f)) break;
    let m;
    if (!numero && (m = f.match(/^pedido\s*(?:n[ºo.]*\s*)?:?\s*([A-Z0-9-]{4,})$/i))) { numero = m[1]; desc = []; continue; }
    if (!fecha && (m = f.match(/^fecha\s*:?\s*(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})/i))) { fecha = `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; desc = []; continue; }
    if ((m = f.match(/^ref(?:erencia)?\.?\s*:?\s*(\S.*)$/i))) {
      const l = { linea: lineas.length + 1, descripcion: desc.join(' ').slice(0, 200), ref: m[1].trim(), precio: 0, cantidad: 1, iva: '' };
      for (let j = i + 1; j < Math.min(filas.length, i + 6); j++) {
        const g = filas[j];
        let k;
        if ((k = g.match(/^(-?[\d.]+,\d{1,2})\s*€$/)) && !l._precio) { l.precio = numeroEs(k[1]); l._precio = true; i = j; continue; }
        if ((k = g.match(/^(\d+(?:,\d+)?)\s*(?:u|ud|uds|unid\.?|unidades)\.?$/i))) { l.cantidad = numeroEs(k[1]); i = j; continue; }
        if ((k = g.match(/^(\d{1,2})\s*%$/))) { l.iva = k[1]; i = j; continue; }
        break;
      }
      delete l._precio;
      if (l.descripcion) lineas.push(l);
      desc = [];
      continue;
    }
    if (CABECERAS_PEDIDO.test(f)) continue;
    desc.push(f);
    if (desc.length > 3) desc.shift();
  }
  return { numero, fecha, lineas };
}

const refNorm = (r) => String(r || '').toUpperCase().replace(/\s+/g, '');

/** Depósito al que más se parecen las referencias del pedido (según el catálogo). */
export function proveedorProbable(lineas, catalogo = []) {
  const refs = new Set(lineas.map((l) => refNorm(l.ref)));
  const n = new Map();
  for (const c of catalogo) if (refs.has(refNorm(c.ref))) n.set(c.proveedor, (n.get(c.proveedor) || 0) + 1);
  return [...n.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || '';
}

/**
 * Cada línea con su producto: «ref» (misma referencia en ese depósito o en la ficha: seguro),
 * «parecido» (sugerencia por el nombre; hay que confirmarla) o «nuevo».
 */
export function emparejarPedido(lineas, proveedor, productos, catalogo = []) {
  const activos = new Map(productos.filter((p) => p.activo).map((p) => [p.id, p]));
  const porRef = new Map(), porRefFicha = new Map();
  const prov = normaliza(proveedor).trim();
  for (const c of catalogo) if (normaliza(c.proveedor).trim() === prov && activos.has(c.producto_id)) porRef.set(refNorm(c.ref), c.producto_id);
  for (const p of activos.values()) if (p.ref_proveedor && (!p.proveedor || normaliza(p.proveedor).trim() === prov)) porRefFicha.set(refNorm(p.ref_proveedor), p.id);
  return lineas.map((l) => {
    const r = refNorm(l.ref);
    const pid = porRef.get(r) || porRefFicha.get(r);
    if (pid) return { ...l, producto_id: pid, estado: 'ref', sugerencias: [] };
    const sug = coincidenciasTexto({ nombre: l.descripcion }, [...activos.values()], catalogo, 5).filter((x) => x.puntos >= 0.35).map((x) => x.p.id);
    return { ...l, producto_id: '', estado: sug.length ? 'parecido' : 'nuevo', sugerencias: sug };
  });
}
