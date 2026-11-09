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
