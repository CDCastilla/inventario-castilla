// Lector de códigos: GS1 DataMatrix / GS1-128 (con o sin paréntesis) y EAN/UPC.
// Devuelve { clave, gtin, lote, caducidad, serie, fabricacion, bruto, tipo }.
// `clave` es lo que identifica al producto en el catálogo: el GTIN en 14 dígitos
// si lo hay; si no, el texto leído. El número de serie (21) nunca forma parte de la clave.

const GS = '\u001d';

// AI → longitud fija (número) o variable (null, hasta FNC1/GS o final).
const AIS = {
  '00': 18, '01': 14, '02': 14, '10': null, '11': 6, '12': 6, '13': 6, '15': 6, '16': 6, '17': 6,
  '20': 2, '21': null, '22': null, '30': null, '37': null,
  '240': null, '241': null, '242': null, '243': null, '250': null, '251': null, '253': null, '254': null,
  '400': null, '401': null, '402': 17, '403': null,
  '410': 13, '411': 13, '412': 13, '413': 13, '414': 13, '415': 13, '416': 13,
  '710': null, '711': null, '712': null, '713': null, '714': null, '715': null,
  '8003': null, '8004': null, '8006': 18, '8020': null, '8200': null
};
// AI de 4 cifras con 6 dígitos de valor (medidas 31nn–36nn, 39nn variable).
function longitudAi(ai) {
  if (Object.prototype.hasOwnProperty.call(AIS, ai)) return { ok: true, len: AIS[ai] };
  if (/^3[1-6]\d\d$/.test(ai)) return { ok: true, len: 6 };
  if (/^39\d\d$/.test(ai)) return { ok: true, len: null };
  return { ok: false };
}

function buscarAi(s, i) {
  for (const n of [2, 3, 4]) {
    const ai = s.slice(i, i + n);
    if (ai.length < n || !/^\d+$/.test(ai)) return null;
    const l = longitudAi(ai);
    if (l.ok) return { ai, len: l.len };
  }
  return null;
}

export function fechaGs1(aammdd) {
  if (!/^\d{6}$/.test(aammdd || '')) return '';
  const a = 2000 + Number(aammdd.slice(0, 2));
  const m = Number(aammdd.slice(2, 4));
  let d = Number(aammdd.slice(4, 6));
  if (m < 1 || m > 12) return '';
  if (d === 0) d = new Date(Date.UTC(a, m, 0)).getUTCDate(); // día 00 = último del mes
  return `${a}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

export function gtin14(n) {
  return String(n).padStart(14, '0');
}

export function digitoControlValido(cod) {
  if (!/^\d{8,14}$/.test(cod)) return false;
  const digs = cod.split('').map(Number);
  const control = digs.pop();
  let suma = 0;
  digs.reverse().forEach((d, i) => { suma += d * (i % 2 === 0 ? 3 : 1); });
  return (10 - (suma % 10)) % 10 === control;
}

function parsearParentesis(s) {
  const campos = {};
  const re = /\((\d{2,4})\)([^()]*)/g;
  let m, alguno = false;
  while ((m = re.exec(s))) { campos[m[1]] = m[2].replace(new RegExp(GS, 'g'), '').trim(); alguno = true; }
  return alguno ? campos : null;
}

function parsearCrudo(s) {
  const campos = {};
  let i = 0;
  while (i < s.length) {
    if (s[i] === GS) { i++; continue; }
    const a = buscarAi(s, i);
    if (!a) return Object.keys(campos).length ? campos : null;
    i += a.ai.length;
    let valor;
    if (a.len) {
      valor = s.slice(i, i + a.len);
      i += a.len;
    } else {
      const fin = s.indexOf(GS, i);
      valor = fin < 0 ? s.slice(i) : s.slice(i, fin);
      i = fin < 0 ? s.length : fin + 1;
    }
    campos[a.ai] = valor;
  }
  return campos;
}

export function leerCodigo(bruto, formato = '') {
  let s = String(bruto || '').replace(/\r|\n/g, '').trim();
  // Identificadores de simbología que algunos lectores anteponen.
  s = s.replace(/^\][A-Za-z]\d/, '');
  if (s.startsWith(GS)) s = s.slice(1);
  const res = { clave: '', gtin: '', lote: '', caducidad: '', serie: '', fabricacion: '', bruto: String(bruto || ''), tipo: 'otro', campos: {} };

  // EAN-8 / UPC-A / EAN-13 / GTIN-14 sueltos.
  if (/^\d{8}$|^\d{12,14}$/.test(s)) {
    res.tipo = formato === 'data_matrix' ? 'gs1' : 'ean';
    res.gtin = gtin14(s);
    res.clave = res.gtin;
    return res;
  }

  const campos = s.includes('(') ? parsearParentesis(s) : parsearCrudo(s);
  if (campos && campos['01'] && /^\d{14}$/.test(campos['01'])) {
    res.tipo = 'gs1';
    res.campos = campos;
    res.gtin = campos['01'];
    res.clave = res.gtin;
    res.lote = (campos['10'] || '').trim();
    res.serie = (campos['21'] || '').trim();
    res.caducidad = fechaGs1(campos['17']);
    res.fabricacion = fechaGs1(campos['11']);
    return res;
  }
  res.clave = s;
  return res;
}
