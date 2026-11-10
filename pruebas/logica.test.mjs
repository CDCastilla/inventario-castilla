import test from 'node:test';
import assert from 'node:assert/strict';
import { stockActual, lotesDe, totalDe, formatoCantidad, avisos, unidadesSalida, finCaducidad } from '../js/logica.js';

const fresas = { id: 'P-FRESA', nombre: 'Fresas', uds_envase: 5, modo: 'unidades', minimo: 4, activo: true };
const guantes = { id: 'P-GUANTE', nombre: 'Guantes S', uds_envase: 1, modo: 'abrir', minimo: 10, activo: true, proveedor: 'Depósito' };

test('stock del servidor + cola pendiente', () => {
  const servidor = [{ producto_id: 'P-GUANTE', lote: '', unidades: 9 }];
  const movs = [{ uuid: 'm1', producto_id: 'P-GUANTE', lote: '', cantidad: 9 }];
  const cola = [
    { tipo: 'movimiento', movimiento: { tipo: 'SALIDA', producto_id: 'P-GUANTE', lote: '', cantidad: 1 } },
    { tipo: 'movimiento', movimiento: { tipo: 'ENTRADA', producto_id: 'P-FRESA', lote: 'L1', caducidad: '2030-01', cantidad: 10 } },
    { tipo: 'movimiento', movimiento: { tipo: 'ANULACION', anula_uuid: 'm1' } },
    { tipo: 'producto', producto: {} }
  ];
  const st = stockActual(servidor, cola, movs);
  assert.equal(totalDe(st, 'P-GUANTE'), -1);
  assert.equal(totalDe(st, 'P-FRESA'), 10);
});

test('lotes ordenados por caducidad (FEFO), sin caducidad al final', () => {
  const st = stockActual([
    { producto_id: 'X', lote: 'B', caducidad: '2030-01', unidades: 1 },
    { producto_id: 'X', lote: 'A', caducidad: '2028-05-01', unidades: 2 },
    { producto_id: 'X', lote: 'C', caducidad: '', unidades: 1 },
    { producto_id: 'X', lote: 'D', caducidad: '2027-01', unidades: 0 }
  ], []);
  assert.deepEqual(lotesDe(st, 'X').map((l) => l.lote), ['A', 'B', 'C']);
});

test('formato de cantidades', () => {
  assert.equal(formatoCantidad(8, guantes), '8 envases');
  assert.equal(formatoCantidad(1, guantes), '1 envase');
  assert.equal(formatoCantidad(13, fresas), '2 env. + 3 ud.');
  assert.equal(formatoCantidad(10, fresas), '2 env. (10 ud.)');
  assert.equal(formatoCantidad(3, fresas), '3 ud.');
  assert.equal(formatoCantidad(84, { uds_envase: 50, modo: 'unidades' }), '1 env. + 34 ud.');
});

test('salida: al abrir descuenta el envase entero', () => {
  assert.equal(unidadesSalida({ modo: 'abrir', uds_envase: 1 }, 2), 2);
  assert.equal(unidadesSalida({ modo: 'abrir', uds_envase: 6 }, 1), 6);
  assert.equal(unidadesSalida(fresas, 2), 2);
});

test('avisos: bajo mínimo, caducidades, gabinetes y copia', () => {
  const datos = {
    productos: [fresas, guantes],
    dotacion: [{ gabinete: 'Gabinete 3', producto_id: 'P-GUANTE', cantidad: 2, dias_revision: 7 }],
    reposiciones: [{ gabinete: 'Gabinete 3', producto_id: 'P-GUANTE', fecha: '2026-11-01T10:00:00Z' }],
    ultimaCopia: '2026-10-30T07:00:00Z'
  };
  const st = stockActual([
    { producto_id: 'P-GUANTE', lote: '', unidades: 8 },
    { producto_id: 'P-FRESA', lote: 'L1', caducidad: '2026-12', unidades: 5 }
  ], []);
  const a = avisos(datos, st, '2026-11-09');
  assert.deepEqual(a.bajoMinimo.map((b) => b.producto.id), ['P-GUANTE']);
  assert.equal(a.caducan.length, 1);
  assert.equal(a.caducan[0].dias, 52);
  assert.equal(a.gabinetes.length, 1);
  assert.equal(a.copiaDias, 10);
  assert.equal(a.copiaVieja, true);
});

test('fin de caducidad por mes', () => {
  assert.equal(finCaducidad('2028-02'), '2028-02-29');
  assert.equal(finCaducidad('2028-02-10'), '2028-02-10');
});

import { posiblesDuplicados, parecidos } from '../js/logica.js';
test('posibles duplicados: nombre parecido, misma referencia y descartes', () => {
  const ps = [
    { id: 'A', nombre: 'Dentasept Tri Enzymatic 5 L', activo: true },
    { id: 'B', nombre: 'DENTASEPT TRI ENZYMATIC 5L ANIOS', activo: true },
    { id: 'C', nombre: 'Guantes látex talla S', activo: true, ref_proveedor: 'GD05NB' },
    { id: 'D', nombre: 'Guantes Santex', activo: true, ref_proveedor: 'gd05nb' },
    { id: 'E', nombre: 'Guantes látex talla M', activo: true },
    { id: 'F', nombre: 'Dentasept Tri Enzymatic 5 L', activo: false }
  ];
  const r = posiblesDuplicados(ps, [], []);
  const pares = r.map((x) => [x.a.id, x.b.id].sort().join(''));
  assert.ok(pares.includes('AB'));
  assert.ok(pares.includes('CD'));
  assert.ok(!pares.some((p) => p.includes('F')), 'los inactivos no salen');
  assert.ok(!pares.includes('CE'), 'tallas distintas no son duplicados');
  const r2 = posiblesDuplicados(ps, [], [['B', 'A']]);
  assert.ok(!r2.map((x) => [x.a.id, x.b.id].sort().join('')).includes('AB'));
  assert.deepEqual(parecidos(ps[0], ps).map((p) => p.id), ['B']);
});

test('coincidenciasTexto: busca por lo leído en la caja (nombre, marca y referencia)', async () => {
  const { coincidenciasTexto } = await import('../js/logica.js');
  const productos = [
    { id: 'A', nombre: 'Colgate duraphat barniz 10 ml', proveedor: 'Broker Dental', ref_proveedor: '091-0533', activo: true },
    { id: 'B', nombre: 'Guantes de Nitrilo sin Polvo 100 uds S', proveedor: 'Bestdent', activo: true },
    { id: 'C', nombre: 'Guantes de Nitrilo sin Polvo 100 uds M', proveedor: 'Bestdent', activo: true },
    { id: 'D', nombre: 'Elite HD+ light rápida', proveedor: 'Proclinic', ref_proveedor: '4376', activo: true },
    { id: 'E', nombre: 'Producto retirado duraphat', activo: false }
  ];
  const catalogo = [{ producto_id: 'A', proveedor: 'Dentaltix', ref: '09VA019', descripcion: 'Duraphat: Barniz de Flúor Desensibilizante' }];
  const r1 = coincidenciasTexto({ nombre: 'Duraphat barniz de flúor', marca: 'Colgate' }, productos, catalogo);
  assert.equal(r1[0].p.id, 'A');
  assert.ok(!r1.some((x) => x.p.id === 'E'));
  const r2 = coincidenciasTexto({ nombre: 'Guantes nitrilo talla M', marca: 'Bestdent' }, productos, catalogo);
  assert.equal(r2[0].p.id, 'C');
  const r3 = coincidenciasTexto({ nombre: 'Silicona', marca: 'Zhermack', ref_proveedor: '09-VA-019' }, productos, catalogo);
  assert.equal(r3[0].p.id, 'A'); assert.ok(r3[0].porRef);
});

test('leerPedido: pedido pegado de Dentaltix', async () => {
  const { leerPedido, proveedorProbable, emparejarPedido } = await import('../js/logica.js');
  const { readFileSync } = await import('node:fs');
  const txt = readFileSync(new URL('./pedido-dentaltix.txt', import.meta.url), 'utf8');
  const p = leerPedido(txt);
  assert.equal(p.numero, '10169186');
  assert.equal(p.fecha, '2026-11-10');
  assert.equal(p.lineas.length, 18);
  assert.deepEqual(p.lineas[0], { linea: 1, descripcion: 'Bioklair Diario: Limpiador de Sistemas de Aspiración (5 L)', ref: '3109', precio: 30.08, cantidad: 1, iva: '21' });
  assert.equal(p.lineas[11].cantidad, 2);
  assert.equal(p.lineas[17].ref, 'DTX-VI801-012');
  assert.equal(p.lineas[17].precio, 0);
  assert.ok(!p.lineas.some((l) => l.ref === 'SHIPPING'));
  // Mismo pedido copiado con tabuladores en una fila
  const p2 = leerPedido('Pedido 123456\nGasas estériles\tReferencia: 270712\t36,85 €\t3u.\t10%');
  assert.equal(p2.lineas[0].cantidad, 3); assert.equal(p2.lineas[0].ref, '270712');

  const catalogo = [
    { proveedor: 'Dentaltix (Dentared)', ref: 'DTX-VI92103', producto_id: 'A3', descripcion: 'Composite Fluido (Jeringa de 1 ml) - A3' },
    { proveedor: 'Dentaltix (Dentared)', ref: 'DTX-VO2269', producto_id: 'PB', descripcion: 'Profluorid Varnish Bubble gum' },
    { proveedor: 'Proclinic', ref: '3109', producto_id: 'OTRO', descripcion: 'Otra cosa' }
  ];
  const productos = [
    { id: 'A3', nombre: 'Composite Fluido (Jeringa de 1 ml) - A3', activo: true },
    { id: 'PB', nombre: 'Profluorid Varnish + BioMin barniz fluoruro Bubble gum', activo: true },
    { id: 'OTRO', nombre: 'Otra cosa', activo: true }
  ];
  const prov = proveedorProbable(p.lineas, catalogo);
  assert.equal(prov, 'Dentaltix (Dentared)');
  const e = emparejarPedido(p.lineas, prov, productos, catalogo);
  assert.equal(e[11].estado, 'ref'); assert.equal(e[11].producto_id, 'A3');
  assert.equal(e[0].estado, 'nuevo'); // la ref 3109 es de otro depósito
  assert.equal(e[8].estado, 'parecido'); assert.equal(e[8].producto_id, ''); // Menta: se sugiere, no se da por hecho
  assert.deepEqual(e[8].sugerencias, ['PB']);
});
