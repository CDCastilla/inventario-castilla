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
