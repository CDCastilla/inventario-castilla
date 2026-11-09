import test from 'node:test';
import assert from 'node:assert/strict';
import { leerCodigo, fechaGs1, digitoControlValido } from '../js/gs1.js';

const GS = '\u001d';

test('Meganest: DataMatrix con serie, lote, caducidad (día 00) y código nacional', () => {
  const r = leerCodigo(']d2010843700869204321414872528838' + GS + '10D05275A' + GS + '17271000' + GS + '7126477199', 'data_matrix');
  assert.equal(r.tipo, 'gs1');
  assert.equal(r.clave, '08437008692043');
  assert.equal(r.serie, '414872528838');
  assert.equal(r.lote, 'D05275A');
  assert.equal(r.caducidad, '2027-10-31');
});

test('Implante Klockner: (01)(17)(10) sin separadores tras campos fijos', () => {
  const r = leerCodigo('010843545163380417291205' + '1069125-2156', 'data_matrix');
  assert.equal(r.clave, '08435451633804');
  assert.equal(r.caducidad, '2029-12-05');
  assert.equal(r.lote, '69125-2156');
});

test('JD Zygoma: (01)(10) GS (11)(17)', () => {
  const r = leerCodigo(GS + '0108056591012580' + '100611269126' + GS + '11260601' + '17310601');
  assert.equal(r.clave, '08056591012580');
  assert.equal(r.lote, '0611269126');
  assert.equal(r.fabricacion, '2026-06-01');
  assert.equal(r.caducidad, '2031-06-01');
});

test('Texto con paréntesis (MTA Angelus)', () => {
  const r = leerCodigo('(01)07898078658222(17)280320(10)79141');
  assert.equal(r.clave, '07898078658222');
  assert.equal(r.caducidad, '2028-03-20');
  assert.equal(r.lote, '79141');
});

test('EAN-13 de guantes y de garrafa → GTIN-14', () => {
  assert.equal(leerCodigo('8437014559484', 'ean_13').clave, '08437014559484');
  assert.equal(leerCodigo('3597610591532', 'ean_13').clave, '03597610591532');
  assert.equal(leerCodigo('8437014559484', 'ean_13').tipo, 'ean');
});

test('El EAN-13 de una caja y el GTIN de su DataMatrix dan la misma clave', () => {
  const dm = leerCodigo('0108437008692043' + '10ABC', 'data_matrix');
  const ean = leerCodigo('8437008692043', 'ean_13');
  assert.equal(dm.clave, ean.clave);
});

test('Código que no es GS1 ni EAN: se usa tal cual', () => {
  const r = leerCodigo('https://ifu.klockner.dental/x', 'qr_code');
  assert.equal(r.tipo, 'otro');
  assert.equal(r.clave, 'https://ifu.klockner.dental/x');
});

test('Fechas GS1', () => {
  assert.equal(fechaGs1('280200'), '2028-02-29');
  assert.equal(fechaGs1('271300'), '');
  assert.equal(fechaGs1('abc'), '');
});

test('Dígito de control EAN', () => {
  assert.ok(digitoControlValido('8437014559484'));
  assert.ok(digitoControlValido('08435451633804'));
  assert.ok(!digitoControlValido('8437014559485'));
});
