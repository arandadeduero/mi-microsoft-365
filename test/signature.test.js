/**
 * Signature helpers (/firma): which concejalía a profile points to, and the
 * generated HTML. Pure: no server, no network.
 */

require('./helpers/env');

const test = require('ava').default;
const { CONCEJALIAS, findConcejaliaKey, buildSignatureHtml } = require('../services/signature');

const PERSON = { name: 'Ana García', position: 'Técnica', email: 'ana@arandadeduero.es', phone: '947 500 100' };

test('the department finds its concejalía ignoring case, accents and "Concejalía de"', t => {
  t.is(findConcejaliaKey('urbanismo', 'Técnica'), 'urbanismo');
  t.is(findConcejaliaKey('Educación', 'Técnica'), 'educacion');
  t.is(findConcejaliaKey('EDUCACION', 'Técnica'), 'educacion');
  t.is(findConcejaliaKey('Concejalía de Cultura', 'Técnica'), 'cultura');
  t.is(findConcejaliaKey('Informática', 'Técnico'), 'informatica');
});

test('a partial match counts only when a single concejalía fits', t => {
  // "Cultura" is itself a concejalía, so the exact one wins over its five services
  t.is(findConcejaliaKey('Cultura', 'Técnica'), 'cultura');
  // "Biblioteca" appears only in one name
  t.is(findConcejaliaKey('Biblioteca', 'Técnica'), 'cultura-biblioteca');
  // "Tesorería" fits several: no guess
  t.is(findConcejaliaKey('Tesorería', 'Técnica'), null);
  t.is(findConcejaliaKey('', 'Técnica'), null);
  t.is(findConcejaliaKey('Nada que ver', 'Técnica'), null);
});

test('a mayor without a department gets the general logo, a councillor is left to choose', t => {
  t.is(findConcejaliaKey('', 'Alcalde'), 'alcaldia');
  t.is(findConcejaliaKey('', 'Alcaldesa'), 'alcaldia');
  t.is(findConcejaliaKey('', 'Concejal de Deportes'), null);
});

test('the signature has the person, the concejalía and the logo from the address given', t => {
  const html = buildSignatureHtml(PERSON, 'urbanismo', 'https://logos.example/');
  t.true(html.includes('Ana García'));
  t.true(html.includes('Técnica'));
  t.true(html.includes('Concejalía de Urbanismo'));
  t.true(html.includes('href="mailto:ana@arandadeduero.es"'));
  t.true(html.includes('📱 947 500 100'));
  t.true(html.includes('src="https://logos.example/blanco/urbanismo.png"'));
});

test('no phone, no phone row; mayors and councillors show no concejalía line', t => {
  t.false(buildSignatureHtml({ ...PERSON, phone: '' }, 'urbanismo', '/x/').includes('📱'));
  for (const position of ['Alcalde', 'Concejala de Cultura']) {
    const html = buildSignatureHtml({ ...PERSON, position }, 'cultura', '/x/');
    t.false(html.includes('font-weight:600'), position);
  }
  const general = buildSignatureHtml({ ...PERSON, position: 'Alcalde' }, 'alcaldia', '/x/');
  t.true(general.includes('src="/x/colores/ayuntamiento.png"'));
});

test('what comes from the profile is escaped', t => {
  const html = buildSignatureHtml({ ...PERSON, name: '<script>alert(1)</script>', position: 'A & B' }, 'urbanismo', '/x/');
  t.false(html.includes('<script>'));
  t.true(html.includes('&lt;script&gt;'));
  t.true(html.includes('A &amp; B'));
});
