/**
 * Internal Requests Tests (/peticiones)
 * The support form (a trial: checked and shown back, sent nowhere) and the link to a Microsoft Form
 */

require('./helpers/env');

const test = require('ava').default;
const request = require('supertest');
const { stubAxios, stubMsal, signInAgent, getCsrfToken, auditEvents } = require('./helpers/stubs');

const http = stubAxios();
stubMsal();

const appConfig = require('../config/appConfig');
const app = require('../server');

let agent;
let csrf;
const post = (body) => agent.post('/peticiones/soporte').set('X-CSRF-Token', csrf).type('form').send(body);
const VALID = {
  category: 'Impresora o escáner', impact: 'No puedo trabajar', subject: 'La impresora no imprime',
  description: 'Sale «error 5» al imprimir desde ayer', device: 'PC-0123', location: 'Casa Consistorial, despacho 2', phone: '947500100',
};

test.before(async () => {
  agent = request.agent(app);
  await signInAgent(agent);
  csrf = await getCsrfToken(agent);
});
test.after.always(() => agent.get('/auth/logout'));
test.beforeEach(() => {
  http.reset();
  appConfig.INTERNAL_REQUESTS_FORM_URL = '';
  http.on('get', '/v1.0/me', { displayName: 'Alice Example', mail: 'alice@arandadeduero.es', department: 'Urbanismo', officeLocation: 'Casa Consistorial, despacho 2', businessPhones: ['947500100'] });
});

test.serial('GET /peticiones shows the support form with the profile filled in, and says it is a trial', async t => {
  const res = await agent.get('/peticiones').expect(200);

  t.regex(res.text, /Soporte informático/);
  t.regex(res.text, /todavía no se envía a nadie/);
  // The person is not typed: it comes from the profile
  t.regex(res.text, /Alice Example<\/strong> \(alice@arandadeduero\.es, Urbanismo\)/);
  t.regex(res.text, /name="location" value="Casa Consistorial, despacho 2"/);
  t.regex(res.text, /name="phone" value="947500100"/);
  for (const name of ['category', 'impact', 'subject', 'description', 'device']) t.true(res.text.includes(`name="${name}"`), name);
  t.regex(res.text, /<option>Impresora o escáner<\/option>/);
  // The menu has it, and with no Form configured there is no card for it
  const menu = res.text.split('<nav ')[1].split('</nav>')[0];
  t.true(menu.includes('href="/peticiones"'));
  t.notRegex(res.text, /Abrir el formulario/);
});

test.serial('the Microsoft Form is offered only when a web address is configured', async t => {
  appConfig.INTERNAL_REQUESTS_FORM_URL = 'https://forms.example/Pages/ResponsePage.aspx?id=abc';
  const res = await agent.get('/peticiones').expect(200);
  t.regex(res.text, /href="https:\/\/forms\.example\/Pages\/ResponsePage\.aspx\?id=abc" target="_blank" rel="noopener"/);

  for (const bad of ['javascript:alert(1)', 'http://forms.example/x', 'forms.example/x']) {
    appConfig.INTERNAL_REQUESTS_FORM_URL = bad;
    t.notRegex((await agent.get('/peticiones').expect(200)).text, /Abrir el formulario/, bad);
  }
});

test.serial('the page still opens, with an empty form, when the profile cannot be read', async t => {
  http.fail('get', '/v1.0/me', 500);
  const res = await agent.get('/peticiones').expect(200);
  t.regex(res.text, /name="location" value=""/);
});

test.serial('a valid support form is shown back, sent nowhere, and only its kind is audited', async t => {
  const res = await post(VALID).expect(200);

  t.regex(res.text, /todavía no se ha enviado a nadie/);
  t.regex(res.text, /La impresora no imprime/);
  t.regex(res.text, /Sale «error 5» al imprimir desde ayer/);
  t.regex(res.text, /De: <strong>[^<]*alice@arandadeduero\.es/);
  // Nothing leaves: no mail, no write of any kind to Graph
  t.is(http.calls.filter(c => c.method !== 'get').length, 0);
  const event = auditEvents('request.support_test').at(-1);
  t.is(event.kind, 'soporte');
  t.false(JSON.stringify(event).includes('impresora'));
});

test.serial('what is typed is escaped, and empty optional fields are not listed', async t => {
  const res = await post({ ...VALID, description: '<script>alert(1)</script>', device: '', phone: '' }).expect(200);
  t.notRegex(res.text, /<script>alert\(1\)<\/script>/);
  t.regex(res.text, /&lt;script&gt;alert\(1\)&lt;\/script&gt;/);
  t.notRegex(res.text, /Equipo afectado/);
});

test.serial('a form with a required field missing, a bad option or too much text is refused', async t => {
  const events = auditEvents('request.support_test').length;
  t.regex((await post({ ...VALID, subject: '  ' }).expect(400)).text, /Falta «Asunto»/);
  t.regex((await post({ ...VALID, category: 'Inventada' }).expect(400)).text, /no es una opción válida/);
  t.regex((await post({ ...VALID, impact: '' }).expect(400)).text, /Falta «Cuánto te afecta»/);
  t.regex((await post({ ...VALID, description: 'x'.repeat(2001) }).expect(400)).text, /demasiado largo/);
  t.is(auditEvents('request.support_test').length, events);
});

test.serial('without a session, /peticiones sends to the login, and posting needs the CSRF token', async t => {
  t.is((await request(app).get('/peticiones').expect(302)).header.location, '/login');
  await agent.post('/peticiones/soporte').type('form').send(VALID).expect(403);
});
