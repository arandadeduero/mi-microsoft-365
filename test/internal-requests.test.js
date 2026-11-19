/**
 * Internal Requests Tests (/peticiones)
 * The support form (sent to IT by email from the person's own mailbox) and the link to a Microsoft Form
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
  http.on('post', '/me/sendMail', {});
  http.on('get', '/v1.0/me', { displayName: 'Alice Example', mail: 'alice@arandadeduero.es', department: 'Urbanismo', officeLocation: 'Casa Consistorial, despacho 2', businessPhones: ['947500100'] });
});

test.serial('GET /peticiones shows the support form with the profile filled in', async t => {
  const res = await agent.get('/peticiones').expect(200);

  t.regex(res.text, /Soporte informático/);
  t.notRegex(res.text, /Prueba\./);
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

test.serial('a valid support form goes to IT from my mailbox, and only its kind is audited', async t => {
  const res = await post(VALID).expect(302);
  t.is(res.header.location, '/peticiones?sent=soporte');

  const [call] = http.callsTo('post', '/me/sendMail');
  t.is(http.callsTo('post', '/me/sendMail').length, 1);
  const { message } = call.data;
  t.is(message.toRecipients[0].emailAddress.address, 'admin@example.test');
  t.is(message.subject, '[Mi Microsoft 365] Soporte informático: La impresora no imprime');
  t.is(message.body.contentType, 'Text');
  for (const line of [
    'pide desde el portal: Soporte informático.', 'Departamento: Urbanismo',
    'Qué tipo de problema es: Impresora o escáner', 'Cuánto te afecta: No puedo trabajar',
    'Qué ocurre: Sale «error 5» al imprimir desde ayer', 'Equipo afectado: PC-0123', 'Teléfono para contactarte: 947500100',
  ]) t.true(message.body.content.includes(line), line);

  const event = auditEvents('request.support').at(-1);
  t.is(event.kind, 'soporte');
  t.false(JSON.stringify(event).includes('impresora'));

  t.regex((await agent.get('/peticiones?sent=soporte').expect(200)).text, /Petición enviada\. Informática la recibe en admin@example\.test/);
  t.notRegex((await agent.get('/peticiones').expect(200)).text, /Petición enviada/);
});

test.serial('optional fields left empty are not in the email, and a subject is always one line', async t => {
  await post({ ...VALID, device: '', phone: '', subject: '  Falla\r\nBcc: otra@x.es   mucho  ' }).expect(302);
  const { message } = http.callsTo('post', '/me/sendMail').at(-1).data;
  t.is(message.subject, '[Mi Microsoft 365] Soporte informático: Falla Bcc: otra@x.es mucho');
  t.notRegex(message.subject, /[\r\n]/);
  t.false(message.body.content.includes('Equipo afectado'));
  t.false(message.body.content.includes('Teléfono para contactarte'));
});

test.serial('if the email cannot be sent, the person is told and nothing is audited', async t => {
  http.fail('post', '/me/sendMail', 500);
  const events = auditEvents('request.support').length;
  const res = await post(VALID).expect(502);
  t.regex(res.text, /No se ha podido enviar la petición\. Escribe directamente a admin@example\.test/);
  t.is(auditEvents('request.support').length, events);
});

test.serial('a form with a required field missing, a bad option or too much text is refused', async t => {
  const events = auditEvents('request.support').length;
  t.regex((await post({ ...VALID, subject: '  ' }).expect(400)).text, /Falta «Asunto»/);
  t.regex((await post({ ...VALID, category: 'Inventada' }).expect(400)).text, /no es una opción válida/);
  t.regex((await post({ ...VALID, impact: '' }).expect(400)).text, /Falta «Cuánto te afecta»/);
  t.regex((await post({ ...VALID, description: 'x'.repeat(2001) }).expect(400)).text, /demasiado largo/);
  t.is(auditEvents('request.support').length, events);
  t.is(http.callsTo('post', '/me/sendMail').length, 0);
});

test.serial('without a session, /peticiones sends to the login, and posting needs the CSRF token', async t => {
  t.is((await request(app).get('/peticiones').expect(302)).header.location, '/login');
  await agent.post('/peticiones/soporte').type('form').send(VALID).expect(403);
});
