/**
 * Admin Pages Tests
 * Licenses, accounts to look at, Microsoft's messages, pending requests,
 * settings and the audit export, for administrators only
 */

require('./helpers/env');

const fs = require('fs');
const path = require('path');
const test = require('ava').default;
const request = require('supertest');
const { stubAxios, stubMsal, signInAgent, getCsrfToken, auditEvents, httpError } = require('./helpers/stubs');
const { reviewAccounts, describeLicenseStock } = require('../services/accountReview');
const { resetMessageCache } = require('../services/messageCenterService');
const { resetGraphCaches } = require('../services/graphService');

const http = stubAxios();
const msal = stubMsal();

const permissions = require('../config/permissions');
const appConfig = require('../config/appConfig');
const requestStore = require('../services/requestStore');
const app = require('../server');

const ROLES = '/me/memberOf/microsoft.graph.directoryRole';
const ADMIN = { username: 'admin@arandadeduero.es', name: 'Ada Admin', localAccountId: 'admin-id' };
const USER = { username: 'alice@arandadeduero.es', name: 'Alice Example', localAccountId: 'alice-id' };

async function signInAs(t, account, { admin }) {
  msal.account = account;
  const agent = request.agent(app);
  await signInAgent(agent);
  t.teardown(() => agent.get('/auth/logout'));
  http.on('get', ROLES, { value: admin ? [{ displayName: 'Global Administrator' }] : [] });
  return agent;
}

test.beforeEach(() => {
  http.reset();
  resetMessageCache();
  resetGraphCaches();
  http.on('get', '/v1.0/me', { displayName: 'Someone' });
  permissions.DOMAIN = 'arandadeduero.es';
});

const day = (offset) => new Date(Date.now() - offset * 86400000).toISOString();

test('the review finds idle, unprotected, expired and mis-licensed accounts, and says what it cannot know', t => {
  const users = [
    { id: 'a', displayName: 'Ana', userPrincipalName: 'ana@x', accountEnabled: true, assignedLicenses: [{}], lastPasswordChangeDateTime: day(100), signInActivity: { lastSignInDateTime: day(120) } },
    { id: 'b', displayName: 'Bea', userPrincipalName: 'bea@x', accountEnabled: true, assignedLicenses: [{}], lastPasswordChangeDateTime: day(10), signInActivity: { lastSignInDateTime: day(200), lastNonInteractiveSignInDateTime: day(1) } },
    { id: 'c', displayName: 'Carl', userPrincipalName: 'carl@x', accountEnabled: false, assignedLicenses: [{}] },
    { id: 'd', displayName: 'Dani', userPrincipalName: 'dani@x', accountEnabled: true, assignedLicenses: [] },
    { id: 'g', displayName: 'Guest', userPrincipalName: 'g#EXT#@x', userType: 'Guest', accountEnabled: true, assignedLicenses: [] },
    { id: 'n', displayName: 'Nuevo', userPrincipalName: 'nuevo@x', accountEnabled: true, assignedLicenses: [{}], createdDateTime: day(200) },
  ];
  const mfa = [{ id: 'a', isMfaRegistered: false }, { id: 'b', isMfaRegistered: true }, { id: 'c', isMfaRegistered: false }];

  const review = reviewAccounts(users, { signInsKnown: true, mfa, passwordMaxAgeDays: 90 });

  t.is(review.total, 5);
  t.is(review.guests, 1);
  t.deepEqual(review.idle.map(p => [p.name, p.note]), [['Ana', 'Último inicio de sesión hace 120 días'], ['Nuevo', 'Nunca ha iniciado sesión']]);
  t.deepEqual(review.withoutMfa.map(p => p.name), ['Ana']);
  t.deepEqual(review.expiredPassword.map(p => p.name), ['Ana']);
  t.deepEqual(review.disabledWithLicense.map(p => p.name), ['Carl']);
  t.deepEqual(review.enabledWithoutLicense.map(p => p.name), ['Dani']);

  const blind = reviewAccounts(users, { signInsKnown: false, mfa: null, passwordMaxAgeDays: null });
  t.is(blind.idle, null);
  t.is(blind.withoutMfa, null);
  t.is(blind.expiredPassword, null);
});

test('license stock says what is running out, named from the catalog', t => {
  const stock = describeLicenseStock([
    { skuPartNumber: 'O365_BUSINESS_PREMIUM', prepaidUnits: { enabled: 100 }, consumedUnits: 95 },
    { skuPartNumber: 'FLOW_FREE', prepaidUnits: { enabled: 10000 }, consumedUnits: 3 },
    { skuPartNumber: 'MYSTERY', prepaidUnits: { enabled: 5 }, consumedUnits: 5 },
    { skuPartNumber: 'NONE', prepaidUnits: { enabled: 0 }, consumedUnits: 0 },
  ], { O365_BUSINESS_PREMIUM: 'Microsoft 365 Empresa Estándar' });

  t.deepEqual(stock.map(s => [s.name, s.left, s.low]), [['MYSTERY', 0, true], ['Microsoft 365 Empresa Estándar', 5, true], ['FLOW_FREE', 9997, false]]);
});

test.serial('the admin pages are for administrators only', async t => {
  const user = await signInAs(t, USER, { admin: false });
  for (const path of ['/admin/cuentas', '/admin/licencias', '/admin/avisos', '/admin/solicitudes', '/admin/salas', '/admin/comisiones', '/admin/ajustes', '/admin/audit.csv']) {
    await user.get(path).expect(403);
  }
  t.pass();
});

test.serial('GET /admin/licencias and /admin/cuentas read Graph and show the lists', async t => {
  const admin = await signInAs(t, ADMIN, { admin: true });
  http.on('get', '/v1.0/subscribedSkus', { value: [{ skuPartNumber: 'O365_BUSINESS_PREMIUM', prepaidUnits: { enabled: 50 }, consumedUnits: 49 }] });
  http.on('get', '/v1.0/users', ({ config }) => (config.params.$select.includes('signInActivity')
    ? { value: [{ id: 'a', displayName: 'Ana Idle', userPrincipalName: 'ana@arandadeduero.es', accountEnabled: true, assignedLicenses: [{}], signInActivity: { lastSignInDateTime: day(300) } }] }
    : { value: [] }));
  http.fail('get', '/userRegistrationDetails', 403);

  const licenses = await admin.get('/admin/licencias').expect(200);
  t.true(licenses.text.includes('Microsoft 365 Empresa Estándar'));
  t.true(licenses.text.includes('Se agotan'));

  const accounts = await admin.get('/admin/cuentas').expect(200);
  t.true(accounts.text.includes('Ana Idle'));
  t.true(accounts.text.includes('hace 300 días'));
  t.true(accounts.text.includes('informe de métodos de autenticación necesita'));
  // Read with the app's token, not the administrator's
  t.is(http.callsTo('get', '/v1.0/users')[0].config.headers.Authorization, 'Bearer app-token');
  t.is(http.callsTo('get', '/v1.0/users')[0].config.params.$top, 500);
});

test.serial('GET /admin/cuentas says why the last sign-in cannot be known, in Graph\'s own words', async t => {
  const admin = await signInAs(t, ADMIN, { admin: true });
  http.on('get', '/v1.0/users', ({ config }) => {
    if (config.params.$select.includes('signInActivity')) {
      const error = httpError(403);
      error.response.data.error.code = 'Authentication_RequestFromNonPremiumTenantOrB2CTenant';
      throw error;
    }
    return { value: [{ id: 'a', displayName: 'Ana', userPrincipalName: 'ana@arandadeduero.es', accountEnabled: true, assignedLicenses: [{}] }] };
  });
  http.fail('get', '/userRegistrationDetails', 403);

  const res = await admin.get('/admin/cuentas').expect(200);

  t.true(res.text.includes('sin licencias Entra ID P1 o P2'));
  t.true(res.text.includes('Graph responde Authentication_RequestFromNonPremiumTenantOrB2CTenant'));
  // The rest of the review still runs on the users read without the property
  t.true(res.text.includes('1 cuentas de personas'));
});

test.serial('GET /admin/avisos lists the message center, or says the permission is missing', async t => {
  const admin = await signInAs(t, ADMIN, { admin: true });
  http.on('get', '/serviceAnnouncement/messages', { value: [{ id: 'MC1', title: 'Cambio en Outlook', category: 'planForChange', services: ['Exchange Online'], isMajorChange: true, actionRequiredByDateTime: '2026-12-01T00:00:00Z', lastModifiedDateTime: '2026-10-01T00:00:00Z' }] });

  const res = await admin.get('/admin/avisos').expect(200);
  t.true(res.text.includes('Cambio en Outlook'));
  t.true(res.text.includes('Cambio importante'));
  t.true(res.text.includes('href="https://admin.microsoft.com/#/MessageCenter/:/messages/MC1"'));

  resetMessageCache();
  http.fail('get', '/serviceAnnouncement/messages', 403);
  const missing = await admin.get('/admin/avisos').expect(200);
  t.true(missing.text.includes('ServiceMessage.Read.All'));
});

test.serial('GET /admin/salas shows a room\'s week with who booked and who goes, read with the app\'s token', async t => {
  const admin = await signInAs(t, ADMIN, { admin: true });
  const SALA = 'sala.plenos@arandadeduero.es';
  http.on('get', '/places/microsoft.graph.room', { value: [{ displayName: 'Sala de Plenos', emailAddress: SALA, capacity: 40 }] });
  const person = (name, response) => ({ emailAddress: { name, address: `${name.toLowerCase()}@arandadeduero.es` }, status: { response } });
  http.on('get', '/calendarView', { value: [
    { subject: 'Pleno', organizer: { emailAddress: { name: 'Olga', address: 'olga@arandadeduero.es' } }, start: { dateTime: '2026-11-17T09:00:00.0000000' }, end: { dateTime: '2026-11-17T10:30:00.0000000' }, showAs: 'tentative', type: 'occurrence', attendees: [person('Bob', 'declined'), { emailAddress: { name: 'Sala de Plenos', address: SALA.toUpperCase() }, status: { response: 'accepted' } }] },
    { subject: 'Feria', isAllDay: true, organizer: { emailAddress: { name: 'Fede' } }, start: { dateTime: '2026-11-21T00:00:00.0000000' }, end: { dateTime: '2026-11-23T00:00:00.0000000' }, attendees: [] },
    { subject: 'Anulada', isCancelled: true, organizer: {}, start: { dateTime: '2026-11-18T09:00:00.0000000' }, end: { dateTime: '2026-11-18T10:00:00.0000000' } },
    { subject: 'Otra semana', organizer: {}, start: { dateTime: '2026-11-23T08:00:00.0000000' }, end: { dateTime: '2026-11-23T09:00:00.0000000' } },
  ] });

  // No room chosen: nothing is read
  const empty = await admin.get('/admin/salas').expect(200);
  t.true(empty.text.includes('Sala de Plenos (40 personas)'));
  t.is(http.callsTo('get', '/calendarView').length, 0);

  // Any day of the week shows the week from its Monday
  const res = await admin.get(`/admin/salas?room=${SALA}&week=2026-11-19`).expect(200);
  const [call] = http.callsTo('get', '/calendarView');
  t.true(call.url.endsWith(`/users/${encodeURIComponent(SALA)}/calendarView`));
  t.is(call.config.headers.Authorization, 'Bearer app-token');
  t.is(call.config.params.startDateTime, '2026-11-15T00:00:00');
  t.is(call.config.params.endDateTime, '2026-11-24T00:00:00');
  t.true(res.text.includes('09:00 – 10:30'));
  t.true(res.text.includes('Reservada por <strong>Olga</strong> (olga@arandadeduero.es)'));
  t.true(res.text.includes('Convocados (1):'));
  t.true(res.text.includes('Bob — ha rechazado'));
  t.true(res.text.includes('Pendiente de aceptar'));
  t.true(res.text.includes('Periódica'));
  // Saturday and Sunday, not the Monday it ends on at 00:00
  t.is(res.text.split('Feria').length - 1, 2);
  t.is(res.text.split('libre todo el día').length - 1, 4);
  t.false(res.text.includes('Anulada'));
  t.false(res.text.includes('Otra semana'));
  t.true(res.text.includes('week=2026-11-09'));
  t.true(res.text.includes('week=2026-11-23'));
});

test.serial('GET /admin/salas only opens mailboxes that are rooms, and a calendar it cannot read is not an empty one', async t => {
  const admin = await signInAs(t, ADMIN, { admin: true });
  http.on('get', '/places/microsoft.graph.room', { value: [{ displayName: 'Sala de Plenos', emailAddress: 'sala.plenos@arandadeduero.es' }] });
  http.fail('get', '/calendarView', 403);

  await admin.get('/admin/salas?room=alcalde@arandadeduero.es').expect(200);
  t.is(http.callsTo('get', '/calendarView').length, 0);

  const res = await admin.get('/admin/salas?room=sala.plenos@arandadeduero.es&week=nonsense').expect(200);
  t.true(res.text.includes('Calendars.Read'));
  t.false(res.text.includes('libre todo el día'));
});

// What Graph answers when everything a commission needs is in place
function stubCommissionSetup() {
  http.on('get', '/me/calendars', { value: [{ id: 'cal-main', name: 'Calendario' }, { id: 'cal-com', name: 'Comisiones y plenos', canEdit: true }] });
  http.on('get', '/places/microsoft.graph.room', { value: [
    { displayName: 'Sala de Plenos', emailAddress: 'sala.plenos@arandadeduero.es', capacity: 40 },
    { displayName: 'Sala Pequeña', emailAddress: 'sala.pequena@arandadeduero.es', capacity: 4 },
    { displayName: 'Sala Sin Aforo', emailAddress: 'sala.sinaforo@arandadeduero.es' },
    { displayName: 'Casa Consistorial - Sala de Juntas', emailAddress: 'sala.juntas@arandadeduero.es', capacity: 12 },
  ] });
  http.on('get', '/v1.0/groups', ({ config }) => ({ value: [{ id: config.params.$filter.includes('eqgobierno') ? 'gov-id' : 'dist-id' }] }));
  http.on('get', '/groups/gov-id/transitiveMembers/microsoft.graph.user', { value: [{ id: 'c1', displayName: 'Carmen Concejala', mail: 'carmen@arandadeduero.es' }] });
  http.on('get', '/groups/dist-id/transitiveMembers/microsoft.graph.user', { value: [
    { id: 'c1', displayName: 'Carmen Concejala', mail: 'carmen@arandadeduero.es' },
    { id: 'c2', displayName: 'Blas Concejal', mail: 'blas@arandadeduero.es' },
    { id: 'c2b', displayName: 'Blas Concejal', mail: 'BLAS@arandadeduero.es' },
    { id: 'c3', displayName: 'Sin Correo', mail: null },
  ] });
  http.on('post', '/me/calendars/cal-com/events', { id: 'event-1' });
}

test.serial('POST /admin/comisiones puts the meeting in the calendar and invites each councillor and the room', async t => {
  const admin = await signInAs(t, ADMIN, { admin: true });
  stubCommissionSetup();

  const page = await admin.get('/admin/comisiones').expect(200);
  t.true(page.text.includes('<option value="Urbanismo">'));
  // The Junta's usual room is offered with it; «Plenos» does not end in «pleno», so the plenary gets none
  t.true(page.text.includes('<option value="JGL" data-room="sala.juntas@arandadeduero.es">'));
  t.true(page.text.includes('<option value="Pleno">'));
  // Nothing asks before sending; a far-off date only gets a notice
  t.false(page.text.includes('data-confirm'));
  t.true(page.text.includes('/js/commissions.js'));
  t.true(page.text.includes('Sala de Plenos (40 personas)'));
  t.false(page.text.includes('Sala Pequeña'));
  t.false(page.text.includes('Sala Sin Aforo'));
  t.true(page.text.includes('Comisiones, Pleno: a quién se invita (2)'));
  t.true(page.text.includes('JGL: a quién se invita (1)'));
  t.true(page.text.includes('Carmen Concejala'));
  t.is(http.callsTo('get', '/v1.0/groups')[0].config.params.$filter, "mail eq 'dist-concejales2327@arandadeduero.es'");

  const csrf = await getCsrfToken(admin);
  const form = { _csrf: csrf, commission: 'Urbanismo', start: '2099-03-02T23:30', room: 'sala.plenos@arandadeduero.es' };
  await admin.post('/admin/comisiones').type('form').send(form).expect(302).expect('Location', '/admin/comisiones');

  const [call] = http.callsTo('post', '/me/calendars/cal-com/events');
  t.is(call.config.headers.Authorization, 'Bearer user-token');
  t.is(call.data.subject, 'Com Urbanismo');
  t.deepEqual(call.data.start, { dateTime: '2099-03-02T23:30:00', timeZone: 'Romance Standard Time' });
  // One hour later, across midnight
  t.deepEqual(call.data.end, { dateTime: '2099-03-03T00:30:00', timeZone: 'Romance Standard Time' });
  t.deepEqual(call.data.attendees.map(a => [a.emailAddress.address.toLowerCase(), a.type]), [
    ['blas@arandadeduero.es', 'required'],
    ['carmen@arandadeduero.es', 'required'],
    ['sala.plenos@arandadeduero.es', 'resource'],
  ]);
  t.deepEqual(auditEvents('admin.commission_called').map(e => [e.commission, e.start, e.room, e.invited]), [['Urbanismo', '2099-03-02T23:30', 'sala.plenos@arandadeduero.es', 2]]);

  const done = await admin.get('/admin/comisiones').expect(200);
  t.true(done.text.includes('Se ha invitado a 2 personas'));

  // A plenary is not a commission: its subject is its name
  await admin.post('/admin/comisiones').type('form').send({ ...form, commission: 'Pleno' }).expect(302);
  t.is(http.callsTo('post', '/me/calendars/cal-com/events')[1].data.subject, 'Pleno');
  // Three hours from 23:30
  t.is(http.callsTo('post', '/me/calendars/cal-com/events')[1].data.end.dateTime, '2099-03-03T02:30:00');

  // The Junta is half an hour, and only for the government team
  await admin.post('/admin/comisiones').type('form').send({ ...form, commission: 'JGL' }).expect(302);
  const jgl = http.callsTo('post', '/me/calendars/cal-com/events')[2].data;
  t.is(jgl.end.dateTime, '2099-03-03T00:00:00');
  t.deepEqual(jgl.attendees.map(a => a.emailAddress.address), ['carmen@arandadeduero.es', 'sala.plenos@arandadeduero.es']);

  // What the form may not choose
  for (const bad of [{ commission: 'Fiestas' }, { room: 'sala.pequena@arandadeduero.es' }, { room: 'alcalde@arandadeduero.es' }, { start: 'mañana' }]) {
    await admin.post('/admin/comisiones').type('form').send({ ...form, ...bad }).expect(400);
  }
  t.is(http.callsTo('post', '/me/calendars/cal-com/events').length, 3);
});

test.serial('/admin/comisiones calls nobody when the calendar or the list of councillors is missing', async t => {
  const admin = await signInAs(t, ADMIN, { admin: true });
  const csrf = await getCsrfToken(admin);
  const form = { _csrf: csrf, commission: 'Cultura', start: '2099-03-02T08:30', room: 'sala.plenos@arandadeduero.es' };

  stubCommissionSetup();
  http.fail('get', '/groups/dist-id/transitiveMembers/microsoft.graph.user', 403);
  const unread = await admin.get('/admin/comisiones').expect(200);
  t.true(unread.text.includes('no se sabe a quién convocar'));
  await admin.post('/admin/comisiones').type('form').send(form).expect(409);
  // The other list is still there, so the Junta can be called
  await admin.post('/admin/comisiones').type('form').send({ ...form, commission: 'JGL' }).expect(302);

  http.reset();
  http.on('get', '/v1.0/me', { displayName: 'Someone' });
  http.on('get', ROLES, { value: [{ displayName: 'Global Administrator' }] });
  stubCommissionSetup();
  http.on('get', '/me/calendars', { value: [{ id: 'cal-main', name: 'Calendario' }] });
  const missing = await admin.get('/admin/comisiones').expect(200);
  t.true(missing.text.includes('No tienes ningún calendario llamado «Comisiones y plenos»'));
  await admin.post('/admin/comisiones').type('form').send(form).expect(409);

  t.is(http.calls.filter(c => c.method === 'post' && c.url.includes('/events')).length, 0);
  t.is(auditEvents('admin.commission_called').filter(e => e.commission === 'Cultura').length, 0);
});

test.serial('GET /admin/solicitudes shows every request still waiting, oldest first', async t => {
  const admin = await signInAs(t, ADMIN, { admin: true });
  requestStore.createRequest({ groupId: 'g1', groupName: 'Obras', requester: { id: 'u1', name: 'Bob', email: 'bob@arandadeduero.es' } });

  const res = await admin.get('/admin/solicitudes').expect(200);
  t.true(res.text.includes('Obras'));
  t.true(res.text.includes('Bob'));
  t.true(res.text.includes('<td>hoy</td>'));
});

test.serial('settings saved from the web take effect at once and are audited', async t => {
  const admin = await signInAs(t, ADMIN, { admin: true });
  const csrf = await getCsrfToken(admin);
  permissions['/blog'] = ['someone'];
  appConfig.SHARED_MAILBOXES = [];

  await admin.post('/admin/ajustes').set('X-CSRF-Token', csrf).type('form').send({ blogUsers: 'Alice\nbob@arandadeduero.es\n', sharedMailboxes: 'Registro@arandadeduero.es\nobras@gmail.com' }).expect(400);
  t.deepEqual(permissions['/blog'], ['someone']);

  const res = await admin.post('/admin/ajustes').set('X-CSRF-Token', csrf).type('form').send({ blogUsers: 'Alice\nbob@arandadeduero.es\n', sharedMailboxes: 'Registro@arandadeduero.es' }).expect(302);
  t.is(res.header.location, '/admin/ajustes');
  t.deepEqual(permissions['/blog'], ['alice', 'bob']);
  t.deepEqual(appConfig.SHARED_MAILBOXES, ['registro@arandadeduero.es']);
  t.true(fs.existsSync(path.join(process.env.DATA_DIR, 'settings.json')));
  t.like(auditEvents('admin.settings_updated').at(-1), { blogUsers: 2, sharedMailboxes: 1 });

  const page = await admin.get('/admin/ajustes').expect(200);
  t.true(page.text.includes('Guardado. Ya está en vigor.'));
  t.true(page.text.includes('>alice\nbob</textarea>'));
});

test.serial('GET /admin/audit.csv exports the filtered log', async t => {
  const admin = await signInAs(t, ADMIN, { admin: true });
  fs.appendFileSync(path.join(process.env.AUDIT_DIR, `${day(0).slice(0, 10)}.log`), JSON.stringify({ time: day(0), event: 'blog.post', actor: 'carol@arandadeduero.es', ip: '203.0.113.9', title: 'Corte' }) + '\n');

  const res = await admin.get('/admin/audit.csv?event=blog.post').expect(200).expect('Content-Type', /text\/csv/);
  t.true(res.text.startsWith('﻿"Fecha";"Evento";"Usuario";"IP";"Detalles"'));
  t.true(res.text.includes('"blog.post";"carol@arandadeduero.es";"203.0.113.9";"{""title"":""Corte""}"'));
  t.false(res.text.includes('auth.sign_in'));
});
