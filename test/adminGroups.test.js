/**
 * Admin Groups Tests (/admin/grupos/*)
 *
 * Covers the list and its filters, creating Microsoft 365 and security groups,
 * the detail page, editing, owners and members, and the groups the portal
 * must leave read-only (dynamic, Exchange lists).
 */

require('./helpers/env');

const test = require('ava').default;
const request = require('supertest');
const { stubAxios, stubMsal, signInAgent, getCsrfToken, auditEvents, httpError } = require('./helpers/stubs');
const { resetGraphCaches } = require('../services/graphService');
const { describeAdminGroup, diffGroupEdit, toMailAlias, validateNewGroup, parseAddressList, MAX_BULK } = require('../services/adminGroups');

const http = stubAxios();
const msal = stubMsal();

const permissions = require('../config/permissions');
const app = require('../server');

const ROLES = '/me/memberOf/microsoft.graph.directoryRole';
const UID_ADMIN = '99999999-9999-9999-9999-999999999999';
const ADMIN = { username: 'admin@arandadeduero.es', name: 'Ada Admin', localAccountId: UID_ADMIN };
const USER = { username: 'alice@arandadeduero.es', name: 'Alice Example', localAccountId: 'alice-id-1' };
const UID_ANA = '11111111-1111-1111-1111-111111111111';
const UID_CARL = '33333333-3333-3333-3333-333333333333';
const GID_INFO = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const GID_DYN = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';
const GID_LIST = 'cccccccc-cccc-cccc-cccc-cccccccccccc';

const GROUPS = {
  [GID_INFO]: { id: GID_INFO, displayName: 'Informática', description: 'Equipo de informática', mail: 'informatica@arandadeduero.es', mailEnabled: true, securityEnabled: false, groupTypes: ['Unified'], visibility: 'Private', createdDateTime: '2026-01-10T08:00:00Z', mailNickname: 'informatica' },
  [GID_DYN]: { id: GID_DYN, displayName: 'Todos los de Obras', mailEnabled: false, securityEnabled: true, groupTypes: ['DynamicMembership'], membershipRule: 'user.department -eq "Obras"', membershipRuleProcessingState: 'On' },
  [GID_LIST]: { id: GID_LIST, displayName: 'Lista prensa', mail: 'prensa@arandadeduero.es', mailEnabled: true, securityEnabled: false, groupTypes: [] },
};

async function signInAs(t, account, { admin }) {
  msal.account = account;
  const agent = request.agent(app);
  await signInAgent(agent);
  t.teardown(() => agent.get('/auth/logout'));
  http.on('get', ROLES, { value: admin ? [{ displayName: 'Global Administrator' }] : [] });
  const csrf = await getCsrfToken(agent);
  return { agent, csrf };
}

function stubGroupPages() {
  http.on('get', '/v1.0/groups', { value: Object.values(GROUPS) });
  for (const g of Object.values(GROUPS)) {
    http.on('get', `/v1.0/groups/${g.id}`, g);
    http.on('get', `/v1.0/groups/${g.id}/owners`, { value: [{ id: UID_ADMIN, displayName: 'Ada Admin', mail: 'admin@arandadeduero.es' }] });
    http.on('get', `/v1.0/groups/${g.id}/members/microsoft.graph.user`, { value: [{ id: UID_ANA, displayName: 'Ana Pérez', mail: 'ana@arandadeduero.es', userType: 'Member' }] });
  }
}

test.beforeEach(() => {
  http.reset();
  resetGraphCaches();
  http.on('get', '/v1.0/me', { displayName: 'Ada Admin' });
  permissions.DOMAIN = 'arandadeduero.es';
});

test('mail aliases are plain ASCII and groups are described by what Graph can do with them', t => {
  t.is(toMailAlias('Informática y Redes!'), 'informatica-y-redes');
  t.is(toMailAlias('  --Ñandú.2026-- '), 'nandu.2026');
  // Long runs of dashes are cut without trouble
  t.is(toMailAlias(`${'-'.repeat(50000)}a${'-'.repeat(50000)}!`), '');
  t.is(toMailAlias('---a---'), 'a');
  t.true(describeAdminGroup(GROUPS[GID_INFO]).manageable);
  t.is(describeAdminGroup(GROUPS[GID_DYN]).kind, 'Seguridad dinámico');
  t.false(describeAdminGroup(GROUPS[GID_DYN]).manageable);
  t.is(describeAdminGroup(GROUPS[GID_LIST]).kind, 'Lista de distribución (Exchange)');
  t.false(describeAdminGroup({ ...GROUPS[GID_INFO], onPremisesSyncEnabled: true }).manageable);
  t.truthy(validateNewGroup({ displayName: '   ' }).error);
  // The name is free; an empty alias becomes gr-<name>, a typed one must start with gr-
  t.is(validateNewGroup({ displayName: 'Informática', type: 'security' }).value.mailNickname, 'gr-informatica');
  t.is(validateNewGroup({ displayName: 'X', mailNickname: 'gr-obras' }).value.mailNickname, 'gr-obras');
  t.is(validateNewGroup({ displayName: 'X', mailNickname: 'GR-Obras' }).value.mailNickname, 'gr-obras');
  for (const alias of ['obras', 'gr', 'gr-', 'xgr-obras', 'Información']) {
    t.regex(validateNewGroup({ displayName: 'Obras', mailNickname: alias }).error, /gr-/, alias);
  }
});

test('a pasted list of addresses is read by lines, commas and semicolons, also as copied from Outlook', t => {
  const { addresses, invalid } = parseAddressList('Ana@arandadeduero.es\r\nLuis Gómez <luis@arandadeduero.es>; pepe@arandadeduero.es,\n\n  ana@arandadeduero.es  \nesto no es un correo\nfalta@');
  // Lowercase, once each, in order
  t.deepEqual(addresses, ['ana@arandadeduero.es', 'luis@arandadeduero.es', 'pepe@arandadeduero.es']);
  t.deepEqual(invalid, ['esto no es un correo', 'falta@']);
  t.deepEqual(parseAddressList(''), { addresses: [], invalid: [] });
  t.deepEqual(parseAddressList(undefined), { addresses: [], invalid: [] });
});

test('editing sends only what changed, and an emptied description becomes null', t => {
  const { patch, changes } = diffGroupEdit(GROUPS[GID_INFO], { displayName: 'Informática', description: '', visibility: 'Public' });
  t.deepEqual(patch, { description: null, visibility: 'Public' });
  t.deepEqual(changes, ['description', 'visibility']);
  t.deepEqual(diffGroupEdit(GROUPS[GID_INFO], { displayName: 'Informática', description: 'Equipo de informática', visibility: 'Private' }).changes, []);
  // A hidden-membership group is not turned public by saving something else
  const hidden = { ...GROUPS[GID_INFO], visibility: 'HiddenMembership' };
  t.deepEqual(diffGroupEdit(hidden, { displayName: 'Informática', description: 'Equipo de informática', visibility: 'Public' }).changes, []);
});

test.serial('non-admins receive 403 on every /admin/grupos route', async t => {
  const { agent, csrf } = await signInAs(t, USER, { admin: false });
  for (const p of ['/admin/grupos', '/admin/grupos-distribucion', '/admin/grupos/nuevo', `/admin/grupos/${GID_INFO}`]) {
    await agent.get(p).expect(403);
  }
  await agent.post('/admin/grupos').set('X-CSRF-Token', csrf).type('form').send({ displayName: 'X' }).expect(403);
  await agent.post(`/admin/grupos/${GID_INFO}/miembros`).set('X-CSRF-Token', csrf).type('form').send({ user: 'a@arandadeduero.es' }).expect(403);
  t.is(http.callsTo('post', '/v1.0/groups').length, 0);
});

test.serial('the list shows every kind of group and filters by text and kind', async t => {
  const { agent } = await signInAs(t, ADMIN, { admin: true });
  stubGroupPages();

  const all = await agent.get('/admin/grupos').expect(200);
  t.regex(all.text, /Informática/);
  t.regex(all.text, /Seguridad dinámico/);
  t.regex(all.text, /Lista de distribución \(Exchange\)/);

  const byText = await agent.get('/admin/grupos').query({ q: 'prensa' }).expect(200);
  t.regex(byText.text, /Lista prensa/);
  t.notRegex(byText.text, /Informática/);

  const byKind = await agent.get('/admin/grupos').query({ tipo: 'Microsoft 365' }).expect(200);
  t.regex(byKind.text, /Informática/);
  t.notRegex(byKind.text, /Lista prensa/);
});

test.serial('the distribution lists page shows only the Exchange lists, with a way to see who is in them', async t => {
  const { agent } = await signInAs(t, ADMIN, { admin: true });
  stubGroupPages();

  const res = await agent.get('/admin/grupos-distribucion').expect(200);
  t.regex(res.text, /Lista prensa/);
  t.regex(res.text, /prensa@arandadeduero\.es/);
  t.regex(res.text, new RegExp(`href="/admin/grupos/${GID_LIST}"`));
  t.notRegex(res.text, /Informática/);
  t.notRegex(res.text, /Todos los de Obras/);
  // What it can and cannot do is said, and the way to make one that can be managed here is offered
  t.regex(res.text, /Aquí solo se pueden consultar/);
  t.regex(res.text, /href="\/admin\/grupos\/nuevo"/);
  t.regex(res.text, /href="\/admin\/grupos-distribucion"/);

  // The list is read-only: its page shows who is in it and offers no change
  const detail = await agent.get(`/admin/grupos/${GID_LIST}`).expect(200);
  t.regex(detail.text, /Ana Pérez/);
  t.notRegex(detail.text, /Añadir miembro/);
});

test.serial('creating a Microsoft 365 group sends the right payload with the admin as owner', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });
  http.on('post', '/v1.0/groups', { id: GID_INFO });
  stubGroupPages();

  const res = await agent.post('/admin/grupos').set('X-CSRF-Token', csrf).type('form')
    .send({ type: 'unified', displayName: 'Informática y Redes', description: 'TIC', visibility: 'Public' })
    .expect(302);
  t.is(res.headers.location, `/admin/grupos/${GID_INFO}`);

  const body = http.callsTo('post', '/v1.0/groups')[0].data;
  t.is(body.mailNickname, 'gr-informatica-y-redes');
  t.deepEqual(body.groupTypes, ['Unified']);
  t.true(body.mailEnabled);
  t.false(body.securityEnabled);
  t.is(body.visibility, 'Public');
  t.deepEqual(body['owners@odata.bind'], [`https://graph.microsoft.com/v1.0/users/${UID_ADMIN}`]);
  t.true(auditEvents('admin.group_created').some(e => e.groupId === GID_INFO && e.type === 'unified'));
});

test.serial('creating a security group sends no group type, no mail and no visibility', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });
  http.on('post', '/v1.0/groups', { id: GID_INFO });

  await agent.post('/admin/grupos').set('X-CSRF-Token', csrf).type('form')
    .send({ type: 'security', displayName: 'Acceso Obras' }).expect(302);

  const body = http.callsTo('post', '/v1.0/groups')[0].data;
  t.deepEqual(body.groupTypes, []);
  t.false(body.mailEnabled);
  t.true(body.securityEnabled);
  t.false('visibility' in body);
  t.false('description' in body);
});

test.serial('an alias without the gr- prefix is refused by the server, and the form asks for it', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });

  const form = (await agent.get('/admin/grupos/nuevo').expect(200)).text;
  t.regex(form, /placeholder="gr-informatica"/);
  t.regex(form, /pattern="gr-\.\+"/);

  for (const mailNickname of ['informatica', 'gr-', 'xgr-informatica']) {
    const res = await agent.post('/admin/grupos').set('X-CSRF-Token', csrf).type('form').send({ type: 'unified', displayName: 'Informática', mailNickname }).expect(302);
    t.is(res.headers.location, '/admin/grupos/nuevo');
    t.regex((await agent.get('/admin/grupos/nuevo')).text, /alias de correo tiene que empezar por «gr-»/);
  }
  t.is(http.callsTo('post', '/v1.0/groups').length, 0);
});

test.serial('a nameless group is not sent, and a repeated alias is explained', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });

  const empty = await agent.post('/admin/grupos').set('X-CSRF-Token', csrf).type('form').send({ displayName: ' ' }).expect(302);
  t.is(empty.headers.location, '/admin/grupos/nuevo');
  t.is(http.callsTo('post', '/v1.0/groups').length, 0);
  t.regex((await agent.get('/admin/grupos/nuevo')).text, /necesita un nombre/);

  const createdBefore = auditEvents('admin.group_created').length;
  http.on('post', '/v1.0/groups', () => {
    const err = httpError(400);
    err.response.data = { error: { message: 'Another object with the same value for property mailNickname already exists.' } };
    throw err;
  });
  await agent.post('/admin/grupos').set('X-CSRF-Token', csrf).type('form').send({ displayName: 'Obras' }).expect(302);
  t.regex((await agent.get('/admin/grupos/nuevo')).text, /Ya existe otro grupo con ese alias/);
  t.is(auditEvents('admin.group_created').length, createdBefore);
});

test.serial('the detail page shows properties, owners and members', async t => {
  const { agent } = await signInAs(t, ADMIN, { admin: true });
  stubGroupPages();

  const res = await agent.get(`/admin/grupos/${GID_INFO}`).expect(200);
  t.regex(res.text, /informatica@arandadeduero\.es/);
  t.regex(res.text, /Ada Admin/);
  t.regex(res.text, /Ana Pérez/);
  t.regex(res.text, /Guardar cambios/);

  await agent.get('/admin/grupos/not-an-id').expect(404);
});

test.serial('editing a group patches only the changed fields and audits it', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });
  stubGroupPages();
  http.on('patch', `/v1.0/groups/${GID_INFO}`, {});

  await agent.post(`/admin/grupos/${GID_INFO}`).set('X-CSRF-Token', csrf).type('form')
    .send({ displayName: 'Informática y TIC', description: 'Equipo de informática', visibility: 'Private' }).expect(302);

  const patches = http.callsTo('patch', `/v1.0/groups/${GID_INFO}`);
  t.is(patches.length, 1);
  t.deepEqual(patches[0].data, { displayName: 'Informática y TIC' });
  t.true(auditEvents('admin.group_updated').some(e => e.groupId === GID_INFO && e.fields.join() === 'displayName'));
});

test.serial('dynamic groups and Exchange lists cannot be changed from the portal', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });
  stubGroupPages();

  const page = await agent.get(`/admin/grupos/${GID_DYN}`).expect(200);
  t.regex(page.text, /Sus miembros los decide una regla/);
  t.regex(page.text, /user\.department -eq/);
  t.notRegex(page.text, /Guardar cambios/);

  for (const gid of [GID_DYN, GID_LIST]) {
    await agent.post(`/admin/grupos/${gid}`).set('X-CSRF-Token', csrf).type('form').send({ displayName: 'Otro' }).expect(302);
    await agent.post(`/admin/grupos/${gid}/miembros`).set('X-CSRF-Token', csrf).type('form').send({ user: 'ana@arandadeduero.es' }).expect(302);
  }
  t.is(http.calls.filter(c => ['patch', 'post'].includes(c.method) && c.url.includes('/groups/')).length, 0);
});

test.serial('members and owners are added by address and removed by ID', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });
  stubGroupPages();
  http.on('get', '/v1.0/users/carl%40arandadeduero.es', { id: UID_CARL, displayName: 'Carl', userPrincipalName: 'carl@arandadeduero.es' });
  http.on('post', `/v1.0/groups/${GID_INFO}/members/$ref`, {});
  http.on('post', `/v1.0/groups/${GID_INFO}/owners/$ref`, {});
  http.on('delete', `/v1.0/groups/${GID_INFO}/members/${UID_ANA}/$ref`, {});
  http.on('delete', `/v1.0/groups/${GID_INFO}/owners/${UID_ADMIN}/$ref`, {});

  const post = (path, form = {}) => agent.post(`/admin/grupos/${GID_INFO}/${path}`).set('X-CSRF-Token', csrf).type('form').send(form).expect(302);

  await post('miembros', { user: 'carl@arandadeduero.es' });
  t.is(http.callsTo('post', `/v1.0/groups/${GID_INFO}/members/$ref`)[0].data['@odata.id'], `https://graph.microsoft.com/v1.0/directoryObjects/${UID_CARL}`);
  await post('propietarios', { user: 'carl@arandadeduero.es' });
  t.is(http.callsTo('post', `/v1.0/groups/${GID_INFO}/owners/$ref`).length, 1);
  await post(`miembros/${UID_ANA}/quitar`);
  await post(`propietarios/${UID_ADMIN}/quitar`);
  t.is(http.callsTo('delete', `/v1.0/groups/${GID_INFO}/members/${UID_ANA}/$ref`).length, 1);
  t.is(http.callsTo('delete', `/v1.0/groups/${GID_INFO}/owners/${UID_ADMIN}/$ref`).length, 1);

  const events = auditEvents().map(e => e.event);
  for (const e of ['admin.group_member_added', 'admin.group_owner_added', 'admin.group_member_removed', 'admin.group_owner_removed']) {
    t.true(events.includes(e), e);
  }

  // An address nobody has: nothing is added
  const before = http.callsTo('post', `/v1.0/groups/${GID_INFO}/members/$ref`).length;
  await post('miembros', { user: 'nadie@arandadeduero.es' });
  t.is(http.callsTo('post', `/v1.0/groups/${GID_INFO}/members/$ref`).length, before);
  t.regex((await agent.get(`/admin/grupos/${GID_INFO}`)).text, /No se ha encontrado ese usuario/);
});

const BULK = (gid) => `/admin/grupos/${gid}/miembros/lote`;

test.serial('the bulk add shows what would happen and changes nothing until it is confirmed', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });
  stubGroupPages();
  http.on('get', '/v1.0/users/ana%40arandadeduero.es', { id: UID_ANA, displayName: 'Ana Pérez', userPrincipalName: 'ana@arandadeduero.es' });
  http.on('get', '/v1.0/users/carl%40arandadeduero.es', { id: UID_CARL, displayName: 'Carl Nuevo', userPrincipalName: 'carl@arandadeduero.es' });

  const res = await agent.post(`${BULK(GID_INFO)}/vista-previa`).set('X-CSRF-Token', csrf).type('form')
    .send({ addresses: 'ana@arandadeduero.es\nCarl <carl@arandadeduero.es>\nnadie@arandadeduero.es\nesto no es un correo' }).expect(200);

  t.regex(res.text, /<strong>1<\/strong> de <strong>4<\/strong> se añadirán/);
  t.regex(res.text, /Ya es miembro/);
  t.regex(res.text, /Se añadirá/);
  t.regex(res.text, /No se ha encontrado/);
  t.regex(res.text, /esto no es un correo/);
  // Only the person who is new travels to the confirmation
  t.regex(res.text, new RegExp(`name="userIds" value="${UID_CARL}"`));
  t.notRegex(res.text, new RegExp(`name="userIds" value="${UID_ANA}"`));
  t.is(http.calls.filter(c => c.method === 'post' && c.url.includes('/groups/')).length, 0);
});

test.serial('confirming adds the ones sent, counts failures, and audits only the numbers', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });
  stubGroupPages();
  http.on('post', `/v1.0/groups/${GID_INFO}/members/$ref`, ({ data }) => {
    if (data['@odata.id'].includes(UID_ANA)) throw httpError(500);
    return {};
  });

  await agent.post(BULK(GID_INFO)).set('X-CSRF-Token', csrf).type('form')
    .send({ userIds: [UID_CARL, UID_ANA, 'not-an-id', UID_CARL] }).expect(302);

  // The repeated and the invalid ones never reach Graph
  const sent = http.callsTo('post', `/v1.0/groups/${GID_INFO}/members/$ref`).map(c => c.data['@odata.id']);
  t.is(sent.length, 2);
  const page = await agent.get(`/admin/grupos/${GID_INFO}`).expect(200);
  t.regex(page.text, /Añadidas 1 persona\./);
  t.regex(page.text, /No se pudo añadir a 1 persona/);
  const event = auditEvents('admin.group_members_bulk_added').at(-1);
  t.like(event, { groupId: GID_INFO, added: 1, failed: 1 });
  // Who was added is not recorded, only how many
  t.false(JSON.stringify(event).includes(UID_CARL) || JSON.stringify(event).includes(UID_ANA));
});

test.serial('the bulk add refuses empty lists, too long lists, read-only groups and non-administrators', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });
  stubGroupPages();
  const post = (path, form) => agent.post(path).set('X-CSRF-Token', csrf).type('form').send(form).expect(302);
  const writes = () => http.calls.filter(c => c.method === 'post' && c.url.includes('/groups/')).length;

  await post(`${BULK(GID_INFO)}/vista-previa`, { addresses: '   ' });
  t.regex((await agent.get(`/admin/grupos/${GID_INFO}`)).text, /Escribe al menos un correo/);

  const many = Array.from({ length: MAX_BULK + 1 }, (_, i) => `persona${i}@arandadeduero.es`).join('\n');
  const lookups = () => http.calls.filter(c => c.method === 'get' && c.url.includes('/v1.0/users/')).length;
  const before = lookups();
  await post(`${BULK(GID_INFO)}/vista-previa`, { addresses: many });
  t.regex((await agent.get(`/admin/grupos/${GID_INFO}`)).text, new RegExp(`Como mucho ${MAX_BULK} personas`));
  t.is(lookups(), before);

  // A dynamic group and an Exchange list are read-only: nothing is looked up or written
  for (const gid of [GID_DYN, GID_LIST]) {
    await post(`${BULK(gid)}/vista-previa`, { addresses: 'ana@arandadeduero.es' });
    await post(BULK(gid), { userIds: [UID_ANA] });
  }
  t.is(writes(), 0);
  t.is(lookups(), before);

  await post(BULK(GID_INFO), {});
  t.regex((await agent.get(`/admin/grupos/${GID_INFO}`)).text, /No había a nadie que añadir/);
});

test.serial('non-administrators cannot use the bulk add', async t => {
  const { agent, csrf } = await signInAs(t, USER, { admin: false });
  await agent.post(`${BULK(GID_INFO)}/vista-previa`).set('X-CSRF-Token', csrf).type('form').send({ addresses: 'a@arandadeduero.es' }).expect(403);
  await agent.post(BULK(GID_INFO)).set('X-CSRF-Token', csrf).type('form').send({ userIds: [UID_ANA] }).expect(403);
  t.is(http.callsTo('post', `/v1.0/groups/${GID_INFO}/members/$ref`).length, 0);
});
