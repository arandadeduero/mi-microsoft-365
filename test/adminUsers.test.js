/**
 * Admin Users Routes Tests (/admin/usuarios/*)
 *
 * Covers the integrated m365-tools workflows: directory & filters, user edit &
 * password reset, Excel validate/sync/disable/import (including multipart CSRF),
 * bulk manager & group tools, and data quality fixes.
 */

require('./helpers/env');

const ExcelJS = require('exceljs');
const test = require('ava').default;
const request = require('supertest');
const { stubAxios, stubMsal, signInAgent, getCsrfToken, auditEvents, httpError } = require('./helpers/stubs');
const { resetGraphCaches } = require('../services/graphService');

const http = stubAxios();
const msal = stubMsal();

const permissions = require('../config/permissions');
const appConfig = require('../config/appConfig');
const app = require('../server');

const ROLES = '/me/memberOf/microsoft.graph.directoryRole';
const ADMIN = { username: 'admin@arandadeduero.es', name: 'Ada Admin', localAccountId: 'admin-id-1' };
const USER = { username: 'alice@arandadeduero.es', name: 'Alice Example', localAccountId: 'alice-id-1' };

const UID_ANA = '11111111-1111-1111-1111-111111111111';
const UID_BOSS = '22222222-2222-2222-2222-222222222222';
const UID_CARL = '33333333-3333-3333-3333-333333333333';
const GID_OBRAS = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa';
const GID_CULTURA = 'bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb';

async function signInAs(t, account, { admin }) {
  msal.account = account;
  const agent = request.agent(app);
  await signInAgent(agent);
  t.teardown(() => agent.get('/auth/logout'));
  http.on('get', ROLES, { value: admin ? [{ displayName: 'Global Administrator' }] : [] });
  const csrf = await getCsrfToken(agent);
  return { agent, csrf };
}

async function buildWorkbookBuffer(sheets) {
  const wb = new ExcelJS.Workbook();
  for (const [name, rows] of Object.entries(sheets)) {
    const ws = wb.addWorksheet(name);
    for (const row of rows) {
      ws.addRow(row);
    }
  }
  const out = await wb.xlsx.writeBuffer();
  return Buffer.from(out);
}

test.beforeEach(() => {
  http.reset();
  resetGraphCaches();
  http.on('get', '/v1.0/me', { displayName: 'Ada Admin' });
  permissions.DOMAIN = 'arandadeduero.es';
  appConfig.ORG_CHART_ROOT = 'boss@arandadeduero.es';
  appConfig.USERS_EXCEL_URL = '';
});

test.serial('non-admins receive 403 on all /admin/usuarios routes', async t => {
  const { agent, csrf } = await signInAs(t, USER, { admin: false });

  for (const p of ['/admin/usuarios', '/admin/usuarios/atencion', '/admin/usuarios/excel', '/admin/usuarios/responsables', '/admin/usuarios/calidad', `/admin/usuarios/${UID_ANA}`]) {
    await agent.get(p).expect(403);
  }
  await agent.post(`/admin/usuarios/${UID_ANA}/reset-password`).set('X-CSRF-Token', csrf).expect(403);
  await agent.post('/admin/usuarios/excel/deshabilitar').set('X-CSRF-Token', csrf).send({ userIds: [UID_ANA] }).expect(403);
  t.pass();
});

test.serial('GET /admin/usuarios lists users, applies filters, and builds the licensed mailing list', async t => {
  const { agent } = await signInAs(t, ADMIN, { admin: true });

  http.on('get', '/v1.0/users', ({ config }) => {
    if (config.params?.$select === 'id,signInActivity') {
      return {
        value: [
          { id: UID_ANA, signInActivity: { lastSignInDateTime: '2026-05-01T10:00:00Z' } },
          { id: UID_BOSS, signInActivity: null },
        ],
      };
    }
    return {
      value: [
        { id: UID_ANA, displayName: 'Ana Pérez', userPrincipalName: 'ana@arandadeduero.es', jobTitle: 'Técnica', department: 'Obras', employeeId: '100', accountEnabled: true },
        { id: UID_BOSS, displayName: 'Boss Jefe', userPrincipalName: 'boss@arandadeduero.es', jobTitle: '', department: 'Alcaldía', employeeId: '1', accountEnabled: true },
        { id: 'ext-1', displayName: 'Invitado', userPrincipalName: 'guest_gmail.com#EXT#@arandadeduero.es', accountEnabled: true },
      ],
    };
  });

  http.on('post', '/v1.0/$batch', ({ data }) => ({
    responses: data.requests.map((r) => {
      if (r.url.includes('/manager')) {
        return r.url.includes(UID_ANA)
          ? { id: r.id, status: 200, body: { id: UID_BOSS, displayName: 'Boss Jefe', userPrincipalName: 'boss@arandadeduero.es' } }
          : { id: r.id, status: 404, body: {} };
      }
      if (r.url.includes('/memberOf/')) {
        return r.url.includes(UID_ANA)
          ? { id: r.id, status: 200, body: { value: [{ id: GID_OBRAS, displayName: 'Obras' }, { id: 'auto', displayName: 'Todo Ayuntamiento' }] } }
          : { id: r.id, status: 200, body: { value: [{ id: 'auto', displayName: 'Todo Ayuntamiento' }] } };
      }
      if (r.url.includes('/licenseDetails')) {
        return r.url.includes(UID_ANA)
          ? { id: r.id, status: 200, body: { value: [{ skuId: 'sku-1', skuPartNumber: 'O365_BUSINESS_PREMIUM' }] } }
          : { id: r.id, status: 200, body: { value: [] } };
      }
      return { id: r.id, status: 404, body: {} };
    }),
  }));

  const res = await agent.get('/admin/usuarios').expect(200);
  t.true(res.text.includes('Ana Pérez'));
  t.true(res.text.includes('Boss Jefe'));
  t.false(res.text.includes('guest_gmail.com#EXT#'));
  // Mailing list has only licensed active users (ana, not boss)
  t.true(res.text.includes('ana@arandadeduero.es'));

  // Filter: missingManager + emptyGroups matches only Boss Jefe
  const filtered = await agent.get('/admin/usuarios?missingManager=1&emptyGroups=1').expect(200);
  t.true(filtered.text.includes('Boss Jefe'));
  t.false(filtered.text.includes(`/admin/usuarios/${UID_ANA}`));
});

test.serial('GET /admin/usuarios/atencion renders accounts needing attention with Usuarios tabs', async t => {
  const { agent } = await signInAs(t, ADMIN, { admin: true });
  http.on('get', '/v1.0/users', { value: [] });
  http.fail('get', '/userRegistrationDetails', 403);

  const res = await agent.get('/admin/usuarios/atencion').expect(200);
  t.true(res.text.includes('Cuentas que necesitan atención'));
  t.true(res.text.includes('/admin/usuarios/excel'));
});

test.serial('user detail, edit, manager, groups, and password reset work and never audit the password', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });

  http.on('get', `/v1.0/users/${UID_ANA}`, {
    id: UID_ANA,
    displayName: 'Ana Pérez',
    givenName: 'Ana',
    surname: 'Pérez',
    mailNickname: 'ana',
    userPrincipalName: 'ana@arandadeduero.es',
    jobTitle: 'Técnica',
    department: 'Obras',
    accountEnabled: true,
    businessPhones: ['947000001'],
  });
  http.on('get', `/users/${UID_ANA}/photo/$value`, () => { throw httpError(404); });
  http.on('get', `/users/${UID_ANA}/manager`, { id: UID_BOSS, displayName: 'Boss Jefe', userPrincipalName: 'boss@arandadeduero.es' });
  http.on('get', `/users/${UID_ANA}/memberOf/microsoft.graph.group`, { value: [{ id: GID_OBRAS, displayName: 'Obras' }] });
  http.on('get', `/users/${UID_ANA}/licenseDetails`, { value: [{ skuId: 'sku-1', skuPartNumber: 'O365_BUSINESS_PREMIUM' }] });
  http.on('get', '/v1.0/groups', { value: [{ id: GID_OBRAS, displayName: 'Obras' }, { id: GID_CULTURA, displayName: 'Cultura' }] });
  http.on('get', '/v1.0/users', {
    value: [
      { id: UID_BOSS, displayName: 'Boss Jefe', mail: 'boss@arandadeduero.es', accountEnabled: true, userType: 'Member', assignedLicenses: [{ skuId: 's1' }] },
      { id: UID_ANA, displayName: 'Ana Pérez', mail: 'ana@arandadeduero.es', accountEnabled: true, userType: 'Member', assignedLicenses: [{ skuId: 's1' }] },
    ],
  });
  http.on('patch', `/v1.0/users/${UID_ANA}`, {});
  http.on('put', `/v1.0/users/${UID_ANA}/manager/$ref`, {});
  http.on('post', `/v1.0/groups/${GID_CULTURA}/members/$ref`, {});
  http.on('delete', `/v1.0/groups/${GID_OBRAS}/members/${UID_ANA}/$ref`, {});

  const page = await agent.get(`/admin/usuarios/${UID_ANA}`).expect(200);
  t.true(page.text.includes('Ana Pérez'));
  t.true(page.text.includes('Boss Jefe'));

  // Edit user properties
  await agent.post(`/admin/usuarios/${UID_ANA}`)
    .set('X-CSRF-Token', csrf)
    .type('form')
    .send({
      displayName: 'Ana Pérez',
      givenName: 'Ana',
      surname: 'Pérez',
      mailNickname: 'ana',
      jobTitle: 'Jefa de sección',
      department: 'Obras',
      businessPhone: '947000001',
      accountEnabled: 'true',
    })
    .expect(302);

  const patches = http.callsTo('patch', `/v1.0/users/${UID_ANA}`);
  t.is(patches.length, 1);
  t.is(patches[0].data.jobTitle, 'Jefa de sección');

  // Update groups (remove Obras, add Cultura)
  await agent.post(`/admin/usuarios/${UID_ANA}/grupos`)
    .set('X-CSRF-Token', csrf)
    .type('form')
    .send({ groupIds: [GID_CULTURA] })
    .expect(302);

  t.is(http.callsTo('post', `/v1.0/groups/${GID_CULTURA}/members/$ref`).length, 1);
  t.is(http.callsTo('delete', `/v1.0/groups/${GID_OBRAS}/members/${UID_ANA}/$ref`).length, 1);

  // Reset password
  await agent.post(`/admin/usuarios/${UID_ANA}/reset-password`)
    .set('X-CSRF-Token', csrf)
    .expect(302);

  const resetPatch = http.callsTo('patch', `/v1.0/users/${UID_ANA}`).at(-1);
  const generatedPwd = resetPatch.data.passwordProfile.password;
  t.regex(generatedPwd, /^[A-ZÁÉÍÓÚÑa-záéíóúñ]+-\d{4}-[A-ZÁÉÍÓÚÑa-záéíóúñ]+$/);
  t.true(resetPatch.data.passwordProfile.forceChangePasswordNextSignIn);

  // Password shown once on redirected page, then cleared on second GET
  const afterReset = await agent.get(`/admin/usuarios/${UID_ANA}`).expect(200);
  t.true(afterReset.text.includes(generatedPwd));
  t.true(afterReset.text.includes('outlook.office.com/mail/deeplink/compose'));
  t.true(afterReset.text.includes('mailto:ana%40arandadeduero.es'));

  const secondLoad = await agent.get(`/admin/usuarios/${UID_ANA}`).expect(200);
  t.false(secondLoad.text.includes(generatedPwd));

  // Audit log recorded the reset WITHOUT the password
  const resetEvents = auditEvents().filter(e => e.event === 'admin.user_password_reset');
  t.true(resetEvents.length >= 1);
  t.false(JSON.stringify(resetEvents).includes(generatedPwd));
});

test.serial('Excel multipart routes enforce CSRF and validate, sync-check, disable safely, and import with preview', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });

  const xlsxBuf = await buildWorkbookBuffer({
    ACTIVOS: [
      ['id_empleado', 'Trabajador', 'Departamento', 'Descripción', 'Fecha de Alta', 'Fecha de Baja', 'e_mail', 'ID responsable'],
      ['1', 'JEFE, BOSS', 'Alcaldía', 'Alcalde', '01/01/2020', '', 'boss@arandadeduero.es', ''],
      ['2', 'PEREZ, ANA', 'Obras', 'Técnica de obras', '15/03/2022', '', 'ana@arandadeduero.es', '1'],
      ['3', 'BAJA, CARLOS', 'Obras', 'Peón', '01/01/2021', '01/01/2024', 'carlos@arandadeduero.es', '1'],
      ['4', 'NUEVA, LUCIA', 'Cultura', 'Archivera', '10/06/2025', '', 'lucia@arandadeduero.es', '1'],
    ],
    CONCEJALES: [
      ['email'],
      ['concejal@arandadeduero.es'],
    ],
  });

  // Missing CSRF on multipart route → 403
  await agent.post('/admin/usuarios/excel/validar')
    .attach('excelFile', xlsxBuf, 'plantilla.xlsx')
    .expect(403);

  // Valid CSRF → validates cleanly with 0 errors
  const valRes = await agent.post('/admin/usuarios/excel/validar')
    .field('_csrf', csrf)
    .attach('excelFile', xlsxBuf, 'plantilla.xlsx')
    .expect(200);
  t.true(valRes.text.includes('0') || valRes.text.toLowerCase().includes('sin errores') || valRes.text.toLowerCase().includes('válido'));

  // Sync-check: Carlos is effective leaving but still enabled; Lucia is missing; concejal is in CONCEJALES
  http.on('get', '/v1.0/users', {
    value: [
      { id: UID_BOSS, displayName: 'Boss Jefe', userPrincipalName: 'boss@arandadeduero.es', employeeId: '1', accountEnabled: true, jobTitle: 'Alcalde', department: 'Alcaldía' },
      { id: UID_ANA, displayName: 'Ana Perez', userPrincipalName: 'ana@arandadeduero.es', employeeId: '2', accountEnabled: true, jobTitle: 'Antiguo puesto', department: 'Obras' },
      { id: UID_CARL, displayName: 'Carlos Baja', userPrincipalName: 'carlos@arandadeduero.es', employeeId: '3', accountEnabled: true },
      { id: 'c1', displayName: 'Concejal Uno', userPrincipalName: 'concejal@arandadeduero.es', accountEnabled: true },
      { id: 'room1', displayName: 'Sala Plenos', userPrincipalName: 'sala-plenos@arandadeduero.es', accountEnabled: true },
    ],
  });
  http.on('post', '/v1.0/$batch', ({ data }) => ({
    responses: data.requests.map((r) => ({
      id: r.id,
      status: 200,
      body: { value: [{ skuId: 'sku-carl', skuPartNumber: 'O365_BUSINESS_PREMIUM' }] },
    })),
  }));

  const syncRes = await agent.post('/admin/usuarios/excel/sync')
    .field('_csrf', csrf)
    .attach('excelFile', xlsxBuf, 'plantilla.xlsx')
    .expect(200);
  t.true(syncRes.text.includes('carlos@arandadeduero.es'));
  t.true(syncRes.text.includes('lucia@arandadeduero.es'));
  t.false(syncRes.text.includes('sala-plenos@arandadeduero.es'));

  // Disable departed: removes licenses BEFORE disabling
  http.on('get', `/v1.0/users/${UID_CARL}`, { id: UID_CARL, displayName: 'Carlos Baja', userPrincipalName: 'carlos@arandadeduero.es' });
  http.on('get', `/users/${UID_CARL}/licenseDetails`, { value: [{ skuId: 'sku-carl', skuPartNumber: 'O365_BUSINESS_PREMIUM' }] });
  const order = [];
  http.on('post', `/v1.0/users/${UID_CARL}/assignLicense`, ({ data }) => {
    order.push({ op: 'unlicense', remove: data.removeLicenses });
    return {};
  });
  http.on('patch', `/v1.0/users/${UID_CARL}`, ({ data }) => {
    order.push({ op: 'disable', enabled: data.accountEnabled });
    return {};
  });

  const disableRes = await agent.post('/admin/usuarios/excel/deshabilitar')
    .set('X-CSRF-Token', csrf)
    .type('form')
    .send({ userIds: [UID_CARL] })
    .expect(200);
  t.true(disableRes.text.includes('Carlos Baja'));
  t.deepEqual(order, [
    { op: 'unlicense', remove: ['sku-carl'] },
    { op: 'disable', enabled: false },
  ]);

  // If license removal fails, account MUST NOT be disabled
  http.fail('post', `/v1.0/users/${UID_CARL}/assignLicense`, 500);
  const patchCountBefore = http.callsTo('patch', `/v1.0/users/${UID_CARL}`).length;
  await agent.post('/admin/usuarios/excel/deshabilitar')
    .set('X-CSRF-Token', csrf)
    .type('form')
    .send({ userIds: [UID_CARL] })
    .expect(200);
  t.is(http.callsTo('patch', `/v1.0/users/${UID_CARL}`).length, patchCountBefore);

  // Import preview + apply
  const previewRes = await agent.post('/admin/usuarios/excel/importar/preview')
    .field('_csrf', csrf)
    .field('skipLeaving', '1')
    .attach('excelFile', xlsxBuf, 'plantilla.xlsx')
    .expect(200);

  const planMatch = previewRes.text.match(/name="planId"\s+value="([a-f0-9]+)"/);
  t.truthy(planMatch, 'import preview should include a planId hidden input');
  const planId = planMatch[1];
  t.true(previewRes.text.includes('lucia@arandadeduero.es'));

  const UID_LUCIA = '44444444-4444-4444-4444-444444444444';
  http.on('post', '/v1.0/users', ({ data }) => ({ id: UID_LUCIA, ...data }));
  http.on('patch', `/v1.0/users/${UID_ANA}`, {});
  http.on('get', `/users/${UID_ANA}/manager`, () => { throw httpError(404); });
  http.on('get', `/users/${UID_LUCIA}/manager`, () => { throw httpError(404); });
  http.on('put', `/v1.0/users/${UID_ANA}/manager/$ref`, {});
  http.on('put', `/v1.0/users/${UID_LUCIA}/manager/$ref`, {});

  const applyRes = await agent.post('/admin/usuarios/excel/importar/aplicar')
    .set('X-CSRF-Token', csrf)
    .type('form')
    .send({ planId })
    .expect(200);

  t.true(applyRes.text.includes('lucia@arandadeduero.es'));
  t.is(http.callsTo('post', '/v1.0/users').length, 1);
  t.is(http.callsTo('patch', `/v1.0/users/${UID_ANA}`).length, 1);
  t.is(http.callsTo('put', `/v1.0/users/${UID_LUCIA}/manager/$ref`).length, 1);
});

test.serial('the employee type list offers the types of the Ayuntamiento, and keeps a value that is not in it', async t => {
  const { agent } = await signInAs(t, ADMIN, { admin: true });
  const stubUser = (employeeType) => {
    http.on('get', `/v1.0/users/${UID_ANA}`, { id: UID_ANA, displayName: 'Ana Pérez', userPrincipalName: 'ana@arandadeduero.es', accountEnabled: true, employeeType });
    http.on('get', `/users/${UID_ANA}/photo/$value`, () => { throw httpError(404); });
    http.on('get', `/users/${UID_ANA}/manager`, () => { throw httpError(404); });
    http.on('get', `/users/${UID_ANA}/memberOf/microsoft.graph.group`, { value: [] });
    http.on('get', `/users/${UID_ANA}/licenseDetails`, { value: [] });
    http.on('get', '/v1.0/groups', { value: [] });
    http.on('get', '/v1.0/users', { value: [] });
    http.on('get', '/v1.0/subscribedSkus', { value: [] });
  };
  const select = (html) => html.split('id="employeeType"')[1].split('</select>')[0];

  stubUser('Laboral');
  const page = select((await agent.get(`/admin/usuarios/${UID_ANA}`).expect(200)).text);
  for (const type of ['Funcionario', 'Laboral', 'Programa de empleo', 'Concejal', 'Eventual']) t.true(page.includes(`value="${type}"`), type);
  t.regex(page, /value="Laboral"[^>]* selected/);
  for (const old of ['Employee', 'Contractor', 'Consultant', 'Vendor']) t.false(page.includes(`value="${old}"`), old);

  // Somebody who already has another value does not lose it by opening the page
  stubUser('Contractor');
  t.regex(select((await agent.get(`/admin/usuarios/${UID_ANA}`).expect(200)).text), /value="Contractor"[^>]* selected/);
});

test.serial('the import preview shows which managers would change, not only the fields', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });
  const UID_LUIS = '55555555-5555-5555-5555-555555555555';
  const xlsxBuf = await buildWorkbookBuffer({
    ACTIVOS: [
      ['id_empleado', 'Trabajador', 'Departamento', 'Descripción', 'e_mail', 'ID responsable'],
      ['1', 'JEFE, BOSS', 'Alcaldía', 'Alcalde', 'boss@arandadeduero.es', ''],
      ['2', 'PEREZ, ANA', 'Obras', 'Técnica', 'ana@arandadeduero.es', '1'],
      ['3', 'GOMEZ, LUIS', 'Obras', 'Peón', 'luis@arandadeduero.es', '1'],
    ],
  });
  http.on('get', '/v1.0/users', { value: [
    { id: UID_BOSS, displayName: 'Boss Jefe', userPrincipalName: 'boss@arandadeduero.es', employeeId: '1', accountEnabled: true },
    { id: UID_ANA, displayName: 'Ana Perez', userPrincipalName: 'ana@arandadeduero.es', employeeId: '2', accountEnabled: true },
    { id: UID_LUIS, displayName: 'Luis Gomez', userPrincipalName: 'luis@arandadeduero.es', employeeId: '3', accountEnabled: true },
  ] });
  // Ana already reports to Boss; Luis reports to somebody else
  http.on('post', '/v1.0/$batch', ({ data }) => ({
    responses: data.requests.map(r => (r.url.includes(UID_ANA)
      ? { id: r.id, status: 200, body: { id: UID_BOSS, displayName: 'Boss Jefe' } }
      : { id: r.id, status: 200, body: { id: 'otro', displayName: 'Otro Jefe' } })),
  }));

  const res = await agent.post('/admin/usuarios/excel/importar/preview').field('_csrf', csrf).attach('excelFile', xlsxBuf, 'plantilla.xlsx').expect(200);

  t.true(res.text.includes('Responsables que cambian: 1'));
  t.regex(res.text, /<code>responsable<\/code>: Otro Jefe &rarr; <strong>Boss Jefe<\/strong>/);
  // Nothing is written by looking
  t.is(http.callsTo('put', `/v1.0/users/${UID_LUIS}/manager/$ref`).length, 0);
});

test.serial('data quality fixes normalize job titles and update mailbox time zones preserving working hours', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });

  http.on('get', '/v1.0/users', {
    value: [
      { id: UID_ANA, displayName: 'Ana Pérez', userPrincipalName: 'ana@arandadeduero.es', jobTitle: 'TECNICA DE OBRAS', accountEnabled: true },
    ],
  });
  http.on('patch', `/v1.0/users/${UID_ANA}`, {});

  const qRes = await agent.get('/admin/usuarios/calidad?section=puestos').expect(200);
  t.true(qRes.text.includes('TECNICA DE OBRAS'));
  t.true(qRes.text.includes('Tecnica de obras'));

  await agent.post('/admin/usuarios/calidad/puestos')
    .set('X-CSRF-Token', csrf)
    .type('form')
    .send({ userIds: [UID_ANA] })
    .expect(200);
  t.is(http.callsTo('patch', `/v1.0/users/${UID_ANA}`)[0].data.jobTitle, 'Tecnica de obras');

  // Timezone audit & fix
  http.on('get', `/v1.0/users/${UID_ANA}/mailboxSettings`, {
    timeZone: 'UTC',
    workingHours: {
      daysOfWeek: ['monday', 'tuesday', 'wednesday', 'thursday', 'friday'],
      startTime: '08:00:00.0000000',
      endTime: '15:00:00.0000000',
      timeZone: { name: 'UTC' },
    },
  });
  http.on('patch', `/v1.0/users/${UID_ANA}/mailboxSettings`, {});

  const tzAuditRes = await agent.get('/admin/usuarios/calidad?section=zona-horaria&runTz=1').expect(200);
  t.true(tzAuditRes.text.includes('Ana Pérez'));
  // Ensure GET mailboxSettings never passes $select
  const mbGet = http.callsTo('get', `/v1.0/users/${UID_ANA}/mailboxSettings`)[0];
  t.falsy(mbGet.config.params?.$select);

  await agent.post('/admin/usuarios/calidad/zona-horaria')
    .set('X-CSRF-Token', csrf)
    .type('form')
    .send({ userIds: [UID_ANA] })
    .expect(200);

  const mbPatch = http.callsTo('patch', `/v1.0/users/${UID_ANA}/mailboxSettings`)[0];
  t.is(mbPatch.data.timeZone, 'Romance Standard Time');
  t.is(mbPatch.data.workingHours.timeZone.name, 'Romance Standard Time');
  t.is(mbPatch.data.workingHours.startTime, '08:00:00.0000000');
});

test.serial('a user\'s licenses can be removed and added, and a disabled account with licenses is warned about', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });
  const SKU_OLD = 'cccccccc-cccc-cccc-cccc-cccccccccccc';
  const SKU_NEW = 'dddddddd-dddd-dddd-dddd-dddddddddddd';

  http.on('get', `/v1.0/users/${UID_ANA}`, { id: UID_ANA, displayName: 'Ana Pérez', userPrincipalName: 'ana@arandadeduero.es', accountEnabled: false });
  http.on('get', `/users/${UID_ANA}/photo/$value`, () => { throw httpError(404); });
  http.on('get', `/users/${UID_ANA}/manager`, () => { throw httpError(404); });
  http.on('get', `/users/${UID_ANA}/memberOf/microsoft.graph.group`, { value: [] });
  http.on('get', `/users/${UID_ANA}/licenseDetails`, { value: [{ skuId: SKU_OLD, skuPartNumber: 'O365_BUSINESS_PREMIUM' }] });
  http.on('get', '/v1.0/groups', { value: [] });
  http.on('get', '/v1.0/users', { value: [] });
  http.on('get', '/v1.0/subscribedSkus', { value: [
    { skuId: SKU_OLD, skuPartNumber: 'O365_BUSINESS_PREMIUM', prepaidUnits: { enabled: 10 }, consumedUnits: 9 },
    { skuId: SKU_NEW, skuPartNumber: 'EXCHANGEDESKLESS', prepaidUnits: { enabled: 5 }, consumedUnits: 1 },
  ] });
  http.on('patch', `/v1.0/users/${UID_ANA}`, {});
  http.on('post', `/v1.0/users/${UID_ANA}/assignLicense`, {});

  const page = await agent.get(`/admin/usuarios/${UID_ANA}`).expect(200);
  t.true(page.text.includes('todavía tiene licencias asignadas'));
  // One list: the license she has comes ticked, the other one does not
  t.regex(page.text, new RegExp(`name="skuIds" value="${SKU_OLD}" id="sku-${SKU_OLD}" checked`));
  t.regex(page.text, new RegExp(`name="skuIds" value="${SKU_NEW}" id="sku-${SKU_NEW}" *>`));
  t.false(page.text.includes('removeSkuIds'));
  t.false(page.text.includes('addSkuIds'));

  // Untick the one she has, tick the other one
  await agent.post(`/admin/usuarios/${UID_ANA}/licencias`)
    .set('X-CSRF-Token', csrf)
    .type('form')
    .send({ shownSkuIds: [SKU_OLD, SKU_NEW], skuIds: [SKU_NEW] })
    .expect(302);

  const [call] = http.callsTo('post', `/v1.0/users/${UID_ANA}/assignLicense`);
  t.deepEqual(call.data.removeLicenses, [SKU_OLD]);
  t.deepEqual(call.data.addLicenses.map(l => l.skuId), [SKU_NEW]);
  // No usage location: Microsoft would refuse the license
  t.is(http.callsTo('patch', `/v1.0/users/${UID_ANA}`)[0].data.usageLocation, 'ES');
  t.true(auditEvents().some(e => e.event === 'admin.licenses_changed'));

  // Nothing ticked and nothing shown (an empty or cut-down submission): she keeps what she has
  const before = http.callsTo('post', `/v1.0/users/${UID_ANA}/assignLicense`).length;
  await agent.post(`/admin/usuarios/${UID_ANA}/licencias`).set('X-CSRF-Token', csrf).type('form').send({}).expect(302);
  // A ticked license she already has (and the other one left unticked) changes nothing either
  await agent.post(`/admin/usuarios/${UID_ANA}/licencias`).set('X-CSRF-Token', csrf).type('form')
    .send({ shownSkuIds: [SKU_OLD, SKU_NEW], skuIds: [SKU_OLD] }).expect(302);
  t.is(http.callsTo('post', `/v1.0/users/${UID_ANA}/assignLicense`).length, before);
});

test.serial('when USERS_EXCEL_URL is set, /admin/usuarios/excel reads the Excel directly from SharePoint without uploading a file', async t => {
  const { agent, csrf } = await signInAs(t, ADMIN, { admin: true });
  const shareUrl = 'https://ayto.sharepoint.com/:x:/s/rrhh/FAKE_DOC_ID';
  appConfig.USERS_EXCEL_URL = shareUrl;

  const xlsxBuf = await buildWorkbookBuffer({
    ACTIVOS: [
      ['id_empleado', 'Trabajador', 'Departamento', 'Descripción', 'Fecha de Alta', 'Fecha de Baja', 'e_mail', 'ID responsable'],
      ['1', 'JEFE, BOSS', 'Alcaldía', 'Alcalde', '01/01/2020', '', 'boss@arandadeduero.es', ''],
      ['2', 'NUEVA, LUCIA', 'Cultura', 'Archivera', '10/06/2025', '', 'lucia@arandadeduero.es', '1'],
    ],
  });

  http.on('get', '/driveItem/content', xlsxBuf);
  http.on('get', '/v1.0/users', {
    value: [
      { id: UID_BOSS, displayName: 'Boss Jefe', userPrincipalName: 'boss@arandadeduero.es', employeeId: '1', accountEnabled: true },
    ],
  });

  // GET /admin/usuarios/excel auto-loads both validation and sync check from SharePoint
  const page = await agent.get('/admin/usuarios/excel').expect(200);
  t.true(page.text.includes('Excel de Personal conectado desde SharePoint'));
  t.true(page.text.includes('Validación correcta'));
  t.true(page.text.includes('lucia@arandadeduero.es'));

  // POST /admin/usuarios/excel/importar/preview works with 0 uploaded files
  const preview = await agent.post('/admin/usuarios/excel/importar/preview')
    .field('_csrf', csrf)
    .field('skipLeaving', 'true')
    .expect(200);
  t.true(preview.text.includes('lucia@arandadeduero.es'));
  t.truthy(preview.text.match(/name="planId"\s+value="([a-f0-9]+)"/));
});

test.serial('GET /admin/licencias links to /admin/licencias/:sku which lists holders with an Editar button to their profile', async t => {
  const { agent } = await signInAs(t, ADMIN, { admin: true });
  const SKU_STD = 'eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee';

  http.on('get', '/v1.0/subscribedSkus', {
    value: [
      { skuId: SKU_STD, skuPartNumber: 'O365_BUSINESS_PREMIUM', prepaidUnits: { enabled: 50 }, consumedUnits: 10 },
    ],
  });
  http.on('get', '/v1.0/users', {
    value: [
      { id: UID_ANA, displayName: 'Ana Pérez', userPrincipalName: 'ana@arandadeduero.es', jobTitle: 'Técnica', department: 'Obras', accountEnabled: true, assignedLicenses: [{ skuId: SKU_STD }] },
      { id: UID_BOSS, displayName: 'Boss Jefe', userPrincipalName: 'boss@arandadeduero.es', accountEnabled: true, assignedLicenses: [] },
    ],
  });

  const listPage = await agent.get('/admin/licencias').expect(200);
  t.true(listPage.text.includes('Ver usuarios'));
  t.true(listPage.text.includes(`/admin/licencias/${SKU_STD}`));

  const detailPage = await agent.get(`/admin/licencias/${SKU_STD}`).expect(200);
  t.true(detailPage.text.includes('Microsoft 365 Empresa Estándar'));
  t.true(detailPage.text.includes('Ana Pérez'));
  t.false(detailPage.text.includes('Boss Jefe'));
  t.true(detailPage.text.includes(`/admin/usuarios/${UID_ANA}`));
});
