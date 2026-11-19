/**
 * Excel Drift Tests
 *
 * The daily read-only check of the users Excel: what it counts, that it keeps
 * only numbers, that a failure is kept too, when it does not run, and the
 * notice it leaves on /admin/usuarios and /admin/usuarios/excel.
 */

require('./helpers/env');

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const test = require('ava').default;
const request = require('supertest');
const { stubAxios, stubMsal, signInAgent, getCsrfToken, auditEvents, httpError } = require('./helpers/stubs');
const { resetGraphCaches } = require('../services/graphService');

const http = stubAxios();
const msal = stubMsal();

const appConfig = require('../config/appConfig');
const permissions = require('../config/permissions');
const app = require('../server');
const { runExcelDriftCheck, readExcelDrift } = require('../services/excelDrift');

const ROLES = '/me/memberOf/microsoft.graph.directoryRole';
const ADMIN = { username: 'admin@arandadeduero.es', name: 'Ada Admin', localAccountId: 'admin-id-1' };
const UID_BOSS = '22222222-2222-2222-2222-222222222222';
const UID_ANA = '11111111-1111-1111-1111-111111111111';
const UID_LUIS = '55555555-5555-5555-5555-555555555555';
const UID_PEDRO = '66666666-6666-6666-6666-666666666666';
const SHARE_URL = 'https://ayto.sharepoint.com/:x:/s/rrhh/FAKE_DOC_ID';
const driftFile = () => path.join(process.env.DATA_DIR, 'excel-drift.json');

async function workbook(sheets) {
  const wb = new ExcelJS.Workbook();
  for (const [name, rows] of Object.entries(sheets)) {
    const ws = wb.addWorksheet(name);
    rows.forEach(row => ws.addRow(row));
  }
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// Boss and Ana are as the Excel says except Ana's job title; Lucia is new; Luis left; Pedro is not in the Excel
async function stubDirectoryAndExcel() {
  const excel = await workbook({
    ACTIVOS: [
      ['id_empleado', 'Trabajador', 'Departamento', 'Descripción', 'Fecha de Baja', 'e_mail', 'ID responsable'],
      ['1', 'JEFE, BOSS', 'Alcaldía', 'Alcalde', '', 'boss@arandadeduero.es', ''],
      ['2', 'PEREZ, ANA', 'Obras', 'Técnica nueva', '', 'ana@arandadeduero.es', '1'],
      ['3', 'GOMEZ, LUIS', 'Obras', 'Peón', '01/01/2020', 'luis@arandadeduero.es', '1'],
      ['4', 'NUEVA, LUCIA', 'Cultura', 'Archivera', '', 'lucia@arandadeduero.es', '1'],
    ],
  });
  http.on('get', '/driveItem/content', excel);
  http.on('get', '/v1.0/users', { value: [
    { id: UID_BOSS, displayName: 'Boss Jefe', givenName: 'Boss', surname: 'Jefe', userPrincipalName: 'boss@arandadeduero.es', employeeId: '1', jobTitle: 'Alcalde', department: 'Alcaldía', accountEnabled: true },
    { id: UID_ANA, displayName: 'Ana Perez', givenName: 'Ana', surname: 'Perez', userPrincipalName: 'ana@arandadeduero.es', employeeId: '2', jobTitle: 'Técnica', department: 'Obras', accountEnabled: true },
    { id: UID_LUIS, displayName: 'Luis Gomez', userPrincipalName: 'luis@arandadeduero.es', employeeId: '3', accountEnabled: true },
    { id: UID_PEDRO, displayName: 'Pedro Suelto', userPrincipalName: 'pedro@arandadeduero.es', employeeId: '9', accountEnabled: true },
  ] });
  // Ana already reports to Boss
  http.on('post', '/v1.0/$batch', ({ data }) => ({
    responses: data.requests.map(r => (r.url.includes(UID_ANA)
      ? { id: r.id, status: 200, body: { id: UID_BOSS, displayName: 'Boss Jefe' } }
      : { id: r.id, status: 404, body: {} })),
  }));
}

test.beforeEach(() => {
  http.reset();
  resetGraphCaches();
  fs.rmSync(driftFile(), { force: true });
  permissions.DOMAIN = 'arandadeduero.es';
  appConfig.USERS_EXCEL_URL = '';
});

test.serial('the check counts what differs and keeps only numbers', async t => {
  appConfig.USERS_EXCEL_URL = SHARE_URL;
  await stubDirectoryAndExcel();

  const result = await runExcelDriftCheck();

  t.like(result, { toDisable: 1, toCreate: 1, notInExcel: 1, fieldChanges: 1, managerChanges: 1, managerNotFound: 0 });
  t.deepEqual(readExcelDrift(), result);
  // No names, no addresses on disk
  const raw = fs.readFileSync(driftFile(), 'utf8');
  for (const text of ['Lucia', 'Luis', 'Pedro', 'Ana', '@arandadeduero.es']) t.false(raw.includes(text), text);
  t.true(auditEvents('excel.drift_checked').some(e => e.toDisable === 1 && e.toCreate === 1));
  // Looking changes nothing
  t.is(http.calls.filter(c => ['post', 'patch', 'put', 'delete'].includes(c.method) && !c.url.endsWith('/$batch') && !c.url.includes('/oauth2/')).length, 0);
});

test.serial('a check that cannot read the file is kept too, saying so', async t => {
  appConfig.USERS_EXCEL_URL = SHARE_URL;
  http.on('get', '/driveItem/content', () => { throw httpError(403); });

  const result = await runExcelDriftCheck();

  t.regex(result.error, /403/);
  t.is(result.toDisable, undefined);
  t.deepEqual(readExcelDrift(), result);
  t.true(auditEvents('excel.drift_failed').some(e => e.status === 403));
});

test.serial('without an Excel nothing runs, and a recent check is not repeated unless forced', async t => {
  t.is(await runExcelDriftCheck(), null);
  t.is(readExcelDrift(), null);

  appConfig.USERS_EXCEL_URL = SHARE_URL;
  await stubDirectoryAndExcel();
  const first = await runExcelDriftCheck();
  const calls = http.calls.length;

  t.deepEqual(await runExcelDriftCheck(), first);
  t.is(http.calls.length, calls);
  await runExcelDriftCheck({ force: true });
  t.true(http.calls.length > calls);
});

test.serial('the notice shows the numbers, says when things match, and admits a failed check', async t => {
  msal.account = ADMIN;
  const agent = request.agent(app);
  await signInAgent(agent);
  t.teardown(() => agent.get('/auth/logout'));
  http.on('get', ROLES, { value: [{ displayName: 'Global Administrator' }] });
  await getCsrfToken(agent);
  http.on('get', '/v1.0/users', { value: [] });

  const write = (data) => fs.writeFileSync(driftFile(), JSON.stringify({ checkedAt: '2026-11-19T03:00:00.000Z', ...data }));

  t.notRegex((await agent.get('/admin/usuarios/excel').expect(200)).text, /comprobación/i);

  write({ toDisable: 2, toCreate: 1, fieldChanges: 5, managerChanges: 3, notInExcel: 0 });
  for (const page of ['/admin/usuarios/excel', '/admin/usuarios']) {
    const res = await agent.get(page).expect(200);
    t.regex(res.text, /no coinciden/, page);
    t.regex(res.text, /2 bajas por aplicar · 1 altas por crear · 5 personas con otros datos · 3 responsables distintos/, page);
    t.notRegex(res.text, /cuentas activas que no están en el Excel/, page);
  }

  write({ toDisable: 0, toCreate: 0, fieldChanges: 0, managerChanges: 0, notInExcel: 0 });
  t.regex((await agent.get('/admin/usuarios/excel')).text, /El Excel y Microsoft 365 coinciden/);

  write({ error: 'Microsoft Graph ha respondido 403' });
  const failed = (await agent.get('/admin/usuarios/excel')).text;
  t.regex(failed, /no ha funcionado: Microsoft Graph ha respondido 403/);
});
