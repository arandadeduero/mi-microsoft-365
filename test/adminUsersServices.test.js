require('./helpers/env');

const ExcelJS = require('exceljs');
const test = require('ava').default;
const {
  cellToString, readExcelBuffer, isEffectiveLeaving, validateExcelRows,
  analyzeSyncCheck, parseImportRows, buildImportPlan, previewManagerChanges,
} = require('../services/adminUsersExcel');
const {
  isRoomUser, generateTempPassword, buildResetCredentialLinks, filterAdminUsers, buildMailingList,
  findJobTitleFixes, analyzeOrgChartIntegrity, classifyMailboxTimezone, buildTimezonePatchPayload,
} = require('../services/adminUsersTools');

const TODAY = new Date(2026, 5, 15);

test('cellToString reads dates in UTC, links, rich text and formulas', t => {
  t.is(cellToString(new Date(Date.UTC(2024, 0, 5))), '05/01/2024');
  t.is(cellToString({ text: ' a ', hyperlink: 'x' }), 'a');
  t.is(cellToString({ richText: [{ text: 'a' }, { text: 'b ' }] }), 'ab');
  t.is(cellToString({ formula: '1+1', result: 2 }), '2');
  t.is(cellToString(null), '');
});

test('readExcelBuffer needs ACTIVOS and reads CONCEJALES when present', async t => {
  const wb = new ExcelJS.Workbook();
  wb.addWorksheet('ACTIVOS').addRows([['id_empleado', 'e_mail'], ['1', 'a@arandadeduero.es']]);
  wb.addWorksheet('CONCEJALES').addRows([['email'], ['c@arandadeduero.es']]);
  const read = await readExcelBuffer(Buffer.from(await wb.xlsx.writeBuffer()));
  t.deepEqual(read.activos, [{ id_empleado: '1', e_mail: 'a@arandadeduero.es' }]);
  t.true(read.hasConcejales);

  const empty = new ExcelJS.Workbook();
  empty.addWorksheet('OTRA');
  await t.throwsAsync(readExcelBuffer(Buffer.from(await empty.xlsx.writeBuffer())), { message: /ACTIVOS/ });
});

test('a departure counts only from the day after Fecha de Baja', t => {
  t.true(isEffectiveLeaving('14/06/2026', TODAY));
  t.false(isEffectiveLeaving('15/06/2026', TODAY));
  t.false(isEffectiveLeaving('16/06/2026', TODAY));
  t.false(isEffectiveLeaving('31/02/2026', TODAY));
});

test('validateExcelRows applies the five rules', t => {
  const result = validateExcelRows([
    { id_empleado: '1', e_mail: 'a@arandadeduero.es', 'ID responsable': '' },
    { id_empleado: '1', e_mail: 'A@arandadeduero.es', 'ID responsable': '9', 'Fecha de Baja': '2026-01-01' },
    { id_empleado: '', e_mail: 'x@gmail.com' },
  ]);
  const ids = result.byRule.map(rule => rule.ruleId);
  t.deepEqual(ids, ['id_empleado-required-unique', 'email-required-domain', 'email-unique', 'id_responsable-exists', 'fecha-de-baja-format']);
  t.is(result.byRule[0].issues[0].row, 3);
});

test('sync check finds missing, not disabled, cloud-only and special accounts', t => {
  const activos = [
    { Trabajador: 'Ana', e_mail: 'ana@arandadeduero.es' },
    { Trabajador: 'Nueva', e_mail: 'nueva@arandadeduero.es' },
    { Trabajador: 'Baja', e_mail: 'baja@arandadeduero.es', 'Fecha de Baja': '01/01/2026' },
    { Trabajador: 'Hoy', e_mail: 'hoy@arandadeduero.es', 'Fecha de Baja': '15/06/2026' },
    { Trabajador: 'SinCorreo', e_mail: '' },
  ];
  const cloud = [
    { id: '1', userPrincipalName: 'ana@arandadeduero.es', accountEnabled: true },
    { id: '2', userPrincipalName: 'baja@arandadeduero.es', accountEnabled: true, displayName: 'Baja' },
    { id: '3', userPrincipalName: 'hoy@arandadeduero.es', accountEnabled: true },
    { id: '4', userPrincipalName: 'concejal@arandadeduero.es', accountEnabled: true },
    { id: '5', userPrincipalName: 'sala-plenos@arandadeduero.es', accountEnabled: true },
    { id: '6', userPrincipalName: 'asac@arandadeduero.es', accountEnabled: true },
    { id: '7', userPrincipalName: 'otro@arandadeduero.es', accountEnabled: true },
  ];
  const result = analyzeSyncCheck(activos, [{ email: 'concejal@arandadeduero.es' }], cloud, 'arandadeduero.es', { today: TODAY });
  t.deepEqual(result.missing.map(i => i.email), ['nueva@arandadeduero.es']);
  t.deepEqual(result.notDisabled.map(i => i.userId), ['2']);
  t.is(result.noEmail.length, 1);
  t.deepEqual(result.cloudOnly.map(i => i.userPrincipalName), ['otro@arandadeduero.es']);
  t.deepEqual(result.specialFound.map(i => i.userPrincipalName), ['asac@arandadeduero.es']);
  t.is(result.totalIssues, 4);
});

test('import parses Spanish headers and plans create, update and ready without re-enabling', t => {
  const parsed = parseImportRows([
    { id_empleado: '2', Trabajador: 'PEREZ GARCIA, ANA', Departamento: 'Obras', 'Descripción': 'Técnica', 'Fecha de Alta': '15/03/2022', e_mail: 'ana@arandadeduero.es', 'ID responsable': '1' },
    { id_empleado: '3', Trabajador: 'OTRA, PERSONA', e_mail: 'otra@arandadeduero.es' },
    { id_empleado: '4', Trabajador: 'NUEVA, LUCIA', e_mail: 'lucia@arandadeduero.es' },
    { id_empleado: '5', Trabajador: 'MAL, CORREO', e_mail: 'mal@gmail.com' },
    { id_empleado: '6', Trabajador: 'BAJA, B', e_mail: 'b@arandadeduero.es', 'Fecha de Baja': '01/01/2020' },
  ], 'arandadeduero.es', { today: TODAY });
  t.is(parsed.validRows.length, 3);
  t.is(parsed.invalidRows.length, 1);
  t.is(parsed.skippedLeaving.length, 1);
  t.is(parsed.validRows[0].payload.displayName, 'Ana Perez Garcia');
  t.is(parsed.validRows[0].payload.employeeHireDate, '2022-03-15T00:00:00.000Z');

  const plan = buildImportPlan(parsed.validRows, [
    { id: 'u1', userPrincipalName: 'ana@arandadeduero.es', displayName: 'Ana Perez Garcia', jobTitle: 'Antiguo', accountEnabled: false },
    { id: 'u2', userPrincipalName: 'otra@arandadeduero.es', displayName: 'Persona Otra', givenName: 'Persona', surname: 'Otra', employeeId: '3', accountEnabled: true },
  ]);
  t.deepEqual(plan.items.map(i => i.action), ['update', 'ready', 'create']);
  t.false('accountEnabled' in plan.items[0].updatePayload);
  t.false('mailNickname' in plan.items[0].updatePayload);
  t.deepEqual(plan.summary, { total: 3, create: 1, update: 1, ready: 1, managerChanges: 1 });
});

test('the import plan says which managers really change', t => {
  const { validRows } = parseImportRows([
    { id_empleado: '2', Trabajador: 'PEREZ, ANA', e_mail: 'ana@arandadeduero.es', 'ID responsable': '1' },
    { id_empleado: '3', Trabajador: 'GOMEZ, LUIS', e_mail: 'luis@arandadeduero.es', 'ID responsable': 'BOSS@arandadeduero.es' },
    { id_empleado: '4', Trabajador: 'NUEVA, LUCIA', e_mail: 'lucia@arandadeduero.es', 'ID responsable': '2' },
    { id_empleado: '5', Trabajador: 'NUEVO, PEDRO', e_mail: 'pedro@arandadeduero.es', 'ID responsable': '99' },
    { id_empleado: '6', Trabajador: 'NUEVA, EVA', e_mail: 'eva@arandadeduero.es', 'ID responsable': 'lucia@arandadeduero.es' },
    { id_empleado: '7', Trabajador: 'SIN, RESPONSABLE', e_mail: 'sr@arandadeduero.es' },
  ], 'arandadeduero.es', { today: TODAY });
  const existing = [
    { id: 'b', userPrincipalName: 'boss@arandadeduero.es', displayName: 'Boss Jefe', employeeId: '1' },
    { id: 'a', userPrincipalName: 'ana@arandadeduero.es', displayName: 'Ana Pérez', employeeId: '2' },
    { id: 'l', userPrincipalName: 'luis@arandadeduero.es', displayName: 'Luis Gómez', employeeId: '3' },
  ];
  const plan = buildImportPlan(validRows, existing);
  previewManagerChanges(plan, existing, new Map([['a', { id: 'b', displayName: 'Boss Jefe' }], ['l', { id: 'x', displayName: 'Otro Jefe' }]]), 'arandadeduero.es');

  const change = (upn) => plan.items.find(item => item.upn === upn).managerChange;
  t.is(change('ana@arandadeduero.es'), undefined); // already has that manager
  t.deepEqual(change('luis@arandadeduero.es'), { status: 'change', from: 'Otro Jefe', to: 'Boss Jefe', newAccount: false });
  t.deepEqual(change('lucia@arandadeduero.es'), { status: 'change', from: null, to: 'Ana Pérez', newAccount: false });
  t.deepEqual(change('pedro@arandadeduero.es'), { status: 'notfound', ref: '99' });
  // A manager who is created by the same file is found too
  t.deepEqual(change('eva@arandadeduero.es'), { status: 'change', from: null, to: 'Lucia Nueva', newAccount: true });
  t.is(change('sr@arandadeduero.es'), undefined);
  // The count is of real changes, no longer of rows that name a manager
  t.is(plan.summary.managerChanges, 3);
  t.is(plan.summary.managerNotFound, 1);
});

test('the credentials email names the welcome guide only when there is a link for it', t => {
  const without = decodeURIComponent(buildResetCredentialLinks('ana@arandadeduero.es', 'Pw-1234-Pw').mailtoUrl);
  t.false(without.includes('guía de bienvenida'));
  t.true(without.includes('password: Pw-1234-Pw'));
  const withGuide = decodeURIComponent(buildResetCredentialLinks('ana@arandadeduero.es', 'Pw-1234-Pw', { welcomeUrl: 'https://guia.example/bienvenida' }).mailtoUrl);
  t.true(withGuide.includes('Sigue la guía de bienvenida en este enlace https://guia.example/bienvenida'));
});

test('passwords and credential links', t => {
  t.regex(generateTempPassword(), /^\S+-\d{4}-\S+$/);
  const links = buildResetCredentialLinks('ana@arandadeduero.es', 'Pw-1234-Pw');
  t.true(links.mailtoUrl.startsWith('mailto:ana%40arandadeduero.es'));
  t.true(links.outlookWebUrl.startsWith('https://outlook.office.com/mail/deeplink/compose'));
});

test('directory filters, mailing list and job title fixes', t => {
  const roomSet = new Set(['reuniones-planta2@x']);
  t.true(isRoomUser({ userPrincipalName: 'reuniones-planta2@x' }, roomSet));
  t.true(isRoomUser({ userPrincipalName: 'sala-plenos@x' }, roomSet));
  t.false(isRoomUser({ userPrincipalName: 'ana@x' }, roomSet));

  const users = [
    { id: '1', displayName: 'Ana', userPrincipalName: 'ana@x', accountEnabled: true, licenses: [{}], groups: [], manager: null, jobTitle: 'TECNICA' },
    { id: '2', displayName: 'Bea', userPrincipalName: 'bea@x', accountEnabled: true, licenses: [], groups: [{ displayName: 'Obras' }], manager: { id: '1' }, jobTitle: 'Jefa' },
    { id: '3', displayName: 'Sala 2', userPrincipalName: 'reuniones-planta2@x', accountEnabled: true, isRoom: true, licenses: [{}], groups: [], manager: null },
  ];
  t.deepEqual(filterAdminUsers(users, { missingManager: true }).map(u => u.id), ['1']);
  t.deepEqual(filterAdminUsers(users, { missingManager: true, includeRooms: true }).map(u => u.id), ['1', '3']);
  t.deepEqual(filterAdminUsers(users, { q: 'bea' }).map(u => u.id), ['2']);
  t.is(buildMailingList(users), 'ana@x');
  t.deepEqual(findJobTitleFixes(users).map(f => [f.current, f.fixed]), [['TECNICA', 'Tecnica']]);
});

test('org chart integrity finds orphans, wrong roots and cycles', t => {
  const users = ['r', 'a', 'b', 'c', 'd', 'e'].map(id => ({ id, displayName: id, userPrincipalName: `${id}@x` }));
  const managers = new Map([['r', null], ['a', 'r'], ['b', null], ['c', 'd'], ['d', 'c'], ['e', 'a']]);
  const result = analyzeOrgChartIntegrity(users, managers, 'r');
  t.deepEqual(result.noManager.map(u => u.id), ['b']);
  t.is(result.cycles.length, 1);
  // The root itself counts as valid, with a and e
  t.is(result.validCount, 3);
});

test('mailbox time zone classification and patch preserve working hours', t => {
  const settings = { timeZone: 'UTC', workingHours: { daysOfWeek: ['monday'], startTime: '08:00:00.0000000', endTime: '15:00:00.0000000', timeZone: { name: 'UTC' } } };
  t.false(classifyMailboxTimezone(settings).ok);
  t.true(classifyMailboxTimezone({ timeZone: 'Romance Standard Time', workingHours: null }).ok);
  const patch = buildTimezonePatchPayload(settings);
  t.is(patch.timeZone, 'Romance Standard Time');
  t.is(patch.workingHours.startTime, '08:00:00.0000000');
  t.deepEqual(buildTimezonePatchPayload({ timeZone: 'UTC' }), { timeZone: 'Romance Standard Time' });
});
