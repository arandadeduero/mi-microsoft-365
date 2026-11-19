/**
 * Excel workflows of Administración → Usuarios (ported from m365-tools):
 * reading the .xlsx, validating it offline, comparing it with the cloud
 * accounts and planning an import. Pure functions apart from the file read.
 */

const ExcelJS = require('exceljs');
const { compareEs } = require('./formatters');
const { SPECIAL_ACCOUNTS, normalizeUpn, isRoomUser } = require('./adminUsersTools');

const DEFAULT_DOMAIN = 'arandadeduero.es';
const DATE_DDMMYYYY = /^(\d{2})\/(\d{2})\/(\d{4})$/;

/**
 * Cell value as a trimmed string. Dates use UTC parts so the time zone of
 * the server never moves them a day.
 */
function cellToString(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) {
    const dd = String(value.getUTCDate()).padStart(2, '0');
    const mm = String(value.getUTCMonth() + 1).padStart(2, '0');
    return `${dd}/${mm}/${value.getUTCFullYear()}`;
  }
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) return value.richText.map(part => part.text || '').join('').trim();
    if ('hyperlink' in value && 'text' in value) return cellToString(value.text);
    if ('formula' in value || 'sharedFormula' in value) return cellToString(value.result);
    if ('text' in value) return cellToString(value.text);
    if ('error' in value) return '';
  }
  return String(value).trim();
}

function sheetToRows(worksheet) {
  const rows = [];
  let headers = null;
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    const values = row.values.slice(1).map(cellToString);
    if (!values.some(Boolean)) return;
    if (!headers) {
      headers = values;
      return;
    }
    const entry = {};
    headers.forEach((header, index) => {
      if (header) entry[header] = values[index] || '';
    });
    rows.push(entry);
  });
  return rows;
}

/**
 * @param {Buffer} buffer - .xlsx contents
 * @returns {Promise<{activos: Array, concejales: Array, hasConcejales: boolean}>}
 */
async function readExcelBuffer(buffer) {
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(buffer);
  } catch {
    throw new Error('No se ha podido leer el archivo: no parece un Excel .xlsx válido.');
  }
  const activosSheet = workbook.getWorksheet('ACTIVOS');
  if (!activosSheet) {
    throw new Error('El archivo Excel debe contener una hoja llamada «ACTIVOS».');
  }
  const concejalesSheet = workbook.getWorksheet('CONCEJALES');
  return {
    activos: sheetToRows(activosSheet),
    concejales: concejalesSheet ? sheetToRows(concejalesSheet) : [],
    hasConcejales: Boolean(concejalesSheet),
  };
}

function parseDdMmYyyy(raw) {
  const match = DATE_DDMMYYYY.exec(String(raw || '').trim());
  if (!match) return null;
  const [, dd, mm, yyyy] = match;
  const day = Number(dd);
  const month = Number(mm);
  const year = Number(yyyy);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  return { day, month, year };
}

/**
 * A departure takes effect the day after "Fecha de Baja": only dates strictly
 * before today count. Today and the future are still active.
 */
function isEffectiveLeaving(fechaBaja, today = new Date()) {
  const parsed = parseDdMmYyyy(fechaBaja);
  if (!parsed) return false;
  const leaving = Date.UTC(parsed.year, parsed.month - 1, parsed.day);
  const now = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate());
  return leaving < now;
}

function parseHireDate(raw) {
  const text = String(raw || '').trim();
  if (!text) return null;
  const slash = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(text);
  if (slash) {
    const day = Number(slash[1]);
    const month = Number(slash[2]);
    const year = Number(slash[3]);
    const date = new Date(Date.UTC(year, month - 1, day));
    if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
    return date.toISOString();
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(text)) {
    const date = new Date(text);
    return Number.isNaN(date.getTime()) ? null : date.toISOString();
  }
  return null;
}

function toTitleCase(text) {
  return String(text || '')
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
    .join(' ');
}

// ---------------------------------------------------------------------------
// Validation (offline)
// ---------------------------------------------------------------------------

/**
 * The five rules of m365-tools over the ACTIVOS rows.
 */
function validateExcelRows(rows, domain = DEFAULT_DOMAIN) {
  const requiredDomain = `@${String(domain).toLowerCase()}`;
  const rules = new Map([
    ['id_empleado-required-unique', { ruleId: 'id_empleado-required-unique', title: 'ID de empleado obligatorio y único', issues: [] }],
    ['email-required-domain', { ruleId: 'email-required-domain', title: 'Correo corporativo obligatorio', issues: [] }],
    ['email-unique', { ruleId: 'email-unique', title: 'Correo electrónico único', issues: [] }],
    ['id_responsable-exists', { ruleId: 'id_responsable-exists', title: 'ID de responsable existente', issues: [] }],
    ['fecha-de-baja-format', { ruleId: 'fecha-de-baja-format', title: 'Formato de Fecha de Baja (dd/mm/yyyy)', issues: [] }],
  ]);
  const add = (ruleId, index, row, column, value, message) => {
    rules.get(ruleId).issues.push({ row: index + 2, worker: row['Trabajador'] || '', column, value: value || '', message });
  };

  const ids = new Set(rows.map(row => (row.id_empleado || '').trim()).filter(Boolean));
  const seenIds = new Map();
  const seenEmails = new Map();

  rows.forEach((row, index) => {
    const id = (row.id_empleado || '').trim();
    if (!id) {
      add('id_empleado-required-unique', index, row, 'id_empleado', '', 'El ID de empleado es obligatorio.');
    } else if (seenIds.has(id)) {
      add('id_empleado-required-unique', index, row, 'id_empleado', id, `ID de empleado duplicado (visto por primera vez en la fila ${seenIds.get(id)}).`);
    } else {
      seenIds.set(id, index + 2);
    }

    const email = (row.e_mail || '').trim();
    if (!email || !email.toLowerCase().endsWith(requiredDomain)) {
      add('email-required-domain', index, row, 'e_mail', email, `El correo es obligatorio y debe terminar en ${requiredDomain}.`);
    }
    if (email) {
      const key = email.toLowerCase();
      if (seenEmails.has(key)) {
        add('email-unique', index, row, 'e_mail', email, `Correo duplicado (visto por primera vez en la fila ${seenEmails.get(key)}).`);
      } else {
        seenEmails.set(key, index + 2);
      }
    }

    const manager = (row['ID responsable'] || '').trim();
    if (manager && !ids.has(manager)) {
      add('id_responsable-exists', index, row, 'ID responsable', manager, 'El ID de responsable no existe entre los ID de empleado del archivo.');
    }

    const baja = (row['Fecha de Baja'] || '').trim();
    if (baja && !DATE_DDMMYYYY.test(baja)) {
      add('fecha-de-baja-format', index, row, 'Fecha de Baja', baja, 'La fecha debe tener el formato dd/mm/aaaa.');
    }
  });

  const byRule = [...rules.values()].filter(rule => rule.issues.length > 0);
  const affected = new Set(byRule.flatMap(rule => rule.issues.map(issue => issue.row)));
  return {
    totalRows: rows.length,
    totalIssues: byRule.reduce((sum, rule) => sum + rule.issues.length, 0),
    affectedRows: affected.size,
    byRule,
  };
}

// ---------------------------------------------------------------------------
// Sync check
// ---------------------------------------------------------------------------

/**
 * Compare the Excel with the active cloud accounts.
 * @param {Array} cloudUsers - Active, internal users from Graph
 */
function analyzeSyncCheck(activosRows, concejalesRows, cloudUsers, domain = DEFAULT_DOMAIN, { today = new Date(), roomEmails = null } = {}) {
  const requiredDomain = `@${String(domain).toLowerCase()}`;
  const mustExist = [];
  const mustBeDisabled = [];
  const noEmail = [];
  const allExcelEmails = new Set();

  activosRows.forEach((row, index) => {
    const email = (row.e_mail || '').trim().toLowerCase();
    const name = row['Trabajador'] || '';
    const base = { row: index + 2, sheet: 'ACTIVOS', name, email, department: row['Departamento'] || '', jobTitle: row['Descripción'] || row['Descripcion'] || '' };
    const baja = (row['Fecha de Baja'] || '').trim();
    if (isEffectiveLeaving(baja, today)) {
      if (email.endsWith(requiredDomain)) {
        mustBeDisabled.push({ ...base, fechaBaja: baja });
        allExcelEmails.add(email);
      }
    } else if (email.endsWith(requiredDomain)) {
      mustExist.push(base);
      allExcelEmails.add(email);
    } else {
      noEmail.push({ ...base, email: email || '(vacío)' });
    }
  });

  for (const row of concejalesRows) {
    const email = String(row.email || row.e_mail || row.Email || '').trim().toLowerCase();
    if (email.endsWith(requiredDomain)) allExcelEmails.add(email);
  }

  const cloudMap = new Map();
  for (const user of cloudUsers) {
    const upn = (user.userPrincipalName || '').toLowerCase();
    if (!upn || upn.includes('#ext#')) continue;
    cloudMap.set(upn, user);
  }

  const missing = mustExist.filter(item => !cloudMap.has(item.email));
  const notDisabled = mustBeDisabled
    .filter(item => cloudMap.has(item.email) && cloudMap.get(item.email).accountEnabled !== false)
    .map(item => {
      const cloud = cloudMap.get(item.email);
      return { ...item, userId: cloud.id, displayName: cloud.displayName || item.name };
    });

  const cloudOnly = [];
  const specialFound = [];
  for (const [upn, user] of cloudMap) {
    if (!upn.endsWith(requiredDomain)) continue;
    if (user.accountEnabled === false) continue;
    if (isRoomUser(user, roomEmails)) continue;
    if (allExcelEmails.has(upn)) continue;
    const entry = { id: user.id, displayName: user.displayName || upn, userPrincipalName: upn, jobTitle: user.jobTitle || '', department: user.department || '' };
    (SPECIAL_ACCOUNTS.has(upn) ? specialFound : cloudOnly).push(entry);
  }

  const byName = (a, b) => compareEs(a.name || a.displayName || '', b.name || b.displayName || '');
  missing.sort(byName);
  noEmail.sort(byName);
  notDisabled.sort(byName);
  cloudOnly.sort(byName);
  specialFound.sort(byName);

  return {
    totalActivos: activosRows.length,
    totalCloud: cloudMap.size,
    missing,
    noEmail,
    notDisabled,
    cloudOnly,
    specialFound,
    totalIssues: missing.length + noEmail.length + notDisabled.length + cloudOnly.length,
  };
}

// ---------------------------------------------------------------------------
// Import
// ---------------------------------------------------------------------------

const PASSTHROUGH_FIELDS = [
  'givenName', 'surname', 'mailNickname', 'mail', 'companyName', 'employeeType', 'officeLocation',
  'mobilePhone', 'streetAddress', 'city', 'state', 'postalCode', 'country', 'usageLocation',
  'preferredLanguage', 'userType',
];

const DIFF_FIELDS = [
  'displayName', 'givenName', 'surname', 'jobTitle', 'department', 'mobilePhone', 'officeLocation',
  'city', 'state', 'postalCode', 'country', 'usageLocation', 'preferredLanguage', 'employeeId',
  'employeeType', 'employeeHireDate',
];

/**
 * Map the Excel rows to Graph payloads. Spanish headers and Graph property
 * names are both accepted, case-insensitively.
 * @returns {{validRows: Array, invalidRows: Array, skippedLeaving: Array}}
 */
function parseImportRows(rows, domain = DEFAULT_DOMAIN, { skipLeaving = true, today = new Date() } = {}) {
  const requiredDomain = `@${String(domain).toLowerCase()}`;
  const validRows = [];
  const invalidRows = [];
  const skippedLeaving = [];

  rows.forEach((raw, index) => {
    const excelRow = index + 2;
    const row = {};
    for (const [key, value] of Object.entries(raw)) row[key.trim().toLowerCase()] = String(value ?? '').trim();
    const pick = (...names) => {
      for (const name of names) if (row[name.toLowerCase()]) return row[name.toLowerCase()];
      return '';
    };

    const payload = {};
    const employeeId = pick('id_empleado', 'employeeId');
    if (employeeId) payload.employeeId = employeeId;
    const department = pick('departamento', 'department');
    if (department) payload.department = department;
    const jobTitle = pick('descripción', 'descripcion', 'jobTitle');
    if (jobTitle) payload.jobTitle = jobTitle;

    for (const field of PASSTHROUGH_FIELDS) {
      if (row[field.toLowerCase()]) payload[field] = row[field.toLowerCase()];
    }
    const businessPhone = pick('businessPhone');
    if (businessPhone) payload.businessPhones = [businessPhone];

    const worker = pick('trabajador');
    if (worker) {
      if (worker.includes(',')) {
        const [surnameRaw, ...rest] = worker.split(',');
        payload.surname = toTitleCase(surnameRaw);
        payload.givenName = toTitleCase(rest.join(','));
        payload.displayName = `${payload.givenName} ${payload.surname}`.trim();
      } else {
        payload.displayName = toTitleCase(worker);
      }
    } else if (pick('displayName')) {
      payload.displayName = pick('displayName');
    }

    const hire = pick('fecha de alta', 'employeeHireDate');
    if (hire) {
      const iso = parseHireDate(hire);
      if (iso) payload.employeeHireDate = iso;
    }

    const emailRaw = pick('e_mail', 'userPrincipalName');
    const upn = emailRaw ? normalizeUpn(emailRaw, domain) : '';
    if (upn) payload.userPrincipalName = upn;
    if (upn && !payload.mailNickname) payload.mailNickname = upn.split('@')[0];

    const enabled = pick('accountEnabled');
    payload.accountEnabled = enabled ? enabled.toLowerCase() !== 'false' : true;
    const hasExplicitEnabled = Boolean(enabled);

    const password = pick('password');
    if (password) {
      payload.passwordProfile = {
        password,
        forceChangePasswordNextSignIn: pick('forceChangePasswordNextSignIn').toLowerCase() !== 'false',
      };
    }

    const fechaBaja = pick('fecha de baja');
    if (skipLeaving && isEffectiveLeaving(fechaBaja, today)) {
      skippedLeaving.push({ excelRow, worker: worker || payload.displayName || '', upn, fechaBaja });
      return;
    }

    const problems = [];
    if (!upn) problems.push('falta el correo (e_mail)');
    else if (!upn.toLowerCase().endsWith(requiredDomain)) problems.push(`el correo debe terminar en ${requiredDomain}`);
    if (!payload.displayName) problems.push('falta el nombre (Trabajador)');
    if (problems.length) {
      invalidRows.push({ excelRow, row: excelRow, worker: worker || '', upn, message: problems.join('; '), error: problems.join('; ') });
      return;
    }

    validRows.push({
      excelRow,
      upn,
      displayName: payload.displayName,
      payload,
      hasExplicitEnabled,
      managerRef: pick('id responsable', 'manager'),
    });
  });

  return { validRows, invalidRows, skippedLeaving };
}

function sameValue(field, current, wanted) {
  if (field === 'employeeId') return String(current ?? '').trim() === String(wanted ?? '').trim();
  if (field === 'employeeHireDate') {
    if (!current) return false;
    return new Date(current).getTime() === new Date(wanted).getTime();
  }
  return (current ?? '') === wanted;
}

/**
 * Decide what an import would do with each row. Existing accounts (active or
 * disabled) are matched by UPN; an existing account is never re-enabled and
 * never gets its password or mail nickname changed.
 */
function buildImportPlan(validRows, existingUsers, { limit = null } = {}) {
  const byUpn = new Map(existingUsers.map(user => [(user.userPrincipalName || '').toLowerCase(), user]));
  const selected = Number.isInteger(limit) && limit > 0 ? validRows.slice(0, limit) : validRows;
  const items = [];
  const summary = { total: selected.length, create: 0, update: 0, ready: 0, managerChanges: 0 };

  for (const row of selected) {
    const existing = byUpn.get(row.upn.toLowerCase());
    const base = { excelRow: row.excelRow, upn: row.upn, displayName: row.displayName, managerRef: row.managerRef || null };
    if (row.managerRef) summary.managerChanges += 1;

    if (!existing) {
      summary.create += 1;
      items.push({ ...base, action: 'create', existingId: null, payload: row.payload, diffs: [] });
      continue;
    }

    const updatePayload = {};
    const diffs = [];
    for (const field of DIFF_FIELDS) {
      if (row.payload[field] === undefined) continue;
      if (!sameValue(field, existing[field], row.payload[field])) {
        updatePayload[field] = row.payload[field];
        diffs.push({ field, from: existing[field] ?? '', to: row.payload[field] });
      }
    }
    if (row.hasExplicitEnabled && row.payload.accountEnabled !== existing.accountEnabled) {
      updatePayload.accountEnabled = row.payload.accountEnabled;
      diffs.push({ field: 'accountEnabled', from: existing.accountEnabled, to: row.payload.accountEnabled });
    }

    if (diffs.length === 0) {
      summary.ready += 1;
      items.push({ ...base, action: 'ready', existingId: existing.id, diffs: [] });
    } else {
      summary.update += 1;
      items.push({ ...base, action: 'update', existingId: existing.id, updatePayload, diffs });
    }
  }

  return { items, summary };
}

/**
 * Say, for every row of a plan that names a manager, whether that manager would
 * change, so the preview shows it (the apply step still decides on its own).
 * The manager is found like the apply step does: a number is an employee ID,
 * anything else is an address; people who are about to be created count too.
 * Sets `item.managerChange` and replaces `summary.managerChanges` (until now the
 * number of rows that named a manager) by the number that really change.
 * @param {{items: Array, summary: Object}} plan - From buildImportPlan
 * @param {Array} existingUsers - Everyone in the directory
 * @param {Map<string, {id: string, displayName: string}|null>} currentManagers - By user ID
 * @param {string} domain
 */
function previewManagerChanges(plan, existingUsers, currentManagers, domain = DEFAULT_DOMAIN) {
  const byEmployeeId = new Map();
  const byUpn = new Map();
  for (const user of existingUsers) {
    if (user.employeeId) byEmployeeId.set(String(user.employeeId).trim(), { id: user.id, displayName: user.displayName });
    if (user.userPrincipalName) byUpn.set(user.userPrincipalName.toLowerCase(), { id: user.id, displayName: user.displayName });
  }
  // Accounts this import creates can be managers of others in the same file
  for (const item of plan.items) {
    if (item.action !== 'create') continue;
    const target = { id: null, displayName: item.displayName, newAccount: true };
    if (item.payload?.employeeId) byEmployeeId.set(String(item.payload.employeeId).trim(), target);
    byUpn.set(item.upn.toLowerCase(), target);
  }

  let changes = 0;
  let notFound = 0;
  for (const item of plan.items) {
    if (!item.managerRef) continue;
    const target = /^\d+$/.test(item.managerRef)
      ? byEmployeeId.get(item.managerRef)
      : byUpn.get(normalizeUpn(item.managerRef, domain).toLowerCase());
    if (!target) {
      item.managerChange = { status: 'notfound', ref: item.managerRef };
      notFound += 1;
      continue;
    }
    const current = item.existingId ? currentManagers.get(item.existingId) : null;
    if (target.id && (target.id === item.existingId || target.id === current?.id)) continue; // nothing to change
    item.managerChange = { status: 'change', from: current?.displayName || null, to: target.displayName, newAccount: Boolean(target.newAccount) };
    changes += 1;
  }
  plan.summary.managerChanges = changes;
  plan.summary.managerNotFound = notFound;
}

module.exports = {
  previewManagerChanges,
  cellToString,
  readExcelBuffer,
  isEffectiveLeaving,
  parseHireDate,
  toTitleCase,
  validateExcelRows,
  analyzeSyncCheck,
  parseImportRows,
  buildImportPlan,
};
