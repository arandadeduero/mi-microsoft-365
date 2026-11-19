/**
 * Pure utilities for the Admin → Users management section.
 *
 * Ports and bug-fixes the pure domain logic from m365-tools (password
 * generation, credential links, user filtering, mailing list, job-title
 * normalization, org-chart integrity check, mailbox timezone audit/patch,
 * and edit-form diffing).
 */

const crypto = require('crypto');
const { compareEs } = require('./formatters');

const AUTO_GROUPS = [
  'Todo Ayuntamiento',
  'Todos los usuarios',
  'Expertos 365',
];

const AUTO_GROUPS_SET = new Set(AUTO_GROUPS);

const SPECIAL_ACCOUNTS = new Set([
  'asac@arandadeduero.es',
  'testasac@arandadeduero.es',
  'equiporedes@arandadeduero.es',
  'info_practicas1@arandadeduero.es',
  'lector01agua@arandadeduero.es',
  'consultorahyk@arandadeduero.es',
  'lector02agua@arandadeduero.es',
]);

const TARGET_TIMEZONE = 'Romance Standard Time';

const WELCOME_GUIDE_DEFAULT_URL =
  'https://aytoarandaduero-my.sharepoint.com/:f:/g/personal/glopez_arandadeduero_es/IgDxAF3YuiHuRILrP8vCkGRnAYrtXe0_jzxtD_l9ZON1CGg?e=86PJVY';

// Readable Word-NNNN-Word Spanish wordlist (meets Entra ID complexity: upper + lower + digits, 12+ chars)
const PASSWORD_WORDS = [
  'Acebo', 'Aguila', 'Alamo', 'Almez', 'Aluda', 'Arbol', 'Arco', 'Arena',
  'Arnal', 'Arroyo', 'Avena', 'Azul', 'Barda', 'Bogar', 'Brama', 'Brasa',
  'Bravo', 'Brisa', 'Bronce', 'Buho', 'Cabal', 'Calma', 'Campo', 'Cardo',
  'Ciervo', 'Cima', 'Cisne', 'Claro', 'Cobro', 'Copa', 'Coral', 'Cueva',
  'Dardo', 'Delta', 'Denso', 'Dicho', 'Dique', 'Durce', 'Ebano', 'Encina',
  'Farro', 'Finca', 'Firma', 'Fisco', 'Flama', 'Flora', 'Flujo', 'Fondo',
  'Forja', 'Freno', 'Grano', 'Grifo', 'Gruta', 'Guila', 'Hierro', 'Hondo',
  'Humo', 'Ibero', 'Ingle', 'Islote', 'Jabali', 'Jarco', 'Junco', 'Kilo',
  'Lanza', 'Largo', 'Lauro', 'Lecho', 'Ledge', 'Limon', 'Lince', 'Linde',
  'Lobos', 'Loma', 'Lunar', 'Macho', 'Manto', 'Marco', 'Marga', 'Marte',
  'Metro', 'Mirlo', 'Monte', 'Moral', 'Mosca', 'Nardo', 'Navío', 'Niebla',
  'Norte', 'Nublo', 'Olivo', 'Ombra', 'Onzas', 'Oveja', 'Pardo', 'Pato',
  'Pecho', 'Pedro', 'Peral', 'Perla', 'Piedra', 'Pinar', 'Plata', 'Plaza',
  'Porto', 'Prisa', 'Prosa', 'Purga', 'Radar', 'Ramal', 'Rango', 'Rasgo',
  'Recto', 'Reina', 'Reloj', 'Remo', 'Resto', 'Rioja', 'Risco', 'Roble',
  'Robre', 'Rocio', 'Rodeo', 'Ronda', 'Rubio', 'Rueda', 'Ruido', 'Salmo',
  'Salto', 'Sauce', 'Selva', 'Serra', 'Siena', 'Sierra', 'Solar', 'Sordo',
  'Sucre', 'Surco', 'Tabla', 'Talon', 'Tarde', 'Techo', 'Tejón', 'Termo',
  'Tierra', 'Tigre', 'Timón', 'Tinto', 'Tiro', 'Torca', 'Torno', 'Torre',
  'Tronco', 'Trueno', 'Tupido', 'Umbral', 'Uncio', 'Vados', 'Valle', 'Vasco',
  'Vega', 'Venas', 'Viento', 'Villa', 'Viña', 'Virgo', 'Vista', 'Volco',
  'Yacer', 'Yegua', 'Zarco', 'Zarza', 'Zorro',
];

/**
 * Append the tenant domain if the UPN input lacks '@'.
 * @param {string|null|undefined} value
 * @param {string} [domain='arandadeduero.es']
 * @returns {string}
 */
function normalizeUpn(value, domain = 'arandadeduero.es') {
  if (value == null) return '';
  const trimmed = String(value).trim();
  if (!trimmed) return '';
  if (trimmed.includes('@') || !domain) return trimmed;
  return `${trimmed}@${domain}`;
}

/**
 * Whether a directory user object is a meeting room mailbox, checked against
 * the email addresses from /places/microsoft.graph.room plus the naming
 * convention fallback (sala-* / salapz*).
 * @param {Object} user - Graph user object
 * @param {Set<string>|null} [roomEmails] - Lowercase room email addresses from getRooms()
 * @returns {boolean}
 */
function isRoomUser(user, roomEmails = null) {
  if (!user) return false;
  if (user.isRoom === true) return true;
  const upn = String(user.userPrincipalName || '').trim().toLowerCase();
  const mail = String(user.mail || '').trim().toLowerCase();
  if (roomEmails && typeof roomEmails.has === 'function') {
    if ((upn && roomEmails.has(upn)) || (mail && roomEmails.has(mail))) {
      return true;
    }
  }
  return upn.startsWith('sala-') || upn.startsWith('salapz');
}

/**
 * Generate a readable temporary password in Word-NNNN-Word format.
 * @returns {string}
 */
function generateTempPassword() {
  const w1 = PASSWORD_WORDS[crypto.randomInt(PASSWORD_WORDS.length)];
  const w2 = PASSWORD_WORDS[crypto.randomInt(PASSWORD_WORDS.length)];
  const num = String(crypto.randomInt(1000, 9999));
  return `${w1}-${num}-${w2}`;
}

/**
 * Build email body, mailto: URL, and Outlook Web compose deep link for new credentials.
 * @param {string} upn
 * @param {string} password
 * @param {object} [options]
 * @returns {{ subject: string, body: string, mailtoUrl: string, outlookWebUrl: string }}
 */
function buildResetCredentialLinks(
  upn,
  password,
  {
    intranetUrl = 'https://aytoarandaduero.sharepoint.com/',
    welcomeUrl = WELCOME_GUIDE_DEFAULT_URL,
  } = {},
) {
  const subject = 'Microsoft365';
  const body = [
    `username: ${upn}`,
    `password: ${password}`,
    'Recuerda que la contraseña de Microsoft365 es diferente a la de tu actual correo corporativo y a la de tu ordenador. Esta contraseña es temporal y deberás cambiarla en tu primer inicio de sesión.',
    '\n',
    `url: ${intranetUrl}`,
    '\n',
    `Sigue la guía de bienvenida en este enlace ${welcomeUrl}`,
  ].join('\n');

  const encodedSubject = encodeURIComponent(subject);
  const encodedBody = encodeURIComponent(body);
  const encodedUpn = encodeURIComponent(upn);

  const mailtoUrl = `mailto:${encodedUpn}?subject=${encodedSubject}&body=${encodedBody}`;
  const outlookWebUrl =
    `https://outlook.office.com/mail/deeplink/compose?to=${encodedUpn}&subject=${encodedSubject}&body=${encodedBody}`;

  return { subject, body, mailtoUrl, outlookWebUrl };
}

/**
 * Filter out tenant-wide auto-assigned groups from a user's group list.
 * Accepts either strings or objects with `displayName`.
 */
function filterAutoGroups(groups) {
  if (!Array.isArray(groups)) return [];
  return groups.filter((g) => {
    const name = typeof g === 'string' ? g : g?.displayName;
    return !AUTO_GROUPS_SET.has(name);
  });
}

/**
 * Filter admin user list by query and data-quality flags.
 * @param {Array<object>} users
 * @param {object} [filters]
 * @returns {Array<object>}
 */
function filterAdminUsers(users, filters = {}) {
  if (!Array.isArray(users)) return [];
  const query = filters.q ? String(filters.q).trim().toLowerCase() : '';

  return users.filter((u) => {
    if (!filters.includeRooms && u.isRoom) return false;
    if (query) {
      const haystacks = [
        u.displayName,
        u.userPrincipalName,
        u.mail,
        u.employeeId,
        u.department,
        u.jobTitle,
      ];
      const matches = haystacks.some(
        (field) => field != null && String(field).toLowerCase().includes(query),
      );
      if (!matches) return false;
    }

    if (filters.missingManager && u.manager) return false;
    if (filters.missingDepartment && u.department && String(u.department).trim() !== '') {
      return false;
    }
    if (filters.missingJobTitle && u.jobTitle && String(u.jobTitle).trim() !== '') {
      return false;
    }
    if (
      filters.neverSignedIn &&
      (u.signInActivity?.lastSignInDateTime || u.signInActivity?.lastSuccessfulSignInDateTime)
    ) {
      return false;
    }
    if (filters.emptyGroups && filterAutoGroups(u.groups).length > 0) {
      return false;
    }

    return true;
  });
}

/**
 * Build a semicolon-joined mailing list of active licensed users sorted alphabetically.
 * @param {Array<object>} users
 * @returns {string}
 */
function buildMailingList(users) {
  if (!Array.isArray(users)) return '';
  return users
    .filter((u) => {
      if (!u || !u.userPrincipalName) return false;
      if (u.accountEnabled === false || u.isRoom) return false;
      const lics = u.licenses !== undefined ? u.licenses : u.assignedLicenses;
      return Array.isArray(lics) && lics.length > 0;
    })
    .map((u) => u.userPrincipalName)
    .sort(compareEs)
    .join(';');
}

/**
 * Convert a string to Sentence case (first letter uppercase, rest lowercase).
 * @param {string|null|undefined} str
 * @returns {string}
 */
function toSentenceCase(str) {
  if (str == null) return '';
  const s = String(str).trim();
  if (!s) return '';
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

/**
 * Find users whose jobTitle is not already in Sentence case.
 * @param {Array<object>} users
 * @returns {Array<{ id: string, displayName: string, upn: string, current: string, fixed: string }>}
 */
function findJobTitleFixes(users) {
  if (!Array.isArray(users)) return [];
  return users
    .filter((u) => u && u.jobTitle && String(u.jobTitle).trim() !== '')
    .map((u) => ({
      id: u.id,
      displayName: u.displayName || '',
      upn: u.userPrincipalName || u.id || '',
      current: u.jobTitle,
      fixed: toSentenceCase(u.jobTitle),
    }))
    .filter((entry) => entry.current !== entry.fixed)
    .sort((a, b) => compareEs(a.displayName, b.displayName) || compareEs(a.upn, b.upn));
}

/**
 * Analyze the manager graph for missing managers, chains not reaching rootId, and cycles.
 * Port of `analyseOrgTree` from m365-tools/src/commands/fix.js.
 *
 * @param {Array<object>} users
 * @param {Map<string, string|object|null>|Record<string, string|object|null>} managerMap
 * @param {string} rootId
 * @returns {{ total: number, validCount: number, noManager: Array, wrongRoot: Array, cycles: Array<Array>, affectedIds: Array<string> }}
 */
function analyzeOrgChartIntegrity(users, managerMap, rootId) {
  const list = Array.isArray(users) ? users : [];
  const byId = new Map(list.map((u) => [u.id, u]));
  const getMgrId = (userId) => {
    const val = managerMap instanceof Map ? managerMap.get(userId) : managerMap?.[userId];
    if (!val) return null;
    return typeof val === 'object' ? (val.id || null) : val;
  };

  const noManager = [];
  const wrongRoot = [];
  const cycles = [];
  const cycleIds = new Set();

  for (const user of list) {
    if (user.id === rootId) continue;

    const managerId = getMgrId(user.id);
    if (!managerId) {
      noManager.push(user);
      continue;
    }

    const visited = new Set([user.id]);
    let current = user.id;
    let reachesRoot = false;
    let hasCycle = false;

    for (let hop = 0; hop < 50; hop++) {
      const nextId = getMgrId(current);
      if (!nextId) break;
      if (nextId === rootId) {
        reachesRoot = true;
        break;
      }
      if (visited.has(nextId)) {
        hasCycle = true;
        break;
      }
      visited.add(nextId);
      current = nextId;
    }

    if (hasCycle) {
      if (cycleIds.has(user.id)) continue;

      const cycleMembers = [user];
      let c = getMgrId(user.id);
      const seen = new Set([user.id]);
      while (c && !seen.has(c) && byId.has(c)) {
        seen.add(c);
        cycleMembers.push(byId.get(c));
        c = getMgrId(c);
      }
      cycles.push(cycleMembers);
      for (const member of cycleMembers) cycleIds.add(member.id);
      continue;
    }

    if (!reachesRoot) {
      wrongRoot.push(user);
    }
  }

  noManager.sort((a, b) => compareEs(a.displayName, b.displayName));
  wrongRoot.sort((a, b) => compareEs(a.displayName, b.displayName));

  const affectedSet = new Set([
    ...noManager.map((u) => u.id),
    ...wrongRoot.map((u) => u.id),
    ...cycles.flat().map((u) => u.id),
  ]);

  const total = list.length;
  const affectedIds = [...affectedSet];
  const validCount = Math.max(0, total - affectedIds.length);

  return {
    total,
    validCount,
    noManager,
    wrongRoot,
    cycles,
    affectedIds,
  };
}

/**
 * Classify whether a user's mailbox timezone and workingHours timezone match targetTz.
 * @param {object|null} settings
 * @param {string} [targetTz=TARGET_TIMEZONE]
 * @returns {{ ok: boolean, tzOk: boolean, whTzOk: boolean, timeZone: string|null, workingHoursTimeZone: string|null }}
 */
function classifyMailboxTimezone(settings, targetTz = TARGET_TIMEZONE) {
  const timeZone = settings?.timeZone ?? null;
  const workingHoursTimeZone =
    settings?.workingHours?.timeZone?.name ??
    settings?.workingHoursTimeZone ??
    null;

  const tzOk = timeZone === targetTz;
  const whTzOk = workingHoursTimeZone === null || workingHoursTimeZone === targetTz;

  return {
    ok: tzOk && whTzOk,
    tzOk,
    whTzOk,
    timeZone,
    workingHoursTimeZone,
  };
}

/**
 * Build the PATCH body for `/users/{id}/mailboxSettings`, preserving workingHours schedule.
 * @param {object|null} settings
 * @param {string} [targetTz=TARGET_TIMEZONE]
 * @returns {object}
 */
function buildTimezonePatchPayload(settings, targetTz = TARGET_TIMEZONE) {
  const payload = { timeZone: targetTz };
  const wh = settings?.workingHours;
  if (wh && typeof wh === 'object') {
    payload.workingHours = {
      daysOfWeek: wh.daysOfWeek,
      startTime: wh.startTime,
      endTime: wh.endTime,
      timeZone: { name: targetTz },
    };
  }
  return payload;
}

const REQUIRED_IDENTITY_FIELDS = ['displayName', 'mailNickname'];

const CLEARABLE_EDIT_FIELDS = [
  'givenName',
  'surname',
  'jobTitle',
  'department',
  'companyName',
  'employeeId',
  'employeeType',
  'mobilePhone',
  'officeLocation',
  'streetAddress',
  'city',
  'state',
  'postalCode',
  'country',
  'usageLocation',
  'preferredLanguage',
];

function parseIsoYmd(raw) {
  const match = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;
  const [, yyyy, mm, dd] = match.map(Number);
  const date = new Date(Date.UTC(yyyy, mm - 1, dd));
  const valid =
    date.getUTCFullYear() === yyyy &&
    date.getUTCMonth() === mm - 1 &&
    date.getUTCDate() === dd;
  return valid ? `${match[1]}-${match[2]}-${match[3]}T00:00:00Z` : null;
}

/**
 * Diff an admin user edit form submission against the current Graph user object.
 * @param {object} currentUser
 * @param {object} body
 * @returns {{ patch: object, changes: Array<{ field: string, old: any, new: any }> }}
 */
function diffUserEdit(currentUser, body) {
  const current = currentUser || {};
  const input = body || {};
  const patch = {};
  const changes = [];

  for (const field of REQUIRED_IDENTITY_FIELDS) {
    if (input[field] === undefined) continue;
    const next = String(input[field]).trim();
    if (next === '') continue; // Never clear required identity fields
    const cur = current[field] == null ? '' : String(current[field]).trim();
    if (next !== cur) {
      patch[field] = next;
      changes.push({ field, old: current[field] ?? null, new: next });
    }
  }

  for (const field of CLEARABLE_EDIT_FIELDS) {
    if (input[field] === undefined) continue;
    const raw = input[field] == null ? '' : String(input[field]).trim();
    const next = raw === '' ? null : raw;
    const curRaw = current[field] == null ? '' : String(current[field]).trim();
    const cur = curRaw === '' ? null : curRaw;
    if (next !== cur) {
      patch[field] = next;
      changes.push({ field, old: cur, new: next });
    }
  }

  if (input.businessPhone !== undefined) {
    const raw = input.businessPhone == null ? '' : String(input.businessPhone).trim();
    const curFirst =
      Array.isArray(current.businessPhones) && current.businessPhones[0]
        ? String(current.businessPhones[0]).trim()
        : '';
    if (raw !== curFirst) {
      patch.businessPhones = raw === '' ? [] : [raw];
      changes.push({ field: 'businessPhones', old: curFirst || null, new: raw || null });
    }
  }

  if (input.employeeHireDate !== undefined) {
    const raw = input.employeeHireDate == null ? '' : String(input.employeeHireDate).trim();
    const curRaw = current.employeeHireDate ? String(current.employeeHireDate).trim() : '';
    if (raw === '') {
      if (curRaw !== '') {
        patch.employeeHireDate = null;
        changes.push({ field: 'employeeHireDate', old: curRaw, new: null });
      }
    } else {
      const ymdIso = parseIsoYmd(raw);
      const nextIso = ymdIso || (!/^\d{4}-\d{2}-\d{2}$/.test(raw) ? raw : null);
      if (nextIso) {
        const nextTime = new Date(nextIso).getTime();
        const curTime = curRaw ? new Date(curRaw).getTime() : NaN;
        if (!Number.isNaN(nextTime) && nextTime !== curTime) {
          patch.employeeHireDate = nextIso;
          changes.push({ field: 'employeeHireDate', old: curRaw || null, new: nextIso });
        }
      }
    }
  }

  if (input.accountEnabled !== undefined) {
    let nextEnabled;
    if (typeof input.accountEnabled === 'boolean') {
      nextEnabled = input.accountEnabled;
    } else {
      const lowered = String(input.accountEnabled).trim().toLowerCase();
      nextEnabled = lowered === 'true' || lowered === '1' || lowered === 'on';
    }
    const curEnabled = Boolean(current.accountEnabled);
    if (nextEnabled !== curEnabled) {
      patch.accountEnabled = nextEnabled;
      changes.push({ field: 'accountEnabled', old: curEnabled, new: nextEnabled });
    }
  }

  return { patch, changes };
}

module.exports = {
  AUTO_GROUPS,
  SPECIAL_ACCOUNTS,
  TARGET_TIMEZONE,
  WELCOME_GUIDE_DEFAULT_URL,
  PASSWORD_WORDS,
  normalizeUpn,
  isRoomUser,
  generateTempPassword,
  buildResetCredentialLinks,
  filterAutoGroups,
  filterAdminUsers,
  buildMailingList,
  toSentenceCase,
  findJobTitleFixes,
  analyzeOrgChartIntegrity,
  classifyMailboxTimezone,
  buildTimezonePatchPayload,
  diffUserEdit,
};
