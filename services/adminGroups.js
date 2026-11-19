/**
 * Pure helpers for /admin/grupos: what kind of group it is, what can be
 * changed from the portal, and how a form becomes a Graph payload.
 */
const { isManageable } = require('./accountInsights');

const VISIBILITIES = {
  Public: 'Público: cualquiera de la organización puede ver su contenido',
  Private: 'Privado: solo ven el contenido sus miembros',
};
// The mail alias of every group the portal creates is gr-<something>; the name is free
const GROUP_PREFIX = 'gr-';
const MAX_NAME = 256;
const MAX_DESCRIPTION = 1024;
const MAX_ALIAS = 64;
// People added to a group in one go
const MAX_BULK = 200;

/**
 * Describe a group for the lists and the detail page.
 * `manageable` is true only when the portal may change it: dynamic groups
 * (members come from a rule), groups synced from an on-premises directory and
 * Exchange lists cannot be edited through Graph.
 */
function describeAdminGroup(group) {
  const types = group.groupTypes || [];
  const dynamic = types.includes('DynamicMembership');
  const unified = types.includes('Unified');
  const synced = group.onPremisesSyncEnabled === true;

  let kind;
  if (unified) kind = dynamic ? 'Microsoft 365 dinámico' : 'Microsoft 365';
  else if (dynamic) kind = 'Seguridad dinámico';
  else if (group.mailEnabled && group.securityEnabled) kind = 'Seguridad con correo (Exchange)';
  else if (group.mailEnabled) kind = 'Lista de distribución (Exchange)';
  else kind = 'Seguridad';

  let readOnlyReason = null;
  if (dynamic) readOnlyReason = 'Sus miembros los decide una regla, no se pueden cambiar a mano.';
  else if (synced) readOnlyReason = 'Viene de un directorio local y solo se cambia allí.';
  else if (!isManageable(group)) readOnlyReason = 'Se gestiona desde el centro de administración de Exchange.';

  return {
    ...group,
    kind,
    isUnified: unified,
    isDynamic: dynamic,
    hasTeam: (group.resourceProvisioningOptions || []).includes('Team'),
    manageable: readOnlyReason === null,
    readOnlyReason,
  };
}

/** Filter the list by free text (name, description, mail) and by kind */
function filterAdminGroups(groups, { q = '', kind = '' } = {}) {
  const needle = String(q).trim().toLowerCase();
  return groups.filter((group) => {
    if (kind && group.kind !== kind) return false;
    if (!needle) return true;
    return [group.displayName, group.description, group.mail]
      .some(text => String(text || '').toLowerCase().includes(needle));
  });
}

/** Turn a name into a mail alias: ASCII lowercase letters, digits, dot, dash, underscore */
function toMailAlias(name) {
  const alias = String(name || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9._-]+/g, '-')
    .slice(0, MAX_ALIAS);
  // Trimmed by hand: a regex for trailing dashes backtracks badly on long runs of them
  let start = 0;
  let end = alias.length;
  while (start < end && '-.'.includes(alias[start])) start += 1;
  while (end > start && '-.'.includes(alias[end - 1])) end -= 1;
  return alias.slice(start, end);
}

/**
 * Check the "new group" form.
 * @returns {{ error: string }|{ value: Object }} value is ready for createGroup
 */
function validateNewGroup(body) {
  const type = body.type === 'security' ? 'security' : 'unified';
  const displayName = String(body.displayName || '').trim();
  const description = String(body.description || '').trim();
  // Left empty, the alias is made from the name; typed, it must already start with gr-
  const alias = toMailAlias(body.mailNickname || `${GROUP_PREFIX}${displayName}`);
  const visibility = body.visibility === 'Public' ? 'Public' : 'Private';

  if (!displayName) return { error: 'El grupo necesita un nombre.' };
  if (displayName.length > MAX_NAME) return { error: `El nombre no puede pasar de ${MAX_NAME} caracteres.` };
  if (description.length > MAX_DESCRIPTION) return { error: `La descripción no puede pasar de ${MAX_DESCRIPTION} caracteres.` };
  if (!alias.startsWith(GROUP_PREFIX) || alias.length === GROUP_PREFIX.length) {
    return { error: `El alias de correo tiene que empezar por «${GROUP_PREFIX}» y seguir con algo, por ejemplo ${GROUP_PREFIX}informatica.` };
  }

  return { value: { type, displayName, description, mailNickname: alias, visibility } };
}

/**
 * Compare the edit form with the current group.
 * @returns {{ patch: Object, changes: string[] }} only the fields that changed
 */
function diffGroupEdit(current, body) {
  const patch = {};
  const changes = [];
  const displayName = String(body.displayName ?? '').trim();
  const description = String(body.description ?? '').trim();

  if (displayName && displayName !== (current.displayName || '')) {
    patch.displayName = displayName;
    changes.push('displayName');
  }
  if (description !== (current.description || '')) {
    patch.description = description || null; // Graph rejects '' on description
    changes.push('description');
  }
  const unified = (current.groupTypes || []).includes('Unified');
  // Only Public and Private are offered: a hidden-membership group keeps what it has
  if (unified && VISIBILITIES[current.visibility] && VISIBILITIES[body.visibility] && body.visibility !== current.visibility) {
    patch.visibility = body.visibility;
    changes.push('visibility');
  }
  return { patch, changes };
}

/** Error text for edit forms; null when fine */
function validateGroupEdit(body) {
  if (!String(body.displayName ?? '').trim()) return 'El grupo necesita un nombre.';
  if (String(body.displayName).trim().length > MAX_NAME) return `El nombre no puede pasar de ${MAX_NAME} caracteres.`;
  if (String(body.description ?? '').trim().length > MAX_DESCRIPTION) return `La descripción no puede pasar de ${MAX_DESCRIPTION} caracteres.`;
  return null;
}

/**
 * Read a pasted list of addresses: one per line, or separated by commas or
 * semicolons, also as copied from Outlook ("Ana Pérez <ana@x.es>").
 * @returns {{ addresses: string[], invalid: string[] }} addresses lowercase and without repeats; invalid are the entries that are not an address
 */
function parseAddressList(text) {
  const addresses = [];
  const invalid = [];
  for (const entry of String(text || '').split(/[\n\r,;]+/)) {
    const raw = entry.trim();
    if (!raw) continue;
    const inside = raw.match(/<([^<>\s]+)>\s*$/);
    const candidate = (inside ? inside[1] : raw).toLowerCase();
    if (!/^[^\s@<>]+@[^\s@<>]+$/.test(candidate)) invalid.push(raw);
    else if (!addresses.includes(candidate)) addresses.push(candidate);
  }
  return { addresses, invalid };
}

/**
 * Sort groups into the ones that need a look: nobody owns them, only one person
 * does, or nobody is in them. A group whose owners and members could not be read
 * is counted apart and never flagged ("could not check" is not "has none").
 * @param {Array} groups - Groups the portal can manage
 * @param {Map<string, {owners: number|null, hasMembers: boolean|null}>} checks - From getGroupsChecks
 */
function reviewGroups(groups, checks) {
  const review = { noOwner: [], oneOwner: [], empty: [], unknown: 0 };
  for (const group of groups) {
    const check = checks.get(group.id);
    if (!check || (check.owners === null && check.hasMembers === null)) {
      review.unknown += 1;
      continue;
    }
    if (check.owners === 0) review.noOwner.push(group);
    else if (check.owners === 1) review.oneOwner.push(group);
    if (check.hasMembers === false) review.empty.push(group);
  }
  return review;
}

module.exports = {
  reviewGroups,
  MAX_BULK,
  parseAddressList,
  GROUP_PREFIX,
  VISIBILITIES,
  describeAdminGroup,
  filterAdminGroups,
  toMailAlias,
  validateNewGroup,
  validateGroupEdit,
  diffGroupEdit,
};
