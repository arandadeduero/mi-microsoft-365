/**
 * Admin Groups Routes (/admin/grupos/*)
 *
 * - /admin/grupos        - List of every group, filter by name and kind
 * - /admin/grupos/nuevo  - Create a Microsoft 365 group or a security group
 * - /admin/grupos/:id    - Properties, owners and members; edit what Graph lets us
 *
 * Everything uses the admin's delegated token, so Entra decides what they may do.
 */
const express = require('express');
const router = express.Router();
const permissions = require('../config/permissions');
const { isAuthenticated } = require('../middleware/auth');
const { userError } = require('../middleware/errorHandler');
const { auditRequest } = require('../services/auditService');
const { compareEs, DIRECTORY_ID } = require('../services/formatters');
const {
  listAdminGroups,
  getAdminGroup,
  createGroup,
  updateGroup,
  getGroupOwners,
  getGroupMembers,
  addGroupOwner,
  removeGroupOwner,
  addUserToGroup,
  removeUserFromGroup,
  getAdminUser,
} = require('../services/graphService');
const {
  GROUP_PREFIX,
  VISIBILITIES,
  describeAdminGroup,
  filterAdminGroups,
  validateNewGroup,
  validateGroupEdit,
  diffGroupEdit,
} = require('../services/adminGroups');

const KINDS = ['Microsoft 365', 'Seguridad', 'Microsoft 365 dinámico', 'Seguridad dinámico', 'Seguridad con correo (Exchange)', 'Lista de distribución (Exchange)'];

function requireAdmin(req, res, next) {
  if (res.locals.isAdmin) return next();
  auditRequest(req, 'permission.denied', { path: req.path, method: req.method });
  next(userError(403, 'Esta página es solo para administradores'));
}

const isValidId = (id) => typeof id === 'string' && DIRECTORY_ID.test(id);
const flashTo = (req, key, value) => { req.session[key] = value; };

function popSessionValue(req, key) {
  const value = req.session && req.session[key] ? req.session[key] : null;
  if (req.session && key in req.session) delete req.session[key];
  return value;
}

function friendlyGraphError(err, fallback) {
  const status = err?.graphStatus || err?.response?.status;
  const detail = err?.graphMessage || err?.response?.data?.error?.message || '';
  if (status === 403) {
    return `${fallback} Microsoft Graph ha denegado el permiso (403). Comprueba tu rol en Entra ID.`;
  }
  if (/mailNickname|same value for property/i.test(detail)) {
    return `${fallback} Ya existe otro grupo con ese alias de correo; elige otro.`;
  }
  return fallback;
}

// A user typed as an address (or a directory ID) becomes a directory user, or null
async function findUser(token, raw) {
  const value = String(raw || '').trim();
  if (!value || (!isValidId(value) && !/^[^\s@]+@[^\s@]+$/.test(value))) return null;
  return getAdminUser(token, value);
}

router.get('/admin/grupos', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const all = (await listAdminGroups(req.session.accessToken)).map(describeAdminGroup);
    const q = String(req.query.q || '');
    const kind = KINDS.includes(req.query.tipo) ? req.query.tipo : '';
    const groups = filterAdminGroups(all, { q, kind });
    res.render('admin-groups', {
      title: 'Grupos',
      currentPage: 'audit',
      groups,
      total: all.length,
      q,
      kind,
      kinds: KINDS,
      flash: popSessionValue(req, 'adminGroupsFlash'),
      errorMsg: popSessionValue(req, 'adminGroupsError'),
    });
  } catch (error) {
    next(error);
  }
});

router.get('/admin/grupos/nuevo', isAuthenticated, requireAdmin, (req, res) => {
  res.render('admin-group-new', {
    title: 'Nuevo grupo',
    currentPage: 'audit',
    domain: permissions.DOMAIN || '',
    form: popSessionValue(req, 'adminGroupForm') || {},
    errorMsg: popSessionValue(req, 'adminGroupsError'),
    visibilities: VISIBILITIES,
    groupPrefix: GROUP_PREFIX,
  });
});

router.post('/admin/grupos', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const checked = validateNewGroup(req.body);
    if (checked.error) {
      flashTo(req, 'adminGroupsError', checked.error);
      flashTo(req, 'adminGroupForm', req.body);
      return res.redirect('/admin/grupos/nuevo');
    }
    const ownerId = req.session.account?.localAccountId;
    if (!isValidId(ownerId)) return next(userError(400, 'No se ha podido identificar tu cuenta.'));

    try {
      const group = await createGroup(req.session.accessToken, checked.value, ownerId);
      auditRequest(req, 'admin.group_created', { groupId: group.id, type: checked.value.type });
      flashTo(req, 'adminGroupsFlash', 'Grupo creado. Tú eres su propietario; añade ahora a los miembros.');
      return res.redirect(`/admin/grupos/${encodeURIComponent(group.id)}`);
    } catch (err) {
      if (err.graphStatus === 401) throw err;
      flashTo(req, 'adminGroupsError', friendlyGraphError(err, 'No se pudo crear el grupo.'));
      flashTo(req, 'adminGroupForm', req.body);
      return res.redirect('/admin/grupos/nuevo');
    }
  } catch (error) {
    next(error);
  }
});

async function loadGroup(req, next) {
  const id = req.params.id;
  if (!isValidId(id)) { next(userError(404, 'Grupo no encontrado')); return null; }
  const group = await getAdminGroup(req.session.accessToken, id);
  if (!group) { next(userError(404, 'Grupo no encontrado')); return null; }
  return describeAdminGroup(group);
}

router.get('/admin/grupos/:id', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const group = await loadGroup(req, next);
    if (!group) return undefined;
    const token = req.session.accessToken;
    const [owners, members] = await Promise.all([
      getGroupOwners(token, group.id),
      getGroupMembers(token, group.id),
    ]);
    return res.render('admin-group-detail', {
      title: group.displayName,
      currentPage: 'audit',
      group,
      owners: [...owners].sort((a, b) => compareEs(a.displayName, b.displayName)),
      members: [...members].sort((a, b) => compareEs(a.displayName, b.displayName)),
      visibilities: VISIBILITIES,
      flash: popSessionValue(req, 'adminGroupsFlash'),
      errorMsg: popSessionValue(req, 'adminGroupsError'),
    });
  } catch (error) {
    return next(error);
  }
});

// Runs one change on a group the portal may edit, then goes back to its page
function groupAction(handler) {
  return async (req, res, next) => {
    try {
      const group = await loadGroup(req, next);
      if (!group) return;
      const back = `/admin/grupos/${encodeURIComponent(group.id)}`;
      if (!group.manageable) {
        flashTo(req, 'adminGroupsError', group.readOnlyReason);
        return res.redirect(back);
      }
      try {
        await handler(req, group);
      } catch (err) {
        if (err.graphStatus === 401) throw err;
        flashTo(req, 'adminGroupsError', friendlyGraphError(err, 'No se pudo completar el cambio.'));
      }
      res.redirect(back);
    } catch (error) {
      next(error);
    }
  };
}

router.post('/admin/grupos/:id', isAuthenticated, requireAdmin, groupAction(async (req, group) => {
  const problem = validateGroupEdit(req.body);
  if (problem) { flashTo(req, 'adminGroupsError', problem); return; }
  const { patch, changes } = diffGroupEdit(group, req.body);
  if (changes.length === 0) { flashTo(req, 'adminGroupsFlash', 'No había cambios que guardar.'); return; }
  await updateGroup(req.session.accessToken, group.id, patch);
  auditRequest(req, 'admin.group_updated', { groupId: group.id, fields: changes });
  flashTo(req, 'adminGroupsFlash', 'Grupo actualizado.');
}));

function userChange({ add, event, done, notFound }) {
  return groupAction(async (req, group) => {
    const token = req.session.accessToken;
    // Adding types an address; removing comes from a button carrying the directory ID
    const user = add ? await findUser(token, req.body.user) : { id: req.params.userId };
    if (!user || !isValidId(user.id)) { flashTo(req, 'adminGroupsError', notFound); return; }
    await event(token, group, user.id);
    auditRequest(req, done.event, { groupId: group.id, userId: user.id });
    flashTo(req, 'adminGroupsFlash', done.message);
  });
}

router.post('/admin/grupos/:id/miembros', isAuthenticated, requireAdmin, userChange({
  add: true,
  event: (token, group, userId) => addUserToGroup(token, group.id, userId),
  done: { event: 'admin.group_member_added', message: 'Miembro añadido.' },
  notFound: 'No se ha encontrado ese usuario. Escribe su correo completo.',
}));

router.post('/admin/grupos/:id/miembros/:userId/quitar', isAuthenticated, requireAdmin, userChange({
  add: false,
  event: (token, group, userId) => removeUserFromGroup(token, group.id, userId),
  done: { event: 'admin.group_member_removed', message: 'Miembro quitado.' },
  notFound: 'Usuario no válido.',
}));

router.post('/admin/grupos/:id/propietarios', isAuthenticated, requireAdmin, userChange({
  add: true,
  event: (token, group, userId) => addGroupOwner(token, group.id, userId),
  done: { event: 'admin.group_owner_added', message: 'Propietario añadido.' },
  notFound: 'No se ha encontrado ese usuario. Escribe su correo completo.',
}));

router.post('/admin/grupos/:id/propietarios/:userId/quitar', isAuthenticated, requireAdmin, userChange({
  add: false,
  event: (token, group, userId) => removeGroupOwner(token, group.id, userId),
  done: { event: 'admin.group_owner_removed', message: 'Propietario quitado.' },
  notFound: 'Usuario no válido.',
}));

module.exports = router;
