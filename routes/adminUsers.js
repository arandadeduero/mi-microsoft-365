/**
 * Admin Users Routes (/admin/usuarios/*)
 *
 * Full user administration suite integrating m365-tools workflows:
 * - /admin/usuarios              - Directory list, search, filters, mailing list
 * - /admin/usuarios/excel/*      - Validate Excel, sync-check + disable departed, import with preview
 * - /admin/usuarios/responsables - Bulk manager assignment, add direct reports to group
 * - /admin/usuarios/calidad      - Fix job titles, fix org chart, fix mailbox time zones
 * - /admin/usuarios/:id          - User detail, full edit, manager, groups, password reset
 */

const crypto = require('crypto');
const express = require('express');
const multer = require('multer');
const router = express.Router();

const appConfig = require('../config/appConfig');
const permissions = require('../config/permissions');
const { SKU_NAMES } = require('../config/licenseCatalog');
const { isAuthenticated } = require('../middleware/auth');
const { verifyCsrfToken } = require('../middleware/csrf');
const { userError } = require('../middleware/errorHandler');
const { auditRequest } = require('../services/auditService');
const { getAppToken } = require('../services/tokenService');
const { reviewAccounts } = require('../services/accountReview');
const { mapWithConcurrency } = require('../services/asyncUtils');
const { compareEs, DIRECTORY_ID, formatDateTime } = require('../services/formatters');
const { readExcelDrift } = require('../services/excelDrift');
const {
  listAdminUsers,
  getUsersForReview,
  getMfaRegistration,
  fetchManagersBatch,
  fetchUserGroupsBatch,
  fetchUserLicensesBatch,
  getAdminUser,
  getUserPhotoById,
  getUserManagerById,
  getUserGroupsById,
  getUserLicensesById,
  getAllTenantGroups,
  getLicensedUsers,
  getRooms,
  createAdminUser,
  updateAdminUser,
  setUserManager,
  removeUserLicenses,
  assignUserLicenses,
  getSubscribedSkus,
  addUserToGroup,
  removeUserFromGroup,
  getUserMailboxSettingsAdmin,
  updateUserMailboxSettingsAdmin,
  downloadSharedFile,
} = require('../services/graphService');
const {
  readExcelBuffer,
  validateExcelRows,
  analyzeSyncCheck,
  parseImportRows,
  buildImportPlan,
  previewManagerChanges,
} = require('../services/adminUsersExcel');
const {
  AUTO_GROUPS,
  TARGET_TIMEZONE,
  normalizeUpn,
  isRoomUser,
  generateTempPassword,
  buildResetCredentialLinks,
  filterAdminUsers,
  buildMailingList,
  toSentenceCase,
  findJobTitleFixes,
  analyzeOrgChartIntegrity,
  classifyMailboxTimezone,
  buildTimezonePatchPayload,
  diffUserEdit,
} = require('../services/adminUsersTools');

const MAX_EXCEL_BYTES = 10 * 1024 * 1024; // 10 MB
const PLAN_TTL_MS = 15 * 60 * 1000; // 15 minutes
const MAX_PLANS = 10;
const AUTO_GROUPS_LOWER = new Set(AUTO_GROUPS.map(g => g.toLowerCase()));

// Short-lived in-memory store for previewed Excel import plans, bound to the admin who created them
const importPlans = new Map();

function pruneImportPlans() {
  const now = Date.now();
  for (const [id, entry] of importPlans) {
    if (now - entry.createdAt > PLAN_TTL_MS) {
      importPlans.delete(id);
    }
  }
  while (importPlans.size > MAX_PLANS) {
    const oldestKey = importPlans.keys().next().value;
    importPlans.delete(oldestKey);
  }
}

function storeImportPlan(ownerId, plan) {
  pruneImportPlans();
  const planId = crypto.randomBytes(16).toString('hex');
  importPlans.set(planId, { ownerId, plan, createdAt: Date.now() });
  return planId;
}

function consumeImportPlan(planId, ownerId) {
  pruneImportPlans();
  const entry = importPlans.get(planId);
  if (!entry || entry.ownerId !== ownerId) return null;
  importPlans.delete(planId);
  return entry.plan;
}

const uploadExcel = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_EXCEL_BYTES },
  fileFilter: (req, file, cb) => {
    const name = String(file.originalname || '').toLowerCase();
    if (!name.endsWith('.xlsx')) {
      return cb(userError(400, 'Solo se aceptan archivos Excel en formato .xlsx.'));
    }
    cb(null, true);
  },
});

function handleExcelUpload(req, res, next) {
  uploadExcel.single('excelFile')(req, res, (err) => {
    if (!err) return next();
    if (err.code === 'LIMIT_FILE_SIZE') {
      return next(userError(400, 'El archivo Excel supera el tamaño máximo permitido (10 MB).'));
    }
    return next(err);
  });
}

function requireAdmin(req, res, next) {
  if (res.locals.isAdmin) {
    return next();
  }
  auditRequest(req, 'permission.denied', { path: req.path, method: req.method });
  next(userError(403, 'Esta página es solo para administradores'));
}

const asArray = (val) => (Array.isArray(val) ? val : (val ? [val] : []));
const isValidId = (id) => typeof id === 'string' && DIRECTORY_ID.test(id);
const currentDomain = () => permissions.DOMAIN || 'arandadeduero.es';
function friendlyGraphError(err, fallback = 'No se ha podido completar la operación en Microsoft 365.') {
  const detail = err?.graphMessage || err?.response?.data?.error?.message || err?.message;
  if (err?.graphStatus === 403 || err?.response?.status === 403) {
    return `${fallback} Microsoft Graph ha denegado el permiso (403). Comprueba tu rol en Entra ID y que los permisos delegados tengan consentimiento de administrador (cierra sesión y vuelve a entrar tras concederlos).${detail ? ` Detalle: ${detail}` : ''}`;
  }
  return detail ? `${fallback} (${detail})` : fallback;
}

function popSessionValue(req, key) {
  const value = req.session && req.session[key] ? req.session[key] : null;
  if (req.session && key in req.session) {
    delete req.session[key];
  }
  return value;
}

function isNonAutoGroup(group) {
  return !AUTO_GROUPS_LOWER.has(String(group?.displayName || '').toLowerCase());
}

async function getRoomEmailSet(accessToken) {
  const places = await getRooms(accessToken);
  return new Set((places || []).map(p => String(p.emailAddress || '').trim().toLowerCase()).filter(Boolean));
}

// ============================================================================
// 1. Directory & Mailing List (GET /admin/usuarios)
// ============================================================================

router.get('/admin/usuarios', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const token = req.session.accessToken;
    const filters = {
      q: String(req.query.q || '').trim(),
      status: req.query.status === 'disabled' ? 'disabled' : 'active',
      missingManager: req.query.missingManager === '1' || req.query.missingManager === 'on',
      missingDepartment: req.query.missingDepartment === '1' || req.query.missingDepartment === 'on',
      missingJobTitle: req.query.missingJobTitle === '1' || req.query.missingJobTitle === 'on',
      neverSignedIn: req.query.neverSignedIn === '1' || req.query.neverSignedIn === 'on',
      emptyGroups: req.query.emptyGroups === '1' || req.query.emptyGroups === 'on',
      includeRooms: req.query.includeRooms === '1' || req.query.includeRooms === 'on',
    };

    let appToken = null;
    try {
      appToken = await getAppToken();
    } catch {
      // Best-effort: sign-in activity stays unknown without app token
    }

    const [{ users: rawUsers, signInsKnown }, roomEmails] = await Promise.all([
      listAdminUsers(token, {
        onlyDisabled: filters.status === 'disabled',
        includeSignInActivity: true,
        appToken,
      }),
      getRoomEmailSet(token),
    ]);

    const userIds = rawUsers.map(u => u.id);
    const [managersMap, groupsMap, licensesMap] = await Promise.all([
      fetchManagersBatch(token, userIds),
      fetchUserGroupsBatch(token, userIds),
      fetchUserLicensesBatch(token, userIds),
    ]);

    const enriched = rawUsers
      .map(u => {
        const allUserGroups = groupsMap.get(u.id) || [];
        const visibleGroups = allUserGroups
          .filter(isNonAutoGroup)
          .sort((a, b) => compareEs(a.displayName || '', b.displayName || ''));
        const licenses = (licensesMap.get(u.id) || []).map(l => ({
          skuId: l.skuId,
          skuPartNumber: l.skuPartNumber,
          name: SKU_NAMES[l.skuPartNumber] || l.skuPartNumber,
        }));
        return {
          ...u,
          isRoom: isRoomUser(u, roomEmails),
          manager: managersMap.get(u.id) || null,
          groups: visibleGroups,
          allGroupsCount: allUserGroups.length,
          licenses,
          lastSignInDateTime: u.signInActivity?.lastSuccessfulSignInDateTime
            || u.signInActivity?.lastSignInDateTime
            || null,
        };
      })
      .sort((a, b) => compareEs(a.displayName || a.userPrincipalName || '', b.displayName || b.userPrincipalName || ''));

    const filtered = filterAdminUsers(enriched, filters);
    const totalCount = filters.includeRooms ? enriched.length : enriched.filter(u => !u.isRoom).length;
    const mailingList = filters.status === 'active' ? buildMailingList(enriched) : '';
    const mailingCount = mailingList ? mailingList.split(';').length : 0;

    res.render('admin-users', {
      title: 'Usuarios',
      currentPage: 'admin',
      activeTab: 'directorio',
      drift: currentDrift(),
      users: filtered,
      totalUsers: totalCount,
      filters,
      signInsKnown,
      mailingList,
      mailingCount,
      flash: popSessionValue(req, 'adminUsersFlash'),
      errorMsg: popSessionValue(req, 'adminUsersError'),
    });
  } catch (error) {
    next(error);
  }
});

/**
 * GET /admin/usuarios/atencion
 * Accounts that need a look, read with the app's token
 */
router.get('/admin/usuarios/atencion', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const appToken = await getAppToken();
    const [listing, mfa, roomEmails] = await Promise.all([
      getUsersForReview(appToken),
      getMfaRegistration(appToken),
      getRoomEmailSet(req.session.accessToken),
    ]);
    const review = listing
      ? reviewAccounts(listing.users, {
        signInsKnown: listing.signInsKnown,
        mfa,
        passwordMaxAgeDays: Number(process.env.PASSWORD_MAX_AGE_DAYS) || null,
        roomEmails,
      })
      : null;

    res.render('admin-accounts', { title: 'Cuentas que necesitan atención', currentPage: 'admin', review, signInsError: listing ? listing.signInsError : null });
  } catch (error) {
    next(error);
  }
});

// ============================================================================
// 2. Excel Workflows (/admin/usuarios/excel/*)
// ============================================================================

// The last daily check of the Excel, ready to show
function currentDrift() {
  const drift = readExcelDrift();
  return drift ? { ...drift, when: formatDateTime(drift.checkedAt) } : null;
}

function renderExcelPage(res, overrides = {}) {
  res.render('admin-users-excel', {
    title: 'Usuarios · Excel',
    currentPage: 'admin',
    activeTab: 'excel',
    excelSection: 'validar',
    excelUrl: appConfig.USERS_EXCEL_URL || '',
    drift: currentDrift(),
    validationResult: null,
    syncResult: null,
    disableResult: null,
    importPreview: null,
    importResult: null,
    skuNames: SKU_NAMES,
    flash: null,
    errorMsg: null,
    ...overrides,
  });
}

async function resolveExcelBuffer(req) {
  if (req.file && req.file.buffer) {
    return req.file.buffer;
  }
  const url = String(appConfig.USERS_EXCEL_URL || '').trim();
  if (!url) return null;
  try {
    return await downloadSharedFile(req.session.accessToken, url);
  } catch (err) {
    if (err.graphStatus === 403 || err.response?.status === 403) {
      const appToken = await getAppToken();
      if (appToken) {
        return await downloadSharedFile(appToken, url);
      }
    }
    throw err;
  }
}

async function buildSyncCheckResult(token, activos, concejales, hasConcejales) {
  const [{ users: cloudUsers }, roomEmails] = await Promise.all([
    listAdminUsers(token, { onlyDisabled: false }),
    getRoomEmailSet(token),
  ]);
  const syncResult = analyzeSyncCheck(activos, concejales, cloudUsers, currentDomain(), { roomEmails });
  syncResult.hasConcejales = hasConcejales;

  // Fetch current licenses for Check 2 accounts so admin can review what will be removed before disabling
  const licenseCounts = new Map();
  if (syncResult.notDisabled.length > 0) {
    const licMap = await fetchUserLicensesBatch(token, syncResult.notDisabled.map(item => item.userId));
    for (const item of syncResult.notDisabled) {
      const lics = licMap.get(item.userId) || [];
      item.licenses = lics.map(l => ({
        skuId: l.skuId,
        skuPartNumber: l.skuPartNumber,
        name: SKU_NAMES[l.skuPartNumber] || l.skuPartNumber,
      }));
      for (const l of item.licenses) {
        licenseCounts.set(l.name, (licenseCounts.get(l.name) || 0) + 1);
      }
    }
  }
  syncResult.licenseSummary = [...licenseCounts.entries()]
    .map(([name, count]) => ({ name, count }))
    .sort((a, b) => b.count - a.count || compareEs(a.name, b.name));

  return syncResult;
}

router.get('/admin/usuarios/excel', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const flash = popSessionValue(req, 'adminUsersFlash');
    let errorMsg = popSessionValue(req, 'adminUsersError');
    let validationResult = null;
    let syncResult = null;

    if (appConfig.USERS_EXCEL_URL && req.query.autoload !== '0') {
      try {
        const buffer = await resolveExcelBuffer(req);
        if (buffer) {
          const { activos, concejales, hasConcejales } = await readExcelBuffer(buffer);
          validationResult = validateExcelRows(activos, currentDomain());
          syncResult = await buildSyncCheckResult(req.session.accessToken, activos, concejales, hasConcejales);
        }
      } catch (err) {
        if (err.graphStatus === 401) return next(err);
        errorMsg = friendlyGraphError(err, 'No se ha podido leer automáticamente el Excel de SharePoint (puedes subir un archivo .xlsx manualmente).');
      }
    }

    renderExcelPage(res, {
      excelSection: String(req.query.section || 'validar'),
      validationResult,
      syncResult,
      flash,
      errorMsg,
    });
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/excel/validar', isAuthenticated, requireAdmin, handleExcelUpload, verifyCsrfToken, async (req, res, next) => {
  try {
    const buffer = await resolveExcelBuffer(req);
    if (!buffer) {
      return renderExcelPage(res, { excelSection: 'validar', errorMsg: 'Selecciona un archivo Excel (.xlsx) para validar.' });
    }
    const { activos } = await readExcelBuffer(buffer);
    const validationResult = validateExcelRows(activos, currentDomain());
    renderExcelPage(res, { excelSection: 'validar', validationResult });
  } catch (error) {
    if (error.graphStatus === 401 || error.status) return next(error);
    renderExcelPage(res, { excelSection: 'validar', errorMsg: friendlyGraphError(error, 'No se ha podido leer el archivo Excel.') });
  }
});

router.post('/admin/usuarios/excel/sync', isAuthenticated, requireAdmin, handleExcelUpload, verifyCsrfToken, async (req, res, next) => {
  try {
    const buffer = await resolveExcelBuffer(req);
    if (!buffer) {
      return renderExcelPage(res, { excelSection: 'sync', errorMsg: 'Selecciona un archivo Excel (.xlsx) para comprobar la sincronización.' });
    }
    const token = req.session.accessToken;
    const { activos, concejales, hasConcejales } = await readExcelBuffer(buffer);
    const syncResult = await buildSyncCheckResult(token, activos, concejales, hasConcejales);

    renderExcelPage(res, { excelSection: 'sync', syncResult });
  } catch (error) {
    if (error.graphStatus === 401 || error.status) return next(error);
    renderExcelPage(res, { excelSection: 'sync', errorMsg: friendlyGraphError(error, 'No se ha podido comprobar la sincronización.') });
  }
});

router.post('/admin/usuarios/excel/deshabilitar', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const token = req.session.accessToken;
    const userIds = asArray(req.body.userIds).filter(isValidId);
    if (userIds.length === 0) {
      return renderExcelPage(res, { excelSection: 'sync', errorMsg: 'No se ha seleccionado ninguna cuenta para deshabilitar.' });
    }

    const results = [];
    let disabledCount = 0;
    let removedLicensesTotal = 0;

    // Sequential and strict order: remove all licenses FIRST, disable ONLY if license removal succeeds
    for (const userId of userIds) {
      let user = null;
      try {
        user = await getAdminUser(token, userId);
        if (!user) {
          results.push({ userId, displayName: userId, upn: '', ok: false, removedLicenses: 0, error: 'Usuario no encontrado' });
          continue;
        }
        const licenses = await getUserLicensesById(token, user.id);
        const skuIds = licenses.map(l => l.skuId).filter(Boolean);
        if (skuIds.length > 0) {
          await removeUserLicenses(token, user.id, skuIds);
          removedLicensesTotal += skuIds.length;
        }
        await updateAdminUser(token, user.id, { accountEnabled: false });
        disabledCount += 1;
        auditRequest(req, 'admin.user_disabled', {
          targetUserId: user.id,
          targetUpn: user.userPrincipalName,
          removedLicenses: licenses.map(l => l.skuPartNumber),
        });
        results.push({
          userId: user.id,
          displayName: user.displayName || user.userPrincipalName,
          upn: user.userPrincipalName,
          ok: true,
          removedLicenses: skuIds.length,
          error: null,
        });
      } catch (err) {
        if (err.graphStatus === 401) throw err;
        results.push({
          userId,
          displayName: user?.displayName || userId,
          upn: user?.userPrincipalName || '',
          ok: false,
          removedLicenses: 0,
          error: friendlyGraphError(err, 'Error al deshabilitar'),
        });
      }
    }

    renderExcelPage(res, {
      excelSection: 'sync',
      disableResult: {
        total: userIds.length,
        disabledCount,
        failedCount: userIds.length - disabledCount,
        removedLicensesTotal,
        results,
      },
      flash: disabledCount > 0
        ? `Se han deshabilitado ${disabledCount} cuenta(s) y retirado ${removedLicensesTotal} licencia(s).`
        : null,
    });
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/excel/importar/preview', isAuthenticated, requireAdmin, handleExcelUpload, verifyCsrfToken, async (req, res, next) => {
  try {
    const buffer = await resolveExcelBuffer(req);
    if (!buffer) {
      return renderExcelPage(res, { excelSection: 'importar', errorMsg: 'Selecciona un archivo Excel (.xlsx) para previsualizar la importación.' });
    }
    const token = req.session.accessToken;
    const rawLimit = Number.parseInt(String(req.body.limit || '').trim(), 10);
    const limit = Number.isInteger(rawLimit) && rawLimit > 0 ? rawLimit : null;
    const skipLeaving = req.body.skipLeaving === '1' || req.body.skipLeaving === 'on' || req.body.skipLeaving === 'true';

    const { activos } = await readExcelBuffer(buffer);
    const parsed = parseImportRows(activos, currentDomain(), { skipLeaving });
    // Load BOTH active and disabled users so disabled accounts are matched properly
    const { users: existingUsers } = await listAdminUsers(token, { onlyDisabled: null });
    const plan = buildImportPlan(parsed.validRows, existingUsers, { domain: currentDomain(), limit });
    // Who their manager is today, to show which managers would change (not only the fields)
    try {
      const withManager = plan.items.filter(item => item.managerRef && item.existingId).map(item => item.existingId);
      previewManagerChanges(plan, existingUsers, await fetchManagersBatch(token, withManager), currentDomain());
    } catch (err) {
      if (err.graphStatus === 401) throw err;
      plan.summary.managerUnknown = true; // the apply step still sets managers
    }
    const ownerId = req.session.account?.localAccountId || req.session.account?.username || 'admin';
    const planId = storeImportPlan(ownerId, plan);

    renderExcelPage(res, {
      excelSection: 'importar',
      importPreview: {
        planId,
        plan,
        invalidRows: parsed.invalidRows,
        skippedLeaving: parsed.skippedLeaving,
        limit,
        skipLeaving,
      },
    });
  } catch (error) {
    if (error.graphStatus === 401 || error.status) return next(error);
    renderExcelPage(res, { excelSection: 'importar', errorMsg: friendlyGraphError(error, 'No se ha podido previsualizar la importación.') });
  }
});

router.post('/admin/usuarios/excel/importar/aplicar', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const token = req.session.accessToken;
    const ownerId = req.session.account?.localAccountId || req.session.account?.username || 'admin';
    const planId = String(req.body.planId || '').trim();
    const plan = consumeImportPlan(planId, ownerId);

    if (!plan) {
      return renderExcelPage(res, {
        excelSection: 'importar',
        errorMsg: 'La vista previa de importación ha caducado o ya se ha aplicado. Vuelve a subir el archivo Excel.',
      });
    }

    // Build lookup maps across all existing users (active + disabled)
    const { users: allUsers } = await listAdminUsers(token, { onlyDisabled: null });
    const byEmployeeId = new Map();
    const byUpn = new Map();
    for (const u of allUsers) {
      if (u.employeeId) byEmployeeId.set(String(u.employeeId).trim(), u);
      if (u.userPrincipalName) byUpn.set(u.userPrincipalName.toLowerCase(), u);
    }

    const created = [];
    const updated = [];
    const ready = [];
    const errors = [];
    const managerErrors = [];
    const touchedUsers = [];

    // Pass 1: Create & Update user accounts
    for (const item of plan.items) {
      if (item.action === 'ready') {
        ready.push({ upn: item.upn, displayName: item.displayName });
        if (item.existingId) touchedUsers.push({ id: item.existingId, upn: item.upn, displayName: item.displayName, managerRef: item.managerRef });
        continue;
      }

      if (item.action === 'create') {
        try {
          const payload = { ...item.payload };
          let generatedPassword = null;
          if (!payload.passwordProfile) {
            generatedPassword = generateTempPassword();
            payload.passwordProfile = {
              password: generatedPassword,
              forceChangePasswordNextSignIn: true,
            };
          } else {
            generatedPassword = payload.passwordProfile.password;
          }
          const createdUser = await createAdminUser(token, payload);
          const createdId = createdUser?.id || null;
          if (createdUser) {
            if (createdUser.employeeId || payload.employeeId) {
              byEmployeeId.set(String(createdUser.employeeId || payload.employeeId).trim(), createdUser);
            }
            byUpn.set(item.upn.toLowerCase(), createdUser);
          }
          auditRequest(req, 'admin.user_created', { targetUserId: createdId, targetUpn: item.upn });
          created.push({
            id: createdId,
            upn: item.upn,
            displayName: item.displayName,
            temporaryPassword: generatedPassword,
          });
          if (createdId) touchedUsers.push({ id: createdId, upn: item.upn, displayName: item.displayName, managerRef: item.managerRef });
        } catch (err) {
          if (err.graphStatus === 401) throw err;
          errors.push({ upn: item.upn, displayName: item.displayName, action: 'create', error: friendlyGraphError(err, 'No se pudo crear') });
        }
        continue;
      }

      if (item.action === 'update') {
        try {
          await updateAdminUser(token, item.existingId || item.upn, item.updatePayload);
          updated.push({ id: item.existingId, upn: item.upn, displayName: item.displayName, diffs: item.diffs });
          if (item.existingId) touchedUsers.push({ id: item.existingId, upn: item.upn, displayName: item.displayName, managerRef: item.managerRef });
        } catch (err) {
          if (err.graphStatus === 401) throw err;
          errors.push({ upn: item.upn, displayName: item.displayName, action: 'update', error: friendlyGraphError(err, 'No se pudo actualizar') });
        }
      }
    }

    // Pass 2: Sync managers (supports managers created in Pass 1!)
    let managersUpdated = 0;
    const domain = currentDomain();
    for (const u of touchedUsers) {
      if (!u.managerRef) continue;
      try {
        let targetManager = null;
        if (/^\d+$/.test(u.managerRef)) {
          targetManager = byEmployeeId.get(u.managerRef) || null;
          if (!targetManager) {
            managerErrors.push({ upn: u.upn, displayName: u.displayName, error: `No se encontró ningún usuario con ID de empleado ${u.managerRef}.` });
            continue;
          }
        } else {
          const mgrUpn = normalizeUpn(u.managerRef, domain).toLowerCase();
          targetManager = byUpn.get(mgrUpn) || await getAdminUser(token, mgrUpn);
          if (!targetManager) {
            managerErrors.push({ upn: u.upn, displayName: u.displayName, error: `No se encontró el responsable ${mgrUpn}.` });
            continue;
          }
        }
        if (targetManager.id === u.id) continue;
        const currentMgr = await getUserManagerById(token, u.id);
        if (!currentMgr || currentMgr.id !== targetManager.id) {
          await setUserManager(token, u.id, targetManager.id);
          managersUpdated += 1;
        }
      } catch (err) {
        if (err.graphStatus === 401) throw err;
        managerErrors.push({ upn: u.upn, displayName: u.displayName, error: friendlyGraphError(err, 'No se pudo asignar responsable') });
      }
    }

    // License check on created & updated accounts
    const checkIds = [...created, ...updated].map(u => u.id).filter(Boolean);
    const unlicensedWarnings = [];
    if (checkIds.length > 0) {
      const licMap = await fetchUserLicensesBatch(token, checkIds);
      for (const entry of [...created, ...updated]) {
        if (entry.id && (licMap.get(entry.id) || []).length === 0) {
          unlicensedWarnings.push({ id: entry.id, upn: entry.upn, displayName: entry.displayName });
        }
      }
    }

    auditRequest(req, 'admin.excel_imported', {
      created: created.length,
      updated: updated.length,
      ready: ready.length,
      managersUpdated,
      errors: errors.length,
      managerErrors: managerErrors.length,
    });

    renderExcelPage(res, {
      excelSection: 'importar',
      importResult: {
        created,
        updated,
        ready,
        managersUpdated,
        errors,
        managerErrors,
        unlicensedWarnings,
      },
      flash: `Importación finalizada: ${created.length} creados, ${updated.length} actualizados, ${ready.length} sin cambios.`,
    });
  } catch (error) {
    next(error);
  }
});

// ============================================================================
// 3. Managers & Groups Bulk Tools (/admin/usuarios/responsables/*)
// ============================================================================

async function loadManagersPageContext(token, query = {}) {
  const [{ users: activeUsers }, allGroups, managersList] = await Promise.all([
    listAdminUsers(token, { onlyDisabled: false }),
    getAllTenantGroups(token),
    getLicensedUsers(token),
  ]);

  const departments = [...new Set(activeUsers.map(u => (u.department || '').trim()).filter(Boolean))].sort(compareEs);
  const jobTitles = [...new Set(activeUsers.map(u => (u.jobTitle || '').trim()).filter(Boolean))].sort(compareEs);
  const sortedGroups = [...allGroups].sort((a, b) => compareEs(a.displayName || '', b.displayName || ''));
  const sortedManagers = [...managersList].sort((a, b) => compareEs(a.displayName || '', b.displayName || ''));

  const filterType = query.filterType === 'jobTitle' ? 'jobTitle' : (query.filterType === 'department' ? 'department' : '');
  const filterValue = String(query.filterValue || '').trim();
  let assignMatched = null;
  if (filterType && filterValue) {
    assignMatched = activeUsers
      .filter(u => String(u[filterType] || '').trim() === filterValue)
      .sort((a, b) => compareEs(a.displayName || '', b.displayName || ''));
  }

  const managerId = isValidId(query.managerId) ? query.managerId : '';
  const groupId = isValidId(query.groupId) ? query.groupId : '';
  let groupManager = null;
  let groupReports = null;
  let groupTarget = null;
  let groupPreview = null;

  if (managerId) {
    groupManager = activeUsers.find(u => u.id === managerId)
      || sortedManagers.find(u => u.id === managerId)
      || null;
    if (groupManager) {
      const managersMap = await fetchManagersBatch(token, activeUsers.map(u => u.id));
      const directReports = activeUsers
        .filter(u => u.id !== managerId && managersMap.get(u.id)?.id === managerId)
        .sort((a, b) => compareEs(a.displayName || '', b.displayName || ''));
      const groupsMap = directReports.length > 0
        ? await fetchUserGroupsBatch(token, directReports.map(u => u.id))
        : new Map();
      groupReports = directReports.map(u => {
        const rawGroups = groupsMap.get(u.id) || [];
        return {
          ...u,
          rawGroupIds: rawGroups.map(g => g.id),
          groups: rawGroups.filter(isNonAutoGroup).sort((a, b) => compareEs(a.displayName || '', b.displayName || '')),
        };
      });
      if (groupId) {
        groupTarget = sortedGroups.find(g => g.id === groupId) || null;
        if (groupTarget) {
          const toAdd = groupReports.filter(u => !u.rawGroupIds.includes(groupId));
          const alreadyIn = groupReports.filter(u => u.rawGroupIds.includes(groupId));
          groupPreview = { toAdd, alreadyIn };
        }
      }
    }
  }

  return {
    departments,
    jobTitles,
    managersList: sortedManagers,
    allGroups: sortedGroups,
    assignFilter: { filterType, filterValue },
    assignMatched,
    groupManager,
    groupReports,
    groupTarget,
    groupPreview,
  };
}

router.get('/admin/usuarios/responsables', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const ctx = await loadManagersPageContext(req.session.accessToken, req.query);
    res.render('admin-users-managers', {
      title: 'Usuarios · Responsables y grupos',
      currentPage: 'admin',
      activeTab: 'responsables',
      ...ctx,
      assignResult: null,
      groupResult: null,
      flash: popSessionValue(req, 'adminUsersFlash'),
      errorMsg: popSessionValue(req, 'adminUsersError'),
    });
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/responsables/asignar', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const token = req.session.accessToken;
    const managerId = String(req.body.managerId || '').trim();
    const userIds = asArray(req.body.userIds).filter(isValidId);

    if (!isValidId(managerId) || userIds.length === 0) {
      const ctx = await loadManagersPageContext(token, req.body);
      return res.render('admin-users-managers', {
        title: 'Usuarios · Responsables y grupos',
        currentPage: 'admin',
        activeTab: 'responsables',
        ...ctx,
        assignResult: null,
        groupResult: null,
        flash: null,
        errorMsg: 'Selecciona un responsable válido y al menos un usuario.',
      });
    }

    const targets = userIds.filter(id => id !== managerId);
    const skippedSelf = userIds.length - targets.length;
    const settled = await mapWithConcurrency(targets, 10, async (uid) => {
      try {
        await setUserManager(token, uid, managerId);
        return { userId: uid, ok: true, error: null };
      } catch (err) {
        if (err.graphStatus === 401) throw err;
        return { userId: uid, ok: false, error: friendlyGraphError(err, 'No se pudo asignar') };
      }
    });

    const updated = settled.filter(r => r.ok).length;
    const failed = settled.filter(r => !r.ok);

    if (updated > 0) {
      auditRequest(req, 'admin.manager_assigned', {
        managerId,
        updatedCount: updated,
        failedCount: failed.length,
      });
    }

    const ctx = await loadManagersPageContext(token, req.body);
    res.render('admin-users-managers', {
      title: 'Usuarios · Responsables y grupos',
      currentPage: 'admin',
      activeTab: 'responsables',
      ...ctx,
      assignResult: { updated, failed, skippedSelf },
      groupResult: null,
      flash: updated > 0 ? `Responsable asignado a ${updated} usuario(s).` : null,
      errorMsg: failed.length > 0 ? `Fallaron ${failed.length} asignaciones.` : null,
    });
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/responsables/grupos', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const token = req.session.accessToken;
    const groupId = String(req.body.groupId || '').trim();
    const userIds = asArray(req.body.userIds).filter(isValidId);

    if (!isValidId(groupId) || userIds.length === 0) {
      const ctx = await loadManagersPageContext(token, req.body);
      return res.render('admin-users-managers', {
        title: 'Usuarios · Responsables y grupos',
        currentPage: 'admin',
        activeTab: 'responsables',
        ...ctx,
        assignResult: null,
        groupResult: null,
        flash: null,
        errorMsg: 'Selecciona un grupo válido y al menos una persona para añadir.',
      });
    }

    let added = 0;
    const failed = [];
    for (const uid of userIds) {
      try {
        await addUserToGroup(token, groupId, uid);
        added += 1;
      } catch (err) {
        if (err.graphStatus === 401) throw err;
        failed.push({ userId: uid, error: friendlyGraphError(err, 'No se pudo añadir al grupo') });
      }
    }

    if (added > 0) {
      auditRequest(req, 'admin.group_members_added', {
        groupId,
        addedCount: added,
        failedCount: failed.length,
      });
    }

    const ctx = await loadManagersPageContext(token, req.body);
    res.render('admin-users-managers', {
      title: 'Usuarios · Responsables y grupos',
      currentPage: 'admin',
      activeTab: 'responsables',
      ...ctx,
      assignResult: null,
      groupResult: { added, failed },
      flash: added > 0 ? `Se han añadido ${added} usuario(s) al grupo.` : null,
      errorMsg: failed.length > 0 ? `No se pudieron añadir ${failed.length} usuario(s).` : null,
    });
  } catch (error) {
    next(error);
  }
});

// ============================================================================
// 4. Data Quality Tools (/admin/usuarios/calidad/*)
// ============================================================================

async function loadQualityPage(token, { section = 'puestos', rootUpn = '', runTz = false } = {}) {
  const effectiveRootUpn = normalizeUpn(rootUpn || appConfig.ORG_CHART_ROOT, currentDomain());
  const state = {
    title: 'Usuarios · Calidad de datos',
    currentPage: 'admin',
    activeTab: 'calidad',
    qualitySection: section,
    jobTitleCandidates: null,
    jobTitleResult: null,
    orgRootUpn: effectiveRootUpn,
    orgRootUser: null,
    orgAnalysis: null,
    managersList: [],
    orgResult: null,
    tzAudit: null,
    tzResult: null,
    targetTimezone: TARGET_TIMEZONE,
    flash: null,
    errorMsg: null,
  };

  if (section === 'puestos') {
    const [{ users }, roomEmails] = await Promise.all([
      listAdminUsers(token, { onlyDisabled: false }),
      getRoomEmailSet(token),
    ]);
    state.jobTitleCandidates = findJobTitleFixes(users.filter(u => !isRoomUser(u, roomEmails)));
  } else if (section === 'organigrama') {
    const [{ users: rawUsers }, managersList, roomEmails] = await Promise.all([
      listAdminUsers(token, { onlyDisabled: false }),
      getLicensedUsers(token),
      getRoomEmailSet(token),
    ]);
    const users = rawUsers.filter(u => !isRoomUser(u, roomEmails));
    state.managersList = [...managersList].sort((a, b) => compareEs(a.displayName || '', b.displayName || ''));
    const rootLower = effectiveRootUpn.toLowerCase();
    let rootUser = users.find(u => (u.userPrincipalName || '').toLowerCase() === rootLower) || null;
    if (!rootUser && effectiveRootUpn) {
      rootUser = await getAdminUser(token, effectiveRootUpn);
    }
    state.orgRootUser = rootUser;
    if (rootUser) {
      const managersMap = await fetchManagersBatch(token, users.map(u => u.id));
      const managerIdMap = new Map();
      for (const [uid, mgr] of managersMap) {
        managerIdMap.set(uid, mgr ? mgr.id : null);
      }
      state.orgAnalysis = analyzeOrgChartIntegrity(users, managerIdMap, rootUser.id);
    } else {
      state.errorMsg = `No se ha encontrado la cuenta raíz del organigrama (${effectiveRootUpn}).`;
    }
  } else if (section === 'zona-horaria' && runTz) {
    const { users } = await listAdminUsers(token, { onlyDisabled: false });
    // Graph only lets an app read other people's mailboxSettings with the
    // application permission MailboxSettings.ReadWrite (delegated tokens only
    // open /me/mailboxSettings or delegated shared mailboxes)
    let mailboxToken = token;
    try {
      mailboxToken = (await getAppToken()) || token;
    } catch {
      mailboxToken = token;
    }

    const alreadyCorrect = [];
    const needsUpdate = [];
    const noMailbox = [];
    const accessDenied = [];
    const errors = [];

    const checks = await mapWithConcurrency(users, 8, async (u) => ({
      user: u,
      res: await getUserMailboxSettingsAdmin(mailboxToken, u.id),
    }));

    for (const { user: u, res } of checks) {
      if (res.status === 'no-mailbox') {
        noMailbox.push(u);
      } else if (res.status === 'access-denied') {
        accessDenied.push(u);
      } else if (res.status === 'error') {
        errors.push({ user: u, errorCode: res.errorCode });
      } else {
        const check = classifyMailboxTimezone(res.settings, TARGET_TIMEZONE);
        if (check.ok) {
          alreadyCorrect.push(u);
        } else {
          needsUpdate.push({
            id: u.id,
            displayName: u.displayName || u.userPrincipalName,
            upn: u.userPrincipalName,
            timeZone: check.timeZone,
            workingHoursTimeZone: check.workingHoursTimeZone,
          });
        }
      }
    }

    needsUpdate.sort((a, b) => compareEs(a.displayName, b.displayName));
    state.tzAudit = {
      total: users.length,
      alreadyCorrect,
      needsUpdate,
      noMailbox,
      accessDenied,
      errors,
    };
  }

  return state;
}

router.get('/admin/usuarios/calidad', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const section = ['puestos', 'organigrama', 'zona-horaria'].includes(req.query.section)
      ? req.query.section
      : 'puestos';
    const state = await loadQualityPage(req.session.accessToken, {
      section,
      rootUpn: String(req.query.rootUpn || '').trim(),
      runTz: req.query.runTz === '1',
    });
    state.flash = popSessionValue(req, 'adminUsersFlash') || state.flash;
    state.errorMsg = popSessionValue(req, 'adminUsersError') || state.errorMsg;
    res.render('admin-users-quality', state);
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/calidad/puestos', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const token = req.session.accessToken;
    const selectedIds = new Set(asArray(req.body.userIds).filter(isValidId));
    const { users } = await listAdminUsers(token, { onlyDisabled: false });
    const candidates = findJobTitleFixes(users).filter(c => selectedIds.has(c.id));

    if (candidates.length === 0) {
      const state = await loadQualityPage(token, { section: 'puestos' });
      state.errorMsg = 'No se ha seleccionado ningún puesto para normalizar.';
      return res.render('admin-users-quality', state);
    }

    const settled = await mapWithConcurrency(candidates, 10, async (c) => {
      try {
        await updateAdminUser(token, c.id, { jobTitle: toSentenceCase(c.current) });
        return { ...c, ok: true, error: null };
      } catch (err) {
        if (err.graphStatus === 401) throw err;
        return { ...c, ok: false, error: friendlyGraphError(err, 'Error al actualizar') };
      }
    });

    const updated = settled.filter(r => r.ok).length;
    const failed = settled.filter(r => !r.ok);
    if (updated > 0) {
      auditRequest(req, 'admin.job_titles_fixed', { updatedCount: updated, failedCount: failed.length });
    }

    const state = await loadQualityPage(token, { section: 'puestos' });
    state.jobTitleResult = { updated, failed };
    state.flash = updated > 0 ? `Se han normalizado ${updated} puesto(s) de trabajo.` : null;
    if (failed.length > 0) state.errorMsg = `No se pudieron actualizar ${failed.length} puesto(s).`;
    res.render('admin-users-quality', state);
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/calidad/organigrama', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const token = req.session.accessToken;
    const userId = String(req.body.userId || '').trim();
    const managerId = String(req.body.managerId || '').trim();
    const rootUpn = String(req.body.rootUpn || '').trim();

    if (!isValidId(userId) || !isValidId(managerId) || userId === managerId) {
      const state = await loadQualityPage(token, { section: 'organigrama', rootUpn });
      state.errorMsg = 'Selecciona un responsable válido distinto del propio usuario.';
      return res.render('admin-users-quality', state);
    }

    let orgResult = null;
    try {
      await setUserManager(token, userId, managerId);
      auditRequest(req, 'admin.manager_assigned', { targetUserId: userId, managerId, source: 'org-chart-fix' });
      orgResult = { ok: true };
    } catch (err) {
      if (err.graphStatus === 401) throw err;
      const state = await loadQualityPage(token, { section: 'organigrama', rootUpn });
      state.errorMsg = friendlyGraphError(err, 'No se pudo asignar el responsable.');
      return res.render('admin-users-quality', state);
    }

    const state = await loadQualityPage(token, { section: 'organigrama', rootUpn });
    state.orgResult = orgResult;
    state.flash = 'Responsable actualizado en el organigrama.';
    res.render('admin-users-quality', state);
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/calidad/zona-horaria', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const token = req.session.accessToken;
    const userIds = asArray(req.body.userIds).filter(isValidId);

    if (userIds.length === 0) {
      const state = await loadQualityPage(token, { section: 'zona-horaria', runTz: false });
      state.errorMsg = 'No se ha seleccionado ningún buzón para corregir.';
      return res.render('admin-users-quality', state);
    }

    let mailboxToken = token;
    try {
      mailboxToken = (await getAppToken()) || token;
    } catch {
      mailboxToken = token;
    }

    let updated = 0;
    const failed = [];

    const results = await mapWithConcurrency(userIds, 6, async (uid) => {
      try {
        // Re-read mailbox settings live so existing workingHours days/times are preserved
        const current = await getUserMailboxSettingsAdmin(mailboxToken, uid);
        if (current.status !== 'ok') {
          return { ok: false, userId: uid, error: current.status === 'access-denied' ? 'Acceso denegado (403)' : 'Sin buzón accesible' };
        }
        const payload = buildTimezonePatchPayload(current.settings, TARGET_TIMEZONE);
        await updateUserMailboxSettingsAdmin(mailboxToken, uid, payload);
        return { ok: true, userId: uid };
      } catch (err) {
        if (err.graphStatus === 401) throw err;
        return { ok: false, userId: uid, error: friendlyGraphError(err, 'No se pudo actualizar la zona horaria') };
      }
    });

    for (const r of results) {
      if (r.ok) updated += 1;
      else failed.push({ userId: r.userId, error: r.error });
    }

    if (updated > 0) {
      auditRequest(req, 'admin.timezones_fixed', { updatedCount: updated, failedCount: failed.length, timeZone: TARGET_TIMEZONE });
    }

    const state = await loadQualityPage(token, { section: 'zona-horaria', runTz: false });
    state.tzResult = { updated, failed };
    state.flash = updated > 0 ? `Zona horaria actualizada a ${TARGET_TIMEZONE} en ${updated} buzón(es).` : null;
    if (failed.length > 0) state.errorMsg = `No se pudieron actualizar ${failed.length} buzón(es).`;
    res.render('admin-users-quality', state);
  } catch (error) {
    next(error);
  }
});

// ============================================================================
// 5. Individual User Detail, Edit, Manager, Groups & Password Reset
//    (Parameterized /:id routes MUST stay after all literal routes above!)
// ============================================================================

router.get('/admin/usuarios/:id', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const id = req.params.id;
    if (!isValidId(id)) {
      return next(userError(404, 'Usuario no encontrado'));
    }

    const token = req.session.accessToken;
    const user = await getAdminUser(token, id);
    if (!user) {
      return next(userError(404, 'Usuario no encontrado'));
    }

    const [photoBuf, manager, userGroups, allGroups, licenses, managersList, skus, roomEmails] = await Promise.all([
      getUserPhotoById(token, user.id),
      getUserManagerById(token, user.id),
      getUserGroupsById(token, user.id),
      getAllTenantGroups(token),
      getUserLicensesById(token, user.id),
      getLicensedUsers(token),
      getSubscribedSkus(token),
      getRoomEmailSet(token),
    ]);
    user.isRoom = isRoomUser(user, roomEmails);

    // Licenses that could be given: null when the tenant's list cannot be read
    const availableSkus = skus
      ? skus
        .filter(sku => (sku.capabilityStatus || 'Enabled') === 'Enabled' && sku.appliesTo !== 'Company')
        .map(sku => ({
          skuId: sku.skuId,
          name: SKU_NAMES[sku.skuPartNumber] || sku.skuPartNumber,
          left: (sku.prepaidUnits?.enabled || 0) - (sku.consumedUnits || 0),
        }))
        .sort((a, b) => compareEs(a.name, b.name))
      : null;

    const userPhoto = photoBuf ? `data:image/jpeg;base64,${photoBuf.toString('base64')}` : null;
    const sortedGroups = [...allGroups].sort((a, b) => compareEs(a.displayName || '', b.displayName || ''));
    const sortedManagers = [...managersList].sort((a, b) => compareEs(a.displayName || '', b.displayName || ''));
    const namedLicenses = licenses.map(l => ({
      skuId: l.skuId,
      skuPartNumber: l.skuPartNumber,
      name: SKU_NAMES[l.skuPartNumber] || l.skuPartNumber,
    }));

    res.render('admin-user-edit', {
      title: user.displayName || user.userPrincipalName,
      currentPage: 'admin',
      activeTab: 'directorio',
      user,
      userPhoto,
      manager,
      userGroupIds: userGroups.map(g => g.id),
      allGroups: sortedGroups,
      licenses: namedLicenses,
      availableSkus,
      managersList: sortedManagers,
      resetResult: popSessionValue(req, 'adminUserResetResult'),
      flash: popSessionValue(req, 'adminUsersFlash'),
      errorMsg: popSessionValue(req, 'adminUsersError'),
    });
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/:id', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const id = req.params.id;
    if (!isValidId(id)) {
      return next(userError(404, 'Usuario no encontrado'));
    }
    const token = req.session.accessToken;
    const current = await getAdminUser(token, id);
    if (!current) {
      return next(userError(404, 'Usuario no encontrado'));
    }

    // Read-only here: the alias, the usage location and the language are not
    // changed from this form (the portal keeps language and region itself)
    const { mailNickname, usageLocation, preferredLanguage, ...editable } = req.body;
    const { patch, changes } = diffUserEdit(current, editable);
    if (changes.length === 0) {
      req.session.adminUsersFlash = 'No había cambios que guardar.';
      return res.redirect(`/admin/usuarios/${encodeURIComponent(id)}`);
    }

    try {
      await updateAdminUser(token, current.id, patch);
      auditRequest(req, 'admin.user_updated', {
        targetUserId: current.id,
        targetUpn: current.userPrincipalName,
        fields: changes.map(c => c.field),
      });
      req.session.adminUsersFlash = `Datos actualizados (${changes.length} campo(s)).`;
    } catch (err) {
      if (err.graphStatus === 401) throw err;
      req.session.adminUsersError = friendlyGraphError(err, 'No se pudieron guardar los cambios.');
    }

    res.redirect(`/admin/usuarios/${encodeURIComponent(id)}`);
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/:id/responsable', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const id = req.params.id;
    const managerId = String(req.body.managerId || '').trim();
    if (!isValidId(id)) return next(userError(404, 'Usuario no encontrado'));

    if (!isValidId(managerId) || managerId === id) {
      req.session.adminUsersError = 'Selecciona un responsable válido distinto del propio usuario.';
      return res.redirect(`/admin/usuarios/${encodeURIComponent(id)}`);
    }

    try {
      await setUserManager(req.session.accessToken, id, managerId);
      auditRequest(req, 'admin.manager_assigned', { targetUserId: id, managerId });
      req.session.adminUsersFlash = 'Responsable actualizado correctamente.';
    } catch (err) {
      if (err.graphStatus === 401) throw err;
      req.session.adminUsersError = friendlyGraphError(err, 'No se pudo actualizar el responsable.');
    }

    res.redirect(`/admin/usuarios/${encodeURIComponent(id)}`);
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/:id/grupos', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const id = req.params.id;
    if (!isValidId(id)) return next(userError(404, 'Usuario no encontrado'));

    const token = req.session.accessToken;
    const desiredIds = new Set(asArray(req.body.groupIds).filter(isValidId));
    const currentGroups = await getUserGroupsById(token, id);
    const currentIds = new Set(currentGroups.map(g => g.id));

    const toAdd = [...desiredIds].filter(gid => !currentIds.has(gid));
    const toRemove = [...currentIds].filter(gid => !desiredIds.has(gid));

    if (toAdd.length === 0 && toRemove.length === 0) {
      req.session.adminUsersFlash = 'No hubo cambios en la pertenencia a grupos.';
      return res.redirect(`/admin/usuarios/${encodeURIComponent(id)}`);
    }

    let added = 0;
    let removed = 0;
    const errors = [];

    for (const gid of toAdd) {
      try {
        await addUserToGroup(token, gid, id);
        added += 1;
      } catch (err) {
        if (err.graphStatus === 401) throw err;
        errors.push(friendlyGraphError(err, 'Error al añadir a grupo'));
      }
    }
    for (const gid of toRemove) {
      try {
        await removeUserFromGroup(token, gid, id);
        removed += 1;
      } catch (err) {
        if (err.graphStatus === 401) throw err;
        errors.push(friendlyGraphError(err, 'Error al quitar de grupo'));
      }
    }

    if (added > 0 || removed > 0) {
      auditRequest(req, 'admin.user_updated', {
        targetUserId: id,
        groupsAdded: added,
        groupsRemoved: removed,
      });
      req.session.adminUsersFlash = `Grupos actualizados (+${added} añadidos, -${removed} quitados).`;
    }
    if (errors.length > 0) {
      req.session.adminUsersError = errors[0];
    }

    res.redirect(`/admin/usuarios/${encodeURIComponent(id)}`);
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/:id/licencias', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const id = req.params.id;
    if (!isValidId(id)) return next(userError(404, 'Usuario no encontrado'));

    const token = req.session.accessToken;
    // The form is one list: what is ticked is what the person should have. Only
    // licenses the form showed (shownSkuIds) can change, so an empty or cut-down
    // submission never strips licenses that were not on the page
    const shown = new Set(asArray(req.body.shownSkuIds).filter(isValidId));
    const wanted = new Set(asArray(req.body.skuIds).filter(sku => isValidId(sku) && shown.has(sku)));
    const back = `/admin/usuarios/${encodeURIComponent(id)}`;

    try {
      const user = await getAdminUser(token, id);
      if (!user) return next(userError(404, 'Usuario no encontrado'));
      const current = new Set((await getUserLicensesById(token, user.id)).map(l => l.skuId));
      const toAdd = [...wanted].filter(sku => !current.has(sku));
      const toRemove = [...current].filter(sku => shown.has(sku) && !wanted.has(sku));

      if (toAdd.length === 0 && toRemove.length === 0) {
        req.session.adminUsersFlash = 'No había cambios de licencias que guardar.';
        return res.redirect(back);
      }

      // Microsoft refuses licenses for people without a usage location
      if (toAdd.length > 0 && !user.usageLocation) {
        await updateAdminUser(token, user.id, { usageLocation: 'ES' });
      }
      await assignUserLicenses(token, user.id, { add: toAdd, remove: toRemove });

      auditRequest(req, 'admin.licenses_changed', {
        targetUserId: user.id,
        targetUpn: user.userPrincipalName,
        added: toAdd,
        removed: toRemove,
      });
      req.session.adminUsersFlash = `Licencias actualizadas (+${toAdd.length} añadidas, -${toRemove.length} quitadas).`;
    } catch (err) {
      if (err.graphStatus === 401) throw err;
      req.session.adminUsersError = friendlyGraphError(err, 'No se pudieron cambiar las licencias.');
    }

    res.redirect(back);
  } catch (error) {
    next(error);
  }
});

router.post('/admin/usuarios/:id/reset-password', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const id = req.params.id;
    if (!isValidId(id)) return next(userError(404, 'Usuario no encontrado'));

    const token = req.session.accessToken;
    const user = await getAdminUser(token, id);
    if (!user) return next(userError(404, 'Usuario no encontrado'));

    const newPassword = generateTempPassword();
    try {
      await updateAdminUser(token, user.id, {
        passwordProfile: {
          password: newPassword,
          forceChangePasswordNextSignIn: true,
        },
      });
    } catch (err) {
      if (err.graphStatus === 401) throw err;
      req.session.adminUsersError = friendlyGraphError(err, 'No se pudo restablecer la contraseña.');
      return res.redirect(`/admin/usuarios/${encodeURIComponent(id)}`);
    }

    // NEVER log the generated password in audit
    auditRequest(req, 'admin.user_password_reset', {
      targetUserId: user.id,
      targetUpn: user.userPrincipalName,
    });

    const links = buildResetCredentialLinks(user.userPrincipalName, newPassword, {
      intranetUrl: appConfig.INTRANET_URL,
      welcomeUrl: appConfig.WELCOME_GUIDE_URL,
    });

    req.session.adminUserResetResult = {
      password: newPassword,
      upn: user.userPrincipalName,
      displayName: user.displayName || user.userPrincipalName,
      ...links,
    };

    res.redirect(`/admin/usuarios/${encodeURIComponent(id)}`);
  } catch (error) {
    next(error);
  }
});

module.exports = router;
