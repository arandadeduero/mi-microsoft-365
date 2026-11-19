/**
 * Excel Drift
 *
 * A daily, read-only check of the users Excel (USERS_EXCEL_URL) against
 * Microsoft 365: people to disable, to create, with other data or another
 * manager. It applies nothing and keeps only numbers (no names), in one JSON
 * file in DATA_DIR that /admin/usuarios and /admin/usuarios/excel show as a
 * notice. Runs with the app's token, since nobody is signed in; if that cannot
 * read the file the notice says so. One instance only, like requestStore.
 */

const fs = require('fs');
const path = require('path');
const appConfig = require('../config/appConfig');
const permissions = require('../config/permissions');
const { audit } = require('./auditService');
const { getAppToken } = require('./tokenService');
const { listAdminUsers, fetchManagersBatch, downloadSharedFile, getRooms } = require('./graphService');
const { readExcelBuffer, analyzeSyncCheck, parseImportRows, buildImportPlan, previewManagerChanges } = require('./adminUsersExcel');

const FILE = 'excel-drift.json';
const EVERY_MS = 24 * 60 * 60 * 1000;
// A restart right after a check does not repeat it
const MIN_GAP_MS = 12 * 60 * 60 * 1000;
const FIRST_RUN_DELAY_MS = 60 * 1000;

const storeFile = () => path.join(process.env.DATA_DIR || path.join(__dirname, '..', 'data'), FILE);
const currentDomain = () => permissions.DOMAIN || 'arandadeduero.es';

/**
 * @returns {Object|null} The last check, or null when there has been none
 */
function readExcelDrift() {
  try {
    return JSON.parse(fs.readFileSync(storeFile(), 'utf8'));
  } catch (error) {
    if (error.code !== 'ENOENT') console.error('Error reading the Excel check:', error.message);
    return null;
  }
}

/**
 * What differs between the Excel and Microsoft 365, as numbers
 * @param {Buffer} buffer - The .xlsx
 * @param {string} token - Token to read the directory with
 */
async function computeExcelDrift(buffer, token) {
  const { activos, concejales } = await readExcelBuffer(buffer);
  const [{ users }, places] = await Promise.all([listAdminUsers(token, { onlyDisabled: null }), getRooms(token)]);
  const roomEmails = new Set((places || []).map(place => String(place.emailAddress || '').trim().toLowerCase()).filter(Boolean));
  const sync = analyzeSyncCheck(activos, concejales, users, currentDomain(), { roomEmails });

  const parsed = parseImportRows(activos, currentDomain(), { skipLeaving: true });
  const plan = buildImportPlan(parsed.validRows, users);
  let managerChanges = null; // null: the current managers could not be read
  let managerNotFound = null;
  try {
    const withManager = plan.items.filter(item => item.managerRef && item.existingId).map(item => item.existingId);
    previewManagerChanges(plan, users, await fetchManagersBatch(token, withManager), currentDomain());
    managerChanges = plan.summary.managerChanges;
    managerNotFound = plan.summary.managerNotFound;
  } catch (error) {
    if (error.graphStatus === 401) throw error;
  }
  return {
    toDisable: sync.notDisabled.length,
    toCreate: sync.missing.length,
    notInExcel: sync.cloudOnly.length,
    fieldChanges: plan.summary.update,
    managerChanges,
    managerNotFound,
    invalidRows: parsed.invalidRows.length,
  };
}

/**
 * Read the Excel and keep what differs. Never throws: a failure is kept too,
 * so the notice can say that the check did not work.
 * @param {{force?: boolean}} [options] - force: also when one was done recently
 * @returns {Promise<Object|null>} The result, or null when there is no Excel configured
 */
async function runExcelDriftCheck({ force = false } = {}) {
  const url = String(appConfig.USERS_EXCEL_URL || '').trim();
  if (!url) return null;
  const last = readExcelDrift();
  if (!force && last && Date.now() - Date.parse(last.checkedAt) < MIN_GAP_MS) return last;

  const checkedAt = new Date().toISOString();
  let result;
  try {
    const token = await getAppToken();
    if (!token) throw new Error('No token for the application');
    result = { checkedAt, ...(await computeExcelDrift(await downloadSharedFile(token, url), token)) };
    await audit('excel.drift_checked', { toDisable: result.toDisable, toCreate: result.toCreate, fieldChanges: result.fieldChanges, managerChanges: result.managerChanges });
  } catch (error) {
    result = { checkedAt, error: error.graphStatus ? `Microsoft Graph ha respondido ${error.graphStatus}` : 'No se ha podido leer o comprobar el archivo' };
    await audit('excel.drift_failed', { status: error.graphStatus || null });
  }
  try {
    fs.mkdirSync(path.dirname(storeFile()), { recursive: true });
    fs.writeFileSync(storeFile(), JSON.stringify(result));
  } catch (error) {
    console.error('Error saving the Excel check:', error.message);
  }
  return result;
}

/**
 * Check once shortly after start-up and then every day (only when an Excel is configured)
 */
function startExcelDriftSchedule() {
  if (!appConfig.USERS_EXCEL_URL) return;
  const run = () => runExcelDriftCheck();
  setTimeout(run, FIRST_RUN_DELAY_MS).unref();
  setInterval(run, EVERY_MS).unref();
}

module.exports = { readExcelDrift, computeExcelDrift, runExcelDriftCheck, startExcelDriftSchedule };
