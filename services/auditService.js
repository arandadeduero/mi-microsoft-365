/**
 * Audit Service
 *
 * Keeps a record of who did what: one JSON object per line, in a file per
 * day under AUDIT_DIR (./audit by default). Passwords, tokens and message
 * contents are never written.
 */

const fs = require('fs');
const path = require('path');

function auditDir() {
  return process.env.AUDIT_DIR || path.join(__dirname, '..', 'audit');
}

let lastEnsuredDir = null;

async function ensureDir(dir) {
  if (lastEnsuredDir === dir) return;
  await fs.promises.mkdir(dir, { recursive: true });
  lastEnsuredDir = dir;
}

/**
 * Record an event
 * @param {string} event - What happened, e.g. 'profile.update'
 * @param {Object} details - Who and what it affected
 * @returns {Promise<void>} Never rejects: a failure to write is logged, so it cannot break the action being recorded
 */
async function audit(event, details = {}) {
  const now = new Date();
  const entry = { time: now.toISOString(), event, ...details };
  const line = JSON.stringify(entry) + '\n';

  try {
    const dir = auditDir();
    const file = path.join(dir, `${entry.time.slice(0, 10)}.log`);
    await ensureDir(dir);
    try {
      await fs.promises.appendFile(file, line);
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
      // Directory may have been removed since last write
      lastEnsuredDir = null;
      await ensureDir(dir);
      await fs.promises.appendFile(file, line);
    }
  } catch (error) {
    console.error('Could not write audit log:', error.message, entry);
  }
}

/**
 * Record an event for a request, adding the signed-in user and client address
 * @param {Object} req - Express request
 * @param {string} event - What happened
 * @param {Object} details - What it affected
 * @returns {Promise<void>}
 */
function auditRequest(req, event, details = {}) {
  return audit(event, {
    actor: req.session?.account?.username || null,
    ip: req.ip,
    ...details,
  });
}

const DAY_FILE = /^\d{4}-\d{2}-\d{2}\.log$/;
const SAFE_RAW_NEEDLE = /^[\x20-\x7e]+$/;

/**
 * Read the audit log, newest entry first
 * @param {Object} options
 * @param {string} [options.day] - Only this day (YYYY-MM-DD); otherwise the most recent days
 * @param {number} [options.days] - How many of the most recent days to read
 * @param {string} [options.actor] - Optional exact actor email (case-insensitive) to filter before retaining in memory
 * @returns {Promise<Array>} Entries; lines that cannot be read are skipped
 */
async function readAudit({ day, days = 7, actor } = {}) {
  const dir = auditDir();
  let files;
  try {
    files = (await fs.promises.readdir(dir)).filter(file => DAY_FILE.test(file)).sort().reverse();
  } catch (error) {
    if (error.code === 'ENOENT') {
      return [];
    }
    throw error;
  }

  // The day is matched against the files that exist, never used to build a path
  files = day ? files.filter(file => file === `${day}.log`) : files.slice(0, days);
  if (files.length === 0) return [];

  const wantedActor = actor ? String(actor).toLowerCase() : null;
  const rawNeedle = wantedActor && SAFE_RAW_NEEDLE.test(wantedActor) && !wantedActor.includes('"') && !wantedActor.includes('\\')
    ? wantedActor
    : null;

  const contents = await Promise.all(
    files.map(file => fs.promises.readFile(path.join(dir, file), 'utf8').catch(() => ''))
  );

  const entries = [];
  for (const content of contents) {
    const lines = content.split('\n');
    for (let i = lines.length - 1; i >= 0; i -= 1) {
      const line = lines[i];
      if (!line) continue;
      if (rawNeedle && !line.toLowerCase().includes(rawNeedle)) continue;
      try {
        const entry = JSON.parse(line);
        if (wantedActor && (entry.actor || '').toLowerCase() !== wantedActor) continue;
        entries.push(entry);
      } catch (error) {
        // A line cut short by a crash; the rest of the file is still good
      }
    }
  }
  return entries;
}

module.exports = { audit, auditRequest, readAudit };
