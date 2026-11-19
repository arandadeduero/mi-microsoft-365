/**
 * What administrators should look at among the accounts. Pure functions.
 */

const { compareEs } = require('./formatters');
const { isRoomUser } = require('./adminUsersTools');

const DAY_MS = 24 * 60 * 60 * 1000;
const IDLE_DAYS = 30;

// The last successful one when Graph gives it (since December 2023); the last
// attempt, interactive or not, otherwise
const lastSignIn = (user) => {
  const activity = user.signInActivity || {};
  const dates = [activity.lastSuccessfulSignInDateTime, activity.lastSignInDateTime, activity.lastNonInteractiveSignInDateTime].filter(Boolean).map(date => new Date(date).getTime());
  return dates.length ? new Date(Math.max(...dates)) : null;
};

const daysSince = (date, now) => Math.floor((now - new Date(date).getTime()) / DAY_MS);

/**
 * @param {Array} users - Users from Graph (see REVIEW_FIELDS)
 * @param {Object} options - signInsKnown, mfa (userRegistrationDetails or null), passwordMaxAgeDays, now
 * @returns {Object} Lists of accounts, each with why it is there
 */
function reviewAccounts(users, { signInsKnown = false, mfa = null, passwordMaxAgeDays = null, roomEmails = null, now = Date.now() } = {}) {
  const mfaById = mfa ? new Map(mfa.map(entry => [entry.id, entry])) : null;
  const person = (user, note) => ({ id: user.id, name: user.displayName || user.userPrincipalName, username: user.userPrincipalName, note });

  const idle = signInsKnown ? [] : null;
  const expiredPassword = passwordMaxAgeDays ? [] : null;
  const disabledWithLicense = [];
  const enabledWithoutLicense = [];
  const withoutMfa = mfa ? [] : null;
  let total = 0;
  let guests = 0;

  for (const user of users) {
    if (user.userType === 'Guest') {
      guests += 1;
      continue;
    }
    if (isRoomUser(user, roomEmails)) continue;
    total += 1;
    const isLicensed = (user.assignedLicenses || []).length > 0;

    if (!user.accountEnabled) {
      if (isLicensed) disabledWithLicense.push(person(user, 'Cuenta deshabilitada con licencia asignada'));
      continue;
    }

    if (!isLicensed) {
      enabledWithoutLicense.push(person(user, 'Sin licencia'));
      continue;
    }

    if (idle) {
      const last = lastSignIn(user);
      if (!last) {
        if (daysSince(user.createdDateTime || now, now) > IDLE_DAYS) {
          idle.push(person(user, 'Nunca ha iniciado sesión'));
        }
      } else {
        const days = daysSince(last, now);
        if (days > IDLE_DAYS) idle.push(person(user, `Último inicio de sesión hace ${days} días`));
      }
    }

    if (expiredPassword && user.lastPasswordChangeDateTime) {
      const days = daysSince(user.lastPasswordChangeDateTime, now);
      if (days > passwordMaxAgeDays) {
        expiredPassword.push(person(user, `Contraseña de hace ${days} días`));
      }
    }

    if (withoutMfa && mfaById.has(user.id) && mfaById.get(user.id).isMfaRegistered === false) {
      withoutMfa.push(person(user, 'Sin segundo factor registrado'));
    }
  }

  const sortByName = (list) => list && list.sort((a, b) => compareEs(a.name, b.name));
  return {
    total,
    guests,
    idle: sortByName(idle),
    expiredPassword: sortByName(expiredPassword),
    disabledWithLicense: sortByName(disabledWithLicense),
    enabledWithoutLicense: sortByName(enabledWithoutLicense),
    withoutMfa: sortByName(withoutMfa),
  };
}

/**
 * The licenses bought, with how many are left
 * @param {Array} skus - subscribedSkus from Graph
 * @param {Object} skuNames - code → name
 */
function describeLicenseStock(skus, skuNames) {
  return skus
    .filter(sku => (sku.prepaidUnits?.enabled || 0) > 0)
    .map((sku) => {
      const total = sku.prepaidUnits.enabled;
      const used = sku.consumedUnits || 0;
      const left = total - used;
      return {
        skuId: sku.skuId || null,
        code: sku.skuPartNumber,
        name: skuNames[sku.skuPartNumber] || sku.skuPartNumber,
        total,
        used,
        left,
        // Running out: none left, or under a tenth
        low: left <= 0 || left < Math.max(2, Math.ceil(total / 10)),
      };
    })
    // Running out first, the fewest left before the rest
    .sort((a, b) => Number(b.low) - Number(a.low) || a.left - b.left || b.total - a.total);
}

module.exports = { reviewAccounts, describeLicenseStock };
