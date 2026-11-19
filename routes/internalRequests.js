/**
 * Internal Requests Routes
 *
 * - GET /peticiones - what a person can ask for inside the Ayuntamiento: the
 *   IT support form (config/internalRequests.js) and a link to a Microsoft Form
 * - POST /peticiones/soporte - check the support form and show what was typed
 *
 * shortcut: it is a trial, so the support form is checked and shown back but
 * sent nowhere and kept nowhere (only that it was filled in is audited); where
 * it goes (email, a list, Planner, the Form) is decided later, in one place:
 * the end of the POST handler.
 */

const express = require('express');
const router = express.Router();
const appConfig = require('../config/appConfig');
const REQUESTS = require('../config/internalRequests');
const { isAuthenticated } = require('../middleware/auth');
const { userError } = require('../middleware/errorHandler');
const { auditRequest } = require('../services/auditService');
const { getUserProfile } = require('../services/graphService');
const { readRequestValues } = require('../services/requestForm');

// What the profile can fill in, by the names config/internalRequests.js uses
function profilePrefill(profile) {
  return {
    officeLocation: profile.officeLocation || '',
    phone: [...(profile.businessPhones || []), profile.mobilePhone].filter(Boolean)[0] || '',
  };
}

// The link to the Microsoft Form comes from .env; only a web address is offered
function formLink() {
  const url = String(appConfig.INTERNAL_REQUESTS_FORM_URL || '').trim();
  return /^https:\/\//i.test(url) ? url : null;
}

router.get('/peticiones', isAuthenticated, async (req, res, next) => {
  try {
    let profile = {};
    try {
      profile = await getUserProfile(req.session.accessToken);
    } catch (error) {
      if (error.graphStatus === 401) throw error; // without the profile the form is just empty
    }
    const prefill = profilePrefill(profile);
    const support = REQUESTS.soporte;
    res.render('internal-requests', {
      title: 'Peticiones internas',
      currentPage: 'requests',
      support,
      prefill: Object.fromEntries(support.fields.filter(field => field.prefill).map(field => [field.name, prefill[field.prefill]])),
      formLink: formLink(),
      contact: { name: profile.displayName || req.session.account.name, email: profile.mail || req.session.account.username, department: profile.department || '' },
    });
  } catch (error) {
    next(error);
  }
});

router.post('/peticiones/soporte', isAuthenticated, async (req, res, next) => {
  try {
    const support = REQUESTS.soporte;
    const values = readRequestValues(support, req.body);
    // What was asked for, not what was typed
    auditRequest(req, support.event, { kind: 'soporte' });
    res.render('internal-request-sent', {
      title: 'Soporte informático',
      currentPage: 'requests',
      request: support,
      values: values.filter(([, value]) => value),
      who: `${req.session.account.name || req.session.account.username} (${req.session.account.username})`,
    });
  } catch (error) {
    next(error instanceof Error ? error : userError(400, 'Revisa el formulario.'));
  }
});

module.exports = router;
