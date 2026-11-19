/**
 * Internal Requests Routes
 *
 * - GET /peticiones - what a person can ask for inside the Ayuntamiento: the
 *   IT support form (config/internalRequests.js) and a link to a Microsoft Form
 * - POST /peticiones/soporte - send the support form to IT_SUPPORT_EMAIL, from the
 *   person's own mailbox so that the answer comes back to them (like /equipo)
 */

const express = require('express');
const router = express.Router();
const appConfig = require('../config/appConfig');
const REQUESTS = require('../config/internalRequests');
const { isAuthenticated } = require('../middleware/auth');
const { userError } = require('../middleware/errorHandler');
const { auditRequest } = require('../services/auditService');
const { getUserProfile, sendEmail } = require('../services/graphService');
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
      sent: req.query.sent === 'soporte',
      supportEmail: appConfig.IT_SUPPORT_EMAIL,
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

    // The department is not typed: it comes from the profile, when it can be read
    let department = '';
    try {
      department = (await getUserProfile(req.session.accessToken)).department || '';
    } catch (error) {
      if (error.graphStatus === 401) throw error;
    }
    const { account } = req.session;
    const body = [
      `${account.name || account.username} (${account.username}) pide desde el portal: ${support.title}.`,
      ...(department ? [`Departamento: ${department}`] : []),
      '',
      ...values.filter(([, value]) => value).map(([label, value]) => `${label}: ${value}`),
    ].join('\n');
    // A subject is one line; what was typed may not be
    const subject = String(req.body.subject || '').replace(/\s+/g, ' ').trim().slice(0, 120);

    try {
      await sendEmail(req.session.accessToken, {
        to: appConfig.IT_SUPPORT_EMAIL,
        subject: `[${appConfig.APP_NAME}] ${support.title}: ${subject}`,
        body,
      });
    } catch (error) {
      if (error.graphStatus === 401) throw error;
      throw userError(502, `No se ha podido enviar la petición. Escribe directamente a ${appConfig.IT_SUPPORT_EMAIL}.`);
    }
    // What was asked for, not what was typed
    auditRequest(req, support.event, { kind: 'soporte' });

    res.redirect('/peticiones?sent=soporte');
  } catch (error) {
    next(error);
  }
});

module.exports = router;
