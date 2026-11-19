/**
 * Signature Route
 *
 * - GET /firma - the person's email signature, made from their profile in the
 *   format of the organization's generator (github.com/arandadeduero/firmas-correo).
 *   Nothing is typed: the only choice is the concejalía, when the profile does
 *   not say it (or says it differently).
 */

const express = require('express');
const router = express.Router();
const appConfig = require('../config/appConfig');
const { isAuthenticated } = require('../middleware/auth');
const { getUserProfile } = require('../services/graphService');
const { CONCEJALIAS, findConcejaliaKey, buildSignatureHtml } = require('../services/signature');

router.get('/firma', isAuthenticated, async (req, res, next) => {
  try {
    const profile = await getUserProfile(req.session.accessToken);

    const person = {
      name: profile.displayName,
      position: profile.jobTitle,
      email: profile.mail || profile.userPrincipalName,
      phone: [...(profile.businessPhones || []), profile.mobilePhone].filter(Boolean)[0] || '',
    };
    // What the profile lacks; only the position is something the person can fill in
    // themselves (read at request time: it depends on PROFILE_EDITABLE_FIELDS)
    const LABELS = { name: 'el nombre', position: 'el puesto', email: 'el correo' };
    const missing = Object.keys(LABELS).filter(field => !person[field]).map(field => ({
      label: LABELS[field],
      canEdit: field === 'position' && appConfig.PROFILE_EDITABLE_FIELDS.includes('jobTitle'),
    }));

    // hasOwn: a plain lookup would also accept "constructor" or "__proto__"
    const chosen = Object.hasOwn(CONCEJALIAS, req.query.concejalia) ? req.query.concejalia : null;
    const key = chosen || findConcejaliaKey(profile.department, profile.jobTitle);

    // Only complete data makes a signature; a half-made one would get sent
    const ready = missing.length === 0 && key !== null;
    res.render('signature', {
      title: 'Mi firma de correo',
      currentPage: 'account',
      concejalias: CONCEJALIAS,
      key,
      department: profile.department || '',
      missing,
      // The logos load from their public address, also in the preview: it is
      // what a recipient gets, and recipients cannot reach this portal
      signatureHtml: ready ? buildSignatureHtml(person, key, appConfig.SIGNATURE_IMAGES_URL) : '',
    });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
