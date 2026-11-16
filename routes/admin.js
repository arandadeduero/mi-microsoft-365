/**
 * Admin Routes
 *
 * Pages for administrators only (users holding an Entra ID directory role):
 * - /admin/audit (and /admin/audit.csv) - The audit log
 * - /admin/cuentas - Accounts that need a look
 * - /admin/licencias - Licenses bought and left
 * - /admin/avisos - Microsoft's message center
 * - /admin/solicitudes - Group requests nobody has decided on
 * - /admin/salas - A room's calendar for a week: who booked it and who goes
 * - /admin/ajustes - The blog list and the shared mailboxes
 */

const express = require('express');
const router = express.Router();
const { isAuthenticated } = require('../middleware/auth');
const { userError } = require('../middleware/errorHandler');
const { auditRequest, readAudit } = require('../services/auditService');
const permissions = require('../config/permissions');
const { SKU_NAMES } = require('../config/licenseCatalog');
const { getAppToken } = require('../services/tokenService');
const { REGION } = require('../config/appConfig');
const { getSubscribedSkus, getUsersForReview, getMfaRegistration, getRooms, getRoomEvents } = require('../services/graphService');
const { reviewAccounts, describeLicenseStock } = require('../services/accountReview');
const { getServiceMessages } = require('../services/messageCenterService');
const { listPending } = require('../services/requestStore');
const { currentSettings, saveSettings } = require('../services/settingsStore');

// Entries shown at most, so a busy week cannot produce an enormous page
const MAX_ENTRIES = 500;

/**
 * Allow only users that checkAdminStatus found to hold a directory role
 */
function requireAdmin(req, res, next) {
  if (res.locals.isAdmin) {
    return next();
  }
  auditRequest(req, 'permission.denied', { path: req.path, method: req.method });
  next(userError(403, 'Esta página es solo para administradores'));
}

/**
 * GET /admin/audit
 * The audit log, newest first: the last 7 days, or one day when given,
 * optionally only one kind of event or one person
 */
router.get('/admin/audit', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const filters = {
      day: String(req.query.day || ''),
      event: String(req.query.event || ''),
      actor: String(req.query.actor || '').trim().toLowerCase(),
    };

    const all = await readAudit({ day: filters.day });
    const matching = all.filter(entry =>
      (!filters.event || entry.event === filters.event) &&
      (!filters.actor || (entry.actor || '').toLowerCase().includes(filters.actor))
    );

    res.render('audit', {
      title: 'Auditoría',
      currentPage: 'audit',
      filters,
      events: [...new Set(all.map(entry => entry.event))].sort(),
      entries: matching.slice(0, MAX_ENTRIES),
      total: matching.length,
    });
  } catch (error) {
    next(error);
  }
});

/**
 * Entries that match the audit filters of the query
 */
async function filteredAudit(query) {
  const filters = {
    day: String(query.day || ''),
    event: String(query.event || ''),
    actor: String(query.actor || '').trim().toLowerCase(),
  };
  const all = await readAudit({ day: filters.day });
  const matching = all.filter(entry =>
    (!filters.event || entry.event === filters.event) &&
    (!filters.actor || (entry.actor || '').toLowerCase().includes(filters.actor))
  );
  return { filters, all, matching };
}

/**
 * GET /admin/audit.csv
 * The same entries as the page, for a spreadsheet
 */
router.get('/admin/audit.csv', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const { matching } = await filteredAudit(req.query);
    const cell = (value) => `"${String(value == null ? '' : value).replace(/"/g, '""')}"`;
    const rows = matching.map(({ time, event, actor, ip, ...details }) => [time, event, actor, ip, JSON.stringify(details)].map(cell).join(';'));
    res.set('Content-Type', 'text/csv; charset=utf-8');
    res.set('Content-Disposition', 'attachment; filename="auditoria.csv"');
    res.send('﻿' + [['Fecha', 'Evento', 'Usuario', 'IP', 'Detalles'].map(cell).join(';'), ...rows].join('\r\n') + '\r\n');
  } catch (error) {
    next(error);
  }
});

/**
 * GET /admin/cuentas
 * Accounts that need a look, read with the app's token
 */
router.get('/admin/cuentas', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const appToken = await getAppToken();
    const listing = await getUsersForReview(appToken);
    const review = listing
      ? reviewAccounts(listing.users, {
        signInsKnown: listing.signInsKnown,
        mfa: await getMfaRegistration(appToken),
        passwordMaxAgeDays: Number(process.env.PASSWORD_MAX_AGE_DAYS) || null,
      })
      : null;

    res.render('admin-accounts', { title: 'Cuentas que necesitan atención', currentPage: 'audit', review, signInsError: listing ? listing.signInsError : null });
  } catch (error) {
    next(error);
  }
});

/**
 * GET /admin/licencias
 */
router.get('/admin/licencias', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const skus = await getSubscribedSkus(req.session.accessToken);
    res.render('admin-licenses', { title: 'Licencias', currentPage: 'audit', licenses: skus && describeLicenseStock(skus, SKU_NAMES) });
  } catch (error) {
    next(error);
  }
});

/**
 * GET /admin/avisos
 */
router.get('/admin/avisos', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    res.render('admin-messages', { title: 'Avisos de Microsoft', currentPage: 'audit', messages: await getServiceMessages() });
  } catch (error) {
    next(error);
  }
});

/**
 * GET /admin/solicitudes
 */
router.get('/admin/solicitudes', isAuthenticated, requireAdmin, (req, res, next) => {
  try {
    const now = Date.now();
    const pending = listPending().map(request => ({ ...request, days: Math.floor((now - new Date(request.createdAt).getTime()) / 86400000) }));
    res.render('admin-requests', { title: 'Solicitudes pendientes', currentPage: 'audit', pending });
  } catch (error) {
    next(error);
  }
});

const DATE = /^\d{4}-\d{2}-\d{2}$/;
const addDays = (date, days) => new Date(Date.parse(`${date}T00:00:00Z`) + days * 86400000).toISOString().slice(0, 10);

/**
 * GET /admin/salas
 * One room's calendar for a week, read with the app's token. The room comes
 * from the list of rooms and nowhere else: that token opens every mailbox.
 */
router.get('/admin/salas', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const places = await getRooms(req.session.accessToken);
    const rooms = places && places
      .filter(place => place.emailAddress)
      .map(place => ({ name: place.displayName || place.emailAddress, email: place.emailAddress, capacity: place.capacity || null }))
      .sort((a, b) => a.name.localeCompare(b.name, 'es'));
    const wanted = String(req.query.room || '').toLowerCase();
    const room = (rooms || []).find(candidate => candidate.email.toLowerCase() === wanted) || null;

    // sv-SE formats as "2026-10-12"
    const today = new Intl.DateTimeFormat('sv-SE', { timeZone: 'Europe/Madrid' }).format(new Date());
    const asked = DATE.test(req.query.week) && !Number.isNaN(Date.parse(req.query.week)) ? req.query.week : today;
    const monday = addDays(asked, -((new Date(`${asked}T00:00:00Z`).getUTCDay() + 6) % 7));

    // The period is asked in UTC, a day wider on each side; the days below keep what is theirs
    const events = room
      ? await getRoomEvents(await getAppToken(), room.email, `${addDays(monday, -1)}T00:00:00`, `${addDays(monday, 8)}T00:00:00`, REGION.timeZone)
      : null;

    const days = events && [0, 1, 2, 3, 4, 5, 6].map(offset => {
      const date = addDays(monday, offset);
      const next = addDays(date, 1);
      return {
        date,
        events: events
          // Times are local and to the minute, so text compares as time; multi-day events land on each of their days
          .filter(event => !event.isCancelled && event.start.dateTime.slice(0, 16) < `${next}T00:00` && event.end.dateTime.slice(0, 16) > `${date}T00:00`)
          .map(event => ({
            from: event.isAllDay ? null : (event.start.dateTime.startsWith(date) ? event.start.dateTime.slice(11, 16) : '00:00'),
            to: event.isAllDay ? null : (event.end.dateTime.startsWith(date) ? event.end.dateTime.slice(11, 16) : '24:00'),
            subject: (event.subject || '').trim(),
            organizer: event.organizer?.emailAddress || {},
            // The room is an attendee of its own bookings
            people: (event.attendees || []).filter(person => (person.emailAddress?.address || '').toLowerCase() !== room.email.toLowerCase()),
            pending: event.showAs === 'tentative',
            isPrivate: event.sensitivity === 'private' || event.sensitivity === 'confidential',
            recurring: Boolean(event.type) && event.type !== 'singleInstance',
            online: Boolean(event.isOnlineMeeting),
            created: event.createdDateTime || null,
          })),
      };
    });

    res.render('admin-rooms', {
      title: 'Ocupación de salas',
      currentPage: 'audit',
      rooms,
      room,
      today,
      monday,
      previous: addDays(monday, -7),
      following: addDays(monday, 7),
      days,
    });
  } catch (error) {
    next(error);
  }
});

const USERNAME = /^[a-z0-9._-]+$/i;
const parseLines = (text) => [...new Set(String(text || '').split(/[\r\n,;]+/).map(line => line.trim().toLowerCase()).filter(Boolean))];

/**
 * GET /admin/ajustes
 */
router.get('/admin/ajustes', isAuthenticated, requireAdmin, (req, res, next) => {
  try {
    const saved = req.session.settingsSaved || false;
    delete req.session.settingsSaved;
    res.render('admin-settings', { title: 'Ajustes', currentPage: 'audit', settings: currentSettings(), domain: permissions.DOMAIN, saved });
  } catch (error) {
    next(error);
  }
});

/**
 * POST /admin/ajustes
 * Save who may publish on the blog and which shared mailboxes exist
 */
router.post('/admin/ajustes', isAuthenticated, requireAdmin, (req, res, next) => {
  try {
    const blogUsers = parseLines(req.body.blogUsers).map(entry => entry.split('@')[0]);
    const sharedMailboxes = parseLines(req.body.sharedMailboxes);
    const badUser = blogUsers.find(user => !USERNAME.test(user));
    if (badUser) throw userError(400, `«${badUser}» no es un nombre de usuario (la parte del correo antes de @).`);
    const badMailbox = sharedMailboxes.find(address => !address.endsWith(`@${permissions.DOMAIN}`) || !/^[^@\s]+@[^@\s]+$/.test(address));
    if (badMailbox) throw userError(400, `«${badMailbox}» no es un buzón del dominio @${permissions.DOMAIN}.`);

    saveSettings({ blogUsers, sharedMailboxes });
    auditRequest(req, 'admin.settings_updated', { blogUsers: blogUsers.length, sharedMailboxes: sharedMailboxes.length });
    req.session.settingsSaved = true;
    res.redirect('/admin/ajustes');
  } catch (error) {
    next(error);
  }
});

module.exports = router;
