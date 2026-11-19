/**
 * Admin Routes
 *
 * Pages for administrators only (users holding an Entra ID directory role):
 * - /admin/auditoria (and /admin/auditoria.csv) - The audit log
 * - /admin/licencias - Licenses bought and left
 * - /admin/avisos - Microsoft's message center
 * - /admin/solicitudes - Group requests nobody has decided on
 * - /admin/salas - A room's calendar for a week: who booked it and who goes
 * - /admin/comisiones - Calling a commission: the meeting, the room and the councillors
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
const { getSubscribedSkus, getRooms, getRoomEvents, getMyCalendars, getListMembers, createCalendarEvent, listAdminUsers } = require('../services/graphService');
const commissions = require('../config/commissions');
const { describeLicenseStock } = require('../services/accountReview');
const { getServiceMessages } = require('../services/messageCenterService');
const { listPending } = require('../services/requestStore');
const { currentSettings, saveSettings } = require('../services/settingsStore');
const { compareEs, madridIsoDate } = require('../services/formatters');

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
 * GET /admin/auditoria
 * The audit log, newest first: the last 7 days, or one day when given,
 * optionally only one kind of event or one person
 */
router.get('/admin/auditoria', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const { filters, all, matching } = await filteredAudit(req.query);

    res.render('audit', {
      title: 'Auditoría',
      currentPage: 'admin',
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
 * GET /admin/audit and /admin/audit.csv
 * The audit log used to be here; old links keep working, with their filters
 */
router.get(['/admin/audit', '/admin/audit.csv'], (req, res) => {
  const search = req.url.includes('?') ? req.url.slice(req.url.indexOf('?')) : '';
  res.redirect(301, `/admin/auditoria${req.path.toLowerCase().endsWith('.csv') ? '.csv' : ''}${search}`);
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
 * GET /admin/auditoria.csv
 * The same entries as the page, for a spreadsheet
 */
router.get('/admin/auditoria.csv', isAuthenticated, requireAdmin, async (req, res, next) => {
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
 * The accounts page now lives under Usuarios; the old address keeps working
 */
router.get('/admin/cuentas', (req, res) => res.redirect(301, '/admin/usuarios/atencion'));

/**
 * GET /admin/licencias
 */
router.get('/admin/licencias', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const skus = await getSubscribedSkus(req.session.accessToken);
    res.render('admin-licenses', { title: 'Licencias', currentPage: 'admin', licenses: skus && describeLicenseStock(skus, SKU_NAMES) });
  } catch (error) {
    next(error);
  }
});

/**
 * GET /admin/licencias/:sku
 * Every account holding a given license SKU, with a direct link to edit it
 */
router.get('/admin/licencias/:sku', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const token = req.session.accessToken;
    const skuParam = String(req.params.sku || '').trim();
    const skus = await getSubscribedSkus(token);
    if (!skus) {
      return next(userError(503, 'Ahora mismo no se pueden leer las licencias.'));
    }

    const stock = describeLicenseStock(skus, SKU_NAMES);
    const matchedSku = skus.find(s => s.skuId === skuParam || s.skuPartNumber === skuParam);
    const license = stock.find(s => (s.skuId && s.skuId === skuParam) || s.code === skuParam)
      || (matchedSku ? {
        skuId: matchedSku.skuId,
        code: matchedSku.skuPartNumber,
        name: SKU_NAMES[matchedSku.skuPartNumber] || matchedSku.skuPartNumber,
        total: matchedSku.prepaidUnits?.enabled || 0,
        used: matchedSku.consumedUnits || 0,
        left: (matchedSku.prepaidUnits?.enabled || 0) - (matchedSku.consumedUnits || 0),
      } : null);

    if (!matchedSku || !license) {
      return next(userError(404, 'Licencia no encontrada'));
    }

    const { users: allUsers } = await listAdminUsers(token, { onlyDisabled: null });
    const holders = allUsers
      .filter(u => (u.assignedLicenses || []).some(l => l.skuId === matchedSku.skuId))
      .sort((a, b) => compareEs(a.displayName || a.userPrincipalName || '', b.displayName || b.userPrincipalName || ''));

    res.render('admin-license-users', {
      title: `Licencia · ${license.name}`,
      currentPage: 'admin',
      license,
      users: holders,
    });
  } catch (error) {
    next(error);
  }
});

/**
 * GET /admin/avisos
 */
router.get('/admin/avisos', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    res.render('admin-messages', { title: 'Avisos de Microsoft', currentPage: 'admin', messages: await getServiceMessages() });
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
    res.render('admin-requests', { title: 'Solicitudes pendientes', currentPage: 'admin', pending });
  } catch (error) {
    next(error);
  }
});

const DATE = /^\d{4}-\d{2}-\d{2}$/;
const addDays = (date, days) => new Date(Date.parse(`${date}T00:00:00Z`) + days * 86400000).toISOString().slice(0, 10);

/**
 * The rooms by name, or null when they cannot be read
 */
async function listRooms(accessToken) {
  const places = await getRooms(accessToken);
  return places && places
    .filter(place => place.emailAddress)
    .map(place => ({ name: place.displayName || place.emailAddress, email: place.emailAddress, capacity: place.capacity || null }))
    .sort((a, b) => compareEs(a.name, b.name));
}

/**
 * GET /admin/salas
 * One room's calendar for a week, read with the app's token. The room comes
 * from the list of rooms and nowhere else: that token opens every mailbox.
 */
router.get('/admin/salas', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const wanted = String(req.query.room || '').toLowerCase();
    const [rooms, appToken] = await Promise.all([
      listRooms(req.session.accessToken),
      wanted ? getAppToken() : null,
    ]);
    const room = (rooms || []).find(candidate => candidate.email.toLowerCase() === wanted) || null;

    const today = madridIsoDate();
    const asked = DATE.test(req.query.week) && !Number.isNaN(Date.parse(req.query.week)) ? req.query.week : today;
    const monday = addDays(asked, -((new Date(`${asked}T00:00:00Z`).getUTCDay() + 6) % 7));

    // The period is asked in UTC, a day wider on each side; the days below keep what is theirs
    const events = room
      ? await getRoomEvents(appToken, room.email, `${addDays(monday, -1)}T00:00:00`, `${addDays(monday, 8)}T00:00:00`, REGION.timeZone)
      : null;

    const roomEmail = room ? room.email.toLowerCase() : '';
    const activeEvents = events && events.filter(event => !event.isCancelled);
    const days = activeEvents && [0, 1, 2, 3, 4, 5, 6].map(offset => {
      const date = addDays(monday, offset);
      const next = addDays(date, 1);
      return {
        date,
        events: activeEvents
          // Times are local and to the minute, so text compares as time; multi-day events land on each of their days
          .filter(event => event.start.dateTime.slice(0, 16) < `${next}T00:00` && event.end.dateTime.slice(0, 16) > `${date}T00:00`)
          .map(event => ({
            from: event.isAllDay ? null : (event.start.dateTime.startsWith(date) ? event.start.dateTime.slice(11, 16) : '00:00'),
            to: event.isAllDay ? null : (event.end.dateTime.startsWith(date) ? event.end.dateTime.slice(11, 16) : '24:00'),
            subject: (event.subject || '').trim(),
            organizer: event.organizer?.emailAddress || {},
            // The room is an attendee of its own bookings
            people: (event.attendees || []).filter(person => (person.emailAddress?.address || '').toLowerCase() !== roomEmail),
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
      currentPage: 'admin',
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

const LOCAL_TIME = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/;

// The list a kind of meeting is called from
const inviteesFor = (kind) => commissions.OTHER_MEETINGS[kind]?.invitees || commissions.INVITEES;

/**
 * What a meeting needs before it can be called: the calendar, the rooms that
 * are big enough and who is in each list. `problem` says what stops every
 * meeting; a list that could not be read has `people: null`, and nothing is
 * assumed about it, so nobody is called to a meeting with no one invited.
 */
async function commissionSetup(accessToken, listAddresses = null) {
  const others = Object.keys(commissions.OTHER_MEETINGS);
  const addresses = listAddresses || [...new Set([commissions.INVITEES, ...others.map(inviteesFor)])];

  const [calendars, allRooms, memberLists] = await Promise.all([
    getMyCalendars(accessToken),
    listRooms(accessToken),
    Promise.all(addresses.map(address => getListMembers(accessToken, address))),
  ]);

  const calendar = (calendars || []).find(candidate => (candidate.name || '').toLowerCase() === commissions.CALENDAR_NAME.toLowerCase()) || null;
  const rooms = (allRooms || []).filter(room => room.capacity >= commissions.MIN_ROOM_CAPACITY);

  const lists = addresses.map((address, index) => {
    const members = memberLists[index];
    return {
      address,
      kinds: [...(address === commissions.INVITEES ? ['Comisiones'] : []), ...others.filter(kind => inviteesFor(kind) === address)],
      // One invitation per address, whatever nested lists repeat
      people: members && [...new Map(members.filter(person => person.mail).map(person => [person.mail.toLowerCase(), { name: person.displayName || person.mail, email: person.mail }])).values()]
        .sort((a, b) => compareEs(a.name, b.name)),
    };
  });

  let problem = null;
  if (!calendars) problem = 'Ahora mismo no se pueden leer tus calendarios.';
  else if (!calendar) problem = `No tienes ningún calendario llamado «${commissions.CALENDAR_NAME}». Las comisiones las convoca quien lo tiene en su buzón.`;
  else if (calendar.canEdit === false) problem = `No puedes escribir en el calendario «${commissions.CALENDAR_NAME}».`;
  else if (rooms.length === 0) problem = `No se ha encontrado ninguna sala para ${commissions.MIN_ROOM_CAPACITY} personas o más.`;

  return { calendar, rooms, lists, problem };
}

/**
 * GET /admin/comisiones
 */
router.get('/admin/comisiones', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const created = req.session.commissionCreated || null;
    delete req.session.commissionCreated;
    res.render('admin-commissions', {
      title: 'Convocar una comisión',
      currentPage: 'admin',
      config: commissions,
      created,
      ...await commissionSetup(req.session.accessToken),
    });
  } catch (error) {
    next(error);
  }
});

/**
 * POST /admin/comisiones
 * Put the meeting in the calendar; Exchange sends the invitations. Who is
 * invited and which rooms may be chosen are read again here, never taken
 * from the form.
 */
router.post('/admin/comisiones', isAuthenticated, requireAdmin, async (req, res, next) => {
  try {
    const commission = [...commissions.COMMISSIONS, ...Object.keys(commissions.OTHER_MEETINGS)].find(name => name === req.body.commission);
    const { calendar, rooms, lists, problem } = await commissionSetup(
      req.session.accessToken,
      commission ? [inviteesFor(commission)] : []
    );
    if (problem) throw userError(409, problem);

    const subject = commissions.COMMISSIONS.includes(commission) ? `Com ${commission}` : commission;
    if (!commission) throw userError(400, 'Elige una comisión de la lista.');
    const list = lists.find(candidate => candidate.address === inviteesFor(commission));
    if (!list.people) throw userError(409, `No se ha podido leer quién está en ${list.address}, así que no se sabe a quién convocar.`);
    if (list.people.length === 0) throw userError(409, `La lista ${list.address} no tiene a nadie con correo.`);
    const invitees = list.people;
    const room = rooms.find(candidate => candidate.email === req.body.room);
    if (!room) throw userError(400, 'Elige una sala de la lista.');
    const start = String(req.body.start || '');
    if (!LOCAL_TIME.test(start) || Number.isNaN(Date.parse(`${start}:00Z`))) throw userError(400, 'Indica el día y la hora de inicio.');
    // Clock arithmetic on the local time, done in UTC so no offset gets in
    const end = new Date(Date.parse(`${start}:00Z`) + (commissions.OTHER_MEETINGS[commission]?.minutes || commissions.DURATION_MINUTES) * 60000).toISOString().slice(0, 19);

    await createCalendarEvent(req.session.accessToken, calendar.id, {
      subject,
      start: { dateTime: `${start}:00`, timeZone: REGION.timeZone },
      end: { dateTime: end, timeZone: REGION.timeZone },
      location: { displayName: room.name, locationEmailAddress: room.email },
      attendees: [
        ...invitees.map(person => ({ emailAddress: { address: person.email, name: person.name }, type: 'required' })),
        { emailAddress: { address: room.email, name: room.name }, type: 'resource' },
      ],
    });

    auditRequest(req, 'admin.commission_called', { commission, start, room: room.email, invited: invitees.length });
    req.session.commissionCreated = { subject, start, room: room.name, invited: invitees.length };
    res.redirect('/admin/comisiones');
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
    res.render('admin-settings', { title: 'Ajustes', currentPage: 'admin', settings: currentSettings(), domain: permissions.DOMAIN, saved });
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
