/**
 * Room Routes
 *
 * - GET /salas - Meeting rooms that are free at a given time
 *
 * The app only looks: booking is done in Outlook, which opens with the room
 * and the time already filled in.
 */

const express = require('express');
const router = express.Router();
const { REGION } = require('../config/appConfig');
const { isAuthenticated } = require('../middleware/auth');
const { getRooms, getSchedules } = require('../services/graphService');
const { compareEs, madridIsoDateTime } = require('../services/formatters');

const DATE = /^\d{4}-\d{2}-\d{2}$/;
const TIME = /^([01]\d|2[0-3]):[0-5]\d$/;
const AVAILABILITY_RANK = new Map([[true, 0], [null, 1], [false, 2]]);

/**
 * Today's date and the next full hour in Madrid, as the form expects them
 */
function defaultSlot(now = new Date()) {
  const [date, time] = madridIsoDateTime(now).split(' ');
  const hour = Math.min(Number(time.slice(0, 2)) + 1, 22);
  const pad = (value) => String(value).padStart(2, '0');
  return { date, start: `${pad(hour)}:00`, end: `${pad(hour + 1)}:00` };
}

/**
 * Whether a room is free for the whole period: true, false, or null when its
 * calendar could not be read
 */
function isFree(schedule) {
  if (!schedule || schedule.error || typeof schedule.availabilityView !== 'string') {
    return null;
  }
  // One digit per quarter of an hour; 0 is free
  return /^0*$/.test(schedule.availabilityView);
}

/**
 * Link that opens a new meeting in Outlook with the room and time filled in
 */
function bookingUrl(room, slot) {
  const query = new URLSearchParams({
    startdt: `${slot.date}T${slot.start}:00`,
    enddt: `${slot.date}T${slot.end}:00`,
    location: room.displayName || '',
    to: room.emailAddress,
  });
  return `https://outlook.office.com/calendar/deeplink/compose?${query}`;
}

/**
 * Where a room is, from what the directory says about it
 */
function whereIs(room) {
  const floor = room.floorLabel || (room.floorNumber != null ? `Planta ${room.floorNumber}` : null);
  return [room.building, floor, room.address?.street].filter(Boolean).join(' · ');
}

/**
 * GET /salas
 * The meeting rooms, with where they are and how many people they take;
 * after a search, also whether each one is free in the period asked for
 */
router.get('/salas', isAuthenticated, async (req, res, next) => {
  try {
    const accessToken = req.session.accessToken;
    const { date, start, end } = req.query;
    const searched = Boolean(date || start || end);
    const valid = DATE.test(date) && TIME.test(start) && TIME.test(end) && start < end;
    const slot = valid ? { date, start, end } : defaultSlot();

    const places = await getRooms(accessToken);
    let rooms = places && places
      .filter(place => place.emailAddress)
      .map(room => ({
        name: room.displayName || room.emailAddress,
        email: room.emailAddress,
        capacity: room.capacity || null,
        where: whereIs(room),
        free: null,
        bookingUrl: valid ? bookingUrl(room, slot) : null,
      }));
    let availabilityKnown = true;

    // Only a search asks the calendars; the list itself needs no date
    if (valid && rooms && rooms.length > 0) {
      const byAddress = new Map();
      try {
        const schedules = await getSchedules(accessToken, rooms.map(room => room.email), `${date}T${start}:00`, `${date}T${end}:00`, REGION.timeZone);
        schedules.forEach(schedule => byAddress.set((schedule.scheduleId || '').toLowerCase(), schedule));
      } catch (error) {
        if (error.graphStatus === 401) throw error;
        availabilityKnown = false;
      }

      rooms = rooms
        .map(room => ({ ...room, free: isFree(byAddress.get(room.email.toLowerCase())) }))
        // Free first, then unknown, then busy
        .sort((a, b) => AVAILABILITY_RANK.get(a.free) - AVAILABILITY_RANK.get(b.free) || compareEs(a.name, b.name));
    } else if (rooms) {
      rooms.sort((a, b) => compareEs(a.name, b.name));
    }

    res.render('rooms', {
      title: 'Salas',
      currentPage: 'rooms',
      slot,
      searched,
      valid,
      roomsReadable: rooms !== null,
      rooms: rooms || [],
      availabilityKnown,
    });
  } catch (error) {
    next(error);
  }
});

module.exports = router;
