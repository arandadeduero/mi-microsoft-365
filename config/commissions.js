/**
 * Commission meetings, as /admin/comisiones calls them
 */
module.exports = {
  // Calendar of the signed-in administrator the meetings go into, found by name
  CALENDAR_NAME: 'Comisiones y plenos',
  // Distribution list whose members are invited, one by one, unless the meeting names its own
  INVITEES: 'dist-concejales2327@arandadeduero.es',
  // Smallest room a commission fits in
  MIN_ROOM_CAPACITY: 5,
  // How long a commission lasts
  DURATION_MINUTES: 60,
  // The subject is "Com <name>", as the meetings already in the calendar
  COMMISSIONS: [
    'Asuntos Económicos',
    'Cultura',
    'Deportes',
    'Juventud',
    'Medio Ambiente',
    'Personal',
    'Promoción',
    'Seguridad Ciudadana',
    'Servicios Sociales',
    'Turismo',
    'Urbanismo',
  ],
  // Called from the same page; their subject is the name alone, and each has its own minutes.
  // `roomEndsWith` picks their usual room in the form (it can still be changed).
  OTHER_MEETINGS: {
    Pleno: { minutes: 180, roomEndsWith: 'pleno' },
    // Only the government team sits on the Junta de Gobierno Local
    JGL: { minutes: 30, invitees: 'dist-eqgobierno2327@arandadeduero.es', roomEndsWith: 'Sala de juntas' },
  },
};
