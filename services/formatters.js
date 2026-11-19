/**
 * Shared Intl formatters and collators.
 *
 * Creating Intl.DateTimeFormat / Intl.Collator once at module load avoids
 * re-allocating ICU objects inside view loops and array sorts, and keeps
 * Madrid time consistent even when the process runs in UTC (Docker/Alpine).
 */

const MADRID = 'Europe/Madrid';

const esCollator = new Intl.Collator('es');
const esBaseCollator = new Intl.Collator('es', { sensitivity: 'base' });

const shortDateFmt = new Intl.DateTimeFormat('es-ES', { timeZone: MADRID });
const longDateFmt = new Intl.DateTimeFormat('es-ES', { day: 'numeric', month: 'long', year: 'numeric', timeZone: MADRID });
const dateTimeFmt = new Intl.DateTimeFormat('es-ES', {
  day: '2-digit', month: '2-digit', year: 'numeric',
  hour: '2-digit', minute: '2-digit',
  timeZone: MADRID,
});
const dateTimeSecFmt = new Intl.DateTimeFormat('es-ES', {
  day: '2-digit', month: '2-digit', year: 'numeric',
  hour: '2-digit', minute: '2-digit', second: '2-digit',
  timeZone: MADRID,
});
const signInDateFmt = new Intl.DateTimeFormat('es-ES', {
  day: 'numeric', month: 'short', year: 'numeric',
  hour: '2-digit', minute: '2-digit',
  timeZone: MADRID,
});
const localCommissionFmt = new Intl.DateTimeFormat('es-ES', {
  timeZone: 'UTC',
  dateStyle: 'full',
  timeStyle: 'short',
});
const utcWeekdayFmt = new Intl.DateTimeFormat('es-ES', {
  timeZone: 'UTC',
  weekday: 'long',
  day: 'numeric',
  month: 'long',
});
const madridIsoDateFmt = new Intl.DateTimeFormat('sv-SE', { timeZone: MADRID });
const madridIsoDateTimeFmt = new Intl.DateTimeFormat('sv-SE', {
  timeZone: MADRID,
  dateStyle: 'short',
  timeStyle: 'short',
});

const compareEs = (a, b) => esCollator.compare(a || '', b || '');
const compareEsBase = (a, b) => esBaseCollator.compare(a || '', b || '');

const formatShortDate = (value) => shortDateFmt.format(new Date(value));
const formatLongDate = (value) => longDateFmt.format(new Date(value));
const formatDateTime = (value) => dateTimeFmt.format(new Date(value));
const formatDateTimeSeconds = (value) => dateTimeSecFmt.format(new Date(value));
const formatSignInDate = (value) => signInDateFmt.format(new Date(value));
const formatLocalCommissionDate = (localStart) => localCommissionFmt.format(new Date(`${localStart}:00Z`));
const formatUtcWeekday = (isoDate) => utcWeekdayFmt.format(new Date(`${isoDate}T00:00:00Z`));
const madridIsoDate = (date = new Date()) => madridIsoDateFmt.format(date);
const madridIsoDateTime = (date = new Date()) => madridIsoDateTimeFmt.format(date);

const DIRECTORY_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

module.exports = {
  DIRECTORY_ID,
  compareEs,
  compareEsBase,
  formatShortDate,
  formatLongDate,
  formatDateTime,
  formatDateTimeSeconds,
  formatSignInDate,
  formatLocalCommissionDate,
  formatUtcWeekday,
  madridIsoDate,
  madridIsoDateTime,
};
