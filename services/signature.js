/**
 * Pure helpers for /firma: which concejalía a person belongs to, and the
 * signature itself. The template is the one of arandadeduero/firmas-correo
 * (`buildSignatureHTML`): a table with inline styles, because that is what
 * mail clients render reliably. Change it there and here together.
 */
const CONCEJALIAS = require('../config/signatureConcejalias');

const HTML_ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
const escapeHtml = (text) => String(text).replace(/[&<>"']/g, char => HTML_ESCAPES[char]);
const fold = (text) => String(text || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();

const logoPath = (key) => (key === 'alcaldia' ? 'colores/ayuntamiento.png' : `blanco/${key}.png`);

// Mayors and councillors do not name a concejalía in their signature
const isCouncilPosition = (position) => /^(alcald|concejal)/i.test(position || '');

/**
 * The concejalía a profile points to, or null when it cannot be told.
 * The department is matched by its name (ignoring case, accents and a leading
 * "Concejalía de"); a partial match counts only when exactly one fits, so
 * "Cultura" never lands on one of its five services by chance.
 */
function findConcejaliaKey(department, position) {
  const wanted = fold(department).replace(/^concejalia (de |del )?/, '');
  if (wanted) {
    const entries = Object.entries(CONCEJALIAS);
    const exact = entries.find(([key, name]) => key === wanted || fold(name) === wanted);
    if (exact) return exact[0];
    const partial = entries.filter(([, name]) => fold(name).includes(wanted));
    if (partial.length === 1) return partial[0][0];
  }
  return /^alcald/i.test(position || '') ? 'alcaldia' : null;
}

/**
 * The signature as HTML.
 * @param {{name: string, position: string, email: string, phone?: string}} person
 * @param {string} key - A key of CONCEJALIAS
 * @param {string} imagesBase - Address the logos are loaded from, ending in "/"
 */
function buildSignatureHtml({ name, position, email, phone }, key, imagesBase) {
  const concejaliaName = `Concejalía de ${CONCEJALIAS[key]}`;
  const showConcejalia = !isCouncilPosition(position) && key !== 'alcaldia';

  const concejalia = showConcejalia
    ? `<div style="color:#000000;font-weight:600;font-size:13px;margin-bottom:8px;">${escapeHtml(concejaliaName)}</div>`
    : '';
  const phoneRow = phone
    ? `<div style="color:#000000;font-size:13px;">📱 ${escapeHtml(phone)}</div>`
    : '';
  const safeEmail = escapeHtml(email);

  return `
<div>--&nbsp;</div>
<table style="font-family:'Outfit',sans-serif;font-size:14px;color:#333;border-collapse:collapse;">
  <tr>
    <td style="padding:10px;padding-left:0;vertical-align:top;">
      <div style="line-height:1.6;">
        <div style="font-weight:bold;font-size:16px;color:#000000;margin-bottom:2px;">${escapeHtml(name)}</div>
        <div style="color:#000000;font-size:13px;margin-bottom:2px;">${escapeHtml(position)}</div>
        ${concejalia}
        <div style="color:#000000;font-size:13px;">
          📧 <a href="mailto:${safeEmail}" style="color:#000000;text-decoration:none;">${safeEmail}</a>
        </div>
        ${phoneRow}
        <div style="color:#666666;font-size:11px;margin-top:8px;">
          <a href="https://www.arandadeduero.es/aviso-legal/" style="color:#666666;text-decoration:none;" target="_blank">Aviso legal</a> · Si no es necesario y te preocupa el medio ambiente, no imprimas este correo electrónico.
        </div>
      </div>
    </td>
  </tr>
  <tr>
    <td style="padding:10px;padding-left:0;padding-top:15px;">
      <img src="${escapeHtml(imagesBase + logoPath(key))}" alt="${escapeHtml(concejaliaName)}" style="display:block;max-width:320px;height:auto;border-radius:8px;">
    </td>
  </tr>
</table>`.trim();
}

module.exports = { CONCEJALIAS, findConcejaliaKey, buildSignatureHtml };
