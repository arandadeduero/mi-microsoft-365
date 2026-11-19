/**
 * Read what a person typed into a form made from a config entry (config/teamRequests.js,
 * config/internalRequests.js): every field has a label, may be required, and a select
 * only takes one of its options.
 */
const { userError } = require('../middleware/errorHandler');

const MAX_FIELD = 2000;

/**
 * @param {{fields: Array}} request - The form's definition
 * @param {Object} body - What was posted
 * @returns {Array<[string, string]>} [label, value] for every field, in order
 * @throws {Error} A 400 saying what is wrong
 */
function readRequestValues(request, body) {
  return request.fields.map((field) => {
    const value = String(body[field.name] || '').trim();
    if (field.required && !value) {
      throw userError(400, `Falta «${field.label}».`);
    }
    if (value.length > MAX_FIELD) {
      throw userError(400, `«${field.label}» es demasiado largo.`);
    }
    if (field.type === 'select' && value && !field.options.includes(value)) {
      throw userError(400, `«${field.label}» no es una opción válida.`);
    }
    return [field.label, value];
  });
}

module.exports = { readRequestValues };
