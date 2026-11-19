/**
 * Application Settings
 *
 * Values shared by several routes that are not specific to MSAL, Graph or
 * WordPress.
 */

const { redirectUri } = require('./authConfig');

// Optional settings of this installation, in the git-ignored config/local.js
let local = {};
try {
  local = require('./local');
} catch (error) {
  if (error.code !== 'MODULE_NOT_FOUND') {
    console.error('Error loading config/local.js:', error);
  }
}

const PROFILE_FIELDS = ['jobTitle', 'mobilePhone', 'businessPhones', 'officeLocation'];

module.exports = {
  // What the portal and the organization are called on every page
  APP_NAME: 'Mi Microsoft 365',
  ORG_NAME: 'Ayuntamiento de Aranda de Duero',

  // The organization's intranet, linked from the home page and the help page
  INTRANET_URL: process.env.INTRANET_URL || 'https://aytoarandaduero.sharepoint.com',

  // Welcome guide linked from admin password-reset credential emails
  WELCOME_GUIDE_URL: process.env.WELCOME_GUIDE_URL || 'https://aytoarandaduero-my.sharepoint.com/:f:/g/personal/glopez_arandadeduero_es/IgDxAF3YuiHuRILrP8vCkGRnAYrtXe0_jzxtD_l9ZON1CGg?e=86PJVY',

  // SharePoint sharing URL of the HR users Excel file (ACTIVOS / CONCEJALES),
  // read directly by /admin/usuarios/excel so the admin does not have to upload it
  USERS_EXCEL_URL: process.env.USERS_EXCEL_URL !== undefined
    ? process.env.USERS_EXCEL_URL
    : 'https://aytoarandaduero.sharepoint.com/:x:/s/gr-personal/IQC5DRK-adTlSJkHPU5LEc43AcZZU0LtY1sVYlcG40AOQqo?e=Z1Hsoc',

  // Address of the person at the top of the org chart. Without them in the
  // directory, the top is whoever has people below and no manager.
  ORG_CHART_ROOT: process.env.ORG_CHART_ROOT || 'alinaje@arandadeduero.es',

  // The organization's own signature generator (github.com/arandadeduero/firmas-correo),
  // shown inside the /firma page. Empty hides the page's embed.
  SIGNATURE_GENERATOR_URL: process.env.SIGNATURE_GENERATOR_URL || 'https://arandadeduero.github.io/firmas-correo/',

  // Recipient of the requests that have nobody else to go to
  ADMIN_EMAIL: process.env.ADMIN_EMAIL || 'glopez@arandadeduero.es',

  // Public address of the app. Links sent by email are built from
  // configuration and not from the request's Host header, which whoever
  // sends the request controls.
  APP_ORIGIN: new URL(redirectUri).origin,

  // Profile fields users may change themselves. Tenants differ in what they
  // let a user edit; the rest are shown read-only.
  PROFILE_FIELDS,
  PROFILE_EDITABLE_FIELDS: process.env.PROFILE_EDITABLE_FIELDS
    ? process.env.PROFILE_EDITABLE_FIELDS.split(',').map(field => field.trim()).filter(field => PROFILE_FIELDS.includes(field))
    : [...PROFILE_FIELDS],

  // Addresses of the shared mailboxes of the organization. Graph cannot list
  // them, so the app checks these one by one for each user.
  SHARED_MAILBOXES: Array.isArray(local.sharedMailboxes) ? [...local.sharedMailboxes] : [],

  // Time zone (as Exchange names it) and language every account is kept on
  REGION: {
    timeZone: 'Romance Standard Time', // Madrid
    locale: 'es-ES',
  },
};
