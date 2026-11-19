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

  // Welcome guide linked from admin password-reset credential emails. No default:
  // it is a sharing link, so it lives in .env only; without it the email leaves the line out
  WELCOME_GUIDE_URL: process.env.WELCOME_GUIDE_URL || '',

  // SharePoint sharing URL of the HR users Excel file (ACTIVOS / CONCEJALES), read
  // directly by /admin/usuarios/excel and checked once a day. No default: it is a
  // sharing link, so it lives in .env only; without it the admin uploads the file
  // and there is no daily check
  USERS_EXCEL_URL: process.env.USERS_EXCEL_URL || '',

  // Microsoft Form offered on /peticiones next to the support form. It is a link
  // to a form of the organization: .env only, not in the code. Empty: not shown
  INTERNAL_REQUESTS_FORM_URL: process.env.INTERNAL_REQUESTS_FORM_URL || '',

  // Address of the person at the top of the org chart. Without them in the
  // directory, the top is whoever has people below and no manager.
  ORG_CHART_ROOT: process.env.ORG_CHART_ROOT || 'alinaje@arandadeduero.es',

  // Where the logos of a copied email signature are loaded from (the
  // organization's github.com/arandadeduero/firmas-correo on GitHub Pages).
  // Recipients cannot reach this portal, so the signature needs a public address.
  SIGNATURE_IMAGES_URL: process.env.SIGNATURE_IMAGES_URL || 'https://arandadeduero.github.io/firmas-correo/images/',

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
