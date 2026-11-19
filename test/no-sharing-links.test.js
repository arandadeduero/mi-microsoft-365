/**
 * Sharing links (SharePoint/OneDrive "/:x:/", "/:f:/"...) give access to files,
 * and this repository is public: they live in .env only (USERS_EXCEL_URL,
 * WELCOME_GUIDE_URL), never in the code or its defaults.
 */

require('./helpers/env');

const fs = require('fs');
const path = require('path');
const test = require('ava').default;

const ROOT = path.join(__dirname, '..');
const SHARING_LINK = /sharepoint\.com\/:[a-z]:\//i;

function sourceFiles(dir) {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((entry) => {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) return sourceFiles(full);
    return /\.(js|ejs)$/.test(entry.name) ? [full] : [];
  });
}

test('no source file carries a SharePoint sharing link', t => {
  const files = ['config', 'middleware', 'routes', 'services', 'views', 'public/js'].flatMap(dir => sourceFiles(path.join(ROOT, dir)))
    .concat(path.join(ROOT, 'server.js'));
  t.true(files.length > 50);
  const offenders = files.filter(file => SHARING_LINK.test(fs.readFileSync(file, 'utf8'))).map(file => path.relative(ROOT, file));
  t.deepEqual(offenders, []);
});

test('the two links have no default', t => {
  const appConfig = require('../config/appConfig');
  t.is(appConfig.USERS_EXCEL_URL, '');
  t.is(appConfig.WELCOME_GUIDE_URL, '');
});
