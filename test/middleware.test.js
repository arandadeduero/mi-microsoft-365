require('./helpers/env');

const test = require('ava').default;
const { stubAxios } = require('./helpers/stubs');

const http = stubAxios();
const { isAuthenticated, redirectIfAuthenticated, checkPermission, checkAdminStatus } = require('../middleware/auth');
const permissions = require('../config/permissions');

test('isAuthenticated redirects to /login if not authenticated', t => {
    const req = { session: {} };
    const res = { redirect: (url) => t.is(url, '/login') };
    const next = () => t.fail('Next should not be called');
    
    isAuthenticated(req, res, next);
});

test('isAuthenticated calls next() if authenticated', t => {
    const req = { session: { accessToken: 'token' } };
    const res = { redirect: () => t.fail('Redirect should not be called') };
    const next = () => t.pass();
    
    isAuthenticated(req, res, next);
});

test('redirectIfAuthenticated redirects to / if already authenticated', t => {
    const req = { session: { accessToken: 'token' } };
    const res = { redirect: (url) => t.is(url, '/') };
    const next = () => t.fail('Next should not be called');
    
    redirectIfAuthenticated(req, res, next);
});

test('redirectIfAuthenticated calls next() if not authenticated', t => {
    const req = { session: {} };
    const res = { redirect: () => t.fail('Redirect should not be called') };
    const next = () => t.pass();
    
    redirectIfAuthenticated(req, res, next);
});

function fakeDeniedResponse(t) {
    return {
        status(code) {
            t.is(code, 403);
            return this;
        },
        send(body) {
            t.is(body, 'Access Denied');
        },
    };
}

function withPermissions(t, overrides) {
    const previous = { ...permissions };
    Object.assign(permissions, overrides);
    t.teardown(() => {
        Object.keys(permissions).forEach(key => delete permissions[key]);
        Object.assign(permissions, previous);
    });
}

test.serial('checkPermission denies requests without a signed-in account', t => {
    t.plan(2);
    const req = { session: {}, path: '/blog', method: 'GET' };

    checkPermission('/blog')(req, fakeDeniedResponse(t), () => t.fail('Next should not be called'));
});

test.serial('checkPermission denies users from another domain', t => {
    t.plan(2);
    withPermissions(t, { DOMAIN: 'arandadeduero.es', '/blog': ['alice'] });
    const req = { session: { account: { username: 'alice@example.com' } }, path: '/blog', method: 'GET' };

    checkPermission('/blog')(req, fakeDeniedResponse(t), () => t.fail('Next should not be called'));
});

test.serial('checkPermission denies users missing from the route list', t => {
    t.plan(2);
    withPermissions(t, { DOMAIN: 'arandadeduero.es', '/blog': ['alice'] });
    const req = { session: { account: { username: 'bob@arandadeduero.es' } }, path: '/blog', method: 'POST' };

    checkPermission('/blog')(req, fakeDeniedResponse(t), () => t.fail('Next should not be called'));
});

test.serial('checkPermission allows users listed for the route', t => {
    t.plan(1);
    withPermissions(t, { DOMAIN: 'arandadeduero.es', '/blog': ['alice'] });
    const req = { session: { account: { username: 'alice@arandadeduero.es' } }, path: '/blog', method: 'GET' };

    checkPermission('/blog')(req, { status: () => t.fail('Status should not be set') }, () => t.pass());
});

test.serial('checkPermission lets nobody in on a key that has no list', t => {
    t.plan(2);
    withPermissions(t, { DOMAIN: 'arandadeduero.es' });
    const req = { session: { account: { username: 'bob@arandadeduero.es' } }, path: '/unlisted', method: 'GET' };

    checkPermission('/unlisted')(req, fakeDeniedResponse(t), () => t.fail('Next should not be called'));
});

test.serial('checkPermission uses its key however the request path is spelled', t => {
    // Express sends /Blog and /blog/ to the handler of /blog
    t.plan(3);
    withPermissions(t, { DOMAIN: 'arandadeduero.es', '/blog': ['alice'] });
    const bob = { session: { account: { username: 'bob@arandadeduero.es' } }, path: '/Blog', method: 'POST' };
    const alice = { session: { account: { username: 'alice@arandadeduero.es' } }, path: '/blog/', method: 'GET' };

    checkPermission('/blog')(bob, fakeDeniedResponse(t), () => t.fail('Next should not be called'));
    checkPermission('/blog')(alice, { status: () => t.fail('Status should not be set') }, () => t.pass());
});

test.serial('checkAdminStatus marks anonymous requests as non-admin without calling Graph', async t => {
    http.reset();
    const res = { locals: {} };

    await checkAdminStatus({ session: {} }, res, () => {});

    t.deepEqual(res.locals, { isAdmin: false, adminRoles: [] });
    t.is(http.calls.length, 0);
});

test.serial('checkAdminStatus exposes the directory roles of the signed-in user', async t => {
    http.reset();
    http.on('get', '/me/memberOf/microsoft.graph.directoryRole', { value: [{ displayName: 'Global Administrator' }] });
    const res = { locals: {} };

    await checkAdminStatus({ session: { accessToken: 'token' } }, res, () => t.pass());

    t.deepEqual(res.locals, { isAdmin: true, adminRoles: ['Global Administrator'] });
});

test.serial('checkAdminStatus treats users without roles as non-admin', async t => {
    http.reset();
    http.on('get', '/me/memberOf/microsoft.graph.directoryRole', { value: [] });
    const res = { locals: {} };

    await checkAdminStatus({ session: { accessToken: 'token' } }, res, () => t.pass());

    t.deepEqual(res.locals, { isAdmin: false, adminRoles: [] });
});

test.serial('checkAdminStatus caches roles in session on normal pages and verifies live on /admin/*', async t => {
    http.reset();
    http.on('get', '/me/memberOf/microsoft.graph.directoryRole', { value: [{ displayName: 'Global Administrator' }] });
    const session = { accessToken: 'token' };

    const res1 = { locals: {} };
    await checkAdminStatus({ path: '/', session }, res1, () => {});
    t.deepEqual(res1.locals, { isAdmin: true, adminRoles: ['Global Administrator'] });
    t.is(http.calls.length, 1);

    // Second regular page reuses session cache without calling Graph
    const res2 = { locals: {} };
    await checkAdminStatus({ path: '/directorio', session }, res2, () => {});
    t.deepEqual(res2.locals, { isAdmin: true, adminRoles: ['Global Administrator'] });
    t.is(http.calls.length, 1);

    // Admin route always verifies live against Graph
    http.on('get', '/me/memberOf/microsoft.graph.directoryRole', { value: [] });
    const res3 = { locals: {} };
    await checkAdminStatus({ path: '/admin/audit', session }, res3, () => {});
    t.deepEqual(res3.locals, { isAdmin: false, adminRoles: [] });
    t.is(http.calls.length, 2);
});

test.serial('checkAdminStatus does not cache failed Graph role lookups', async t => {
    http.reset();
    http.fail('get', '/me/memberOf/microsoft.graph.directoryRole', 500);
    const session = { accessToken: 'token' };

    const res1 = { locals: {} };
    await checkAdminStatus({ path: '/', session }, res1, () => {});
    t.deepEqual(res1.locals, { isAdmin: false, adminRoles: [] });
    t.is(session.adminStatus, undefined);

    http.on('get', '/me/memberOf/microsoft.graph.directoryRole', { value: [{ displayName: 'Global Administrator' }] });
    const res2 = { locals: {} };
    await checkAdminStatus({ path: '/', session }, res2, () => {});
    t.deepEqual(res2.locals, { isAdmin: true, adminRoles: ['Global Administrator'] });
});
