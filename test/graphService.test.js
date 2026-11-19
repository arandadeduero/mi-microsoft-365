/**
 * Microsoft Graph Service Tests
 * Unit tests for services/graphService.js with axios stubbed out
 */

require('./helpers/env');

const test = require('ava').default;
const { stubAxios } = require('./helpers/stubs');

const http = stubAxios();
const graph = require('../services/graphService');

test.beforeEach(() => {
  http.reset();
  graph.resetGraphCaches();
});

test.serial('getUserProfile sends the bearer token and returns the profile', async t => {
  http.on('get', '/v1.0/me', { displayName: 'Alice' });

  const profile = await graph.getUserProfile('token-123');

  t.deepEqual(profile, { displayName: 'Alice' });
  t.is(http.calls[0].config.headers.Authorization, 'Bearer token-123');
});

test.serial('getUserProfile throws a generic error that keeps the Graph status', async t => {
  http.fail('get', '/v1.0/me', 401);

  const error = await t.throwsAsync(() => graph.getUserProfile('token'), { message: 'Failed to fetch user profile' });
  t.is(error.graphStatus, 401);
});

test.serial('getUserPhoto returns a buffer when the user has a photo', async t => {
  http.on('get', '/me/photo/$value', Buffer.from('jpeg-bytes'));

  const photo = await graph.getUserPhoto('token');

  t.true(Buffer.isBuffer(photo));
  t.is(photo.toString(), 'jpeg-bytes');
  t.is(http.calls[0].config.responseType, 'arraybuffer');
});

test.serial('getUserPhoto returns null when there is no photo or Graph fails', async t => {
  t.is(await graph.getUserPhoto('token'), null);

  http.fail('get', '/me/photo/$value', 500);
  t.is(await graph.getUserPhoto('token'), null);
});

test.serial('getUserPhotoById fetches another user photo and returns null when missing or unreadable', async t => {
  t.is(await graph.getUserPhotoById('token', ''), null);
  t.is(http.calls.length, 0);

  http.on('get', '/users/u1/photo/$value', Buffer.from('png-bytes'));
  const photo = await graph.getUserPhotoById('token', 'u1');
  t.true(Buffer.isBuffer(photo));
  t.is(photo.toString(), 'png-bytes');
  t.is(http.calls[0].config.responseType, 'arraybuffer');

  t.is(await graph.getUserPhotoById('token', 'missing'), null);

  http.fail('get', '/users/u2/photo/$value', 500);
  t.is(await graph.getUserPhotoById('token', 'u2'), null);
});

test.serial('getUserManager returns null when no manager is assigned', async t => {
  t.is(await graph.getUserManager('token'), null);

  http.on('get', '/me/manager', { id: 'm1', displayName: 'Boss' });
  t.deepEqual(await graph.getUserManager('token'), { id: 'm1', displayName: 'Boss' });
});

test.serial('getUserGroups returns an empty list when Graph fails', async t => {
  http.fail('get', '/me/memberOf', 403);

  t.deepEqual(await graph.getUserGroups('token'), []);
});

test.serial('getAllUsers requests a sorted page of users and returns [] on failure', async t => {
  http.on('get', '/v1.0/users', { value: [{ id: 'u1' }] });

  t.deepEqual(await graph.getAllUsers('token'), [{ id: 'u1' }]);
  t.is(http.calls[0].config.params.$orderby, 'displayName');
  t.is(http.calls[0].config.params.$top, 999);

  http.fail('get', '/v1.0/users', 500);
  t.deepEqual(await graph.getAllUsers('token'), []);
});

test.serial('getLicensedUsers keeps only active licensed members and strips license metadata', async t => {
  http.on('get', '/v1.0/users', {
    value: [
      { id: 'u1', displayName: 'Ana', accountEnabled: true, userType: 'Member', assignedLicenses: [{ skuId: 'sku-1' }] },
      { id: 'u2', displayName: 'Sala', accountEnabled: false, userType: 'Member', assignedLicenses: [] },
      { id: 'u3', displayName: 'Buzón', accountEnabled: true, userType: 'Member', assignedLicenses: [] },
      { id: 'u4', displayName: 'Deshabilitado', accountEnabled: false, userType: 'Member', assignedLicenses: [{ skuId: 'sku-1' }] },
      { id: 'u5', displayName: 'Invitado', accountEnabled: true, userType: 'Guest', assignedLicenses: [{ skuId: 'sku-1' }] },
    ],
  });

  t.deepEqual(await graph.getLicensedUsers('token'), [{ id: 'u1', displayName: 'Ana' }]);
  t.true(http.calls[0].config.params.$select.includes('assignedLicenses'));
});

test.serial('getAllUsers follows @odata.nextLink until the last page', async t => {
  http.on('get', '/v1.0/users', {
    value: [{ id: 'u1' }],
    '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=page2',
  });
  http.on('get', '$skiptoken=page2', {
    value: [{ id: 'u2' }],
    '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=page3',
  });
  http.on('get', '$skiptoken=page3', { value: [{ id: 'u3' }] });

  t.deepEqual(await graph.getAllUsers('token'), [{ id: 'u1' }, { id: 'u2' }, { id: 'u3' }]);
  t.is(http.calls.length, 3);
  // Later pages are requested with the token and timeout, without repeating the query
  t.is(http.calls[2].config.headers.Authorization, 'Bearer token');
  t.is(http.calls[2].config.params, undefined);
  t.true(http.calls[2].config.timeout > 0);
});

test.serial('getUserGroups and getAllTenantGroups read every page', async t => {
  http.on('get', '/me/memberOf', { value: [{ id: 'g1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/memberOf?$skiptoken=m2' });
  http.on('get', '$skiptoken=m2', { value: [{ id: 'g2' }] });
  http.on('get', '/v1.0/groups', { value: [{ id: 'g1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/groups?$skiptoken=t2' });
  http.on('get', '$skiptoken=t2', { value: [{ id: 'g2' }, { id: 'g3' }] });

  t.deepEqual(await graph.getUserGroups('token'), [{ id: 'g1' }, { id: 'g2' }]);
  t.deepEqual(await graph.getAllTenantGroups('token'), [{ id: 'g1' }, { id: 'g2' }, { id: 'g3' }]);
});

test.serial('getUsersWithManager expands the manager in the same request and sorts by name', async t => {
  http.on('get', '/v1.0/users', {
    value: [
      { id: 'u2', displayName: 'Boss' },
      { id: 'u1', displayName: 'Alice', manager: { '@odata.type': '#microsoft.graph.user', id: 'u2' } },
    ],
  });

  t.deepEqual(await graph.getUsersWithManager('token'), [
    { id: 'u1', displayName: 'Alice', managerId: 'u2' },
    { id: 'u2', displayName: 'Boss', managerId: null },
  ]);
  t.is(http.calls.length, 1);
  t.is(http.calls[0].config.params.$expand, 'manager($select=id)');
});

test.serial('getUsersWithManager throws when Graph fails', async t => {
  http.fail('get', '/v1.0/users', 403);

  await t.throwsAsync(() => graph.getUsersWithManager('token'), { message: 'Failed to fetch users' });
});

test.serial('findUserByEmail filters by mail or user principal name', async t => {
  http.on('get', '/v1.0/users', { value: [{ id: 'u1', mail: 'alice@example.test' }] });

  t.deepEqual(await graph.findUserByEmail('token', 'alice@example.test'), { id: 'u1', mail: 'alice@example.test' });
  t.is(
    http.calls[0].config.params.$filter,
    "mail eq 'alice@example.test' or userPrincipalName eq 'alice@example.test'"
  );
  t.true(http.calls[0].config.params.$select.includes('lastPasswordChangeDateTime'));
});

test.serial('findUserByEmail escapes quotes and returns null when nobody matches', async t => {
  http.on('get', '/v1.0/users', { value: [] });

  t.is(await graph.findUserByEmail('token', "o'brien@example.test"), null);
  t.true(http.calls[0].config.params.$filter.startsWith("mail eq 'o''brien@example.test' or "));
});

test.serial('findUserByEmail throws when Graph fails', async t => {
  http.fail('get', '/v1.0/users', 500);

  await t.throwsAsync(() => graph.findUserByEmail('token', 'alice@example.test'), { message: 'Failed to look up user' });
});

test.serial('getUserForReset returns the user, or null when it no longer exists', async t => {
  http.on('get', '/users/u1', { id: 'u1', lastPasswordChangeDateTime: '2026-01-01T00:00:00Z' });

  t.deepEqual(await graph.getUserForReset('token', 'u1'), { id: 'u1', lastPasswordChangeDateTime: '2026-01-01T00:00:00Z' });
  t.is(await graph.getUserForReset('token', 'gone'), null);

  http.fail('get', '/users/u1', 500);
  await t.throwsAsync(() => graph.getUserForReset('token', 'u1'), { message: 'Failed to fetch user' });
});

test.serial('getGroupsWithMembership flags the groups the user belongs to', async t => {
  http.on('get', '/v1.0/groups', { value: [{ id: 'g1' }, { id: 'g2' }, { id: 'g3' }] });
  http.on('get', '/me/memberOf', { value: [{ id: 'g2' }] });

  const result = await graph.getGroupsWithMembership('token');

  t.deepEqual(result.allGroups, [
    { id: 'g1', isMember: false },
    { id: 'g2', isMember: true },
    { id: 'g3', isMember: false },
  ]);
  t.is(result.memberCount, 1);
  t.is(result.totalCount, 3);
});

test.serial('getGroupsWithMembership fails when the tenant groups cannot be read', async t => {
  http.fail('get', '/v1.0/groups', 403);
  http.on('get', '/me/memberOf', { value: [] });

  await t.throwsAsync(() => graph.getGroupsWithMembership('token'), { message: 'Failed to fetch tenant groups' });
});

test.serial('getAllUserData combines profile, photo, groups and manager', async t => {
  http.on('get', '/v1.0/me', { displayName: 'Alice' });
  http.on('get', '/me/memberOf', { value: [{ id: 'g1' }] });

  t.deepEqual(await graph.getAllUserData('token'), {
    profile: { displayName: 'Alice' },
    photo: null,
    groups: [{ id: 'g1' }],
    manager: null,
  });
});

test.serial('sendEmail posts a plain text message to /me/sendMail', async t => {
  http.on('post', '/me/sendMail', {});

  await graph.sendEmail('token', { to: 'bob@example.test', subject: 'Hi', body: 'Hello' });

  const { message } = http.calls[0].data;
  t.is(message.subject, 'Hi');
  t.deepEqual(message.body, { contentType: 'Text', content: 'Hello' });
  t.deepEqual(message.toRecipients, [{ emailAddress: { address: 'bob@example.test' } }]);
  t.false('attachments' in message);
});

test.serial('sendEmail maps attachments to Graph file attachments', async t => {
  http.on('post', '/me/sendMail', {});

  await graph.sendEmail('token', {
    to: 'bob@example.test',
    subject: 'Hi',
    body: 'Hello',
    attachments: [{ name: 'a.pdf', contentType: 'application/pdf', contentBytes: 'QUJD' }],
  });

  const { message } = http.calls[0].data;
  t.true(message.hasAttachments);
  t.deepEqual(message.attachments, [{
    '@odata.type': '#microsoft.graph.fileAttachment',
    name: 'a.pdf',
    contentType: 'application/pdf',
    contentBytes: 'QUJD',
  }]);
});

test.serial('sendEmail accepts several recipients', async t => {
  http.on('post', '/me/sendMail', {});

  await graph.sendEmail('token', { to: ['a@example.test', 'b@example.test'], subject: 'Hi', body: 'Hello' });

  t.deepEqual(http.calls[0].data.message.toRecipients, [
    { emailAddress: { address: 'a@example.test' } },
    { emailAddress: { address: 'b@example.test' } },
  ]);
});

test.serial('getGroupOwners returns the owners, or [] when they cannot be read', async t => {
  http.on('get', '/groups/g1/owners', { value: [{ id: 'u1', mail: 'a@example.test' }] });

  t.deepEqual(await graph.getGroupOwners('token', 'g1'), [{ id: 'u1', mail: 'a@example.test' }]);
  t.deepEqual(await graph.getGroupOwners('token', 'g2'), []);
});

test.serial('getOwnedGroups asks only for groups and throws on failure', async t => {
  http.on('get', '/me/ownedObjects/microsoft.graph.group', { value: [{ id: 'g1' }] });
  t.deepEqual(await graph.getOwnedGroups('token'), [{ id: 'g1' }]);

  http.fail('get', '/me/ownedObjects/microsoft.graph.group', 500);
  await t.throwsAsync(() => graph.getOwnedGroups('token'), { message: 'Failed to fetch owned groups' });
});

test.serial('addGroupMember posts a reference to the user and tolerates existing members', async t => {
  http.on('post', '/groups/g1/members/$ref', {});

  await graph.addGroupMember('owner-token', 'g1', 'u1');

  t.deepEqual(http.calls[0].data, { '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/u1' });
  t.is(http.calls[0].config.headers.Authorization, 'Bearer owner-token');

  http.on('post', '/groups/g1/members/$ref', () => {
    const error = new Error('Bad Request');
    error.response = { status: 400, data: { error: { message: 'One or more added object references already exist' } } };
    throw error;
  });
  await t.notThrowsAsync(() => graph.addGroupMember('owner-token', 'g1', 'u1'));

  http.fail('post', '/groups/g1/members/$ref', 403);
  const error = await t.throwsAsync(() => graph.addGroupMember('owner-token', 'g1', 'u1'), { message: 'Failed to add group member' });
  t.is(error.graphStatus, 403);
});

test.serial('getAuthenticationMethods returns null, not an empty list, when it cannot read them', async t => {
  http.on('get', '/me/authentication/methods', { value: [{ id: 'm1' }] });
  t.deepEqual(await graph.getAuthenticationMethods('token'), [{ id: 'm1' }]);

  http.fail('get', '/me/authentication/methods', 403);
  t.is(await graph.getAuthenticationMethods('token'), null);
});

test.serial('getUserDevices merges registered and owned devices', async t => {
  http.on('get', '/me/registeredDevices/microsoft.graph.device', { value: [{ id: 'd1' }, { id: 'd2' }] });
  http.on('get', '/me/ownedDevices/microsoft.graph.device', { value: [{ id: 'd2' }, { id: 'd3' }] });

  t.deepEqual((await graph.getUserDevices('token')).map(d => d.id), ['d1', 'd2', 'd3']);

  http.fail('get', '/me/ownedDevices/microsoft.graph.device', 500);
  t.deepEqual(await graph.getUserDevices('token'), []);
});

test.serial('sendEmail throws when Graph rejects the message', async t => {
  http.fail('post', '/me/sendMail', 403);

  await t.throwsAsync(
    () => graph.sendEmail('token', { to: 'bob@example.test', subject: 'Hi', body: 'Hello' }),
    { message: 'Failed to send email' }
  );
});

test.serial('updateUserProfile patches /me and throws on failure', async t => {
  http.on('patch', '/v1.0/me', {});

  await graph.updateUserProfile('token', { jobTitle: 'Dev' });
  t.deepEqual(http.calls[0].data, { jobTitle: 'Dev' });

  http.fail('patch', '/v1.0/me', 400);
  await t.throwsAsync(() => graph.updateUserProfile('token', {}), { message: 'Failed to update user profile' });
});

test.serial('getJoinedTeamsAndChannels keeps a team whose channels cannot be read', async t => {
  http.on('get', '/me/joinedTeams', { value: [{ id: 't1', displayName: 'One' }, { id: 't2', displayName: 'Two' }] });
  http.on('get', '/teams/t1/channels', { value: [{ id: 'c1', displayName: 'General' }] });
  http.fail('get', '/teams/t2/channels', 403);

  t.deepEqual(await graph.getJoinedTeamsAndChannels('token'), [
    { id: 't1', displayName: 'One', channels: [{ id: 'c1', displayName: 'General' }] },
    { id: 't2', displayName: 'Two', channels: [] },
  ]);
});

test.serial('getJoinedTeamsAndChannels throws when the teams cannot be read', async t => {
  http.fail('get', '/me/joinedTeams', 500);

  await t.throwsAsync(() => graph.getJoinedTeamsAndChannels('token'), { message: 'Failed to fetch teams' });
});

test.serial('getUserAdminRoles returns the role names', async t => {
  http.on('get', '/me/memberOf/microsoft.graph.directoryRole', {
    value: [{ displayName: 'Global Administrator' }, { displayName: 'User Administrator' }],
  });

  t.deepEqual(await graph.getUserAdminRoles('token'), ['Global Administrator', 'User Administrator']);
});

test.serial('getUserAdminRoles treats a Graph failure as no roles', async t => {
  http.fail('get', '/me/memberOf/microsoft.graph.directoryRole', 500);

  t.deepEqual(await graph.getUserAdminRoles('token'), []);
});

test.serial('resetUserPassword patches the password profile of the given user', async t => {
  http.on('patch', '/users/u1', {});

  await graph.resetUserPassword('app-token', 'u1', 'N3w-Passw0rd');

  t.deepEqual(http.calls[0].data, {
    passwordProfile: { password: 'N3w-Passw0rd', forceChangePasswordNextSignIn: false },
  });
  t.is(http.calls[0].config.headers.Authorization, 'Bearer app-token');
});

test.serial('resetUserPassword throws when Graph rejects the change', async t => {
  http.fail('patch', '/users/u1', 403);

  await t.throwsAsync(() => graph.resetUserPassword('app-token', 'u1', 'x'), { message: 'Failed to reset password' });
});

test.serial('hasUserPhoto checks photo metadata without downloading bytes', async t => {
  http.on('get', '/me/photo', { id: 'default' });
  t.true(await graph.hasUserPhoto('token'));
  t.is(http.calls[0].config.params.$select, 'id');

  http.fail('get', '/me/photo', 404);
  t.false(await graph.hasUserPhoto('token'));
});

test.serial('GET requests retry on 429 honoring Retry-After', async t => {
  let attempts = 0;
  http.on('get', '/v1.0/me', () => {
    attempts += 1;
    if (attempts === 1) {
      const error = new Error('Too Many Requests');
      error.response = { status: 429, headers: { 'retry-after': '0' }, data: {} };
      throw error;
    }
    return { displayName: 'Alice' };
  });

  const profile = await graph.getUserProfile('token');
  t.deepEqual(profile, { displayName: 'Alice' });
  t.is(attempts, 2);
});

test.serial('getServicePrincipal and getRooms cache successful lookups', async t => {
  http.on('get', '/servicePrincipals/sp1', { id: 'sp1', displayName: 'App One' });
  http.on('get', '/places/microsoft.graph.room', { value: [{ displayName: 'Sala 1', emailAddress: 'sala1@x' }] });

  const firstSp = await graph.getServicePrincipal('token', 'sp1');
  const secondSp = await graph.getServicePrincipal('token', 'sp1');
  t.deepEqual(firstSp, secondSp);
  t.is(http.callsTo('get', '/servicePrincipals/sp1').length, 1);

  const firstRooms = await graph.getRooms('token');
  const secondRooms = await graph.getRooms('token');
  t.deepEqual(firstRooms, secondRooms);
  t.is(http.callsTo('get', '/places/microsoft.graph.room').length, 1);
});

test.serial('graphBatch splits requests into chunks of 20 and returns a Map keyed by id', async t => {
  t.deepEqual(await graph.graphBatch('token', []), new Map());
  t.is(http.calls.length, 0);

  const requests = Array.from({ length: 25 }, (_, i) => ({
    id: String(i),
    method: 'GET',
    url: `/users/u${i}`,
  }));

  http.on('post', '/v1.0/$batch', (call) => ({
    responses: call.data.requests.map(r => ({
      id: r.id,
      status: 200,
      body: { id: `u${r.id}` },
    })),
  }));

  const results = await graph.graphBatch('token', requests);

  t.is(http.calls.length, 2);
  t.is(http.calls[0].data.requests.length, 20);
  t.is(http.calls[1].data.requests.length, 5);
  t.is(results.size, 25);
  t.deepEqual(results.get('0'), { status: 200, body: { id: 'u0' } });
  t.deepEqual(results.get('24'), { status: 200, body: { id: 'u24' } });
});

test.serial('graphBatch throws graphError when a chunk request fails', async t => {
  http.fail('post', '/v1.0/$batch', 503);

  const error = await t.throwsAsync(
    () => graph.graphBatch('token', [{ id: '0', method: 'GET', url: '/users/u1' }]),
    { message: 'Failed to execute batch request' }
  );
  t.is(error.graphStatus, 503);
});

test.serial('listAdminUsers filters out #EXT# users and applies accountEnabled filter modes', async t => {
  http.on('get', '/v1.0/users', {
    value: [
      { id: 'u1', userPrincipalName: 'alice@example.test', accountEnabled: true },
      { id: 'u2', userPrincipalName: 'guest_partner#EXT#@example.onmicrosoft.com', accountEnabled: true },
      { id: 'u3', userPrincipalName: 'other_guest#ext#@example.onmicrosoft.com', accountEnabled: true },
    ],
  });

  const active = await graph.listAdminUsers('token');
  t.deepEqual(active.users.map(u => u.id), ['u1']);
  t.false(active.signInsKnown);
  t.is(http.calls[0].config.params.$filter, 'accountEnabled eq true');
  t.is(http.calls[0].config.params.$top, 999);

  await graph.listAdminUsers('token', { onlyDisabled: true });
  t.is(http.calls[1].config.params.$filter, 'accountEnabled eq false');

  await graph.listAdminUsers('token', { onlyDisabled: null });
  t.is(http.calls[2].config.params.$filter, undefined);
});

test.serial('listAdminUsers merges signInActivity with appToken and falls back on failure', async t => {
  http.on('get', '/v1.0/users', (call) => {
    if (call.config.params.$select === 'id,signInActivity') {
      return {
        value: [
          { id: 'u1', signInActivity: { lastSignInDateTime: '2026-10-01T10:00:00Z' } },
        ],
      };
    }
    return {
      value: [
        { id: 'u1', userPrincipalName: 'alice@example.test' },
        { id: 'u2', userPrincipalName: 'bob@example.test' },
      ],
    };
  });

  const ok = await graph.listAdminUsers('user-token', { includeSignInActivity: true, appToken: 'app-token' });
  t.true(ok.signInsKnown);
  t.deepEqual(ok.users[0].signInActivity, { lastSignInDateTime: '2026-10-01T10:00:00Z' });
  t.is(ok.users[1].signInActivity, undefined);
  t.is(http.calls[1].config.headers.Authorization, 'Bearer app-token');
  t.is(http.calls[1].config.params.$top, 500);

  http.on('get', '/v1.0/users', (call) => {
    if (call.config.params.$select === 'id,signInActivity') {
      const error = new Error('Forbidden');
      error.response = { status: 403, data: { error: { code: 'Authorization_RequestDenied' } } };
      throw error;
    }
    return { value: [{ id: 'u1', userPrincipalName: 'alice@example.test' }] };
  });

  const fallback = await graph.listAdminUsers('user-token', { includeSignInActivity: true, appToken: 'app-token' });
  t.false(fallback.signInsKnown);
  t.deepEqual(fallback.users.map(u => u.id), ['u1']);
});

test.serial('fetchManagersBatch, fetchUserGroupsBatch and fetchUserLicensesBatch map batch responses', async t => {
  http.on('post', '/v1.0/$batch', (call) => ({
    responses: call.data.requests.map(req => {
      if (req.url.includes('/u1/manager')) {
        return { id: req.id, status: 200, body: { id: 'm1', displayName: 'Boss', userPrincipalName: 'boss@example.test', extra: 'x' } };
      }
      if (req.url.includes('/u2/manager')) {
        return { id: req.id, status: 404, body: { error: { code: 'Request_ResourceNotFound' } } };
      }
      if (req.url.includes('/u1/memberOf/')) {
        return { id: req.id, status: 200, body: { value: [{ id: 'g1', displayName: 'IT', extra: 1 }] } };
      }
      if (req.url.includes('/u2/memberOf/')) {
        return { id: req.id, status: 403, body: {} };
      }
      if (req.url.includes('/u1/licenseDetails')) {
        return { id: req.id, status: 200, body: { value: [{ skuId: 's1', skuPartNumber: 'SPE_E3', extra: 1 }] } };
      }
      return { id: req.id, status: 500, body: {} };
    }),
  }));

  const managers = await graph.fetchManagersBatch('token', ['u1', 'u2']);
  t.deepEqual(managers.get('u1'), { id: 'm1', displayName: 'Boss', userPrincipalName: 'boss@example.test' });
  t.is(managers.get('u2'), null);

  const groups = await graph.fetchUserGroupsBatch('token', ['u1', 'u2']);
  t.deepEqual(groups.get('u1'), [{ id: 'g1', displayName: 'IT' }]);
  t.deepEqual(groups.get('u2'), []);

  const licenses = await graph.fetchUserLicensesBatch('token', ['u1', 'u2']);
  t.deepEqual(licenses.get('u1'), [{ skuId: 's1', skuPartNumber: 'SPE_E3' }]);
  t.deepEqual(licenses.get('u2'), []);
});

test.serial('getAdminUser returns user, null on 404, and throws on other errors', async t => {
  http.on('get', '/users/u1', { id: 'u1', displayName: 'Alice', mailNickname: 'alice' });

  const user = await graph.getAdminUser('token', 'u1');
  t.deepEqual(user, { id: 'u1', displayName: 'Alice', mailNickname: 'alice' });
  t.true(http.calls[0].config.params.$select.includes('mailNickname'));
  t.true(http.calls[0].config.params.$select.includes('assignedLicenses'));

  t.is(await graph.getAdminUser('token', 'missing'), null);

  http.fail('get', '/users/u1', 500);
  const error = await t.throwsAsync(() => graph.getAdminUser('token', 'u1'), { message: 'Failed to fetch admin user' });
  t.is(error.graphStatus, 500);
});

test.serial('getUserManagerById, getUserGroupsById and getUserLicensesById handle success and errors', async t => {
  http.on('get', '/users/u1/manager', { id: 'm1', displayName: 'Boss' });
  t.deepEqual(await graph.getUserManagerById('token', 'u1'), { id: 'm1', displayName: 'Boss' });
  t.is(await graph.getUserManagerById('token', 'missing'), null);
  http.fail('get', '/users/u1/manager', 500);
  t.is(await graph.getUserManagerById('token', 'u1'), null);

  http.on('get', '/users/u1/memberOf/microsoft.graph.group', {
    value: [{ id: 'g1', displayName: 'One' }],
    '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users/u1/memberOf/microsoft.graph.group?$skiptoken=p2',
  });
  http.on('get', '$skiptoken=p2', { value: [{ id: 'g2', displayName: 'Two' }] });
  t.deepEqual(await graph.getUserGroupsById('token', 'u1'), [
    { id: 'g1', displayName: 'One' },
    { id: 'g2', displayName: 'Two' },
  ]);
  http.fail('get', '/users/u1/memberOf/microsoft.graph.group', 403);
  t.deepEqual(await graph.getUserGroupsById('token', 'u1'), []);

  http.on('get', '/users/u1/licenseDetails', { value: [{ skuId: 's1', skuPartNumber: 'E3' }] });
  t.deepEqual(await graph.getUserLicensesById('token', 'u1'), [{ skuId: 's1', skuPartNumber: 'E3' }]);
  http.fail('get', '/users/u1/licenseDetails', 403);
  t.deepEqual(await graph.getUserLicensesById('token', 'u1'), []);
});

test.serial('createAdminUser and updateAdminUser preserve Graph status and message on failure', async t => {
  http.on('post', '/v1.0/users', { id: 'new-1', displayName: 'New User' });
  const created = await graph.createAdminUser('token', { displayName: 'New User' });
  t.deepEqual(created, { id: 'new-1', displayName: 'New User' });
  t.deepEqual(http.calls[0].data, { displayName: 'New User' });

  http.on('post', '/v1.0/users', () => {
    const error = new Error('Bad Request');
    error.response = { status: 400, data: { error: { message: 'UserPrincipalName already exists' } } };
    throw error;
  });
  const createErr = await t.throwsAsync(
    () => graph.createAdminUser('token', {}),
    { message: 'Failed to create user' }
  );
  t.is(createErr.graphStatus, 400);
  t.is(createErr.graphMessage, 'UserPrincipalName already exists');

  http.on('patch', '/users/u1', {});
  await graph.updateAdminUser('token', 'u1', { jobTitle: 'Lead' });
  t.deepEqual(http.calls.at(-1).data, { jobTitle: 'Lead' });

  http.on('patch', '/users/u1', () => {
    const error = new Error('Forbidden');
    error.response = { status: 403, data: { error: { message: 'Insufficient privileges' } } };
    throw error;
  });
  const updateErr = await t.throwsAsync(
    () => graph.updateAdminUser('token', 'u1', { jobTitle: 'Lead' }),
    { message: 'Failed to update user' }
  );
  t.is(updateErr.graphStatus, 403);
  t.is(updateErr.graphMessage, 'Insufficient privileges');
});

test.serial('setUserManager puts the manager @odata.id reference', async t => {
  http.on('put', '/users/u1/manager/$ref', {});

  await graph.setUserManager('token', 'u1', 'm1');
  t.deepEqual(http.calls[0].data, { '@odata.id': 'https://graph.microsoft.com/v1.0/users/m1' });

  http.on('put', '/users/u1/manager/$ref', () => {
    const error = new Error('Bad Request');
    error.response = { status: 400, data: { error: { message: 'Cannot set self as manager' } } };
    throw error;
  });
  const err = await t.throwsAsync(
    () => graph.setUserManager('token', 'u1', 'u1'),
    { message: 'Failed to set user manager' }
  );
  t.is(err.graphStatus, 400);
  t.is(err.graphMessage, 'Cannot set self as manager');
});

test.serial('removeUserLicenses skips empty lists and posts removeLicenses otherwise', async t => {
  await graph.removeUserLicenses('token', 'u1', []);
  await graph.removeUserLicenses('token', 'u1', null);
  t.is(http.calls.length, 0);

  http.on('post', '/users/u1/assignLicense', {});
  await graph.removeUserLicenses('token', 'u1', ['sku-1', 'sku-2']);
  t.deepEqual(http.calls[0].data, { addLicenses: [], removeLicenses: ['sku-1', 'sku-2'] });

  http.on('post', '/users/u1/assignLicense', () => {
    const error = new Error('Bad Request');
    error.response = { status: 400, data: { error: { message: 'License inherited from group' } } };
    throw error;
  });
  const err = await t.throwsAsync(
    () => graph.removeUserLicenses('token', 'u1', ['sku-1']),
    { message: 'Failed to remove user licenses' }
  );
  t.is(err.graphStatus, 400);
  t.is(err.graphMessage, 'License inherited from group');
});

test.serial('addUserToGroup and removeUserFromGroup ignore 409/400 already-exist and 404 respectively', async t => {
  http.on('post', '/groups/g1/members/$ref', {});
  await graph.addUserToGroup('token', 'g1', 'u1');
  t.deepEqual(http.calls[0].data, { '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/u1' });

  http.fail('post', '/groups/g1/members/$ref', 409);
  await t.notThrowsAsync(() => graph.addUserToGroup('token', 'g1', 'u1'));

  http.on('post', '/groups/g1/members/$ref', () => {
    const error = new Error('Bad Request');
    error.response = { status: 400, data: { error: { message: 'One or more added object references already exist' } } };
    throw error;
  });
  await t.notThrowsAsync(() => graph.addUserToGroup('token', 'g1', 'u1'));

  http.fail('post', '/groups/g1/members/$ref', 403);
  const addErr = await t.throwsAsync(
    () => graph.addUserToGroup('token', 'g1', 'u1'),
    { message: 'Failed to add user to group' }
  );
  t.is(addErr.graphStatus, 403);

  http.on('delete', '/groups/g1/members/u1/$ref', {});
  await graph.removeUserFromGroup('token', 'g1', 'u1');

  http.fail('delete', '/groups/g1/members/u1/$ref', 404);
  await t.notThrowsAsync(() => graph.removeUserFromGroup('token', 'g1', 'u1'));

  http.fail('delete', '/groups/g1/members/u1/$ref', 403);
  const removeErr = await t.throwsAsync(
    () => graph.removeUserFromGroup('token', 'g1', 'u1'),
    { message: 'Failed to remove user from group' }
  );
  t.is(removeErr.graphStatus, 403);
});

test.serial('getUserMailboxSettingsAdmin never sends $select and categorizes status codes', async t => {
  http.on('get', '/users/u1/mailboxSettings', { timeZone: 'Romance Standard Time', automaticRepliesSetting: { status: 'disabled' } });
  const ok = await graph.getUserMailboxSettingsAdmin('token', 'u1');
  t.deepEqual(ok, {
    status: 'ok',
    settings: { timeZone: 'Romance Standard Time', automaticRepliesSetting: { status: 'disabled' } },
  });
  t.is(http.calls[0].config.params, undefined);

  // 404 -> no-mailbox
  const notFound = await graph.getUserMailboxSettingsAdmin('token', 'no-mbx');
  t.is(notFound.status, 'no-mailbox');
  t.is(notFound.settings, null);

  // 400 -> no-mailbox
  http.fail('get', '/users/u1/mailboxSettings', 400);
  const badReq = await graph.getUserMailboxSettingsAdmin('token', 'u1');
  t.is(badReq.status, 'no-mailbox');

  // MailboxNotEnabledForRESTAPI -> no-mailbox
  http.on('get', '/users/u1/mailboxSettings', () => {
    const error = new Error('Not Found');
    error.response = { status: 404, data: { error: { code: 'MailboxNotEnabledForRESTAPI' } } };
    throw error;
  });
  const notEnabled = await graph.getUserMailboxSettingsAdmin('token', 'u1');
  t.deepEqual(notEnabled, { status: 'no-mailbox', settings: null, errorCode: 'MailboxNotEnabledForRESTAPI' });

  // 403 -> access-denied
  http.on('get', '/users/u1/mailboxSettings', () => {
    const error = new Error('Forbidden');
    error.response = { status: 403, data: { error: { code: 'ErrorAccessDenied' } } };
    throw error;
  });
  const denied = await graph.getUserMailboxSettingsAdmin('token', 'u1');
  t.deepEqual(denied, { status: 'access-denied', settings: null, errorCode: 'ErrorAccessDenied' });

  // 500 -> error
  http.fail('get', '/users/u1/mailboxSettings', 500);
  const serverErr = await graph.getUserMailboxSettingsAdmin('token', 'u1');
  t.is(serverErr.status, 'error');
  t.is(serverErr.settings, null);
});

test.serial('updateUserMailboxSettingsAdmin patches mailbox settings and throws on failure', async t => {
  http.on('patch', '/users/u1/mailboxSettings', {});
  await graph.updateUserMailboxSettingsAdmin('token', 'u1', { automaticRepliesSetting: { status: 'alwaysEnabled' } });
  t.deepEqual(http.calls[0].data, { automaticRepliesSetting: { status: 'alwaysEnabled' } });

  http.on('patch', '/users/u1/mailboxSettings', () => {
    const error = new Error('Forbidden');
    error.response = { status: 403, data: { error: { message: 'Access is denied' } } };
    throw error;
  });
  const err = await t.throwsAsync(
    () => graph.updateUserMailboxSettingsAdmin('token', 'u1', {}),
    { message: 'Failed to update mailbox settings' }
  );
  t.is(err.graphStatus, 403);
  t.is(err.graphMessage, 'Access is denied');
});
