/**
 * Group Request Routes
 *
 * Asking to join a group, and the owners of the group deciding:
 * - POST /api/group-access-request - Ask to join the selected groups
 * - GET /groups/requests - Requests for my groups, and the ones I made
 * - POST /groups/requests/:id/approve - Add the requester to the group
 * - POST /groups/requests/:id/reject - Turn the request down
 *
 * An approval adds the member with the token of the owner who approves, so
 * the app can only do what that owner could do themselves.
 */

const express = require('express');
const router = express.Router();
const appConfig = require('../config/appConfig');
const { isAuthenticated } = require('../middleware/auth');
const { userError } = require('../middleware/errorHandler');
const { auditRequest } = require('../services/auditService');
const { mapWithConcurrency } = require('../services/asyncUtils');
const requestStore = require('../services/requestStore');
const { isManageable } = require('../services/accountInsights');
const {
  getGroupsWithMembership,
  getGroupOwners,
  getOwnedGroups,
  addGroupMember,
  sendEmail
} = require('../services/graphService');

function currentUser(req) {
  const account = req.session.account || {};
  return {
    id: account.localAccountId,
    name: account.name || account.username,
    email: account.username,
  };
}

/**
 * Send an email without letting a failure undo what it reports
 */
async function notify(req, emailData) {
  try {
    await sendEmail(req.session.accessToken, emailData);
  } catch (error) {
    console.error('Error sending notification:', error.message);
  }
}

/**
 * POST /api/group-access-request
 * Ask to join the selected groups. Each request goes to the owners of its
 * group; groups without owners, or that Graph cannot manage, go to the admin
 * mailbox as a single email.
 */
router.post('/api/group-access-request', isAuthenticated, async (req, res, next) => {
  try {
    const { selectedGroups } = req.body;

    if (!Array.isArray(selectedGroups) || selectedGroups.length === 0) {
      const err = new Error('Group information is required');
      err.status = 400;
      return next(err);
    }

    const accessToken = req.session.accessToken;
    const requester = currentUser(req);

    // Fetch group details, never trusting names sent by the browser
    const groupsData = await getGroupsWithMembership(accessToken);
    const wantedIds = new Set(selectedGroups);
    const groups = groupsData.allGroups.filter(g => wantedIds.has(g.id) && !g.isMember);

    const ownersByGroup = await mapWithConcurrency(groups, 6, group =>
      requester.id && isManageable(group) ? getGroupOwners(accessToken, group.id) : []
    );

    const forAdmin = [];
    const ownerTargets = [];

    for (const [index, group] of groups.entries()) {
      const owners = ownersByGroup[index].map(owner => owner.mail || owner.userPrincipalName).filter(Boolean);
      if (owners.length === 0) {
        forAdmin.push(group);
      } else {
        ownerTargets.push({ group, owners });
      }
    }

    const stored = requestStore.createRequests(ownerTargets.map(({ group }) => ({
      groupId: group.id,
      groupName: group.displayName,
      requester,
    })));
    const sentToOwners = ownerTargets.length;

    // Asking twice does not write to the owners twice
    const toNotify = stored
      .map((entry, index) => ({ ...entry, ...ownerTargets[index] }))
      .filter(entry => entry.created);

    await mapWithConcurrency(toNotify, 3, async ({ request, group, owners }) => {
      await Promise.all([
        auditRequest(req, 'group_request.created', { requestId: request.id, groupId: group.id }),
        notify(req, {
          to: owners,
          subject: `[${appConfig.APP_NAME}] Solicitud de acceso a ${group.displayName}`,
          body: `${requester.name} (${requester.email}) solicita unirse al grupo "${group.displayName}", del que eres propietario.\n\nPuedes aprobar o rechazar la solicitud aquí:\n${appConfig.APP_ORIGIN}/groups/requests`,
        }),
      ]);
    });

    if (forAdmin.length > 0) {
      let body = `The user ${requester.name} (${requester.email}) is requesting access to the following groups:\n\n`;
      forAdmin.forEach((group, index) => {
        body += `${index + 1}. ${group.displayName} (ID: ${group.id})\n`;
      });

      await sendEmail(accessToken, {
        to: appConfig.ADMIN_EMAIL,
        subject: '[m365-admin] Group Access Request',
        body,
      });
      await auditRequest(req, 'request.group_access', { groups: forAdmin.map(g => g.id) });
    }

    res.json({
      success: true,
      message: 'Group access request sent successfully',
      sentToOwners,
      sentToAdmin: forAdmin.length,
    });
  } catch (error) {
    console.error('Error sending group access request:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to send group access request'
    });
  }
});

/**
 * GET /groups/requests
 * Pending requests for the groups I own, and the requests I have made
 */
router.get('/groups/requests', isAuthenticated, async (req, res, next) => {
  try {
    const ownedGroups = await getOwnedGroups(req.session.accessToken);
    const { pending, mine } = requestStore.listForUser(ownedGroups.map(g => g.id), currentUser(req).id);

    res.render('group-requests', {
      title: 'Solicitudes de grupos',
      currentPage: 'groups',
      pending,
      mine,
    });
  } catch (error) {
    next(error);
  }
});

/**
 * Load the request named in the URL and check that the signed-in user may
 * decide on it: it must be pending and for a group they own
 */
async function loadRequestToResolve(req) {
  const request = requestStore.getRequest(req.params.id);

  if (!request || request.status !== 'pending') {
    throw userError(404, 'La solicitud no existe o ya está resuelta');
  }

  const ownedGroups = await getOwnedGroups(req.session.accessToken);
  if (!ownedGroups.some(group => group.id === request.groupId)) {
    await auditRequest(req, 'permission.denied', { path: req.path, method: req.method });
    throw userError(403, 'Solo los propietarios del grupo pueden resolver esta solicitud');
  }

  return request;
}

/**
 * POST /groups/requests/:id/approve
 * Add the requester to the group
 */
router.post('/groups/requests/:id/approve', isAuthenticated, async (req, res, next) => {
  try {
    const request = await loadRequestToResolve(req);

    await addGroupMember(req.session.accessToken, request.groupId, request.requester.id);
    requestStore.resolveRequest(request.id, 'approved', currentUser(req).email);
    await Promise.all([
      auditRequest(req, 'group_request.approved', {
        requestId: request.id,
        groupId: request.groupId,
        member: request.requester.email,
      }),
      notify(req, {
        to: request.requester.email,
        subject: `[${appConfig.APP_NAME}] Ya formas parte de ${request.groupName}`,
        body: `Tu solicitud para unirte al grupo "${request.groupName}" ha sido aprobada.`,
      }),
    ]);

    res.redirect('/groups/requests');
  } catch (error) {
    next(error);
  }
});

/**
 * POST /groups/requests/:id/reject
 * Turn the request down
 */
router.post('/groups/requests/:id/reject', isAuthenticated, async (req, res, next) => {
  try {
    const request = await loadRequestToResolve(req);

    requestStore.resolveRequest(request.id, 'rejected', currentUser(req).email);
    await Promise.all([
      auditRequest(req, 'group_request.rejected', {
        requestId: request.id,
        groupId: request.groupId,
        member: request.requester.email,
      }),
      notify(req, {
        to: request.requester.email,
        subject: `[${appConfig.APP_NAME}] Solicitud de acceso a ${request.groupName}`,
        body: `Tu solicitud para unirte al grupo "${request.groupName}" no ha sido aprobada. Si crees que es un error, habla con los propietarios del grupo.`,
      }),
    ]);

    res.redirect('/groups/requests');
  } catch (error) {
    next(error);
  }
});

module.exports = router;
