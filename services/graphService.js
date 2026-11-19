/**
 * Microsoft Graph API Service
 *
 * Handles all Microsoft Graph API calls with consistent auth headers,
 * timeouts, pagination bounds, 429 retry for idempotent reads, and
 * low-churn TTL caches (service principals, meeting rooms).
 */

const axios = require('axios');
const graphConfig = require('../config/graphConfig');
const { mapWithConcurrency, createTtlCache } = require('./asyncUtils');
const { compareEs } = require('./formatters');

const GRAPH_API_BASE = 'https://graph.microsoft.com/v1.0';
const DEFAULT_TIMEOUT_MS = 15 * 1000;
const LARGE_TIMEOUT_MS = 30 * 1000;
const MAX_PAGES = 100;
const MAX_429_RETRIES = 2;
const MAX_RETRY_WAIT_MS = 5 * 1000;

const servicePrincipalCache = createTtlCache({ ttlMs: 60 * 60 * 1000, maxSize: 500 });
const roomsCache = createTtlCache({ ttlMs: 5 * 60 * 1000, maxSize: 1 });

function resetGraphCaches() {
  servicePrincipalCache.clear();
  roomsCache.clear();
}

/**
 * Build the error thrown to callers, keeping the HTTP status Graph answered
 * with so that an expired or revoked token (401) can be told apart
 * @param {string} message - Message for the caller
 * @param {Error} error - Original axios error
 * @returns {Error}
 */
function graphError(message, error) {
  const err = new Error(message);
  err.graphStatus = error.response?.status;
  if (error.response?.data?.error?.message) {
    err.graphMessage = error.response.data.error.message;
  }
  return err;
}

function retryAfterMs(error, attempt) {
  const raw = error.response?.headers?.['retry-after'] ?? error.response?.headers?.['Retry-After'];
  const seconds = raw != null ? Number(raw) : Number.NaN;
  if (Number.isFinite(seconds) && seconds >= 0) {
    return seconds * 1000;
  }
  return 250 * (attempt + 1);
}

const wait = (ms) => new Promise(resolve => setTimeout(resolve, ms));

/**
 * Send one authenticated request to Microsoft Graph with a timeout and
 * bounded Retry-After handling for throttled GETs.
 */
async function graphRequest(method, url, accessToken, { data, params, headers, responseType, timeout = DEFAULT_TIMEOUT_MS } = {}) {
  const config = {
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...headers,
    },
    timeout,
  };
  if (params) config.params = params;
  if (responseType) config.responseType = responseType;

  const canRetry = method === 'get';
  for (let attempt = 0; ; attempt += 1) {
    try {
      if (method === 'get' || method === 'delete') {
        return await axios[method](url, config);
      }
      return await axios[method](url, data, config);
    } catch (error) {
      if (canRetry && error.response?.status === 429 && attempt < MAX_429_RETRIES) {
        const delayMs = retryAfterMs(error, attempt);
        if (delayMs <= MAX_RETRY_WAIT_MS) {
          if (delayMs > 0) await wait(delayMs);
          continue;
        }
      }
      throw error;
    }
  }
}

/**
 * Read a Graph collection to the end, following @odata.nextLink
 * @param {string} url - Collection endpoint
 * @param {string} accessToken - Valid access token
 * @param {Object} [opts] - params, headers, timeout, maxPages
 * @returns {Promise<Array>} Items of every page
 */
async function getAllPages(url, accessToken, { params, headers, timeout = DEFAULT_TIMEOUT_MS, maxPages = MAX_PAGES } = {}) {
  let response = await graphRequest('get', url, accessToken, { params, headers, timeout });
  const items = [...(response.data.value || [])];
  let pages = 1;

  while (response.data['@odata.nextLink'] && pages < maxPages) {
    // The next link already carries the query, so only headers and timeout are sent again
    response = await graphRequest('get', response.data['@odata.nextLink'], accessToken, { headers, timeout });
    items.push(...(response.data.value || []));
    pages += 1;
  }

  return items;
}

/**
 * Get user profile information with extended fields
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @returns {Promise<Object>} User profile data
 */
async function getUserProfile(accessToken) {
  try {
    const response = await graphRequest('get', graphConfig.userProfile.endpoint, accessToken);
    return response.data;
  } catch (error) {
    console.error('Error fetching user profile:', error.response?.data || error.message);
    throw graphError('Failed to fetch user profile', error);
  }
}

// What the directory shows about a person
const DIRECTORY_USER_FIELDS = 'id,displayName,givenName,surname,mail,jobTitle,department,officeLocation,businessPhones,mobilePhone';

/**
 * Get one person of the directory, with the signed-in user's token
 * @param {string} accessToken - Valid access token
 * @param {string} userId - Directory id, already checked for its format
 * @returns {Promise<Object|null>} The user, or null when there is none
 */
/**
 * Get active, licensed member accounts (excludes rooms, shared mailboxes,
 * unlicensed aliases, disabled accounts and external guests).
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @returns {Promise<Array>} Array of licensed users sorted by displayName
 */
async function getLicensedUsers(accessToken) {
  try {
    const users = await getAllPages(`${GRAPH_API_BASE}/users`, accessToken, {
      params: {
        $select: 'id,displayName,mail,userPrincipalName,jobTitle,department,assignedLicenses,accountEnabled,userType',
        $orderby: 'displayName',
        $top: 999,
      },
    });
    return users
      .filter(user =>
        user.accountEnabled !== false
        && user.userType !== 'Guest'
        && Array.isArray(user.assignedLicenses)
        && user.assignedLicenses.length > 0
      )
      .map(({ assignedLicenses, accountEnabled, userType, ...user }) => user);
  } catch (error) {
    console.error('Error fetching licensed users:', error.response?.data || error.message);
    return [];
  }
}

async function getUserById(accessToken, userId) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}`, accessToken, {
      params: { $select: DIRECTORY_USER_FIELDS },
    });
    return response.data;
  } catch (error) {
    if (error.response?.status === 404) return null;
    console.error(`Error fetching user ${userId}:`, error.response?.data || error.message);
    throw graphError('Failed to fetch user', error);
  }
}

/**
 * The people whose manager is the signed-in user
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array|null>} Users, or null when they could not be read
 */
async function getDirectReports(accessToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/me/directReports/microsoft.graph.user`, accessToken, {
      params: { $select: DIRECTORY_USER_FIELDS, $top: 999 },
    });
  } catch (error) {
    console.error('Error fetching direct reports:', error.response?.data || error.message);
    return null;
  }
}

/**
 * The licenses the organization has bought, and how many are in use
 * @param {string} accessToken - Valid access token (an administrator's)
 * @returns {Promise<Array|null>} subscribedSkus, or null when they could not be read
 */
async function getSubscribedSkus(accessToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/subscribedSkus`, accessToken);
  } catch (error) {
    console.error('Error fetching subscribed SKUs:', error.response?.data || error.message);
    return null;
  }
}

// What the accounts review looks at; signInActivity needs Entra ID P1
const REVIEW_FIELDS = 'id,displayName,userPrincipalName,accountEnabled,userType,lastPasswordChangeDateTime,assignedLicenses,createdDateTime';

/**
 * Every account, for the administrators' review, read with the app's token
 * @param {string} appToken - App-only token
 * @returns {Promise<{users: Array, signInsKnown: boolean, signInsError: string|null}|null>} The users, whether
 *   their last sign-in could be read and, if not, Graph's error code; null when nothing could be read
 */
async function getUsersForReview(appToken) {
  let signInsError = null;
  try {
    // Graph serves at most 500 users per page when signInActivity is selected
    const users = await getAllPages(`${GRAPH_API_BASE}/users`, appToken, {
      params: { $select: `${REVIEW_FIELDS},signInActivity`, $top: 500 },
      timeout: LARGE_TIMEOUT_MS,
    });
    return { users, signInsKnown: true, signInsError };
  } catch (error) {
    signInsError = error.response?.data?.error?.code || error.message;
    console.error('Error fetching users with sign-in activity, trying without:', error.response?.data || error.message);
  }
  try {
    const users = await getAllPages(`${GRAPH_API_BASE}/users`, appToken, {
      params: { $select: REVIEW_FIELDS, $top: 999 },
      timeout: LARGE_TIMEOUT_MS,
    });
    return { users, signInsKnown: false, signInsError };
  } catch (error) {
    console.error('Error fetching users for review:', error.response?.data || error.message);
    return null;
  }
}

/**
 * Whether each person has registered a second factor, from the
 * authentication methods report, read with the app's token
 * @param {string} appToken - App-only token
 * @returns {Promise<Array|null>} userRegistrationDetails, or null when unreadable
 */
async function getMfaRegistration(appToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/reports/authenticationMethods/userRegistrationDetails`, appToken, {
      params: { $select: 'id,userPrincipalName,isMfaRegistered,isMfaCapable', $top: 999 },
      timeout: LARGE_TIMEOUT_MS,
    });
  } catch (error) {
    console.error('Error fetching MFA registration:', error.response?.data || error.message);
    return null;
  }
}

/**
 * Get all users in the organization for selection dropdowns
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @returns {Promise<Array>} Array of users with basic info
 */
async function getAllUsers(accessToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/users`, accessToken, {
      params: {
        $select: 'id,displayName,mail,jobTitle,department,officeLocation,businessPhones,mobilePhone',
        $orderby: 'displayName',
        $top: 999, // Largest page Graph allows
      },
    });
  } catch (error) {
    console.error('Error fetching all users:', error.response?.data || error.message);
    return [];
  }
}

/**
 * Get active, licensed member accounts (excludes rooms, shared mailboxes,
 * unlicensed aliases, disabled accounts and external guests).
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @returns {Promise<Array>} Array of licensed users sorted by displayName
 */
async function getLicensedUsers(accessToken) {
  try {
    const users = await getAllPages(`${GRAPH_API_BASE}/users`, accessToken, {
      params: {
        $select: 'id,displayName,mail,userPrincipalName,jobTitle,department,assignedLicenses,accountEnabled,userType',
        $orderby: 'displayName',
        $top: 999,
      },
    });
    return users
      .filter(user =>
        user.accountEnabled !== false
        && user.userType !== 'Guest'
        && Array.isArray(user.assignedLicenses)
        && user.assignedLicenses.length > 0
      )
      .map(({ assignedLicenses, accountEnabled, userType, ...user }) => user);
  } catch (error) {
    console.error('Error fetching licensed users:', error.response?.data || error.message);
    return [];
  }
}

/**
 * Send an email via Microsoft Graph
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @param {Object} emailData - Email data (to, subject, body); `to` is one address or a list
 * @returns {Promise<void>}
 */
async function sendEmail(accessToken, emailData) {
  try {
    const { to, subject, body, attachments } = emailData;
    const recipients = Array.isArray(to) ? to : [to];

    const message = {
      message: {
        subject,
        body: {
          contentType: 'Text',
          content: body,
        },
        toRecipients: recipients.map(address => ({
          emailAddress: { address },
        })),
      },
    };

    if (attachments && attachments.length > 0) {
      message.message.hasAttachments = true;
      message.message.attachments = attachments.map(att => ({
        '@odata.type': '#microsoft.graph.fileAttachment',
        name: att.name,
        contentType: att.contentType,
        contentBytes: att.contentBytes, // base64 encoded
      }));
    }

    await graphRequest('post', `${GRAPH_API_BASE}/me/sendMail`, accessToken, {
      data: message,
      headers: { 'Content-Type': 'application/json' },
    });

    console.log(`Email sent successfully to ${recipients.join(', ')}`);
  } catch (error) {
    console.error('Error sending email:', error.response?.data || error.message);
    throw graphError('Failed to send email', error);
  }
}

/**
 * Get user profile photo
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @returns {Promise<Buffer|null>} Photo data as buffer, or null if no photo exists
 */
async function getUserPhoto(accessToken) {
  try {
    const response = await graphRequest('get', graphConfig.userPhoto.endpoint, accessToken, {
      responseType: 'arraybuffer',
    });
    return Buffer.from(response.data, 'binary');
  } catch (error) {
    // 404 means user has no profile photo - this is normal
    if (error.response?.status === 404) {
      console.log('No profile photo found for user');
      return null;
    }
    console.error('Error fetching user photo:', error.response?.data || error.message);
    return null;
  }
}

/**
 * Get another user's profile photo by directory ID
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @param {string} userId - Directory ID or UPN of the target user
 * @returns {Promise<Buffer|null>} Photo data as buffer, or null if no photo exists or unreadable
 */
async function getUserPhotoById(accessToken, userId) {
  if (!userId) return null;
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}/photo/$value`, accessToken, {
      responseType: 'arraybuffer',
    });
    return Buffer.from(response.data, 'binary');
  } catch (error) {
    if (error.response?.status === 404) {
      return null;
    }
    console.error(`Error fetching user photo for ${userId}:`, error.response?.data || error.message);
    return null;
  }
}

/**
 * Check whether the signed-in user has a profile photo without downloading its bytes
 * @param {string} accessToken - Valid access token
 * @returns {Promise<boolean>}
 */
/**
 * Get another user's profile photo by directory ID
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @param {string} userId - Directory ID or UPN of the target user
 * @returns {Promise<Buffer|null>} Photo data as buffer, or null if no photo exists or unreadable
 */
async function getUserPhotoById(accessToken, userId) {
  if (!userId) return null;
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}/photo/$value`, accessToken, {
      responseType: 'arraybuffer',
    });
    return Buffer.from(response.data, 'binary');
  } catch (error) {
    if (error.response?.status === 404) {
      return null;
    }
    console.error(`Error fetching user photo for ${userId}:`, error.response?.data || error.message);
    return null;
  }
}

async function hasUserPhoto(accessToken) {
  try {
    await graphRequest('get', `${GRAPH_API_BASE}/me/photo`, accessToken, {
      params: { $select: 'id' },
    });
    return true;
  } catch (error) {
    if (error.response?.status !== 404) {
      console.error('Error checking user photo:', error.response?.data || error.message);
    }
    return false;
  }
}

/**
 * Get user's group memberships
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @returns {Promise<Array>} Array of groups the user belongs to
 */
async function getUserGroups(accessToken) {
  try {
    return await getAllPages(graphConfig.groupMemberships.endpoint, accessToken, {
      params: { $select: 'id,displayName', $top: 999 },
    });
  } catch (error) {
    console.error('Error fetching user groups:', error.response?.data || error.message);
    // Return empty array instead of throwing - groups are optional
    return [];
  }
}

/**
 * Get user's manager information
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @returns {Promise<Object|null>} Manager data, or null if no manager assigned
 */
async function getUserManager(accessToken) {
  try {
    const response = await graphRequest('get', graphConfig.manager.endpoint, accessToken);
    return response.data;
  } catch (error) {
    // 404 means user has no manager assigned - this is normal
    if (error.response?.status === 404) {
      console.log('No manager found for user');
      return null;
    }
    console.error('Error fetching user manager:', error.response?.data || error.message);
    return null;
  }
}

/**
 * Get all groups in the tenant
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @returns {Promise<Array>} Array of all groups in the tenant
 */
async function getAllTenantGroups(accessToken) {
  try {
    // No $orderby: Graph answers "Invalid search request" for it on groups
    // with large pages in some tenants, so the list is sorted here instead
    const groups = await getAllPages(`${GRAPH_API_BASE}/groups`, accessToken, {
      params: {
        $select: 'id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes',
        $top: 999,
      },
    });
    return groups.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || '', 'es'));
  } catch (error) {
    console.error('Error fetching all tenant groups:', error.response?.data || error.message);
    throw graphError('Failed to fetch tenant groups', error);
  }
}

/**
 * Get all tenant groups with user membership status
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @returns {Promise<Object>} Object with all groups and user's groups
 */
async function getGroupsWithMembership(accessToken) {
  try {
    const [allGroups, userGroups] = await Promise.all([
      getAllTenantGroups(accessToken),
      getUserGroups(accessToken),
    ]);

    const userGroupIds = new Set(userGroups.map(g => g.id));
    const groupsWithStatus = allGroups.map(group => ({
      ...group,
      isMember: userGroupIds.has(group.id),
    }));

    return {
      allGroups: groupsWithStatus,
      userGroups,
      memberCount: userGroups.length,
      totalCount: allGroups.length,
    };
  } catch (error) {
    console.error('Error fetching groups with membership:', error.message);
    throw error;
  }
}

/**
 * Get the owners of a group
 * @param {string} accessToken - Valid access token
 * @param {string} groupId - Group ID
 * @param {Object} [options] - Optional { top } to cap the page when only counting (e.g. top: 2)
 * @returns {Promise<Array>} Owners, or [] if the group has none or they cannot be read
 */
async function getGroupOwners(accessToken, groupId, { top } = {}) {
  const url = `${GRAPH_API_BASE}/groups/${encodeURIComponent(groupId)}/owners`;
  try {
    if (top) {
      const response = await graphRequest('get', url, accessToken, {
        params: { $select: 'id', $top: top },
      });
      return response.data.value || [];
    }
    return await getAllPages(url, accessToken, {
      params: { $select: 'id,displayName,mail,userPrincipalName', $top: 999 },
    });
  } catch (error) {
    console.error(`Error fetching owners of group ${groupId}:`, error.response?.data || error.message);
    // Return empty array instead of throwing - a group without owners is handled by its caller
    return [];
  }
}

/**
 * Get the groups the signed-in user owns
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array>} Groups (id and name)
 */
async function getOwnedGroups(accessToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/me/ownedObjects/microsoft.graph.group`, accessToken, {
      params: {
        $select: 'id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes',
        $top: 999,
      },
    });
  } catch (error) {
    console.error('Error fetching owned groups:', error.response?.data || error.message);
    throw graphError('Failed to fetch owned groups', error);
  }
}

/**
 * Add a user to a group, with the rights of whoever the token belongs to
 * @param {string} accessToken - Valid access token of an owner of the group
 * @param {string} groupId - Group ID
 * @param {string} userId - Directory ID of the new member
 * @returns {Promise<void>} Also resolves when the user was already a member
 */
async function addGroupMember(accessToken, groupId, userId) {
  try {
    await graphRequest('post', `${GRAPH_API_BASE}/groups/${encodeURIComponent(groupId)}/members/$ref`, accessToken, {
      data: { '@odata.id': `${GRAPH_API_BASE}/directoryObjects/${encodeURIComponent(userId)}` },
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    const message = error.response?.data?.error?.message || '';
    if (error.response?.status === 400 && message.includes('already exist')) {
      return;
    }
    console.error(`Error adding member to group ${groupId}:`, error.response?.data || error.message);
    throw graphError('Failed to add group member', error);
  }
}

/**
 * Get the people in a group (members that are groups or devices are left out)
 * @param {string} accessToken - Valid access token
 * @param {string} groupId - Group ID
 * @returns {Promise<Array>} Users, with userType telling guests apart
 */
async function getGroupMembers(accessToken, groupId) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/groups/${encodeURIComponent(groupId)}/members/microsoft.graph.user`, accessToken, {
      params: { $select: 'id,displayName,mail,userPrincipalName,userType', $top: 999 },
    });
  } catch (error) {
    console.error(`Error fetching members of group ${groupId}:`, error.response?.data || error.message);
    throw graphError('Failed to fetch group members', error);
  }
}

/**
 * Remove a user from a group, with the rights of whoever the token belongs to
 * @param {string} accessToken - Valid access token of an owner of the group
 * @param {string} groupId - Group ID
 * @param {string} userId - Directory ID of the member
 * @returns {Promise<void>} Also resolves when the user was not a member
 */
async function removeGroupMember(accessToken, groupId, userId) {
  try {
    await graphRequest('delete', `${GRAPH_API_BASE}/groups/${encodeURIComponent(groupId)}/members/${encodeURIComponent(userId)}/$ref`, accessToken);
  } catch (error) {
    if (error.response?.status === 404) {
      return;
    }
    console.error(`Error removing member from group ${groupId}:`, error.response?.data || error.message);
    throw graphError('Failed to remove group member', error);
  }
}

/**
 * Close every session of the signed-in user, on every device and app. They
 * have to sign in again everywhere, this app included.
 * @param {string} accessToken - Valid access token
 * @returns {Promise<void>}
 */
async function revokeSignInSessions(accessToken) {
  try {
    await graphRequest('post', `${GRAPH_API_BASE}/me/revokeSignInSessions`, accessToken, {
      data: {},
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error('Error revoking sign-in sessions:', error.response?.data || error.message);
    throw graphError('Failed to revoke sign-in sessions', error);
  }
}

/**
 * Get who the signed-in user shares their calendar with
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array|null>} calendarPermissions, or null if they cannot be read
 */
async function getCalendarPermissions(accessToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/me/calendar/calendarPermissions`, accessToken);
  } catch (error) {
    console.error('Error fetching calendar permissions:', error.response?.data || error.message);
    // null and not []: "could not check" must not look like "shared with nobody"
    return null;
  }
}

/**
 * Share the signed-in user's calendar with somebody
 * @param {string} accessToken - Valid access token
 * @param {string} address - Who, by email address
 * @param {string} role - freeBusyRead, limitedRead, read or write
 * @returns {Promise<Object>} The permission Graph created
 */
async function shareCalendar(accessToken, address, role) {
  try {
    const response = await graphRequest('post', `${GRAPH_API_BASE}/me/calendar/calendarPermissions`, accessToken, {
      data: { emailAddress: { address }, role },
    });
    return response.data;
  } catch (error) {
    console.error('Error sharing calendar:', error.response?.data || error.message);
    throw graphError('Failed to share calendar', error);
  }
}

/**
 * Take back somebody's access to the signed-in user's calendar
 * @param {string} accessToken - Valid access token
 * @param {string} permissionId - Id of the calendarPermission
 * @returns {Promise<void>}
 */
async function removeCalendarPermission(accessToken, permissionId) {
  try {
    await graphRequest('delete', `${GRAPH_API_BASE}/me/calendar/calendarPermissions/${encodeURIComponent(permissionId)}`, accessToken);
  } catch (error) {
    console.error('Error removing calendar permission:', error.response?.data || error.message);
    throw graphError('Failed to remove calendar permission', error);
  }
}

// People Graph answers about in one presence request
const PRESENCES_PER_REQUEST = 650;

/**
 * The Teams presence of several people
 * @param {string} accessToken - Valid access token
 * @param {Array<string>} userIds - Directory ids
 * @returns {Promise<Array|null>} One entry per person (id, availability, activity), or null when it could not be read
 */
async function getPresences(accessToken, userIds) {
  try {
    const chunks = [];
    for (let i = 0; i < userIds.length; i += PRESENCES_PER_REQUEST) {
      chunks.push(userIds.slice(i, i + PRESENCES_PER_REQUEST));
    }
    const pages = await mapWithConcurrency(chunks, 4, async (ids) => {
      const response = await graphRequest('post', `${GRAPH_API_BASE}/communications/getPresencesByUserId`, accessToken, {
        data: { ids },
      });
      return response.data.value || [];
    });
    return pages.flat();
  } catch (error) {
    console.error('Error fetching presences:', error.response?.data || error.message);
    // null, not []: "could not check" must not look like "everybody offline"
    return null;
  }
}

/**
 * The delegated permissions apps hold on the signed-in user's account: what
 * the user consented to, and what an administrator consented to for everybody
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array|null>} The grants, or null when they could not be read
 */
async function getMyAppGrants(accessToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/me/oauth2PermissionGrants`, accessToken, {
      params: { $top: 999 },
    });
  } catch (error) {
    console.error('Error fetching app grants:', error.response?.data || error.message);
    return null;
  }
}

/**
 * An application of the directory, by the id of its service principal.
 * Tenant-wide and low-churn, so successful lookups are cached for 1 hour.
 * @param {string} accessToken - Valid access token
 * @param {string} id - Service principal id, already checked for its format
 * @returns {Promise<Object|null>} Name and publisher, or null when it cannot be read
 */
async function getServicePrincipal(accessToken, id) {
  return servicePrincipalCache.get(id, async () => {
    try {
      const response = await graphRequest('get', `${GRAPH_API_BASE}/servicePrincipals/${encodeURIComponent(id)}`, accessToken, {
        params: { $select: 'id,appId,displayName,publisherName,verifiedPublisher,homepage' },
      });
      return response.data;
    } catch (error) {
      console.error(`Error fetching service principal ${id}:`, error.response?.data || error.message);
      return null;
    }
  });
}

// Pages of the drive listing read at most, so a huge OneDrive does not hang the page
const DRIVE_PAGES_MAX = 25;

/**
 * The items of the user's OneDrive that are shared with somebody
 * @param {string} accessToken - Valid access token
 * @returns {Promise<{items: Array, complete: boolean}|null>} The shared items, and whether the whole drive was read; null when it could not be
 */
async function getSharedDriveItems(accessToken) {
  try {
    let response = await graphRequest('get', `${GRAPH_API_BASE}/me/drive/root/delta`, accessToken, {
      params: { $select: 'id,name,webUrl,shared,file,folder,deleted,parentReference', $top: 1000 },
      timeout: LARGE_TIMEOUT_MS,
    });
    const items = [];
    let pages = 1;
    const keep = (page) => items.push(...(page.value || []).filter(item => item.shared && !item.deleted));
    keep(response.data);
    while (response.data['@odata.nextLink'] && pages < DRIVE_PAGES_MAX) {
      response = await graphRequest('get', response.data['@odata.nextLink'], accessToken, {
        timeout: LARGE_TIMEOUT_MS,
      });
      keep(response.data);
      pages += 1;
    }
    return { items, complete: !response.data['@odata.nextLink'] };
  } catch (error) {
    console.error('Error listing shared drive items:', error.response?.data || error.message);
    return null;
  }
}

/**
 * Who an item of the user's OneDrive is shared with
 * @param {string} accessToken - Valid access token
 * @param {string} itemId - Drive item id, already checked for its format
 * @returns {Promise<Array|null>} Its permissions, or null when they could not be read
 */
async function getDriveItemPermissions(accessToken, itemId) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/me/drive/items/${encodeURIComponent(itemId)}/permissions`, accessToken);
  } catch (error) {
    console.error(`Error fetching permissions of item ${itemId}:`, error.response?.data || error.message);
    return null;
  }
}

/**
 * Stop sharing an item of the user's OneDrive with somebody, or through a link
 * @param {string} accessToken - Valid access token
 * @param {string} itemId - Drive item id, already checked for its format
 * @param {string} permissionId - Permission id, already checked for its format
 * @returns {Promise<void>}
 */
async function removeDriveItemPermission(accessToken, itemId, permissionId) {
  try {
    await graphRequest('delete', `${GRAPH_API_BASE}/me/drive/items/${encodeURIComponent(itemId)}/permissions/${encodeURIComponent(permissionId)}`, accessToken);
  } catch (error) {
    console.error(`Error removing permission ${permissionId} of item ${itemId}:`, error.response?.data || error.message);
    throw graphError('Failed to remove sharing', error);
  }
}

/**
 * Whether the signed-in user can open somebody else's mailbox (a shared one)
 * @param {string} accessToken - Valid access token
 * @param {string} address - Address of the mailbox
 * @returns {Promise<boolean|null>} null when Graph answered neither yes nor no
 */
async function canOpenMailbox(accessToken, address) {
  try {
    await graphRequest('get', `${GRAPH_API_BASE}/users/${encodeURIComponent(address)}/mailFolders/inbox`, accessToken, {
      params: { $select: 'id' },
    });
    return true;
  } catch (error) {
    const status = error.response?.status;
    if (status === 403 || status === 404) {
      return false;
    }
    console.error(`Error checking access to mailbox ${address}:`, error.response?.data || error.message);
    return null;
  }
}

/**
 * Get the meeting rooms of the organization (cached for 5 minutes)
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array|null>} Rooms, or null if they cannot be read
 */
async function getRooms(accessToken) {
  return roomsCache.get('rooms', async () => {
    try {
      return await getAllPages(`${GRAPH_API_BASE}/places/microsoft.graph.room`, accessToken, {
        params: { $top: 100 },
      });
    } catch (error) {
      console.error('Error fetching rooms:', error.response?.data || error.message);
      return null;
    }
  });
}

// Mailboxes Graph accepts in one getSchedule request
const SCHEDULES_PER_REQUEST = 20;

/**
 * Get the free/busy information of several mailboxes for a period
 * @param {string} accessToken - Valid access token
 * @param {Array<string>} addresses - Mailboxes (people or rooms)
 * @param {string} start - Local date and time, e.g. 2026-10-12T10:00:00
 * @param {string} end - Local date and time
 * @param {string} timeZone - Time zone of start and end, as Exchange names it
 * @returns {Promise<Array>} One scheduleInformation per mailbox
 */
async function getSchedules(accessToken, addresses, start, end, timeZone) {
  try {
    const chunks = [];
    for (let index = 0; index < addresses.length; index += SCHEDULES_PER_REQUEST) {
      chunks.push(addresses.slice(index, index + SCHEDULES_PER_REQUEST));
    }
    const pages = await mapWithConcurrency(chunks, 4, async (schedules) => {
      const response = await graphRequest('post', `${GRAPH_API_BASE}/me/calendar/getSchedule`, accessToken, {
        data: {
          schedules,
          startTime: { dateTime: start, timeZone },
          endTime: { dateTime: end, timeZone },
          availabilityViewInterval: 15,
        },
        headers: { 'Content-Type': 'application/json' },
      });
      return response.data.value || [];
    });
    return pages.flat();
  } catch (error) {
    console.error('Error fetching schedules:', error.response?.data || error.message);
    throw graphError('Failed to fetch schedules', error);
  }
}

/**
 * Get what is in a room's calendar for a period, with who booked and who goes
 * @param {string} appToken - App-only token: no user may open a room's mailbox
 * @param {string} address - The room's mailbox
 * @param {string} start - UTC date and time, e.g. 2026-10-12T00:00:00
 * @param {string} end - UTC date and time
 * @param {string} timeZone - Time zone the events come back in, as Exchange names it
 * @returns {Promise<Array|null>} Events, or null if they cannot be read
 */
async function getRoomEvents(appToken, address, start, end, timeZone) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/users/${encodeURIComponent(address)}/calendarView`, appToken, {
      params: {
        startDateTime: start,
        endDateTime: end,
        $select: 'subject,organizer,start,end,isAllDay,isCancelled,showAs,sensitivity,type,isOnlineMeeting,attendees,createdDateTime',
        $orderby: 'start/dateTime',
        $top: 100,
      },
      headers: { Prefer: `outlook.timezone="${timeZone}"` },
    });
  } catch (error) {
    console.error('Error fetching room events:', error.response?.data || error.message);
    return null;
  }
}

/**
 * Get the calendars of the signed-in user
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array|null>} Calendars, or null if they cannot be read
 */
async function getMyCalendars(accessToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/me/calendars`, accessToken, {
      params: { $select: 'id,name,canEdit,canShare,canViewPrivateItems,isDefaultCalendar,owner', $top: 100 },
    });
  } catch (error) {
    console.error('Error fetching calendars:', error.response?.data || error.message);
    return null;
  }
}

/**
 * Get the people inside a distribution list or group, nested groups included
 * @param {string} accessToken - Valid access token
 * @param {string} address - The list's email address
 * @returns {Promise<Array|null>} Users, or null if the list is not found or cannot be read
 */
async function getListMembers(accessToken, address) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/groups`, accessToken, {
      params: { $filter: `mail eq '${address.replace(/'/g, "''")}'`, $select: 'id', $top: 1 },
    });
    const [group] = response.data.value || [];
    if (!group) return null;
    return await getAllPages(`${GRAPH_API_BASE}/groups/${encodeURIComponent(group.id)}/transitiveMembers/microsoft.graph.user`, accessToken, {
      // Graph only casts transitive members as an advanced query
      params: { $select: 'id,displayName,mail', $top: 999, $count: true },
      headers: { ConsistencyLevel: 'eventual' },
    });
  } catch (error) {
    console.error(`Error fetching members of ${address}:`, error.response?.data || error.message);
    return null;
  }
}

/**
 * Create an event in one of the signed-in user's calendars; Exchange invites its attendees
 * @param {string} accessToken - Valid access token
 * @param {string} calendarId - Calendar ID
 * @param {Object} event - The event, as Graph takes it
 * @returns {Promise<Object>} The created event
 */
async function createCalendarEvent(accessToken, calendarId, event) {
  try {
    const response = await graphRequest('post', `${GRAPH_API_BASE}/me/calendars/${encodeURIComponent(calendarId)}/events`, accessToken, {
      data: event,
      headers: { 'Content-Type': 'application/json' },
    });
    return response.data;
  } catch (error) {
    console.error('Error creating calendar event:', error.response?.data || error.message);
    throw graphError('Failed to create calendar event', error);
  }
}

/**
 * Get the sign-in methods the signed-in user has registered
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array|null>} Methods, or null if they cannot be read
 */
async function getAuthenticationMethods(accessToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/me/authentication/methods`, accessToken);
  } catch (error) {
    console.error('Error fetching authentication methods:', error.response?.data || error.message);
    // null and not []: "could not check" must not look like "has none"
    return null;
  }
}

/**
 * Get the devices registered to or owned by the signed-in user
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array>} Devices, or [] if they cannot be read
 */
async function getUserDevices(accessToken) {
  try {
    const opts = {
      params: {
        $select: 'id,displayName,operatingSystem,operatingSystemVersion,approximateLastSignInDateTime',
        $top: 999,
      },
    };
    const [registered, owned] = await Promise.all([
      getAllPages(`${GRAPH_API_BASE}/me/registeredDevices/microsoft.graph.device`, accessToken, opts),
      getAllPages(`${GRAPH_API_BASE}/me/ownedDevices/microsoft.graph.device`, accessToken, opts),
    ]);

    // Most devices appear in both lists
    const byId = new Map([...registered, ...owned].map(device => [device.id, device]));
    return [...byId.values()];
  } catch (error) {
    console.error('Error fetching user devices:', error.response?.data || error.message);
    return [];
  }
}

/**
 * Get all user data in one call (convenience function)
 * @param {string} accessToken - Valid access token for Microsoft Graph
 * @returns {Promise<Object>} Complete user data object
 */
async function getAllUserData(accessToken) {
  try {
    const [profile, photo, groups, manager] = await Promise.all([
      getUserProfile(accessToken),
      getUserPhoto(accessToken),
      getUserGroups(accessToken),
      getUserManager(accessToken),
    ]);

    return { profile, photo, groups, manager };
  } catch (error) {
    console.error('Error fetching all user data:', error.message);
    throw error;
  }
}

/**
 * Update user profile information
 * @param {string} accessToken - Valid access token
 * @param {Object} profileData - Data to update
 * @returns {Promise<void>}
 */
async function updateUserProfile(accessToken, profileData) {
  try {
    await graphRequest('patch', `${GRAPH_API_BASE}/me`, accessToken, {
      data: profileData,
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error('Error updating user profile:', error.response?.data || error.message);
    throw graphError('Failed to update user profile', error);
  }
}

/**
 * Replace the profile photo of the signed-in user
 * @param {string} accessToken - Valid access token
 * @param {Buffer} photo - Image data
 * @param {string} contentType - image/jpeg or image/png
 * @returns {Promise<void>}
 */
async function updateUserPhoto(accessToken, photo, contentType) {
  try {
    await graphRequest('put', graphConfig.userPhoto.endpoint, accessToken, {
      data: photo,
      headers: { 'Content-Type': contentType },
    });
  } catch (error) {
    console.error('Error updating user photo:', error.response?.data || error.message);
    throw graphError('Failed to update user photo', error);
  }
}

/**
 * Get the account details shown on the account page
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Object>} Password change date, alternate addresses, phones and creation date
 */
async function getAccountStatus(accessToken) {
  try {
    const response = await graphRequest('get', graphConfig.userProfile.endpoint, accessToken, {
      params: { $select: 'lastPasswordChangeDateTime,otherMails,createdDateTime,mobilePhone,businessPhones' },
    });
    return response.data;
  } catch (error) {
    console.error('Error fetching account status:', error.response?.data || error.message);
    throw graphError('Failed to fetch account status', error);
  }
}

/**
 * Get the licenses assigned to the signed-in user
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array>} License details, or [] if they cannot be read
 */
async function getUserLicenses(accessToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/me/licenseDetails`, accessToken);
  } catch (error) {
    console.error('Error fetching user licenses:', error.response?.data || error.message);
    // Return empty array instead of throwing - licenses are optional
    return [];
  }
}

/**
 * Get the groups the signed-in user belongs to, with what is needed to tell
 * what kind of group each one is
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array>} Groups (directory roles are not included)
 */
async function getMyGroups(accessToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/me/memberOf/microsoft.graph.group`, accessToken, {
      params: {
        $select: 'id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,resourceProvisioningOptions',
        $top: 999,
      },
    });
  } catch (error) {
    console.error('Error fetching my groups:', error.response?.data || error.message);
    throw graphError('Failed to fetch groups', error);
  }
}

/**
 * Get the inbox rules of the signed-in user
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array|null>} Rules, or null if they cannot be read
 */
async function getInboxRules(accessToken) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/me/mailFolders/inbox/messageRules`, accessToken);
  } catch (error) {
    console.error('Error fetching inbox rules:', error.response?.data || error.message);
    // null and not []: "could not check" must not look like "has none"
    return null;
  }
}

/**
 * Get how much OneDrive space the signed-in user has and uses
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Object|null>} Quota (total, used, remaining in bytes), or null if unknown
 */
async function getDriveQuota(accessToken) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/me/drive`, accessToken, {
      params: { $select: 'quota' },
    });
    return response.data.quota || null;
  } catch (error) {
    console.error('Error fetching drive quota:', error.response?.data || error.message);
    return null;
  }
}

/**
 * Get the most recent sign-ins of a user from the directory's sign-in log.
 * Users cannot read their own, so this needs an app-only token.
 * @param {string} appToken - App-only access token
 * @param {string} userId - Directory ID of the user
 * @returns {Promise<Array>} Sign-ins, newest first
 */
async function getRecentSignIns(appToken, userId) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/auditLogs/signIns`, appToken, {
      params: {
        $filter: `userId eq '${userId.replace(/'/g, "''")}'`,
        $orderby: 'createdDateTime desc',
        $top: 25,
      },
    });
    return response.data.value || [];
  } catch (error) {
    console.error('Error fetching sign-ins:', error.response?.data || error.message);
    throw graphError('Failed to fetch sign-ins', error);
  }
}

/**
 * Get the mailbox settings of the signed-in user
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Object>} mailboxSettings (time zone, language, working hours...)
 */
async function getMailboxSettings(accessToken) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/me/mailboxSettings`, accessToken);
    return response.data;
  } catch (error) {
    console.error('Error fetching mailbox settings:', error.response?.data || error.message);
    throw graphError('Failed to fetch mailbox settings', error);
  }
}

/**
 * Change mailbox settings of the signed-in user
 * @param {string} accessToken - Valid access token
 * @param {Object} settings - The settings to change; others are left as they are
 * @returns {Promise<void>}
 */
async function updateMailboxSettings(accessToken, settings) {
  try {
    await graphRequest('patch', `${GRAPH_API_BASE}/me/mailboxSettings`, accessToken, {
      data: settings,
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error('Error updating mailbox settings:', error.response?.data || error.message);
    throw graphError('Failed to update mailbox settings', error);
  }
}

/**
 * Get the preferred language of the signed-in user's account
 * @param {string} accessToken - Valid access token
 * @returns {Promise<string|null>} Locale such as es-ES, or null if none is set
 */
async function getPreferredLanguage(accessToken) {
  try {
    const response = await graphRequest('get', graphConfig.userProfile.endpoint, accessToken, {
      params: { $select: 'preferredLanguage' },
    });
    return response.data.preferredLanguage || null;
  } catch (error) {
    console.error('Error fetching preferred language:', error.response?.data || error.message);
    throw graphError('Failed to fetch preferred language', error);
  }
}

/**
 * Get the time zone of the signed-in user's mailbox
 * @param {string} accessToken - Valid access token
 * @returns {Promise<string>} Time zone name as Exchange knows it, or UTC if none is set
 */
async function getMailboxTimeZone(accessToken) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/me/mailboxSettings/timeZone`, accessToken);
    return response.data.value || 'UTC';
  } catch (error) {
    console.error('Error fetching mailbox time zone:', error.response?.data || error.message);
    throw graphError('Failed to fetch mailbox settings', error);
  }
}

/**
 * Get the automatic replies (out of office) setting of the signed-in user,
 * with its dates expressed in the given time zone
 * @param {string} accessToken - Valid access token
 * @param {string} timeZone - Time zone from getMailboxTimeZone
 * @returns {Promise<Object>} automaticRepliesSetting
 */
async function getAutomaticReplies(accessToken, timeZone) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/me/mailboxSettings/automaticRepliesSetting`, accessToken, {
      headers: { Prefer: `outlook.timezone="${timeZone}"` },
    });
    return response.data;
  } catch (error) {
    console.error('Error fetching automatic replies:', error.response?.data || error.message);
    throw graphError('Failed to fetch automatic replies', error);
  }
}

/**
 * Change the automatic replies (out of office) setting of the signed-in user
 * @param {string} accessToken - Valid access token
 * @param {Object} setting - automaticRepliesSetting
 * @returns {Promise<void>}
 */
async function updateAutomaticReplies(accessToken, setting) {
  try {
    await updateMailboxSettings(accessToken, { automaticRepliesSetting: setting });
  } catch (error) {
    const err = new Error('Failed to update automatic replies');
    err.graphStatus = error.graphStatus;
    throw err;
  }
}

/**
 * Get user extended directory information (alias of getUserProfile)
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Object>} User extended data
 */
function getUserExtendedInfo(accessToken) {
  return getUserProfile(accessToken);
}

/**
 * Get joined teams and their channels (with bounded concurrency)
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array>} Array of teams with channels
 */
async function getJoinedTeamsAndChannels(accessToken) {
  try {
    const teams = await getAllPages(`${GRAPH_API_BASE}/me/joinedTeams`, accessToken, {
      params: { $select: 'id,displayName,description' },
    });

    return await mapWithConcurrency(teams, 6, async (team) => {
      try {
        const channels = await getAllPages(`${GRAPH_API_BASE}/teams/${encodeURIComponent(team.id)}/channels`, accessToken, {
          params: { $select: 'id,displayName,description,membershipType,webUrl' },
        });
        return { ...team, channels };
      } catch (error) {
        console.error(`Error fetching channels for team ${team.id}:`, error.message);
        return { ...team, channels: [] };
      }
    });
  } catch (error) {
    console.error('Error fetching joined teams:', error.response?.data || error.message);
    throw graphError('Failed to fetch teams', error);
  }
}

/**
 * Get administrative directory roles for the user, returning null when Graph fails
 * so callers can distinguish a failed lookup from "no roles".
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array<string>|null>}
 */
async function fetchUserAdminRoles(accessToken) {
  try {
    const roles = await getAllPages(`${GRAPH_API_BASE}/me/memberOf/microsoft.graph.directoryRole`, accessToken, {
      params: { $select: 'displayName' },
    });
    return roles.map(role => role.displayName);
  } catch (error) {
    console.error('Error fetching admin roles:', error.response?.data || error.message);
    return null;
  }
}

/**
 * Get administrative directory roles for the user
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array>} Array of admin role names
 */
async function getUserAdminRoles(accessToken) {
  return (await fetchUserAdminRoles(accessToken)) || [];
}

/**
 * Get every user in the organization together with the id of their manager
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array>} Users sorted by name, each with a managerId (or null)
 */
async function getUsersWithManager(accessToken) {
  try {
    const users = await getAllPages(`${GRAPH_API_BASE}/users`, accessToken, {
      params: {
        $select: 'id,displayName,mail,jobTitle,department',
        // One request per page instead of one manager request per user
        $expand: 'manager($select=id)',
        $top: 999,
      },
    });

    return users
      .map(({ manager, ...user }) => ({ ...user, managerId: manager ? manager.id : null }))
      .sort((a, b) => compareEs(a.displayName, b.displayName));
  } catch (error) {
    console.error('Error fetching users with manager:', error.response?.data || error.message);
    throw graphError('Failed to fetch users', error);
  }
}

// Fields needed to issue and check a password reset link
const RESET_USER_FIELDS = 'id,mail,userPrincipalName,otherMails,lastPasswordChangeDateTime';

/**
 * Find a user by mail address or user principal name
 * @param {string} accessToken - Valid access token
 * @param {string} email - Address to look for
 * @returns {Promise<Object|null>} The user, or null if nobody has that address
 */
async function findUserByEmail(accessToken, email) {
  try {
    // Single quotes are doubled inside an OData string literal
    const literal = email.replace(/'/g, "''");
    const response = await graphRequest('get', `${GRAPH_API_BASE}/users`, accessToken, {
      params: {
        $filter: `mail eq '${literal}' or userPrincipalName eq '${literal}'`,
        $select: RESET_USER_FIELDS,
      },
    });
    return (response.data.value || [])[0] || null;
  } catch (error) {
    console.error('Error looking up user by email:', error.response?.data || error.message);
    throw graphError('Failed to look up user', error);
  }
}

/**
 * Get the fields of a user that a password reset link depends on
 * @param {string} accessToken - Valid access token
 * @param {string} userId - User ID
 * @returns {Promise<Object|null>} The user, or null if it no longer exists
 */
async function getUserForReset(accessToken, userId) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}`, accessToken, {
      params: { $select: RESET_USER_FIELDS },
    });
    return response.data;
  } catch (error) {
    if (error.response?.status === 404) return null;
    console.error(`Error fetching user ${userId}:`, error.response?.data || error.message);
    throw graphError('Failed to fetch user', error);
  }
}

/**
 * Reset user password via Microsoft Graph API
 * @param {string} accessToken - Valid access token with appropriate permissions
 * @param {string} userId - User ID
 * @param {string} newPassword - New password
 * @returns {Promise<void>}
 */
async function resetUserPassword(accessToken, userId, newPassword) {
  try {
    const passwordProfile = {
      passwordProfile: {
        password: newPassword,
        forceChangePasswordNextSignIn: false,
      },
    };
    await graphRequest('patch', `${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}`, accessToken, {
      data: passwordProfile,
      headers: { 'Content-Type': 'application/json' },
    });
    console.log(`Password reset successfully for user ${userId}`);
  } catch (error) {
    console.error('Error resetting password:', error.response?.data || error.message);
    throw graphError('Failed to reset password', error);
  }
}

// Microsoft Graph $batch accepts at most 20 sub-requests per call
const GRAPH_BATCH_LIMIT = 20;

const ADMIN_USER_LIST_FIELDS = [
  'id', 'userPrincipalName', 'displayName', 'givenName', 'surname', 'mail',
  'jobTitle', 'department', 'companyName', 'officeLocation', 'mobilePhone',
  'streetAddress', 'city', 'state', 'postalCode', 'country', 'usageLocation',
  'preferredLanguage', 'accountEnabled', 'userType', 'employeeId',
  'employeeType', 'employeeHireDate', 'lastPasswordChangeDateTime', 'assignedLicenses',
].join(',');

const ADMIN_USER_DETAIL_FIELDS = [
  'id', 'userPrincipalName', 'displayName', 'givenName', 'surname', 'mail',
  'mailNickname', 'jobTitle', 'department', 'officeLocation', 'mobilePhone',
  'businessPhones', 'usageLocation', 'accountEnabled', 'city', 'country',
  'postalCode', 'state', 'streetAddress', 'companyName', 'employeeId',
  'employeeType', 'employeeHireDate', 'preferredLanguage', 'userType',
  'lastPasswordChangeDateTime', 'assignedLicenses',
].join(',');

const NO_MAILBOX_ERROR_RE = /MailboxNotEnabledForRESTAPI|ResourceNotFound|ErrorMailboxStoreUnavailable|BadRequest/;

/**
 * Execute Graph JSON batch requests in chunks of 20
 * @param {string} accessToken - Valid access token
 * @param {Array<{id: string, method: string, url: string, headers?: Object, body?: any}>} requests
 * @returns {Promise<Map<string, {status: number, body: any}>>} Responses keyed by sub-request id
 */
async function graphBatch(accessToken, requests) {
  const results = new Map();
  if (!requests || requests.length === 0) return results;

  for (let i = 0; i < requests.length; i += GRAPH_BATCH_LIMIT) {
    const chunk = requests.slice(i, i + GRAPH_BATCH_LIMIT);
    try {
      const response = await graphRequest('post', `${GRAPH_API_BASE}/$batch`, accessToken, {
        data: { requests: chunk },
        headers: { 'Content-Type': 'application/json' },
        timeout: LARGE_TIMEOUT_MS,
      });
      for (const item of response.data?.responses || []) {
        results.set(String(item.id), { status: item.status, body: item.body });
      }
    } catch (error) {
      console.error('Error executing Graph batch request:', error.response?.data || error.message);
      throw graphError('Failed to execute batch request', error);
    }
  }

  return results;
}

/**
 * List internal users for admin management, optionally merging signInActivity
 * @param {string} accessToken - Valid access token
 * @param {Object} [options]
 * @param {boolean|null} [options.onlyDisabled=false] - false: enabled only, true: disabled only, null: all users
 * @param {boolean} [options.includeSignInActivity=false] - Whether to enrich with signInActivity via appToken
 * @param {string|null} [options.appToken=null] - App-only token for signInActivity lookup
 * @returns {Promise<{users: Array, signInsKnown: boolean}>}
 */
async function listAdminUsers(accessToken, { onlyDisabled = false, includeSignInActivity = false, appToken = null } = {}) {
  const params = {
    $select: ADMIN_USER_LIST_FIELDS,
    $top: 999,
  };
  if (onlyDisabled !== null) {
    params.$filter = onlyDisabled ? 'accountEnabled eq false' : 'accountEnabled eq true';
  }

  let rawUsers;
  try {
    rawUsers = await getAllPages(`${GRAPH_API_BASE}/users`, accessToken, {
      params,
      timeout: LARGE_TIMEOUT_MS,
    });
  } catch (error) {
    console.error('Error fetching admin users:', error.response?.data || error.message);
    throw graphError('Failed to fetch admin users', error);
  }

  let users = rawUsers.filter(u => !String(u.userPrincipalName || '').toLowerCase().includes('#ext#'));
  let signInsKnown = false;

  if (includeSignInActivity && appToken) {
    try {
      const activityRows = await getAllPages(`${GRAPH_API_BASE}/users`, appToken, {
        params: { $select: 'id,signInActivity', $top: 500 },
        timeout: LARGE_TIMEOUT_MS,
      });
      const byId = new Map(activityRows.map(row => [row.id, row.signInActivity || null]));
      users = users.map(user => (
        byId.has(user.id) ? { ...user, signInActivity: byId.get(user.id) } : user
      ));
      signInsKnown = true;
    } catch (error) {
      console.error('Error fetching sign-in activity for admin users:', error.response?.data || error.message);
      signInsKnown = false;
    }
  }

  return { users, signInsKnown };
}

/**
 * Fetch managers for multiple users using $batch
 * @param {string} accessToken - Valid access token
 * @param {Array<string>} userIds - Directory IDs
 * @returns {Promise<Map<string, {id: string, displayName: string, userPrincipalName: string}|null>>}
 */
async function fetchManagersBatch(accessToken, userIds) {
  const result = new Map();
  if (!userIds || userIds.length === 0) return result;

  const requests = userIds.map((uid, i) => ({
    id: String(i),
    method: 'GET',
    url: `/users/${encodeURIComponent(uid)}/manager?$select=id,displayName,userPrincipalName`,
  }));
  const responses = await graphBatch(accessToken, requests);

  userIds.forEach((uid, i) => {
    const res = responses.get(String(i));
    if (res && res.status === 200 && res.body) {
      result.set(uid, {
        id: res.body.id,
        displayName: res.body.displayName,
        userPrincipalName: res.body.userPrincipalName,
      });
    } else {
      result.set(uid, null);
    }
  });

  return result;
}

/**
 * Fetch group memberships for multiple users using $batch
 * @param {string} accessToken - Valid access token
 * @param {Array<string>} userIds - Directory IDs
 * @returns {Promise<Map<string, Array<{id: string, displayName: string}>>>}
 */
async function fetchUserGroupsBatch(accessToken, userIds) {
  const result = new Map();
  if (!userIds || userIds.length === 0) return result;

  const requests = userIds.map((uid, i) => ({
    id: String(i),
    method: 'GET',
    url: `/users/${encodeURIComponent(uid)}/memberOf/microsoft.graph.group?$select=id,displayName&$top=999`,
  }));
  const responses = await graphBatch(accessToken, requests);

  userIds.forEach((uid, i) => {
    const res = responses.get(String(i));
    if (res && res.status === 200 && Array.isArray(res.body?.value)) {
      result.set(uid, res.body.value.map(g => ({ id: g.id, displayName: g.displayName })));
    } else {
      result.set(uid, []);
    }
  });

  return result;
}

/**
 * Fetch license details for multiple users using $batch
 * @param {string} accessToken - Valid access token
 * @param {Array<string>} userIds - Directory IDs
 * @returns {Promise<Map<string, Array<{skuId: string, skuPartNumber: string}>>>}
 */
async function fetchUserLicensesBatch(accessToken, userIds) {
  const result = new Map();
  if (!userIds || userIds.length === 0) return result;

  const requests = userIds.map((uid, i) => ({
    id: String(i),
    method: 'GET',
    url: `/users/${encodeURIComponent(uid)}/licenseDetails?$select=skuId,skuPartNumber`,
  }));
  const responses = await graphBatch(accessToken, requests);

  userIds.forEach((uid, i) => {
    const res = responses.get(String(i));
    if (res && res.status === 200 && Array.isArray(res.body?.value)) {
      result.set(uid, res.body.value.map(l => ({ skuId: l.skuId, skuPartNumber: l.skuPartNumber })));
    } else {
      result.set(uid, []);
    }
  });

  return result;
}

/**
 * Get full admin details for a single user by directory ID or UPN
 * @param {string} accessToken - Valid access token
 * @param {string} idOrUpn - User ID or UPN
 * @returns {Promise<Object|null>} User object, or null on 404
 */
async function getAdminUser(accessToken, idOrUpn) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/users/${encodeURIComponent(idOrUpn)}`, accessToken, {
      params: { $select: ADMIN_USER_DETAIL_FIELDS },
    });
    return response.data;
  } catch (error) {
    if (error.response?.status === 404) return null;
    console.error('Error fetching admin user %s:', encodeURIComponent(idOrUpn), error.response?.data || error.message);
    throw graphError('Failed to fetch admin user', error);
  }
}

/**
 * Get a user's manager by user ID or UPN
 * @param {string} accessToken - Valid access token
 * @param {string} idOrUpn - User ID or UPN
 * @returns {Promise<Object|null>} Manager object, or null on 404 or error
 */
async function getUserManagerById(accessToken, idOrUpn) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/users/${encodeURIComponent(idOrUpn)}/manager`, accessToken, {
      params: { $select: 'id,displayName,userPrincipalName,jobTitle,mail' },
    });
    return response.data;
  } catch (error) {
    if (error.response?.status !== 404) {
      console.error(`Error fetching manager for user ${idOrUpn}:`, error.response?.data || error.message);
    }
    return null;
  }
}

/**
 * Get a user's group memberships by user ID
 * @param {string} accessToken - Valid access token
 * @param {string} userId - User ID
 * @returns {Promise<Array>} Groups, or [] on error
 */
async function getUserGroupsById(accessToken, userId) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}/memberOf/microsoft.graph.group`, accessToken, {
      params: { $select: 'id,displayName', $top: 999 },
    });
  } catch (error) {
    console.error(`Error fetching groups for user ${userId}:`, error.response?.data || error.message);
    return [];
  }
}

/**
 * Get a user's license details by user ID
 * @param {string} accessToken - Valid access token
 * @param {string} userId - User ID
 * @returns {Promise<Array>} License details, or [] on error
 */
async function getUserLicensesById(accessToken, userId) {
  try {
    return await getAllPages(`${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}/licenseDetails`, accessToken, {
      params: { $select: 'skuId,skuPartNumber' },
    });
  } catch (error) {
    console.error(`Error fetching licenses for user ${userId}:`, error.response?.data || error.message);
    return [];
  }
}

/**
 * Create a new directory user
 * @param {string} accessToken - Valid access token
 * @param {Object} payload - User creation payload
 * @returns {Promise<Object>} Created user
 */
async function createAdminUser(accessToken, payload) {
  try {
    const response = await graphRequest('post', `${GRAPH_API_BASE}/users`, accessToken, {
      data: payload,
      headers: { 'Content-Type': 'application/json' },
    });
    return response.data;
  } catch (error) {
    console.error('Error creating user:', error.response?.data || error.message);
    throw graphError('Failed to create user', error);
  }
}

/**
 * Update a directory user by ID or UPN
 * @param {string} accessToken - Valid access token
 * @param {string} idOrUpn - User ID or UPN
 * @param {Object} payload - Fields to patch
 * @returns {Promise<void>}
 */
async function updateAdminUser(accessToken, idOrUpn, payload) {
  try {
    await graphRequest('patch', `${GRAPH_API_BASE}/users/${encodeURIComponent(idOrUpn)}`, accessToken, {
      data: payload,
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error(`Error updating user ${idOrUpn}:`, error.response?.data || error.message);
    throw graphError('Failed to update user', error);
  }
}

/**
 * Assign a manager to a user
 * @param {string} accessToken - Valid access token
 * @param {string} userId - User ID
 * @param {string} managerId - Manager directory ID
 * @returns {Promise<void>}
 */
async function setUserManager(accessToken, userId, managerId) {
  try {
    await graphRequest('put', `${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}/manager/$ref`, accessToken, {
      data: { '@odata.id': `${GRAPH_API_BASE}/users/${encodeURIComponent(managerId)}` },
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error(`Error setting manager for user ${userId}:`, error.response?.data || error.message);
    throw graphError('Failed to set user manager', error);
  }
}

/**
 * Remove licenses from a user
 * @param {string} accessToken - Valid access token
 * @param {string} userId - User ID
 * @param {Array<string>} skuIds - SKU IDs to remove
 * @returns {Promise<void>}
 */
async function removeUserLicenses(accessToken, userId, skuIds) {
  if (!skuIds || skuIds.length === 0) return;
  try {
    await graphRequest('post', `${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}/assignLicense`, accessToken, {
      data: { addLicenses: [], removeLicenses: skuIds },
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error(`Error removing licenses for user ${userId}:`, error.response?.data || error.message);
    throw graphError('Failed to remove user licenses', error);
  }
}

/**
 * Add and remove licenses of a user in one call
 * @param {string} accessToken - Valid access token
 * @param {string} userId - User directory ID
 * @param {{add?: string[], remove?: string[]}} changes - SKU ids
 * @returns {Promise<void>}
 */
async function assignUserLicenses(accessToken, userId, { add = [], remove = [] } = {}) {
  if (add.length === 0 && remove.length === 0) return;
  try {
    await graphRequest('post', `${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}/assignLicense`, accessToken, {
      data: { addLicenses: add.map(skuId => ({ skuId, disabledPlans: [] })), removeLicenses: remove },
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error(`Error assigning licenses for user ${userId}:`, error.response?.data || error.message);
    throw graphError('Failed to assign user licenses', error);
  }
}

/**
 * Add a user to a group, ignoring already-member conflicts (409 or 400 "already exist")
 * @param {string} accessToken - Valid access token
 * @param {string} groupId - Group ID
 * @param {string} userId - User directory ID
 * @returns {Promise<void>}
 */
async function addUserToGroup(accessToken, groupId, userId) {
  try {
    await graphRequest('post', `${GRAPH_API_BASE}/groups/${encodeURIComponent(groupId)}/members/$ref`, accessToken, {
      data: { '@odata.id': `${GRAPH_API_BASE}/directoryObjects/${encodeURIComponent(userId)}` },
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    const status = error.response?.status;
    const message = error.response?.data?.error?.message || '';
    if (status === 409 || (status === 400 && message.includes('already exist'))) {
      return;
    }
    console.error('Error adding user %s to group %s:', encodeURIComponent(userId), encodeURIComponent(groupId), error.response?.data || error.message);
    throw graphError('Failed to add user to group', error);
  }
}

/**
 * Remove a user from a group, ignoring 404 (not a member)
 * @param {string} accessToken - Valid access token
 * @param {string} groupId - Group ID
 * @param {string} userId - User directory ID
 * @returns {Promise<void>}
 */
async function removeUserFromGroup(accessToken, groupId, userId) {
  try {
    await graphRequest('delete', `${GRAPH_API_BASE}/groups/${encodeURIComponent(groupId)}/members/${encodeURIComponent(userId)}/$ref`, accessToken);
  } catch (error) {
    if (error.response?.status === 404) {
      return;
    }
    console.error(`Error removing user ${userId} from group ${groupId}:`, error.response?.data || error.message);
    throw graphError('Failed to remove user from group', error);
  }
}

/**
 * Read another user's mailbox settings (never passes $select, which Graph rejects with 400)
 * @param {string} accessToken - Valid access token
 * @param {string} userId - User ID
 * @returns {Promise<{status: 'ok'|'no-mailbox'|'access-denied'|'error', settings: Object|null, errorCode?: string}>}
 */
async function getUserMailboxSettingsAdmin(accessToken, userId) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}/mailboxSettings`, accessToken);
    return { status: 'ok', settings: response.data };
  } catch (error) {
    const httpStatus = error.response?.status;
    const errorCode = error.response?.data?.error?.code || (httpStatus ? String(httpStatus) : error.message);
    if (httpStatus === 404 || httpStatus === 400 || NO_MAILBOX_ERROR_RE.test(errorCode)) {
      return { status: 'no-mailbox', settings: null, errorCode };
    }
    if (httpStatus === 403 || errorCode === 'ErrorAccessDenied') {
      return { status: 'access-denied', settings: null, errorCode };
    }
    console.error(`Error fetching mailbox settings for ${userId}:`, error.response?.data || error.message);
    return { status: 'error', settings: null, errorCode };
  }
}

/**
 * Update another user's mailbox settings
 * @param {string} accessToken - Valid access token
 * @param {string} userId - User ID
 * @param {Object} payload - Mailbox settings to patch
 * @returns {Promise<void>}
 */
async function updateUserMailboxSettingsAdmin(accessToken, userId, payload) {
  try {
    await graphRequest('patch', `${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}/mailboxSettings`, accessToken, {
      data: payload,
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error(`Error updating mailbox settings for ${userId}:`, error.response?.data || error.message);
    throw graphError('Failed to update mailbox settings', error);
  }
}

/**
 * Encode a SharePoint or OneDrive sharing link as a Graph shareId (u!base64url)
 * @param {string} sharingUrl
 * @returns {string}
 */
function encodeSharingUrl(sharingUrl) {
  const base64 = Buffer.from(String(sharingUrl || '').trim(), 'utf8')
    .toString('base64')
    .replace(/=+$/, '')
    .replace(/\//g, '_')
    .replace(/\+/g, '-');
  return `u!${base64}`;
}

/**
 * Download the raw bytes of a file shared in SharePoint or OneDrive via Graph
 * @param {string} accessToken - Valid access token
 * @param {string} sharingUrl - SharePoint/OneDrive sharing URL
 * @returns {Promise<Buffer>}
 */
async function downloadSharedFile(accessToken, sharingUrl) {
  const shareId = encodeSharingUrl(sharingUrl);
  try {
    const response = await graphRequest(
      'get',
      `${GRAPH_API_BASE}/shares/${encodeURIComponent(shareId)}/driveItem/content`,
      accessToken,
      { responseType: 'arraybuffer', timeout: LARGE_TIMEOUT_MS }
    );
    return Buffer.from(response.data, 'binary');
  } catch (error) {
    console.error('Error downloading shared file from Graph:', error.response?.data || error.message);
    throw graphError('Failed to download shared file', error);
  }
}

const ADMIN_GROUP_LIST_FIELDS = 'id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,visibility,createdDateTime,membershipRule,resourceProvisioningOptions,onPremisesSyncEnabled';
const ADMIN_GROUP_DETAIL_FIELDS = `${ADMIN_GROUP_LIST_FIELDS},mailNickname,classification,expirationDateTime,renewedDateTime,membershipRuleProcessingState,isAssignableToRole,onPremisesLastSyncDateTime`;

/**
 * Every group of the tenant with what the admin list shows
 * @param {string} accessToken - Valid access token
 * @returns {Promise<Array>} Groups sorted by name
 */
async function listAdminGroups(accessToken) {
  try {
    const groups = await getAllPages(`${GRAPH_API_BASE}/groups`, accessToken, {
      params: { $select: ADMIN_GROUP_LIST_FIELDS, $top: 999 },
    });
    return groups.sort((a, b) => (a.displayName || '').localeCompare(b.displayName || '', 'es'));
  } catch (error) {
    console.error('Error listing groups for admin:', error.response?.data || error.message);
    throw graphError('Failed to list groups', error);
  }
}

/**
 * One group with every property the detail page shows
 * @param {string} accessToken - Valid access token
 * @param {string} groupId - Group ID
 * @returns {Promise<Object|null>} The group, or null on 404
 */
async function getAdminGroup(accessToken, groupId) {
  try {
    const response = await graphRequest('get', `${GRAPH_API_BASE}/groups/${encodeURIComponent(groupId)}`, accessToken, {
      params: { $select: ADMIN_GROUP_DETAIL_FIELDS },
    });
    return response.data;
  } catch (error) {
    if (error.response?.status === 404) return null;
    console.error('Error fetching group for admin:', error.response?.data || error.message);
    throw graphError('Failed to fetch group', error);
  }
}

/**
 * Create a Microsoft 365 group or a plain security group, owned by the creator
 * @param {string} accessToken - Valid access token
 * @param {{type: 'unified'|'security', displayName: string, description: string, mailNickname: string, visibility: string}} group
 * @param {string} ownerId - Directory ID of the first owner
 * @returns {Promise<Object>} The created group
 */
async function createGroup(accessToken, group, ownerId) {
  const unified = group.type === 'unified';
  const data = {
    displayName: group.displayName,
    mailNickname: group.mailNickname,
    mailEnabled: unified,
    securityEnabled: !unified,
    groupTypes: unified ? ['Unified'] : [],
    'owners@odata.bind': [`${GRAPH_API_BASE}/users/${encodeURIComponent(ownerId)}`],
  };
  if (group.description) data.description = group.description;
  if (unified) data.visibility = group.visibility;
  try {
    const response = await graphRequest('post', `${GRAPH_API_BASE}/groups`, accessToken, {
      data,
      headers: { 'Content-Type': 'application/json' },
    });
    return response.data;
  } catch (error) {
    console.error('Error creating group:', error.response?.data || error.message);
    throw graphError('Failed to create group', error);
  }
}

/**
 * Change a group's own properties (name, description, visibility)
 * @param {string} accessToken - Valid access token
 * @param {string} groupId - Group ID
 * @param {Object} patch - Only the properties to change
 * @returns {Promise<void>}
 */
async function updateGroup(accessToken, groupId, patch) {
  try {
    await graphRequest('patch', `${GRAPH_API_BASE}/groups/${encodeURIComponent(groupId)}`, accessToken, {
      data: patch,
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    console.error('Error updating group:', error.response?.data || error.message);
    throw graphError('Failed to update group', error);
  }
}

/**
 * Make a user an owner of a group (already being one is fine)
 * @param {string} accessToken - Valid access token
 * @param {string} groupId - Group ID
 * @param {string} userId - User directory ID
 * @returns {Promise<void>}
 */
async function addGroupOwner(accessToken, groupId, userId) {
  try {
    await graphRequest('post', `${GRAPH_API_BASE}/groups/${encodeURIComponent(groupId)}/owners/$ref`, accessToken, {
      data: { '@odata.id': `${GRAPH_API_BASE}/users/${encodeURIComponent(userId)}` },
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (error) {
    const message = error.response?.data?.error?.message || '';
    if (error.response?.status === 400 && message.includes('already exist')) return;
    console.error('Error adding group owner:', error.response?.data || error.message);
    throw graphError('Failed to add group owner', error);
  }
}

/**
 * Take a user off a group's owners (not being one is fine)
 * @param {string} accessToken - Valid access token
 * @param {string} groupId - Group ID
 * @param {string} userId - User directory ID
 * @returns {Promise<void>}
 */
async function removeGroupOwner(accessToken, groupId, userId) {
  try {
    await graphRequest('delete', `${GRAPH_API_BASE}/groups/${encodeURIComponent(groupId)}/owners/${encodeURIComponent(userId)}/$ref`, accessToken);
  } catch (error) {
    if (error.response?.status === 404) return;
    console.error('Error removing group owner:', error.response?.data || error.message);
    throw graphError('Failed to remove group owner', error);
  }
}

module.exports = {
  resetGraphCaches,
  getUserProfile,
  getUserPhoto,
  getUserPhotoById,
  hasUserPhoto,
  getUserGroups,
  getUserManager,
  getAllUserData,
  getAllTenantGroups,
  getGroupsWithMembership,
  getGroupOwners,
  getOwnedGroups,
  addGroupMember,
  getGroupMembers,
  removeGroupMember,
  revokeSignInSessions,
  getCalendarPermissions,
  canOpenMailbox,
  getRooms,
  getSchedules,
  getRoomEvents,
  getMyCalendars,
  getListMembers,
  createCalendarEvent,
  getAuthenticationMethods,
  getUserDevices,
  getAllUsers,
  listAdminGroups,
  getAdminGroup,
  createGroup,
  updateGroup,
  addGroupOwner,
  removeGroupOwner,
  getLicensedUsers,
  getUserById,
  getDirectReports,
  getSubscribedSkus,
  getUsersForReview,
  getMfaRegistration,
  getPresences,
  getMyAppGrants,
  getServicePrincipal,
  getSharedDriveItems,
  getDriveItemPermissions,
  removeDriveItemPermission,
  shareCalendar,
  removeCalendarPermission,
  getUsersWithManager,
  findUserByEmail,
  getUserForReset,
  sendEmail,
  updateUserProfile,
  updateUserPhoto,
  getAccountStatus,
  getUserLicenses,
  getMyGroups,
  getInboxRules,
  getDriveQuota,
  getRecentSignIns,
  getMailboxSettings,
  updateMailboxSettings,
  getPreferredLanguage,
  getMailboxTimeZone,
  getAutomaticReplies,
  updateAutomaticReplies,
  getUserExtendedInfo,
  getJoinedTeamsAndChannels,
  fetchUserAdminRoles,
  getUserAdminRoles,
  resetUserPassword,
  graphBatch,
  listAdminUsers,
  fetchManagersBatch,
  fetchUserGroupsBatch,
  fetchUserLicensesBatch,
  getAdminUser,
  getUserManagerById,
  getUserGroupsById,
  getUserLicensesById,
  createAdminUser,
  updateAdminUser,
  setUserManager,
  removeUserLicenses,
  assignUserLicenses,
  addUserToGroup,
  removeUserFromGroup,
  getUserMailboxSettingsAdmin,
  updateUserMailboxSettingsAdmin,
  encodeSharingUrl,
  downloadSharedFile,
};
