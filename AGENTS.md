# AGENTS.md

Internal Microsoft 365 self-service portal for the Ayuntamiento de Aranda de Duero: Express 5 + EJS, signed in with Entra ID (MSAL), data from Microsoft Graph. UI text and the README are Spanish; code, comments, logs and this file are English.

## Never touch `.env`

`.env` holds the real credentials and is not in git, so nothing can restore it. Do not write, overwrite, move, delete or regenerate it, and do not read or print its values — including indirectly (redirection, `cp`/`mv`, scripts, test setup, temporary swaps). Pass alternative envs via CLI vars or a file outside the project.

- Tests never read it: `test/helpers/env.js` sets `NODE_ENV=test`, and the app skips `dotenv` in that mode. Keep it that way; `test/routes.test.js` fails if `.env` loads during tests.
- When running commands against a copy of the project, use absolute paths for writes and never assume `cd` succeeded.
- New settings go in `.env.example` and the README; ask the user to add them to their `.env`.

## Commands

```bash
npm run dev                                       # nodemon, NODE_ENV=development
npm start                                         # node server.js
npm test                                          # NODE_ENV=test ava
NODE_ENV=test npx ava test/graphService.test.js   # one file
NODE_ENV=test npx ava --match='*reset link*'      # by title
```

No linter, formatter, build step or coverage report. Node >= 25 (`engines`, Dockerfile, CI). CommonJS throughout.

## Layout

- `server.js` — builds and exports the app; only listens when run directly. Middleware order: `securityHeaders` → body parsers → static → file-backed session → `refreshAccessToken` → `checkAdminStatus` → routers → 404 → `expiredSessionHandler` → error handler. Refuses to start in production without `SESSION_SECRET`.
- `middleware/securityHeaders.js` & `public/vendor/` — CSP (`default-src 'self'`, `script-src 'self'`, `frame-ancestors 'none'`, `frame-src 'none'`; `img-src` also allows the origin of `SIGNATURE_IMAGES_URL`, for the signature logos only), `nosniff`, `Referrer-Policy`, HSTS in production. **No inline JS** (page scripts in `public/js/<page>.js`, no `onclick=`, confirmations use `data-confirm` + `/js/confirm.js`, mutating `fetch` uses `<meta name="csrf-token">` + `/js/csrf.js`). **Pages load nothing from external servers** (no CDN, web fonts, or image services); vendor libs (Bootstrap, Outfit, d3) live in `public/vendor/` with version in filename and are documented in `public/vendor/README.md`. `test/routes.test.js` enforces both.
- **Branding & menu:** `APP_NAME`, `ORG_NAME`, `INTRANET_URL` (`config/appConfig.js`) are exposed to all views as `appName`, `orgName`, `intranetUrl` — never hardcode them in views. The nav menu lives only in `views/partials/header.ejs` (`server.js` injects `currentUser`, `currentPath`, `canPublish`, `hasSharedMailboxes`; `checkAdminStatus` sets `isAdmin`). Pages set `currentPage` (`profile`, `account`, `directory`, `organigrama`, `groups`, `teams`, `rooms`, `blog`, `help`, `admin`; every page of Administración uses `admin`). Sign-out lives only in the user menu.
- `routes/auth.js` — `/auth/signin`, `/auth/callback`, `/auth/logout`.
- `routes/index.js` — profile, groups, teams, `/organigrama` (`services/orgChart.js` builds one tree under `ORG_CHART_ROOT` + unassigned list; `public/js/organigrama.js` renders with d3-org-chart; plain-text initials, no avatar service), `/api/*`, and public password reset.
- `routes/directory.js` — `/directorio`, CSV export, `/directorio/:id` (contact card + local QR via `qrcode`), `/directorio/:id/vcard` (`services/vcard.js`).
- `routes/account.js` — `/account`, `/account/calendars` (calendar sharing + calendars shared with me), `/account/groups`, `/account/sign-ins` (+ `/report`), `/account/mailboxes`, `/profile/me/auto-reply`, `/profile/me/photo`. Logic in `services/accountInsights.js` (pure) and `config/licenseCatalog.js`.
- `routes/activity.js` — `/account/activity`: user's own audit entries labeled by `config/auditEvents.js` (add a label for every new audit event).
- `routes/apps.js` — `/account/apps`: delegated app grants described by `services/appGrants.js` (pure); read-only (revoking links to Microsoft).
- `routes/sharing.js` — `/account/sharing`: OneDrive delta walk (capped at 25 pages / 150 shared items), permissions classified by `services/sharing.js` (pure); shows only anonymous links and external people, re-reads permissions before deletion. **A person with no address is not an outsider** (same rule as calendar sharing: no false alarms).
- `routes/ownedGroups.js` & `routes/groupRequests.js` — `/groups/owned` (owner member management) and group join requests stored in `services/requestStore.js` (single JSON file in `DATA_DIR`).
- `routes/rooms.js` — `/salas`: meeting rooms (client-sorted by `public/js/rooms.js`) and free/busy search; booking links to Outlook.
- `routes/team.js` — `/equipo`: `/me/directReports` and four IT request forms defined in `config/teamRequests.js` (sent as plain-text email from user's mailbox; audit logs only the kind).
- `routes/blog.js` — `/blog`: emails a post to Jetpack Post-by-Email through the signed-in user's mailbox (`services/wpService.js` fetches WordPress categories).
- `routes/help.js` — `/ayuda` + contact form; all copy lives in `config/helpGuide.js` and `config/helpTopics.js`.
- `routes/signature.js` — `/firma`: the signature is generated server-side from the profile (`services/signature.js`, pure; template mirrors `buildSignatureHTML` of `arandadeduero/firmas-correo`). Nothing is typed; the only input is the concejalía (`config/signatureConcejalias.js`, matched from `department`, `?concejalia=` overrides). Preview and copied HTML load the logos from `SIGNATURE_IMAGES_URL` (GitHub Pages) on purpose: recipients cannot reach the portal, and the preview then shows what they will get. This is the one external load the pages make. No signature unless name, position, email and concejalía are all known.
- `routes/debug.js` — `/debug` («Datos técnicos»): raw `/me` and session account, never tokens.
- `routes/adminUsers.js` — `/admin/usuarios/*` (port of `arandadeduero/m365-tools`; `/admin/usuarios/atencion` is the accounts-to-review page (app token + `services/accountReview.js`; uncomputable lists are `null`, shown as "cannot know"); `/admin/cuentas` only redirects to it). Uses the **admin's delegated token** (Entra enforces their role), not the app token, except sign-in activity. Pure logic in `services/adminUsersExcel.js` (xlsx via `exceljs`, sheets `ACTIVOS`/`CONCEJALES`, UTC dates) and `services/adminUsersTools.js`. Excel uploads are multipart (`MULTIPART_ROUTES` + `verifyCsrfToken`); imports are previewed, kept as an in-memory plan bound to the admin for 15 min, then applied. Disabling departures removes licenses first and disables only if that succeeds. `GET .../mailboxSettings` must never carry `$select`. Generated passwords are shown once (session flash) and never audited. Parameterised `/:id` routes must stay after the literal ones.
- `routes/adminGroups.js` — `/admin/grupos` (list + filters, `/nuevo`, `/:id` detail, edit, owners, members). Delegated admin token only; creates Microsoft 365 and plain security groups (creator bound as owner via `owners@odata.bind`). Pure logic in `services/adminGroups.js`: `describeAdminGroup` marks dynamic, on-premises-synced and Exchange groups read-only (`manageable: false`, every POST refuses them). The mail alias of a new group must start with `gr-` (`GROUP_PREFIX`; the name is free). `validateNewGroup` checks it server-side and the form by `pattern`; an empty alias is made as `gr-<name>`. **No dynamic groups here: the tenant has no Entra ID P1.** Logs never include IDs or names (CodeQL log-injection). Literal routes (`/nuevo`) stay before `/:id`.
- `routes/admin.js` — `/admin/*` (`res.locals.isAdmin`): `auditoria` (+ CSV; `/admin/audit` redirects, keeping the query), `licencias` (`/subscribedSkus`), `avisos` (`services/messageCenterService.js`, app token, cached), `solicitudes` (`requestStore.listPending`), `salas` (one room's week via app token; room MUST be in `getRooms()` since the app token can read any mailbox; unreadable calendar is `null`, never "free"), `comisiones` (`config/commissions.js`; creates event in admin's own calendar found by name with delegated token; re-reads invitees/rooms on POST and aborts if a list cannot be read), `ajustes` (`services/settingsStore.js`: `DATA_DIR/settings.json` overrides `config/local.js` blog list and shared mailboxes in place at start and on save).
- Services & config: `services/graphService.js` (all Graph calls), `services/tokenService.js` (MSAL), `services/resetTokenService.js` (reset link tokens), `services/emailService.js` (AWS SES, reset emails only), `services/regionalSettings.js` (enforces Madrid time + Spanish at `/auth/callback`), `services/serviceHealthService.js` (M365 incidents banner, app token, cached), `config/permissions.js` (domain and per-route allow-lists, overridden by git-ignored `config/local.js`, template `config/local.example.js`).

## Things that are easy to get wrong

- **Two kinds of token.** Signed-in routes use `req.session.accessToken` (delegated, `scopes` in `config/authConfig.js`). Background/public flows (password reset, sign-in log, incidents, admin reviews/rooms) use `getAppToken()` (application permissions in Entra).
- **Token renewal lives in the session.** At sign-in the MSAL cache (with refresh token) is serialized into `req.session.tokenCache`. `refreshAccessToken` renews when < 5 min remain using a fresh MSAL client loaded from that cache — never pass user tokens through the shared client in `tokenService` (its cache is process-wide). On failure, regenerate (empty) the session rather than destroying it so later middleware can still read `req.session`.
- **Graph errors carry `graphStatus`.** Service functions throw generic `Failed to ...` errors built via `graphError()` preserving Graph's HTTP status; `expiredSessionHandler` turns 401 into a redirect to `/login`. Copy `graphStatus` if wrapping these errors.
- **Adding a delegated scope** requires updating `scopes` in `config/authConfig.js` and granting it in Entra; existing sessions keep old consent until re-sign-in.
- **Service functions fail in two ways.** Optional lookups (`getUserPhoto`, `getUserManager`, `getUserGroups`, `getAllUsers`, `getUserAdminRoles`) swallow errors and return `null`/`[]`; required ones throw. Check before relying on a throw.
- **Graph lists are paged.** Always use `getAllPages()` for collections; a single `axios.get` silently truncates at 100 items.
- **`checkPermission('/blog')` takes its key explicitly, never from `req.path`.** Express routes `/Blog` and `/blog/` to `/blog`, so `req.path` lookup bypassed the allow-list. An unconfigured key denies everyone (`/blog` defaults to empty until configured in `config/local.js` or `/admin/ajustes`). `MULTIPART_ROUTES` in `middleware/csrf.js` normalises paths for the same reason.
- **OAuth callback validates session `state`.** `/auth/signin` stores random `oauthState`; `/auth/callback` rejects mismatches (400, login CSRF), deletes it, and regenerates the session before saving tokens.
- **Profile edits send a per-field diff.** `POST /profile/me/edit` diffs against current profile, sends emptied boxes as `null` (Graph rejects `''`), and issues one `PATCH` per field so one 403 (e.g., admin phone) doesn't fail the rest; refused fields are stored in `req.session.profileRefused` and shown with a notice.
- **Locked profile fields are skipped on POST.** Fields absent from `PROFILE_EDITABLE_FIELDS` or in `req.session.profileRefused` render as disabled inputs (not posted by browsers); reading absence as emptied would clear them. `lockedProfileFields()` decides for both view and handler.
- **Read mutable `config/appConfig.js` properties at request time** (`appConfig.SHARED_MAILBOXES`, `appConfig.PROFILE_EDITABLE_FIELDS`, `appConfig.ADMIN_EMAIL`) instead of destructuring at module load, so tests can override them.
- **d3-org-chart only flags nodes.** `setCentered`, `setHighlighted`, `setExpanded` require `render()` to draw; `_expanded` means the node itself is visible, not its children. Escape card names with `escapeHtml`.
- **`checkAdminStatus` caches roles in `req.session.adminStatus` for 5 minutes** on regular pages (skipping cache on failed lookups) and always verifies live against Graph on `/admin/*` routes.
- **Errors:** Pages call `next(err)` with `err.status` (`views/error.ejs`). Messages are hidden in production unless `err.expose` is set (see `userError()` in `routes/index.js`). `/api/*` returns `{ success: false, error }` JSON; `/blog` POST and permission denials return plain text.
- **Env vars read at require time:** `ADMIN_EMAIL` and `REDIRECT_URI` (`routes/index.js`), `RESET_RATE_LIMIT` (`middleware/rateLimit.js`). `TOKEN_SECRET` is read on use so the app boots without it and only password reset fails.

## CSRF and the audit log: rules for new routes

- **Every form needs `<input type="hidden" name="_csrf" value="<%= csrfToken() %>">`** and every non-GET `fetch` needs the `X-CSRF-Token` header; global `csrfProtection` rejects anything else with 403.
- **Multipart uploads** cannot be parsed before multer: list the route in `MULTIPART_ROUTES` (`middleware/csrf.js`) *and* call `verifyCsrfToken` immediately after multer (see `/blog`, `/profile/me/photo`). Listing without calling leaves it unprotected.
- **`csrfToken()` starts a session.** Only call it in views with forms so anonymous GETs don't create session files (`/reset-password` is unavoidable).
- **Audit mutations** via `auditRequest(req, 'area.action', details)` after success. Log what changed, never secrets/content (no passwords, tokens, reset links, message bodies). `audit()` never throws.

## Account pages

- **"Could not check" is not "has none".** Sign-in methods, inbox rules, and drive quota return `null` when Graph refuses; views state this rather than showing an empty list, so failures never trigger false "no MFA" warnings or hide forwarding rules.
- **Warnings err towards silence.** Unknown sign-in methods count as MFA, and only *enabled* rules count as external forwarding — avoid alert fatigue.
- **Sign-in log uses the app token**, so the target user MUST come strictly from `req.session.account.localAccountId`, never from request parameters.
- **Reporting a sign-in ("No fui yo") order matters:** re-read the entry by ID from the log (never trust form details) → email IT using the user's token → revoke sessions → write audit entry → regenerate session. The result page reports which steps actually succeeded.
- **Shared mailboxes** cannot be listed by Graph. `canOpenMailbox` probes each inbox in `sharedMailboxes` with the user's token: 200 = yes, 403/404 = no (missing `Mail.Read.Shared` also returns 403), anything else = unchecked. Inaccessible mailboxes are never shown.
- **Rooms have three states:** free only if `availabilityView` is all zeros; any `error` or missing schedule is "sin comprobar" (still links to Outlook).
- **Regional enforcement never blocks sign-in.** `enforceRegionalSettings` catches all errors, preserves working days/hours, and only fixes time zone/language.
- **Licenses & copy:** Add new SKUs/plans in `config/licenseCatalog.js` (unknown SKUs show their code; unknown plans are omitted). Keep short plain-Spanish explanations of M365 concepts on pages.

## Group requests and the status banner

- **Owner actions & approvals** always use the owner's delegated token (never `getAppToken()`) and verify ownership server-side on every request (`loadOwnedGroup` for member reads/changes; `loadRequestToResolve` for pending requests). Validate directory IDs before interpolating into Graph URLs.
- **Exactly one owner is the warning** (not `< 2`), because `getGroupOwners` returns `[]` on error.
- **Manageable groups only** (`isManageable`: M365 groups and non-mail-enabled security groups). Others or ownerless groups fall back to emailing `ADMIN_EMAIL`.
- **`requestStore` is synchronous on purpose** (read+write with no `await` prevents in-process races; not safe for multiple instances).
- **Status banner (`getServiceIssues()`)** never rejects, has a 3s timeout, and caches failures for 1 hour so missing permissions aren't retried every page.

## Password reset, and why it is built this way

- `POST /reset-password` responds immediately and runs user lookup + email in the background so timing and response are identical for known/unknown addresses. Never `await` it or expose failures.
- Reset URLs use `REDIRECT_URI`, never the untrusted `Host` header.
- Tokens are stateless AES-256-GCM embedding `lastPasswordChangeDateTime`; `checkResetToken` verifies it against Graph on every use so links die on password change (Graph lag leaves a brief theoretical window).
- Both steps are rate-limited per IP in memory (`TRUST_PROXY` required behind a proxy).
- Reset emails go to the first `otherMails` address, falling back to the account's own mailbox. The form shows only the account being reset, never the destination address.

## Not verified against a real tenant

Everything is tested against stubs. Verify these once against the real tenant:

- Org chart: `/users?$expand=manager($select=id)` returning all users with manager.
- Password reset: app permissions reading `lastPasswordChangeDateTime` and seeing it update on reset (if unreadable, `sendResetLink` refuses and logs `reset.link_failed`).
- Groups & owners: `/groups/{id}/owners`, `/me/ownedObjects/microsoft.graph.group`, `/groups/{id}/members/microsoft.graph.user` (`userType`), owner add/remove with delegated `GroupMember.ReadWrite.All`, "already exist" error wording, `/me/memberOf/microsoft.graph.group` (`resourceProvisioningOptions`).
- Account security: `/me/authentication/methods` (`UserAuthenticationMethod.Read`), `/me/registeredDevices` & `/me/ownedDevices` (`microsoft.graph.device`), `/me/mailFolders/inbox/messageRules`, `/me/drive?$select=quota` (`Files.Read`), `/auditLogs/signIns` filtered by `userId` (app `AuditLog.Read.All`, needs Entra ID P1/P2) returning entry `id`, `POST /me/revokeSignInSessions` (delegated `User.RevokeSessions.All`) for non-admins.
- Calendar sharing: `/me/calendar/calendarPermissions` (`Calendars.ReadWrite`), org-wide entry having no address, `POST` `{ emailAddress: { address }, role }` and `DELETE .../{id}` for normal users.
- Regional & profile: `PATCH /me/mailboxSettings` accepting `timeZone` + `language` + `workingHours` together, normal users updating `preferredLanguage`, auto-reply dates with `Prefer: outlook.timezone`, `PUT /me/photo/$value` (`User.ReadWrite`), `MailboxSettings.ReadWrite` consent.
- Service health & admin: `/admin/serviceAnnouncement/issues?$filter=isResolved eq false` (app `ServiceHealth.Read.All`), `/admin/serviceAnnouncement/messages` (app `ServiceMessage.Read.All`), `/subscribedSkus` (delegated `Directory.Read.All`), `/users?$select=...,signInActivity` (app `AuditLog.Read.All`, P1), `/reports/authenticationMethods/userRegistrationDetails` (app `UserAuthenticationMethod.Read.All`), `config/licenseCatalog.js` plan names vs real licenses.
- Rooms & commissions: `/places/microsoft.graph.room` (`Place.Read.All`), `getSchedule` with 20 rooms (`Calendars.Read`), Outlook compose deep link (`/calendar/deeplink/compose`), `/users/{room}/calendarView` (app `Calendars.Read`, `Prefer: outlook.timezone`, subject/attendee retention), `/groups?$filter=mail eq '...'` finding Exchange distribution lists, `/groups/{id}/transitiveMembers/microsoft.graph.user` (`ConsistencyLevel: eventual`), `POST /me/calendars/{id}/events` (`Calendars.ReadWrite`) inviting a room as `resource`.
- Apps, sharing, presence, mailboxes: `/me/oauth2PermissionGrants` (delegated `Directory.Read.All`) + `/servicePrincipals/{id}`, `/me/drive/root/delta` (`shared`, `deleted`) + `/me/drive/items/{id}/permissions` (`link.scope`, `grantedToIdentitiesV2`) + `DELETE` (delegated `Files.ReadWrite`), `POST /communications/getPresencesByUserId` (delegated `Presence.Read.All`, admin consent, up to 650 IDs), `/users/{address}/mailFolders/inbox` (`Mail.Read.Shared`, 200 vs 403/404).
- Silent token renewal with real refresh tokens (MSAL mechanics tested in `test/tokenService.test.js` via fake endpoint).

## Tests

Ava + supertest, one worker per file; nothing may reach the network:

- `test/helpers/env.js` — require first in every test file: sets `NODE_ENV=test`, pins env vars, points `AUDIT_DIR`/`DATA_DIR` to per-file temp dirs, and sets `RESET_RATE_LIMIT` high (only `test/rate-limit.test.js` lowers it).
- `test/helpers/stubs.js` — `stubAxios()` (matches method + URL suffix, records calls, unmatched → 404), `stubMsal()` (controls token lifetime/renewal failure and `msal.account` to switch users; Graph stub answers are test-controlled vars), `stubSes()`, `waitFor()`, `getCsrfToken(agent)`, `auditEvents()`. Create stubs **before** requiring `../server` (`wpService` calls `axios.create` at load).
- Tests sharing a stub use `test.serial` and reset it in `beforeEach` (also call `resetGraphCaches()`, `resetCategoryCache()`, `resetMessageCache()` where those endpoints are stubbed).
- **Sessions & CSRF:** Sign in via `signInAgent(agent)` (runs real `/auth/signin` + `/auth/callback` with stubbed MSAL) using a supertest agent. POSTs require a cookie + CSRF token (`getCsrfToken(agent)` as header or `_csrf`). Log agents out (`/auth/logout`) after fetching pages with forms so session files aren't left in `sessions/`. Use a dedicated session (`freshSession` in `test/authenticated.test.js`) for profile-refusal tests.
- **Call counting:** Sign-in triggers regional-settings calls (clear `http.calls` after sign-in or stub already-correct settings; see `test/csrf.test.js`, `test/token-refresh.test.js`). Incident lookup runs once per process on the first signed-in page — filter out `/serviceAnnouncement/`.
- `test/tokenService.test.js` runs real MSAL against a fake endpoint via `msalConfig.system.networkClient`.
- Tests relying on permissions must set `permissions.DOMAIN` / `permissions['/blog']` explicitly so local `config/local.js` doesn't affect them.

## Deployment

`Dockerfile` (multi-stage, `node:25-alpine`) builds and pushes to GHCR via `.github/workflows/docker-publish.yml` on push to `main`; `.github/workflows/test.yml` runs `npm test`. Persist `sessions/`, `audit/` (`AUDIT_DIR`), and `data/` (`DATA_DIR`) as volumes across container replacements. Run a single instance: sessions, in-memory rate limits, and `requestStore` are all per-process or per-directory.
