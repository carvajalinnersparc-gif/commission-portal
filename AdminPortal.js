// ============================================================
// Inner SPARC Realty Corporation — Agent Portal
// AdminPortal.gs — Admin panel backend (Agents, User Access)
//
// The Admin panel is scoped to managing sales agents and their portal
// access — not developer records or sales/commission data, which is why
// there is no Developers or Dashboard/Sales surface here. Developer rows
// are still maintained directly in the "Developers" sheet.
//
// Called directly from Portal.html's "Admin" view (visible only when
// action=admin or the logged-in user's role is Admin). Every function
// here re-checks the session server-side via pa_isAdmin_() (Auth.gs)
// before touching any data — the client-side Admin tab is only a
// convenience, never the security boundary.
//
// Agents/Developers/Sales data and Portal user access/roles now live in
// the same spreadsheet (TARGET_SHEET_ID, see Code.gs) — the commission
// data and the Auth Database were merged into one file. This file just
// adds an admin-authorization layer on top of the existing read/write
// functions in those two files — it does not duplicate their logic.
// ============================================================

/** Summary counts for the Admin panel landing state. */
function pa_adminOverview(sessionId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  pa_reconcileAuthTables_(); // Auth.gs — self-heal before counting

  return {
    ok: true,
    counts: {
      agents: getAgents().length,
      users: pa_adminListUsers_().length,
      // Sales Agent names (Sales sheets, column B) with no matching Agent
      // Access alias — i.e. commission rows currently invisible to any
      // agent because of a name mismatch. See PortalData.gs's
      // pa_findUnmatchedSalesAgentNames_() for the full breakdown; call
      // pa_adminFindUnmatchedSalesAgentNames(sessionId) to fetch it.
      unmatchedSalesAgentNames: pa_findUnmatchedSalesAgentNames_().length
    }
  };
}

// ── Agents ───────────────────────────────────────────────────
function pa_adminGetAgents(sessionId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;
  pa_reconcileAuthTables_(); // Auth.gs — self-heal before listing (e.g. a
                             // manually deleted Agent Access row removes
                             // the matching roster row before it's shown)
  return { ok: true, agents: getAgents() };
}

function pa_adminAddAgent(sessionId, data) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;
  if (!data || !String(data.fullName || '').trim()) {
    return { ok: false, error: 'Full Name is required.' };
  }
  const result = addAgent(data);
  return Object.assign({ ok: result.success === true }, result);
}

function pa_adminUpdateAgent(sessionId, data) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;
  if (!data || !data.originalId) {
    return { ok: false, error: 'Missing Agent ID.' };
  }
  const result = updateAgent(data);
  return Object.assign({ ok: result.success === true }, result);
}

function pa_adminDeleteAgent(sessionId, agentId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;
  const result = deleteAgent(agentId);
  return Object.assign({ ok: result.success === true }, result);
}

// ── Users (portal access) ───────────────────────────────────
// Most Agent Access CRUD (list/add/update/delete) lives in Auth.gs
// alongside the rest of the session/credentials machinery it shares
// helpers with (pa_adminListUsers, pa_adminAddUser, pa_adminUpdateUser,
// pa_adminDeleteUser). pa_adminSetUserPassword below is grouped here
// instead, next to the rest of the Admin-panel-only actions.

/** Admin: sets a portal user's password directly, without an email
 *  round-trip. This is the Admin-panel counterpart to the self-service
 *  "Set Password" / "Forgot Password?" links (pa_requestPasswordAction,
 *  Auth.gs) — useful when a user can't access their inbox yet, or when
 *  an admin wants to hand someone a password in person. Takes effect
 *  immediately: no token, no email, and it silently overwrites whatever
 *  password (if any) the user had before. The new password is stored
 *  through the exact same salted-hash path (pa_setPasswordHash_,
 *  Auth.gs) as a self-service set/reset, so there is no way to tell
 *  afterward which route was used to set it. */
function pa_adminSetUserPassword(sessionId, email, newPassword) {
  const auth = pa_isAdmin_(sessionId); // Auth.gs
  if (!auth.ok) return auth;

  const target = pa_normEmail_(email); // Auth.gs
  if (!target) return { ok: false, error: 'Email is required.' };
  if (!pa_getAccessRecord_(target)) return { ok: false, error: 'User not found.' }; // Auth.gs
  if (!newPassword || newPassword.length < 8) {
    return { ok: false, error: 'Password must be at least 8 characters.' };
  }

  pa_setPasswordHash_(target, newPassword); // Auth.gs
  return { ok: true };
}