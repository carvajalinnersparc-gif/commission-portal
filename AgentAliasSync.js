// ============================================================
// AgentAliasSync.gs — imports Canonical Name / Aliases from the CMS
// spreadsheet's "Agent Aliases" tab into the "Agent Access" tab, and keeps
// them in sync on a schedule.
//
//   SOURCE  CMS_SHEET_ID (Code.gs) → tab "Agent Aliases"
//           columns (found by header name): Canonical Name | Aliases | Email
//   TARGET  TARGET_SHEET_ID (Code.gs) → tab "Agent Access"
//           C Canonical Name, G Aliases  (matched to a row by Email, B)
//
// Rules
//   • Rows are matched by Email. If the source row has NO email, it falls
//     back to a unique Canonical-Name/alias match.
//   • It never creates users, roles or logins. A source row with no
//     matching Agent Access row is only reported in the log (adding login
//     access is a deliberate admin action).
//   • Never touches User ID, Email, Role, Active or Status (F).
//   • Aliases are separated by semicolons (";") in the source, same as in
//     Agent Access, because names such as "FLORES, JULIE FE" contain commas.
//
// SETUP (once)
//   1. Add this file to the Apps Script project (needs Auth.gs updated for
//      the Status/Aliases column layout: Aliases = G).
//   2. Run syncAgentAliasesFromCms() once by hand and check View > Logs.
//   3. Run installAgentAliasSyncTrigger() once to keep it updated
//      automatically (default: every hour). Run
//      removeAgentAliasSyncTrigger() to stop it.
// ============================================================

const AGENT_ALIAS_SOURCE_SHEET_ = "Agent Aliases";

// "merge"  → adds any new names from the source; never removes an alias that
//            is already in Agent Access (admin-panel edits are safe).
// "mirror" → Agent Access Canonical Name + Aliases become exactly what the
//            source says (names removed from the source are removed here too).
const AGENT_ALIAS_SYNC_MODE_ = "merge";

const AGENT_ALIAS_SYNC_EVERY_HOURS_ = 1; // 1, 2, 4, 6, 8 or 12

/** Entry point — run manually, or via the time trigger below. */
function syncAgentAliasesFromCms() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('syncAgentAliasesFromCms: another run is in progress — skipped.');
    return { ok: false, error: 'Another sync is already running.' };
  }
  try {
    const res = pa_syncAgentAliasesFromCms_();
    Logger.log('syncAgentAliasesFromCms: ' + JSON.stringify(res, null, 2));
    return res;
  } finally {
    lock.releaseLock();
  }
}

/** Admin-panel hook (no UI wired yet): same sync, admin session required. */
function pa_adminSyncAgentAliases(sessionId) {
  const auth = pa_isAdmin_(sessionId); // Auth.gs
  if (!auth.ok) return auth;
  return syncAgentAliasesFromCms();
}

function pa_syncAgentAliasesFromCms_() {
  // ── 1. Read the source tab ────────────────────────────────
  const src = getCmsSpreadsheet().getSheetByName(AGENT_ALIAS_SOURCE_SHEET_); // Code.gs
  if (!src) return { ok: false, error: 'Tab "' + AGENT_ALIAS_SOURCE_SHEET_ + '" not found in the CMS spreadsheet.' };

  const srcLastRow = src.getLastRow();
  const srcLastCol = Math.max(src.getLastColumn(), 3);
  if (srcLastRow < 2) return { ok: true, note: 'Source tab has no data rows.' };

  const hdr = src.getRange(1, 1, 1, srcLastCol).getValues()[0].map(h => String(h || '').trim().toLowerCase());
  const iCanon = hdr.findIndex(h => h.indexOf('canonical') === 0);
  const iAlias = hdr.findIndex(h => h.indexOf('alias') === 0);
  const iEmail = hdr.findIndex(h => h.indexOf('email') === 0);
  if (iCanon === -1 || iAlias === -1 || iEmail === -1) {
    return { ok: false, error: 'Source headers must include "Canonical Name", "Aliases" and "Email". Found: ' + hdr.join(' | ') };
  }

  // Collapse duplicate emails in the source into one entry (names are unioned).
  const entries = [];      // {email, canonical, aliases:[...]}
  const entryByEmail = {};
  let blankRows = 0;
  src.getRange(2, 1, srcLastRow - 1, srcLastCol).getValues().forEach(r => {
    const email = pa_normEmail_(r[iEmail]);
    const canonical = String(r[iCanon] || '').replace(/\s+/g, ' ').trim();
    const aliases = pa_parseAliasesCsv_(r[iAlias]); // Auth.gs — splits on ';'
    if (!email && !canonical && !aliases.length) { blankRows++; return; }
    if (email && entryByEmail[email]) {
      const e = entryByEmail[email];
      if (!e.canonical) e.canonical = canonical;
      e.aliases = e.aliases.concat(aliases);
      return;
    }
    const entry = { email: email, canonical: canonical, aliases: aliases };
    entries.push(entry);
    if (email) entryByEmail[email] = entry;
  });

  // ── 2. Read Agent Access ──────────────────────────────────
  const sh = pa_agentAccessSheet_(); // Auth.gs — also guarantees the Aliases header (G)
  const last = sh.getLastRow();
  if (last < 2) return { ok: true, note: 'Agent Access has no rows to update.' };

  const C = AA_COL_.CANONICAL, G = AA_COL_.ALIASES, EMAIL = AA_COL_.EMAIL;
  const rows = sh.getRange(2, 1, last - 1, G + 1).getValues();

  const rowByEmail = {};
  const rowsByName = {};
  rows.forEach((r, i) => {
    const em = pa_normEmail_(r[EMAIL]);
    if (em) rowByEmail[em] = i;
    pa_buildAccessNames_(r[C], r[G]).aliases.forEach(n => {
      const k = pa_nameKey_(n);
      (rowsByName[k] = rowsByName[k] || []).push(i);
    });
  });

  // ── 3. Compute updates ────────────────────────────────────
  const mirror = AGENT_ALIAS_SYNC_MODE_ === 'mirror';
  const report = {
    ok: true,
    mode: AGENT_ALIAS_SYNC_MODE_,
    sourceRows: entries.length,
    blankSourceRows: blankRows,
    matchedByEmail: 0,
    matchedByName: 0,
    updated: 0,
    unchanged: 0,
    notInAgentAccess: [],
    ambiguousNoEmail: [],
    canonicalDiffers: []
  };
  const changedRows = {}; // row index -> true

  entries.forEach(e => {
    let idx;
    if (e.email) {
      idx = rowByEmail[e.email];
      if (idx === undefined) { report.notInAgentAccess.push(e.email + (e.canonical ? ' (' + e.canonical + ')' : '')); return; }
      report.matchedByEmail++;
    } else {
      const hits = rowsByName[pa_nameKey_(e.canonical)] || [];
      const uniq = hits.filter((v, k) => hits.indexOf(v) === k);
      if (uniq.length === 1) { idx = uniq[0]; report.matchedByName++; }
      else { report.ambiguousNoEmail.push(e.canonical || '(blank name)'); return; }
    }

    const r = rows[idx];
    const curCanon = String(r[C] || '').trim();
    const curAlias = String(r[G] || '').trim();
    const srcNames = pa_buildAccessNames_(e.canonical, e.aliases.join('; '));

    let names;
    if (mirror) {
      // Source is the truth: its canonical name + its aliases, nothing else.
      names = pa_buildAccessNames_(srcNames.canonicalName || curCanon, srcNames.extraAliases.join('; '));
    } else {
      // Additive: keep our canonical name (unless blank) and every alias we
      // already have; add the source's canonical name + aliases on top.
      const keep = curCanon || srcNames.canonicalName;
      names = pa_buildAccessNames_(keep, [curAlias, srcNames.aliases.join('; ')].filter(Boolean).join('; '));
      const curFirst = pa_buildAccessNames_(curCanon, '').canonicalName; // first name only (handles un-migrated "A; B" cells)
      if (curFirst && srcNames.canonicalName && pa_nameKey_(curFirst) !== pa_nameKey_(srcNames.canonicalName)) {
        report.canonicalDiffers.push(e.email + ': Agent Access="' + curCanon + '" vs CMS="' + srcNames.canonicalName + '" (CMS name added as an alias)');
      }
    }

    const newCanon = names.canonicalName;
    const newAlias = names.extraAliases.join('; ');
    if (newCanon === curCanon && newAlias === curAlias) { report.unchanged++; return; }

    r[C] = newCanon;
    r[G] = newAlias;
    changedRows[idx] = true;
    report.updated++;
  });

  // ── 4. Write only the changed cells (C and G) ─────────────
  Object.keys(changedRows).forEach(k => {
    const i = Number(k);
    sh.getRange(i + 2, C + 1).setValue(rows[i][C]);
    sh.getRange(i + 2, G + 1).setValue(rows[i][G]);
  });

  return report;
}

// ── Scheduling ──────────────────────────────────────────────
/** Run ONCE from the editor: syncs automatically every AGENT_ALIAS_SYNC_EVERY_HOURS_ hour(s). */
function installAgentAliasSyncTrigger() {
  removeAgentAliasSyncTrigger();
  ScriptApp.newTrigger('syncAgentAliasesFromCms')
    .timeBased()
    .everyHours(AGENT_ALIAS_SYNC_EVERY_HOURS_)
    .create();
  Logger.log('Agent alias sync scheduled: every ' + AGENT_ALIAS_SYNC_EVERY_HOURS_ + ' hour(s).');
}

function removeAgentAliasSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncAgentAliasesFromCms') ScriptApp.deleteTrigger(t);
  });
}