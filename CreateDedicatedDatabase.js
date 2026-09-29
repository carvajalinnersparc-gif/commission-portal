// ============================================================
// Inner SPARC Realty Corporation — Agent Portal
// CreateDedicatedDatabase.gs
//
// The Auth Database (Agent Access, Credentials, Login Tokens, Sessions)
// used to live in its own dedicated spreadsheet. It has since been
// merged into the same spreadsheet as the commission data
// (TARGET_SHEET_ID, see Code.gs) — Auth.gs's pa_authSpreadsheet_() now
// just calls getTargetSpreadsheet() instead of opening a separate file.
//
// USE mergeAuthTabsIntoTargetSpreadsheet() BELOW for that merged setup:
// it copies whichever of the four auth tabs aren't already present in
// TARGET_SHEET_ID over from the old Auth Database spreadsheet
// (OLD_AUTH_SHEET_ID, below) — real data and formatting included, not
// just empty placeholder tabs. It skips tabs that already exist in
// TARGET_SHEET_ID, so it's safe to run more than once — e.g. if you've
// already hand-copied some tabs over yourself.
//
// HOW TO RUN (first-time setup / recovering real data):
//   1. Make sure TARGET_SHEET_ID at the top of Code.gs points at the
//      spreadsheet you want everything to live in.
//   2. Select "mergeAuthTabsIntoTargetSpreadsheet" in the toolbar
//      dropdown and click Run.
//   3. Approve the permissions prompt if asked.
//   4. Check View > Logs — it prints which tabs it copied real data
//      into vs. built fresh vs. already had.
//
// HOW TO RUN (schema repair — missing tab or column later on):
//   Select "ensureAuthTablesAndColumns" in the toolbar dropdown and
//   click Run. See the function's own comment below for details on
//   exactly what it checks and fixes.
//
// createAuthDatabase() below (which builds a brand-new, separate
// spreadsheet) is kept only for reference/rollback — it is not part of
// the current merged setup and nothing in Auth.gs calls it anymore.
// ============================================================

// The old, separate Auth Database spreadsheet that the four auth tabs
// used to live in, before everything was consolidated into
// TARGET_SHEET_ID (Code.gs). mergeAuthTabsIntoTargetSpreadsheet() below
// copies the real data out of this spreadsheet — it does NOT rebuild
// tabs from scratch, so nothing gets lost or reset to sample rows.
const OLD_AUTH_SHEET_ID = "1gvq8IKtgXcXQ_1vhAfdF8DBKgjTmO1QgYfwpKKe4iHU";

/** Idempotent: copies any of the four auth tabs that are missing from
 *  TARGET_SHEET_ID (Code.gs) over from OLD_AUTH_SHEET_ID, preserving
 *  their real data and formatting. Tabs that already exist in
 *  TARGET_SHEET_ID — including any you already copied over by hand —
 *  are left untouched and are not overwritten. If a tab is missing from
 *  OLD_AUTH_SHEET_ID too, it falls back to building a fresh empty tab
 *  (via the same builders CreateDedicatedDatabase used to use) so setup
 *  still completes. */
function mergeAuthTabsIntoTargetSpreadsheet() {
  const target = getTargetSpreadsheet(); // defined in Code.gs
  const source = SpreadsheetApp.openById(OLD_AUTH_SHEET_ID);
  const builders = {
    "Agent Access": buildAgentAccessSheet_,
    "Credentials": buildCredentialsSheet_,
    "Login Tokens": buildLoginTokensSheet_,
    "Sessions": buildSessionsSheet_
  };

  const copied = [], skipped = [], builtFresh = [];
  Object.keys(builders).forEach(name => {
    if (target.getSheetByName(name)) {
      skipped.push(name);
      return;
    }

    const sourceSheet = source.getSheetByName(name);
    if (sourceSheet) {
      const newSheet = sourceSheet.copyTo(target);
      newSheet.setName(name);
      // copyTo() sometimes drops frozen rows on older Sheets versions,
      // so re-apply them explicitly from the source.
      const frozen = sourceSheet.getFrozenRows();
      if (frozen > 0) newSheet.setFrozenRows(frozen);
      copied.push(name);
    } else {
      builders[name](target);
      builtFresh.push(name);
    }
  });

  Logger.log("Target spreadsheet: " + target.getName() + " (" + target.getId() + ")");
  Logger.log("Source (old Auth Database): " + source.getName() + " (" + source.getId() + ")");
  Logger.log(copied.length ? ("Copied real data from old spreadsheet: " + copied.join(", ")) : "Copied real data: (none)");
  Logger.log(builtFresh.length ? ("Not found in old spreadsheet, built fresh/empty instead: " + builtFresh.join(", ")) : "Built fresh: (none)");
  Logger.log(skipped.length ? ("Already present in target, left as-is: " + skipped.join(", ")) : "Already present, left as-is: (none)");

  return { copied: copied, builtFresh: builtFresh, skipped: skipped };
}

// ══════════════════════════════════════════════════════════
// SCHEMA REPAIR — run this any time you suspect a tab or column is
// missing from TARGET_SHEET_ID (e.g. someone deleted a column, or a
// tab was copied over from an older version of the Auth Database with
// fewer fields). Safe to run repeatedly.
//
// HOW TO RUN:
//   Select "ensureAuthTablesAndColumns" in the toolbar dropdown and
//   click Run, then check View > Logs for a summary.
//
// What it does, per tab:
//   - Tab missing entirely  -> creates it fresh (via the same builders
//     used above), fully populated with the correct headers.
//   - Tab exists, missing one or more expected columns -> inserts each
//     missing column AT THE CORRECT POSITION (shifting columns after
//     it to the right), so column order still matches what Auth.gs /
//     PortalData.gs expect. Existing data in other columns is never
//     touched or reordered.
//   - Tab exists, has a column with the right name but in the WRONG
//     position -> does NOT move it (moving a column could desync it
//     from data rows read by fixed index elsewhere in the code). It's
//     only logged as a warning for you to check by hand.
// ══════════════════════════════════════════════════════════
function ensureAuthTablesAndColumns() {
  const ss = getTargetSpreadsheet(); // defined in Code.gs
  const schema = {
    "Agent Access": {
      headers: ["User ID", "Email", "Canonical Name", "Role", "Active", "Status", "Aliases"],
      // Column C used to be "Aliases". Rename it in place instead of inserting
      // a new column (which would shift Role/Active and break fixed-index reads).
      legacyHeaders: { "Canonical Name": ["Aliases (comma-separated)", "Aliases", "Aliases (semicolon-separated)"] },
      builder: buildAgentAccessSheet_
    },
    "Credentials": {
      headers: ["Email", "Password Hash", "Password Set", "Last Login", "Created Date"],
      builder: buildCredentialsSheet_
    },
    "Login Tokens": {
      headers: ["Token", "Email", "Type", "Created At", "Expires At", "Used"],
      builder: buildLoginTokensSheet_
    },
    "Sessions": {
      headers: ["Session ID", "Email", "Created At", "Expires At"],
      builder: buildSessionsSheet_
    }
  };

  const report = { tablesCreated: [], columnsAdded: {}, columnsRenamed: {}, columnsMisplaced: {}, alreadyOk: [] };

  Object.keys(schema).forEach(name => {
    const def = schema[name];
    const sh = ss.getSheetByName(name);

    if (!sh) {
      def.builder(ss);
      report.tablesCreated.push(name);
      return;
    }

    const lastCol = Math.max(sh.getLastColumn(), 1);
    let currentHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
    let touched = false;

    def.headers.forEach((expected, idx) => {
      const col = idx + 1; // 1-based column for this header in the schema
      if (currentHeaders[idx] === expected) return; // already correct in place

      // A known legacy header sitting in THIS position -> rename it in place.
      const legacyNames = (def.legacyHeaders && def.legacyHeaders[expected]) || [];
      if (legacyNames.some(l => String(currentHeaders[idx] || "").toLowerCase() === l.toLowerCase())) {
        (report.columnsRenamed[name] = report.columnsRenamed[name] || []).push(currentHeaders[idx] + " → " + expected);
        sh.getRange(1, col).setValue(expected);
        currentHeaders[idx] = expected;
        return;
      }

      const elsewhereIdx = currentHeaders.indexOf(expected);
      if (elsewhereIdx !== -1) {
        (report.columnsMisplaced[name] = report.columnsMisplaced[name] || []).push(
          expected + " — found in column " + (elsewhereIdx + 1) + ", expected column " + col + " (left as-is, not moved)"
        );
        return;
      }

      // Truly missing from this tab — insert it at the correct position.
      if (col > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), col - sh.getMaxColumns());
      else sh.insertColumnBefore(col);
      sh.getRange(1, col).setValue(expected)
        .setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
      currentHeaders.splice(col - 1, 0, expected); // keep local copy in sync for the rest of this loop
      (report.columnsAdded[name] = report.columnsAdded[name] || []).push(expected);
      touched = true;
    });

    sh.setFrozenRows(1);
    if (touched) sh.autoResizeColumns(1, def.headers.length);
    if (!report.columnsAdded[name] && !report.columnsRenamed[name] && !report.columnsMisplaced[name]) report.alreadyOk.push(name);
  });

  Logger.log("Target spreadsheet: " + ss.getName() + " (" + ss.getId() + ")");
  Logger.log(report.tablesCreated.length
    ? "Created missing tabs: " + report.tablesCreated.join(", ")
    : "Created missing tabs: (none)");
  Logger.log(Object.keys(report.columnsAdded).length
    ? "Added missing columns: " + Object.keys(report.columnsAdded).map(t => t + " [" + report.columnsAdded[t].join(", ") + "]").join("; ")
    : "Added missing columns: (none)");
  Logger.log(Object.keys(report.columnsRenamed).length
    ? "Renamed columns: " + Object.keys(report.columnsRenamed).map(t => t + " [" + report.columnsRenamed[t].join(", ") + "]").join("; ")
    : "Renamed columns: (none)");
  Logger.log(Object.keys(report.columnsMisplaced).length
    ? "WARNING — misplaced columns (not moved, check manually): " + Object.keys(report.columnsMisplaced).map(t => t + " [" + report.columnsMisplaced[t].join(" | ") + "]").join("; ")
    : "Misplaced columns: (none)");
  Logger.log(report.alreadyOk.length
    ? "Already correct: " + report.alreadyOk.join(", ")
    : "Already correct: (none)");

  return report;
}

// ── Legacy: creates a brand-new, SEPARATE spreadsheet ─────────
// Superseded by mergeAuthTabsIntoTargetSpreadsheet() above. Kept only
// for reference — nothing in this project calls this anymore.

function createAuthDatabase() {
  const ss = SpreadsheetApp.create("Inner SPARC Realty — Auth Database");

  buildReadmeSheet_(ss);
  buildAgentAccessSheet_(ss);
  buildCredentialsSheet_(ss);
  buildLoginTokensSheet_(ss);
  buildSessionsSheet_(ss);

  // The default "Sheet1" that SpreadsheetApp.create() adds automatically.
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet) ss.deleteSheet(defaultSheet);

  ss.setActiveSheet(ss.getSheetByName("README"));
  ss.moveActiveSheet(1);

  const id = ss.getId();
  const url = ss.getUrl();
  Logger.log("Created auth database.");
  Logger.log("Spreadsheet ID: " + id);
  Logger.log("URL: " + url);

  return { id: id, url: url };
}

// ── README ───────────────────────────────────────────────────
function buildReadmeSheet_(ss) {
  const sh = ss.insertSheet("README");
  sh.setColumnWidths(1, 2, 340);
  const rows = [
    ["Inner SPARC Realty — Auth Database", ""],
    ["", ""],
    ["This spreadsheet holds ONLY authentication and authorization data.", ""],
    ["Spreadsheet ID:", ss.getId()],
    ["", ""],
    ["Tab", "Purpose"],
    ["Agent Access", "AUTHORIZATION. Maps a login email to the agent's Canonical Name and any Aliases that appear in the Sales Agent column, plus a role and an active flag. This is what pa_authorize_() in PortalData.gs is designed to read."],
    ["Credentials", "AUTHENTICATION. One row per login email: password hash, whether a password has been set, last login, created date. Matches the shape pa_getProfile() in PortalData.gs already reads."],
    ["Login Tokens", "AUTHENTICATION. One-time tokens for password setup, password reset, and passwordless login links (the flows wired up in Portal.html: setup-password / reset-password / login-link)."],
    ["Sessions", "AUTHENTICATION. Active logged-in sessions (session ID, email, created/expires). Added so Auth.gs has somewhere to persist a session after a successful login, and so pa_validateSession()/pa_logout() have something to check."],
    ["", ""],
    ["SETUP STEP", "Paste this spreadsheet's ID into AUTH_SHEET_ID at the top of Auth.gs, then deploy the project as a web app so login links resolve correctly."],
    ["NOTE", "Developers, Agents, and Sales data live in a separate spreadsheet (TARGET_SHEET_ID in Code.gs) — they aren't needed here since they don't affect who can log in or what they're authorized to see."]
  ];
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).merge().setFontSize(14).setFontWeight("bold");
  sh.getRange(6, 1, 1, 2).setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
  sh.getRange(11, 1, 1, 2).setFontWeight("bold").setBackground("#fff3cd");
  sh.getRange(12, 1, 1, 2).setFontWeight("bold").setBackground("#eef1ff");
  sh.setFrozenRows(6);
}

// ── Agent Access (AUTHORIZATION) ────────────────────────────
// Read by pa_getAccessRecord_() in Auth.gs. "User ID" is a stable
// identifier independent of email (in case an agent's email changes).
// "Canonical Name" is the one official name; "Aliases" (semicolon-separated)
// lets one login also cover other name spellings that show up in the Sales
// Agent column across years. Aliases is column G (after Status in F) so Role/Active stay D/E.
function buildAgentAccessSheet_(ss) {
  const sh = ss.insertSheet("Agent Access");
  const headers = ["User ID", "Email", "Canonical Name", "Role", "Active", "Status", "Aliases"];
  writeHeader_(sh, headers);
  sh.getRange(2, 1, 1, headers.length).setValues(
    [["USR-001", "agent@example.com", "Sample Agent", "Agent", true, "", ""]]
  );
  const roleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Agent", "Supervisor", "Manager", "Admin"], true).build();
  sh.getRange(2, 4, 500, 1).setDataValidation(roleRule);
  const activeRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange(2, 5, 500, 1).setDataValidation(activeRule);
}

// ── Credentials (AUTHENTICATION) ────────────────────────────
// Column order matches what pa_getProfile() in PortalData.gs already reads:
// r[0]=Email, r[2]=Password Set, r[3]=Last Login, r[4]=Created Date.
function buildCredentialsSheet_(ss) {
  const sh = ss.insertSheet("Credentials");
  const headers = ["Email", "Password Hash", "Password Set", "Last Login", "Created Date"];
  writeHeader_(sh, headers);
  const boolRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange(2, 3, 500, 1).setDataValidation(boolRule);
  sh.getRange("D2:E").setNumberFormat("MM/dd/yyyy hh:mm a");
  sh.hideColumns(2); // password hash — keep it out of casual view
}

// ── Login Tokens (AUTHENTICATION) ───────────────────────────
// Backs the setup-password / reset-password / login-link email flows
// wired into Portal.html (INITIAL_ACTION = setup-password | reset-password
// | verify). "login-link" is the passwordless magic-link type — a user
// clicking that emailed link goes straight into a session via
// pa_verifyLoginToken() (Auth.gs), no password involved at all. Keep this
// allowed-values list in sync with every token `type` string Auth.gs's
// pa_createToken_()/pa_consumeToken_() calls actually use, or a legitimate
// token write here will get silently rejected by the validation below.
function buildLoginTokensSheet_(ss) {
  const sh = ss.insertSheet("Login Tokens");
  const headers = ["Token", "Email", "Type", "Created At", "Expires At", "Used"];
  writeHeader_(sh, headers);
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["setup-password", "reset-password", "login-link"], true).build();
  sh.getRange(2, 3, 500, 1).setDataValidation(typeRule);
  sh.getRange("D2:E").setNumberFormat("MM/dd/yyyy hh:mm a");
  const usedRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sh.getRange(2, 6, 500, 1).setDataValidation(usedRule);
}

// ── Sessions (AUTHENTICATION) ────────────────────────────────
// One row per logged-in session. Read/written by pa_createSession_(),
// pa_validateSession_(), and pa_deleteSession_() in Auth.gs.
function buildSessionsSheet_(ss) {
  const sh = ss.insertSheet("Sessions");
  const headers = ["Session ID", "Email", "Created At", "Expires At"];
  writeHeader_(sh, headers);
  sh.getRange("C2:D").setNumberFormat("MM/dd/yyyy hh:mm a");
}

// ── Shared header styling ────────────────────────────────────
function writeHeader_(sh, headers) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
}