// ============================================================
// Inner SPARC Realty Corporation — Commission Monitoring System v2.0
// Code.gs — Google Apps Script Backend
// ============================================================

// ── Target Spreadsheet ─────────────────────────────────────
// As of the Agent Portal auth merge, this spreadsheet also houses the
// Auth Database tabs (Agent Access, Credentials, Login Tokens, Sessions —
// see Auth.gs). pa_authSpreadsheet_() in Auth.gs reads this same
// spreadsheet via getTargetSpreadsheet() below rather than keeping its
// own separate ID, so there is exactly one place to update if this
// spreadsheet is ever moved/copied again.
const TARGET_SHEET_ID = "1TJmB5JpmEdHQqWXjlBq4DmSkbAP3oJZiSaWbEu0_j5M";

function getTargetSpreadsheet() {
  return SpreadsheetApp.openById(TARGET_SHEET_ID);
}

// ── CMS Spreadsheet (Inner SPARC CMS) ──────────────────────
// The auth migration (see Auth.gs / CreateDedicatedDatabase.gs) only ever
// copied Agent Access, Credentials, Login Tokens, and Sessions into
// TARGET_SHEET_ID. It did NOT migrate historical Sales data — the sales
// CRUD functions further down this file all read/write
// "<currentYear> Sales" (currently "2026 Sales") inside TARGET_SHEET_ID,
// which IS the live, actively-edited sheet for the current year. But
// 2020 Sales .. 2025 Sales were never copied over and still only exist in
// the separate "Inner SPARC CMS" spreadsheet below, which is also the
// source used to pivot/normalize agent names against the "Sales Agent"
// text in those older tabs. Anything that needs FULL sales history
// (not just the current year) must read historical years from
// getCmsSpreadsheet() and merge them with the current year from
// getTargetSpreadsheet() — see pa_getSalesForAliases_() / pa_getAllSalesYears_()
// in PortalData.gs for the merge logic.
const CMS_SHEET_ID = "17woAyRfBJgQMhZ_SYarjKZwRNpO_2xlV49sQmzHrF6c";
//const CMS_SHEET_ID = "1TJVunI2XHpUEBqqGn0jrJLAOkJpDjq0G4ZOFQwXzxqg";

function getCmsSpreadsheet() {
  return SpreadsheetApp.openById(CMS_SHEET_ID);
}

// ── Sheet Name Constants ────────────────────────────────────
const SHEET_DEVELOPERS = "Developers";
const SHEET_AGENTS = "Agents";

// ── Entry Point ──
// Routes every request (Agent Portal + Admin panel) through the same
// "Portal" template. "admin" is just another protected deep link, like
// dashboard/sales/profile — the Admin nav tab and view inside Portal.html
// only render/act for sessions whose role is Admin, which is enforced
// server-side by every pa_admin* function in Auth.gs (the client-side
// hiding of the tab is a convenience, not the security boundary).
//
// NOTE (fixed): this used to special-case ?action=admin with
// HtmlService.createHtmlOutputFromFile(), which serves the file's raw text
// WITHOUT evaluating GAS template scriptlets. Portal.html's <script> block
// relies on <?!= JSON.stringify(initialAction) ?> being evaluated to inject
// the initial route — skipping that left the literal "<?!= ... ?>" text in
// the page's JavaScript, which is a syntax error and broke the entire admin
// page (blank screen / console error). Routing "admin" through the same
// createTemplateFromFile() path as everything else fixes that.
// NOTE: "verify" is the passwordless login-link action — a user who
// requests a login link (Portal.html's "Email me a login link", Auth.gs's
// pa_requestLoginLink) gets an emailed ?action=verify&token=... URL, and
// Portal.html's handleInitialRoute() calls pa_verifyLoginToken(token) to
// consume it and start a session directly, no password involved. This was
// previously whitelisted just so a lingering old link could be told
// explicitly "this link type is no longer supported" rather than falling
// through to resolvedAction="login" with initialToken never even set —
// which looked identical to the page simply being broken. Keep "verify" in
// this list either way: without it, doGet() silently drops the token
// before Portal.html ever sees it.
const PORTAL_ACTIONS_ = ["login", "setup-password", "reset-password", "verify",
                         "dashboard", "sales", "sale", "profile", "admin",
                         "reportcard", "salesperf", "emails"];

// Portal.html is fully self-contained — its own <style> block and its own
// <script> block, plus the 3 server-injected route consts. It does NOT
// include Portal_CSS.html or Portal_JS.html: include() below is defined but
// never called. Those two files are stale/unused leftovers from an earlier
// split-file layout and are NOT part of what's served — edits to them have
// no effect on the running app. Edit Portal.html directly, and delete
// Portal_CSS.html / Portal_JS.html from the project (or clearly mark them
// deprecated) so they stop being mistaken for live code.
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : null;

  // ── Email open-tracking pixel ──────────────────────────────
  // Handled BEFORE the Portal routing below and OUTSIDE the PORTAL_ACTIONS_
  // whitelist on purpose: this is the only doGet() request that must NOT
  // return the Portal HTML template. It's hit directly by mail clients
  // loading the <img> pixel embedded in a broadcast email (EmailCampaign.gs),
  // so it has to (a) work with zero session/auth — a mail client can't log
  // in — and (b) always return quickly and successfully, even for a
  // garbage/expired id, so no mail client ever shows a broken-image icon
  // any differently than it would for a legitimate one. See EmailCampaign.gs
  // ("EMAIL OPEN TRACKING" section) for the actual logging logic —
  // ec_handleTrackingRequest_ is intentionally the only thing called here.
  if (action === 'track') {
    return ec_handleTrackingRequest_(e); // EmailCampaign.gs
  }

  // Unrecognized/missing actions fall back to "login" inside Portal.html.
  const resolvedAction = (action && PORTAL_ACTIONS_.indexOf(action) !== -1) ? action : "login";
  const tmpl = HtmlService.createTemplateFromFile("Portal");
  tmpl.initialAction = resolvedAction;
  // Optional script property overrides the default; set it to 'off' to disable.
  const analyticsOverride = PropertiesService.getScriptProperties().getProperty('GA4_MEASUREMENT_ID');
  const analyticsId = String(analyticsOverride === null ? 'G-F9V2PHXFBB' : analyticsOverride).trim();
  tmpl.analyticsMeasurementId = /^G-[A-Z0-9]+$/.test(analyticsId) ? analyticsId : '';
  tmpl.analyticsPageUrl = ScriptApp.getService().getUrl() || 'https://script.google.com';
  tmpl.initialToken = String((e && e.parameter && e.parameter.token) || "").trim();
  tmpl.initialSaleId = (e && e.parameter && e.parameter.id) || "";
  // Session fallback for page refreshes when localStorage isn't reachable
  // (e.g. sandboxed webapp iframe) — Portal.html reads this back into
  // INITIAL_SESSION and round-trips it via ?session=... on history.replace().
  tmpl.initialSession = (e && e.parameter && e.parameter.session) || "";
  return tmpl.evaluate()
    .setTitle(resolvedAction === "admin" ? "Inner SPARC Realty — Admin" : "Inner SPARC Realty — Agent Portal")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ── ON EDIT TRIGGER - Automatically updates Company Net Comm when relevant cells change,
//    and auto-provisions Credentials/Agents rows when a user is added to Agent Access,
//    and cascades email changes across related tables ──
function onEditInstallable(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  const editedColumn = range.getColumn();
  const editedRow = range.getRow();

  // Admin typed/pasted a row directly into Agent Access -> give it the
  // same Credentials (+ Agents roster, if role is "Agent") provisioning
  // that pa_addAgentAccess_() gives a programmatic add. See Auth.gs.
  if (sheetName === "Agent Access") {
    pa_onAgentAccessEdit_(e); // defined in Auth.gs
    return;
  }

  // Single-cell email edit in Agents roster (column C = Email) -> cascade
  // the change to Agent Access and Credentials. For multi-cell edits,
  // reconciliation will catch any resulting mismatches.
  if (sheetName === SHEET_AGENTS && editedColumn === 3 && editedRow >= 2 &&
      range.getNumRows() === 1 && e.oldValue !== undefined && e.value !== undefined) {
    const oldEmail = String(e.oldValue || '').trim().toLowerCase();
    const newEmail = String(e.value || '').trim().toLowerCase();
    if (oldEmail && oldEmail.indexOf('@') !== -1 && newEmail && newEmail.indexOf('@') !== -1 && oldEmail !== newEmail) {
      try {
        pa_cascadeEmailChange_(oldEmail, newEmail, 'agents'); // Auth.gs
      } catch (err) {
        Logger.log('onEdit: Agents email cascade failed: ' + err.toString());
      }
    }
  }

  const currentYear = new Date().getFullYear();
  const targetSheetName = currentYear + " Sales";
  
  // Only process edits on the current year's Sales sheet
  if (sheetName !== targetSheetName) return;
  
  // Skip header row
  if (editedRow < 2) return;
  
  // Columns that affect Company Net Comm calculation
  const AFFECTED_COLUMNS = [11, 12, 14, 16, 30]; // K(11), L(12), N(14), P(16), AD(30)
  // Developer Comm Rate (K), Company Gross (L), Withholding Tax (N), Agent Comm Rate (P), Cash Advance (AD)
  
  if (AFFECTED_COLUMNS.includes(editedColumn)) {
    // Update Company Net Comm for the edited row
    updateCompanyNetCommForRow(editedRow, sheet);
  }
}

function installEditTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onEditInstallable') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(TARGET_SHEET_ID)
    .onEdit()
    .create();
}

// ── Helper Functions ──
function getSheetData(sheetName, startRow = 2) {
  try {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sh = ss.getSheetByName(sheetName);
    if (!sh) {
      console.error(`Sheet ${sheetName} not found`);
      return [];
    }
    
    const lastRow = sh.getLastRow();
    if (lastRow < startRow) return [];
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const rows = sh.getRange(startRow, 1, lastRow - startRow + 1, sh.getLastColumn()).getValues();
    
    return rows.map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let value = row[i];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "MM/dd/yyyy");
        }
        const cleanHeader = String(h || '').trim();
        obj[cleanHeader] = value;
      });
      return obj;
    });
  } catch (error) {
    console.error(`Error reading sheet ${sheetName}:`, error.toString());
    return [];
  }
}

function appendRow(sheetName, rowData) {
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  sh.appendRow(rowData);
}

function generateId(prefix) {
  return prefix + "-" + new Date().getTime().toString().slice(-6);
}

// ── Auto-generate Sale IDs for existing data ──
function ensureSaleIds(sheetName) {
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  
  if (!sh) return { success: false, error: "Sales sheet not found" };
  
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: true, message: "No data rows" };
  
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const saleIdColumnIndex = headers.findIndex(h => h === 'Sale ID');
  
  if (saleIdColumnIndex === -1) {
    return { success: false, error: "Sale ID column not found" };
  }
  
  let updatedCount = 0;
  
  for (let i = 2; i <= lastRow; i++) {
    const saleId = sh.getRange(i, saleIdColumnIndex + 1).getValue();
    if (!saleId || saleId === '') {
      const newSaleId = generateId("SALE");
      sh.getRange(i, saleIdColumnIndex + 1).setValue(newSaleId);
      updatedCount++;
    }
  }
  
  return { success: true, updatedCount };
}

// ── Agents ──
function getAgents() {
  try {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    let sh = ss.getSheetByName(SHEET_AGENTS);
    
    if (!sh) {
      console.log("Agents sheet not found, creating it...");
      sh = ss.insertSheet(SHEET_AGENTS);
      const headers = ["Agent ID", "Full Name", "Email", "Phone", "License No.", "Date Hired", "Status"];
      sh.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
      sh.setFrozenRows(1);
      sh.appendRow(["AGT-001", "Sample Agent", "agent@example.com", "09123456789", "PRC-12345", new Date().toLocaleDateString(), "Active"]);
    }
    
    if (sh.getLastRow() < 2) return [];
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    
    return rows.map(row => {
      const agent = {};
      headers.forEach((header, index) => {
        let value = row[index];
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "MM/dd/yyyy");
        }
        if (header === 'Phone' && typeof value === 'number') {
          value = '0' + value.toString();
        }
        agent[header] = value;
      });
      return agent;
    });
  } catch (error) {
    console.error("Error in getAgents:", error.toString());
    return [];
  }
}

/** Inserts a new Agents-roster row directly below the current last data
 *  row (row 1 = header, so a roster of 5 agents gets the new row at row 7,
 *  right after row 6) rather than sheet.appendRow(), which lands on the
 *  sheet's absolute last used row and can drift down over time if stray
 *  formatting/blank rows ever end up below the real data. Mirrors
 *  pa_insertAgentAccessRowBelowLast_ (Auth.gs) for the Agent Access sheet. */
function insertAgentRowBelowLast_(sh, rowValues) {
  const lastRow = Math.max(sh.getLastRow(), 1); // at least the header row
  const insertAt = lastRow + 1;
  sh.insertRowAfter(lastRow);
  const newRange = sh.getRange(insertAt, 1, 1, rowValues.length);
  // insertRowAfter() copies the formatting of the row above (bold/blue when
  // that's the header row, or whatever a previous row happened to carry) —
  // reset explicitly so new agent rows always render plain.
  newRange.setBackground(null).setFontWeight("normal");
  newRange.setValues([rowValues]);
}

/** Admin: creates a new Agent row with an auto-generated Agent ID, inserted
 *  directly below the last existing row (see insertAgentRowBelowLast_).
 *  If an email is given, also grants that agent Agent Portal login access
 *  (an "Agent Access" + "Credentials" row, via pa_ensureAgentAccessForAgent_
 *  in Auth.gs) — so adding someone here always leaves a matching entry in
 *  User Access, skipped (not duplicated) if one already exists. */
function addAgent(data) {
  try {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    let sh = ss.getSheetByName(SHEET_AGENTS);
    if (!sh) {
      getAgents(); // creates the sheet (with headers + sample row) if missing
      sh = ss.getSheetByName(SHEET_AGENTS);
    }

    const agentId = generateId("AGT");
    insertAgentRowBelowLast_(sh, [
      agentId,
      data.fullName || '',
      data.email || '',
      data.phone || '',
      data.license || '',
      data.dateHired || new Date().toLocaleDateString(),
      data.status || 'Active'
    ]);

    const email = String(data.email || '').trim();
    let accessWarning = null;
    if (email) {
      try {
        const provisioned = pa_ensureAgentAccessForAgent_(email, data.fullName || ''); // Auth.gs
        // Only auto-send the "set up your password" email the moment portal
        // access is actually newly granted — not on every edit/re-add, and
        // never for an email that already had access (they may already
        // have a password set; re-sending would be a surprise reset link).
        // This is the ONLY place that email was previously ever triggered
        // from — nothing called pa_requestPasswordAction on the add path,
        // so new agents were provisioned with login access but never told
        // it existed. The admin's only recourse was the "Send password
        // setup link" paper-plane button — which lives on the Admin > Users
        // tab, not here on the Agents tab where this data.email came from.
        if (provisioned.agentAccess === 'added') {
          try {
            pa_requestPasswordAction(email, 'setup'); // Auth.gs — same flow as the manual "send setup link" button
          } catch (e) {
            console.error('addAgent: pa_requestPasswordAction failed for ' + email + ' — ' + e.toString());
            accessWarning = 'Agent added, but the setup-password email could not be sent automatically. Use the paper-plane icon on Admin > Users for ' + email + '.';
          }
        }
      } catch (e) {
        console.error('addAgent: pa_ensureAgentAccessForAgent_ failed for ' + email + ' — ' + e.toString());
        // This used to be swallowed into console.error only — invisible from
        // the Admin UI, so a real provisioning failure looked identical to
        // success. Surface it so the admin knows to check Admin > Users.
        accessWarning = 'Agent added, but portal access could not be granted automatically (' + e.toString() + '). Add them manually on Admin > Users.';
      }
    }

    return { success: true, agentId: agentId, warning: accessWarning };
  } catch (error) {
    console.error("Error in addAgent:", error.toString());
    return { success: false, error: error.toString() };
  }
}

/** Admin: updates an existing Agent row, matched by original Agent ID. */
function updateAgent(data) {
  try {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sh = ss.getSheetByName(SHEET_AGENTS);
    if (!sh) return { success: false, error: "Agents sheet not found" };

    const lastRow = sh.getLastRow();
    let targetRow = -1;
    for (let i = 2; i <= lastRow; i++) {
      if (sh.getRange(i, 1).getValue() === data.originalId) { targetRow = i; break; }
    }
    if (targetRow === -1) return { success: false, error: "Agent not found" };

    sh.getRange(targetRow, 2).setValue(data.fullName || '');   // B: Full Name
    sh.getRange(targetRow, 3).setValue(data.email || '');      // C: Email
    sh.getRange(targetRow, 4).setValue(data.phone || '');      // D: Phone
    sh.getRange(targetRow, 5).setValue(data.license || '');    // E: License No.
    sh.getRange(targetRow, 6).setValue(data.dateHired || '');  // F: Date Hired
    sh.getRange(targetRow, 7).setValue(data.status || 'Active'); // G: Status

    return { success: true };
  } catch (error) {
    console.error("Error in updateAgent:", error.toString());
    return { success: false, error: error.toString() };
  }
}

/** Admin: deletes an Agent row by Agent ID. Also revokes portal access
 *  (Agent Access + Credentials + Sessions) for this agent's email, if any
 *  — otherwise they'd still be able to log in after being removed from
 *  the roster, and they'd keep showing up on the Admin panel's Users tab.
 *  Only ever touches a role "Agent" access row (see
 *  pa_cascadeDeleteAgentAccessIfRoleAgent_ in Auth.gs) — a Supervisor/
 *  Manager/Admin account is never affected just because a same-email
 *  Agents roster row was deleted. */
function deleteAgent(agentId) {
  try {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sh = ss.getSheetByName(SHEET_AGENTS);
    if (!sh) return { success: false, error: "Agents sheet not found" };

    const data = sh.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === agentId) {
        const email = String(data[i][2] || '').trim(); // C: Email
        sh.deleteRow(i + 1);

        if (email) {
          try {
            pa_cascadeDeleteAgentAccessIfRoleAgent_(email); // Auth.gs
          } catch (e) {
            console.error('deleteAgent: access cascade failed for ' + email + ' — ' + e.toString());
          }
        }
        return { success: true };
      }
    }
    return { success: false, error: "Agent not found" };
  } catch (error) {
    console.error("Error in deleteAgent:", error.toString());
    return { success: false, error: error.toString() };
  }
}

/** Removes the Agents roster row matching `email`, if any — the reverse
 *  direction of the cascade above. Used by pa_adminDeleteUser() (Auth.gs)
 *  so removing someone's portal access on the Users tab also removes them
 *  from the Agents tab, and by pa_reconcileAuthTables_() for the same sync
 *  after a manual sheet edit. A missing match is not an error, same as
 *  deleteAgent()'s own "not found" case. */
function deleteAgentByEmail_(email) {
  const target = String(email || '').trim().toLowerCase();
  if (!target) return { success: false, error: 'No email given' };
  const agent = getAgents().find(a => String(a['Email'] || '').trim().toLowerCase() === target);
  if (!agent) return { success: false, error: 'Agent not found' };
  return deleteAgent(agent['Agent ID']);
}

// ── Developers ──
function getDevelopers() {
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const sh = ss.getSheetByName(SHEET_DEVELOPERS);
  
  if (!sh) return [];
  
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  
  // Get all columns including new columns (A through I)
  const rows = sh.getRange(2, 1, lastRow - 1, 9).getValues();
  
  return rows.map((row, index) => ({
    'Developer ID': 'DEV-' + (index + 1).toString().padStart(3, '0'),
    'Developer Name': row[0] || '',
    'Developer Comm Rate (%)': row[1] || 0,
    'Agent Comm Rate (%)': row[2] || 0,
    'TSP Formula': row[3] || '',
    'VAT (%)': row[4] || 0,
    'Withholding Tax (%)': row[5] || 0,
    'LMF (%)': row[6] || 0,
    'TIN': row[7] || '',
    'Developer Address': row[8] || '',
    'Project': row[0] || ''
  }));
}

function updateDeveloper(data) {
  try {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sh = ss.getSheetByName(SHEET_DEVELOPERS);
    
    if (!sh) {
      return { success: false, error: "Developers sheet not found" };
    }
    
    const lastRow = sh.getLastRow();
    let targetRow = -1;
    
    // Find the row by original developer name
    for (let i = 2; i <= lastRow; i++) {
      const name = sh.getRange(i, 1).getValue();
      if (name === data.originalName) {
        targetRow = i;
        break;
      }
    }
    
    if (targetRow === -1) {
      // If not found by original name, try by new name
      for (let i = 2; i <= lastRow; i++) {
        const name = sh.getRange(i, 1).getValue();
        if (name === data.developerName) {
          targetRow = i;
          break;
        }
      }
    }
    
    if (targetRow === -1) {
      return { success: false, error: "Developer not found" };
    }
    
    // Update all columns (A through I)
    sh.getRange(targetRow, 1).setValue(data.developerName);        // A: Developer Name
    sh.getRange(targetRow, 2).setValue(data.devCommRate);          // B: Dev Comm Rate
    sh.getRange(targetRow, 3).setValue(data.agentCommRate);        // C: Agent Comm Rate
    // Column 4 (D: TSP Formula) - keep as is
    sh.getRange(targetRow, 5).setValue(data.vatRate);              // E: VAT
    sh.getRange(targetRow, 6).setValue(data.withholdingRate);      // F: Withholding Tax
    sh.getRange(targetRow, 7).setValue(data.lmfRate);              // G: LMF
    sh.getRange(targetRow, 8).setValue(data.tin || '');            // H: TIN
    sh.getRange(targetRow, 9).setValue(data.developerAddress || ''); // I: Developer Address
    
    return { success: true };
  } catch (error) {
    console.error("Error in updateDeveloper:", error.toString());
    return { success: false, error: error.toString() };
  }
}

function addDeveloper(data) {
  appendRow(SHEET_DEVELOPERS, [
    data.developerName,      // A
    data.devCommRate,        // B
    data.agentCommRate,      // C
    "",                      // D: TSP Formula (empty)
    data.vatRate,            // E
    data.withholdingRate,    // F
    data.lmfRate,            // G
    data.tin || '',          // H: TIN
    data.address || ''       // I: Developer Address
  ]);
  return { success: true };
}

function deleteDeveloper(devName) {
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const sh = ss.getSheetByName(SHEET_DEVELOPERS);
  const data = sh.getDataRange().getValues();
  
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === devName) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: "Developer not found" };
}

// ── Get Sales with Full Commission Data using 0-based Column Indexing ──
function getSales(year) {
  const sheetName = year + " Sales";
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  let sh = ss.getSheetByName(sheetName);
  
  if (!sh) {
    Logger.log(`Sheet ${sheetName} not found`);
    return [];
  }
  
  ensureSaleIds(sheetName);
  
  if (sh.getLastRow() < 2) return [];
  
  // Read columns A-U (21 columns, indices 0-20)
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 21).getValues();
  
  return rows.map(row => {
    let reservationDate = row[0];
    if (reservationDate instanceof Date) {
      reservationDate = Utilities.formatDate(reservationDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
    }
    
    return {
      'Reservation Date': reservationDate || '',
      'Sales Agent': row[1] || '',
      'Developer': row[2] || '',
      'Project': row[3] || '',
      'Unit ID': row[4] || '',
      "Buyer's Name": row[5] || '',
      'DP Term': row[6] || '',
      'ACTIVE SALES AMOUNT': parseFloat(row[7]) || 0,
      'Computed TSP': parseFloat(row[9]) || 0,
      'Company Gross Comm (₱)': parseFloat(row[11]) || 0,
      'Net VAT (₱)': parseFloat(row[12]) || 0,  // This is Net AFTER VAT deduction
      'Withholding Tax (Company) (₱)': parseFloat(row[13]) || 0,
      'Company Net Comm (₱)': parseFloat(row[14]) || 0,
      'Agent Gross Comm (₱)': parseFloat(row[16]) || 0,
      'VAT Share (Agent) (₱)': parseFloat(row[17]) || 0,
      'Withholding Tax (Agent) (₱)': parseFloat(row[18]) || 0,
      'Agent Net Comm (₱)': parseFloat(row[19]) || 0,
      'Developer Comm Rate (%)': parseFloat(row[10]) || 0,
      'Agent Comm Rate (%)': parseFloat(row[15]) || 0,
      'Sale ID': row[20] || ''
    };
  });
}

// ── Get Commissions Data for a Specific Year using Column Indexing ──
function getCommissionsForYear(year) {
  const sheetName = year + " Sales";
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  
  if (!sh || sh.getLastRow() < 2) return [];
  
  // Read columns A-U (21 columns)
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 21).getValues();
  
  return rows.map(row => {
    // Column Index Mapping (0-based):
    // 0: Reservation Date, 1: Sales Agent, 2: Developer, 3: Project, 4: Unit ID
    // 5: Buyer's Name, 6: DP Term, 7: ACTIVE SALES AMOUNT, 8: Agent ID
    // 9: Computed TSP, 10: Developer Comm Rate (%), 11: Company Gross Comm (₱)
    // 12: Net VAT (₱), 13: Withholding Tax (Company) (₱), 14: Company Net Comm (₱)
    // 15: Agent Comm Rate (%), 16: Agent Gross Comm (₱), 17: VAT Share (Agent) (₱)
    // 18: Withholding Tax (Agent) (₱), 19: Agent Net Comm (₱), 20: Sale ID
    
    let reservationDate = row[0];
    if (reservationDate instanceof Date) {
      reservationDate = Utilities.formatDate(reservationDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
    }
    
    const agentNetComm = parseFloat(row[19]) || 0;
    let releasedAmount = 0;
    let receivablesAmount = 0;
    let commissionPercentage = '0%';
    
    // Try to get released commission from column V (index 21) if it exists
    // Otherwise, we'll calculate from available data
    try {
      // Check if column V exists (index 21)
      const allColumns = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
      if (allColumns.length > 21) {
        // Column V exists (index 21)
        const fullRow = sh.getRange(targetRow + 1, 1, 1, sh.getLastColumn()).getValues()[0];
        releasedAmount = parseFloat(fullRow[21]) || 0;
      }
    } catch (e) {
      // Column V doesn't exist, use default
      releasedAmount = 0;
    }
    
    receivablesAmount = agentNetComm - releasedAmount;
    
    if (agentNetComm > 0) {
      const percentValue = (releasedAmount / agentNetComm) * 100;
      commissionPercentage = percentValue.toFixed(1) + '%';
      if (percentValue === 100) commissionPercentage = '100%';
    }
    
    return {
      'Sale ID': row[20] || '',
      'Reservation Date': reservationDate || '',
      'Sales Agent': row[1] || '',
      'Developer': row[2] || '',
      'Project': row[3] || '',
      'Unit ID': row[4] || '',
      'DP Term': row[6] || '',
      'Computed TSP': parseFloat(row[9]) || 0,
      'Developer Comm Rate (%)': parseFloat(row[10]) || 0,
      'Company Gross Comm (₱)': parseFloat(row[11]) || 0,
      'Net VAT (₱)': parseFloat(row[12]) || 0,
      'Withholding Tax (Company) (₱)': parseFloat(row[13]) || 0,
      'Company Net Comm (₱)': parseFloat(row[14]) || 0,
      'Agent Comm Rate (%)': parseFloat(row[15]) || 0,
      'Agent Gross Comm (₱)': parseFloat(row[16]) || 0,
      'VAT Share (Agent) (₱)': parseFloat(row[17]) || 0,
      'Withholding Tax (Agent) (₱)': parseFloat(row[18]) || 0,
      'Agent Net Comm (₱)': agentNetComm,
      'Milestone Ratio': '--',
      'DP Tranche': '--',
      'Loan Takeout Tranche': '--',
      'Released Commission': releasedAmount,
      'Receivables': receivablesAmount,
      'Commission Status': commissionPercentage,
      'DP Release Amounts': '',
      'Loan Takeout Amount': 0,
      'Cash Advance': 0
    };
  }).filter(c => c['Sale ID']);
}

function recordSale(data) {
  const currentYear = new Date().getFullYear();
  const sheetName = currentYear + " Sales";
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  let sh = ss.getSheetByName(sheetName);
  
  if (!sh) {
    sh = ss.insertSheet(sheetName);
    const headers = ["Reservation Date", "Sales Agent", "Developer", "Project", "Unit ID", "Buyer's Name", 
                     "DP Term", "ACTIVE SALES AMOUNT", "Agent ID", "Computed TSP", "Developer Comm Rate (%)", 
                     "Company Gross Comm (₱)", "Net VAT (₱)", "Withholding Tax (Company) (₱)", 
                     "Company Net Comm (₱)", "Agent Comm Rate (%)", "Agent Gross Comm (₱)", 
                     "VAT Share (Agent) (₱)", "Withholding Tax (Agent) (₱)", "Agent Net Comm (₱)", 
                     "Sale ID", "Milestone Ratio", "DP Tranche", "Loan Takeout Tranche"];
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
    sh.setFrozenRows(1);
  }
  
  const tsp = parseFloat(data.computedTSP) || 0;
  const devCommRate = parseFloat(data.devCommRate) || 0;
  const agentCommRate = parseFloat(data.agentCommRate) || 0;
  const vatRate = parseFloat(data.vatRate) || 12;
  const withholdingRate = parseFloat(data.withholdingRate) || 5;
  
  const companyGross = tsp * devCommRate / 100;
  const vatAmount = companyGross * vatRate / 100;
  const companyWithholdingTax = companyGross * withholdingRate / 100;
  const companyNet = companyGross - vatAmount - companyWithholdingTax;
  
  const agentGross = tsp * agentCommRate / 100;
  const vatShareAmount = agentGross * vatRate / 100;
  const agentWithholdingTax = agentGross * withholdingRate / 100;
  const agentNet = agentGross - vatShareAmount - agentWithholdingTax;
  
  const saleId = generateId("SALE");
  const lastRow = sh.getLastRow();
  const nextRow = lastRow + 1;
  
  const rowData = [
    data.reservationDate,
    data.agentName,
    data.developerName,
    data.project,
    data.unitId,
    data.buyerName,
    data.dpTerm,
    data.activeSalesAmount,
    data.agentId,
    data.computedTSP,
    devCommRate,
    companyGross,
    vatAmount,
    companyWithholdingTax,
    companyNet,
    agentCommRate,
    agentGross,
    vatShareAmount,
    agentWithholdingTax,
    agentNet,
    saleId,
    "--",
    "--",
    "--"
  ];
  
  sh.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
  
  return {
    success: true,
    saleId,
    companyGross,
    vatAmount,
    companyWithholdingTax,
    companyNet,
    agentGross,
    vatShareAmount,
    agentWithholdingTax,
    agentNet
  };
}

// ── Commissions ──
function getCommissions(year) {
  const sheetName = year + " Sales";
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  
  if (!sh || sh.getLastRow() < 2) return [];
  
  // Read all columns (A through AE = 31 columns, indices 0-30)
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 31).getValues();
  
  return rows.map(row => {
    // 0-based column mapping:
    // 0: A - Reservation Date
    // 1: B - Sales Agent
    // 2: C - Developer
    // 3: D - Project
    // 4: E - Unit ID
    // 5: F - Buyer's Name
    // 6: G - DP Term
    // 7: H - ACTIVE SALES AMOUNT
    // 8: I - Agent ID
    // 9: J - Computed TSP
    // 10: K - Developer Comm Rate (%)
    // 11: L - Company Gross Comm (₱)
    // 12: M - Net VAT (₱)
    // 13: N - Withholding Tax (Company) (₱)
    // 14: O - Company Net Comm (₱)
    // 15: P - Agent Comm Rate (%)
    // 16: Q - Agent Gross Comm (₱)
    // 17: R - VAT Share (Agent) (₱)
    // 18: S - Withholding Tax (Agent) (₱)
    // 19: T - Agent Net Comm (₱)
    // 20: U - Sale ID
    // 21: V - Milestone Ratio
    // 22: W - DP Tranche
    // 23: X - Loan Takeout Tranche*
    // 24: Y - Released Commission*
    // 25: Z - Receivables*
    // 26: AA - Commission Status (%)*
    // 27: AB - DP Release Amount
    // 28: AC - Loan Takeout Release Amount
    // 29: AD - Cash Advance
    // 30: AE - Commission Status* (duplicate - use column 26 instead)
    
    const agentNetComm = parseFloat(row[19]) || 0;
    
    // Get released commission from column Y (index 24)
    const releasedAmount = parseFloat(row[24]) || 0;
    
    // Get receivables from column Z (index 25)
    let receivablesAmount = parseFloat(row[25]) || 0;
    
    // Get commission status from column AA (index 26)
    let commissionPercentage = row[26] || '0%';
    
    // Get DP release amounts from column AB (index 27)
    let dpReleaseAmounts = [];
    const dpReleaseRaw = row[27] || '';
    if (dpReleaseRaw) {
      dpReleaseAmounts = dpReleaseRaw.toString().split(',').map(v => parseFloat(v.trim()));
    }
    
    // Get loan takeout amount from column AC (index 28)
    const loanTakeoutAmount = parseFloat(row[28]) || 0;
    
    // Get cash advance from column AD (index 29)
    const cashAdvance = parseFloat(row[29]) || 0;
    
    // If status is missing but we have released amount, calculate it
    if (!commissionPercentage || commissionPercentage === '0%' || commissionPercentage === '') {
      if (agentNetComm > 0) {
        const percentValue = (releasedAmount / agentNetComm) * 100;
        commissionPercentage = percentValue.toFixed(1) + '%';
        if (percentValue >= 99.99) commissionPercentage = '100%';
      }
    }
    
    // If receivables is 0 but we have released amount, calculate it
    if (receivablesAmount === 0 && agentNetComm > 0) {
      receivablesAmount = agentNetComm - releasedAmount;
    }
    
    return {
      'Sale ID': row[20] || '',
      'Reservation Date': row[0] instanceof Date ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "MM/dd/yyyy") : row[0] || '',
      'Sales Agent': row[1] || '',
      'Developer': row[2] || '',
      'Project': row[3] || '',
      'Unit ID': row[4] || '',
      "Buyer's Name": row[5] || '',
      'DP Term': row[6] || '',
      'Computed TSP': parseFloat(row[9]) || 0,
      'Developer Comm Rate (%)': parseFloat(row[10]) || 0,
      'Company Gross Comm (₱)': parseFloat(row[11]) || 0,
      'Net VAT (₱)': parseFloat(row[12]) || 0,
      'Withholding Tax (Company) (₱)': parseFloat(row[13]) || 0,
      'Company Net Comm (₱)': parseFloat(row[14]) || 0,
      'Agent Comm Rate (%)': parseFloat(row[15]) || 0,
      'Agent Gross Comm (₱)': parseFloat(row[16]) || 0,
      'VAT Share (Agent) (₱)': parseFloat(row[17]) || 0,
      'Withholding Tax (Agent) (₱)': parseFloat(row[18]) || 0,
      'Agent Net Comm (₱)': agentNetComm,
      'Milestone Ratio': row[21] || '--',
      'DP Tranche': row[22] || '--',
      'Loan Takeout Tranche': row[23] || '--',
      'Released Commission': releasedAmount,
      'Receivables': receivablesAmount,
      'Commission Status': commissionPercentage,
      'DP Release Amounts': dpReleaseAmounts,
      'Loan Takeout Amount': loanTakeoutAmount,
      'Cash Advance': cashAdvance
    };
  }).filter(c => c['Sale ID']);
}

// ── CORE FUNCTION: Update Company Net Comm for a specific row ──
function updateCompanyNetCommForRow(row, sheet) {
  try {
    // Column definitions (1-based indices)
    const COL_DEV_COMM_RATE = 11;      // Column K
    const COL_COMPANY_GROSS = 12;      // Column L
    const COL_WITHHOLDING = 14;        // Column N
    const COL_AGENT_COMM_RATE = 16;    // Column P
    const COL_CASH_ADVANCE = 30;       // Column AD
    const COL_COMPANY_NET = 15;        // Column O
    
    // Read values
    const companyGross = parseFloat(sheet.getRange(row, COL_COMPANY_GROSS).getValue()) || 0;
    const withholding = parseFloat(sheet.getRange(row, COL_WITHHOLDING).getValue()) || 0;
    const devRate = parseFloat(sheet.getRange(row, COL_DEV_COMM_RATE).getValue()) || 0;
    const agentRate = parseFloat(sheet.getRange(row, COL_AGENT_COMM_RATE).getValue()) || 0;
    const cashAdvance = parseFloat(sheet.getRange(row, COL_CASH_ADVANCE).getValue()) || 0;
    
    // Calculate base Company Net Comm (Gross - Withholding)
    let baseNet = companyGross - withholding;
    
    // Deduct the FULL cash advance from Company Net Comm
    let finalNet = baseNet - cashAdvance;
    finalNet = Math.max(0, Math.round(finalNet * 100) / 100);
    
    // Update the cell
    sheet.getRange(row, COL_COMPANY_NET).setValue(finalNet);
    
    return {
      success: true,
      row: row,
      baseNet: baseNet,
      cashAdvance: cashAdvance,
      finalNet: finalNet
    };
  } catch (error) {
    console.error(`Error updating row ${row}:`, error.toString());
    return { success: false, error: error.toString() };
  }
}

// ── Update Cash Advance (preserved for HTML interface calls) ──
function updateCashAdvance(saleId, cashAdvance) {
  try {
    const currentYear = new Date().getFullYear();
    const sheetName = currentYear + " Sales";
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sh = ss.getSheetByName(sheetName);
    
    if (!sh) {
      return { success: false, error: "Sales sheet not found" };
    }
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const saleIdColumnIndex = headers.findIndex(h => h === 'Sale ID');
    const cashAdvanceIndex = headers.findIndex(h => h === 'Cash Advance');
    const devCommRateIndex = headers.findIndex(h => h === 'Developer Comm Rate (%)');
    const agentCommRateIndex = headers.findIndex(h => h === 'Agent Comm Rate (%)');
    
    if (saleIdColumnIndex === -1 || cashAdvanceIndex === -1) {
      return { success: false, error: "Required columns not found" };
    }
    
    const lastRow = sh.getLastRow();
    let targetRow = -1;
    
    for (let i = 2; i <= lastRow; i++) {
      const id = sh.getRange(i, saleIdColumnIndex + 1).getValue();
      if (id === saleId) {
        targetRow = i;
        break;
      }
    }
    
    if (targetRow === -1) {
      return { success: false, error: "Sale not found" };
    }
    
    const devCommRate = parseFloat(sh.getRange(targetRow, devCommRateIndex + 1).getValue()) || 0;
    const agentCommRate = parseFloat(sh.getRange(targetRow, agentCommRateIndex + 1).getValue()) || 0;
    
    // Calculate agent's cash advance portion (what actually gets deducted from agent's commission)
    let agentCashAdvance = 0;
    let companyCashAdvance = 0;
    
    if (devCommRate > 0 && cashAdvance > 0) {
      agentCashAdvance = (cashAdvance / devCommRate) * agentCommRate;
      agentCashAdvance = Math.round(agentCashAdvance * 100) / 100;
      companyCashAdvance = cashAdvance - agentCashAdvance;
    } else {
      agentCashAdvance = cashAdvance;
      companyCashAdvance = 0;
    }
    
    // Update cash advance amount
    sh.getRange(targetRow, cashAdvanceIndex + 1).setValue(cashAdvance);
    
    // Store agent cash advance for tracking
    const agentCAIndex = headers.findIndex(h => h === 'Agent Cash Advance');
    if (agentCAIndex !== -1) {
      sh.getRange(targetRow, agentCAIndex + 1).setValue(agentCashAdvance);
    }
    
    // Update Company Net Comm (deduct the company portion only)
    const companyGrossIndex = headers.findIndex(h => h === 'Company Gross Comm (₱)');
    const withholdingIndex = headers.findIndex(h => h === 'Withholding Tax (Company) (₱)');
    const companyNetIndex = headers.findIndex(h => h === 'Company Net Comm (₱)');
    
    let finalCompanyNet = 0;
    
    if (companyGrossIndex !== -1 && withholdingIndex !== -1 && companyNetIndex !== -1) {
      const companyGross = parseFloat(sh.getRange(targetRow, companyGrossIndex + 1).getValue()) || 0;
      const withholding = parseFloat(sh.getRange(targetRow, withholdingIndex + 1).getValue()) || 0;
      
      // Calculate base Company Net Comm (Gross - Withholding)
      let baseCompanyNet = companyGross - withholding;
      
      // Deduct ONLY the company portion from Company Net Comm
      finalCompanyNet = baseCompanyNet - companyCashAdvance;
      finalCompanyNet = Math.max(0, Math.round(finalCompanyNet * 100) / 100);
      
      // Update Company Net Comm
      sh.getRange(targetRow, companyNetIndex + 1).setValue(finalCompanyNet);
    }
    
    // Recalculate DP and Loan amounts (deduct ONLY agent portion from first DP release)
    const milestoneRatioIndex = headers.findIndex(h => h === 'Milestone Ratio');
    const dpTrancheIndex = headers.findIndex(h => h === 'DP Tranche');
    const agentNetCommIndex = headers.findIndex(h => h === 'Agent Net Comm (₱)');
    const dpReleaseAmountIndex = headers.findIndex(h => h === 'DP Release Amount');
    const loanTakeoutAmountIndex = headers.findIndex(h => h === 'Loan Takeout Release Amount');
    
    // Agent Net Comm remains unchanged
    const agentNetComm = parseFloat(sh.getRange(targetRow, agentNetCommIndex + 1).getValue()) || 0;
    const milestoneRatio = sh.getRange(targetRow, milestoneRatioIndex + 1).getValue() || '80/20';
    const dpTranche = parseInt(sh.getRange(targetRow, dpTrancheIndex + 1).getValue()) || 4;
    
    const [dpPercent, loanPercent] = milestoneRatio.split('/').map(p => parseFloat(p));
    const dpTotalAmount = agentNetComm * (dpPercent / 100);
    const loanAmount = agentNetComm * (loanPercent / 100);
    
    let dpPerInstallment = dpTotalAmount / dpTranche;
    let dpReleaseAmounts = Array(dpTranche).fill(dpPerInstallment);
    
    // Deduct ONLY the AGENT portion from the first DP release
    if (dpTranche > 0 && agentCashAdvance > 0) {
      dpReleaseAmounts[0] = dpPerInstallment - agentCashAdvance;
      if (dpReleaseAmounts[0] < 0) dpReleaseAmounts[0] = 0;
    }
    
    // Round all amounts to 2 decimal places
    dpReleaseAmounts = dpReleaseAmounts.map(amt => Math.round(amt * 100) / 100);
    
    const dpReleaseFormatted = dpReleaseAmounts.map(v => v.toFixed(2)).join(', ');
    
    console.log(`DP Release Amounts: ${dpReleaseFormatted}`);
    console.log(`Agent Cash Advance Deducted: ${agentCashAdvance}`);
    
    if (dpReleaseAmountIndex !== -1) {
      sh.getRange(targetRow, dpReleaseAmountIndex + 1).setValue(dpReleaseFormatted);
    }
    if (loanTakeoutAmountIndex !== -1) {
      let finalLoanAmount = loanAmount;
      if (dpTranche === 0 && agentCashAdvance > 0) {
        finalLoanAmount = loanAmount - agentCashAdvance;
        if (finalLoanAmount < 0) finalLoanAmount = 0;
      }
      sh.getRange(targetRow, loanTakeoutAmountIndex + 1).setValue(finalLoanAmount);
    }
    
    return { 
      success: true, 
      agentCashAdvance: agentCashAdvance,
      companyCashAdvance: companyCashAdvance,
      finalCompanyNet: finalCompanyNet || 0,
      agentNetComm: agentNetComm
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function getCashAdvanceBreakdown() {
  const currentYear = new Date().getFullYear();
  const sheetName = currentYear + " Sales";
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  
  if (!sh) {
    console.log(`Sheet ${sheetName} not found`);
    return;
  }
  
  const lastRow = sh.getLastRow();
  
  console.log("\n=== CASH ADVANCE BREAKDOWN ===\n");
  console.log("Row | Sale ID | Agent | Cash Advance | Agent CA | Company CA");
  console.log("-".repeat(80));
  
  for (let i = 2; i <= lastRow; i++) {
    const cashAdvance = parseFloat(sh.getRange(i, 30).getValue()) || 0; // Column AD
    
    if (cashAdvance > 0) {
      const saleId = sh.getRange(i, 21).getValue(); // Column U
      const agent = sh.getRange(i, 2).getValue(); // Column B
      const devRate = parseFloat(sh.getRange(i, 11).getValue()) || 0; // Column K
      const agentRate = parseFloat(sh.getRange(i, 16).getValue()) || 0; // Column P
      
      let agentCA = 0;
      let companyCA = 0;
      
      if (devRate > 0) {
        agentCA = (cashAdvance / devRate) * agentRate;
        agentCA = Math.round(agentCA * 100) / 100;
        companyCA = cashAdvance - agentCA;
      } else {
        agentCA = cashAdvance;
        companyCA = 0;
      }
      
      console.log(`${i} | ${saleId} | ${agent} | ₱${cashAdvance} | ₱${agentCA} | ₱${companyCA}`);
    }
  }
  
  console.log("\n✅ Done!");
}

// ── Update Milestone Configuration ──
function updateMilestoneConfig(saleId, ratio, dpTranche, loanTranche) {
  try {
    const currentYear = new Date().getFullYear();
    const sheetName = currentYear + " Sales";
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sh = ss.getSheetByName(sheetName);
    
    if (!sh) {
      return { success: false, error: "Sales sheet not found" };
    }
    
    // Use column indices directly (0-based)
    // 20: U - Sale ID
    // 19: T - Agent Net Comm (₱)
    // 21: V - Milestone Ratio
    // 22: W - DP Tranche
    // 23: X - Loan Takeout Tranche*
    // 27: AB - DP Release Amount
    // 28: AC - Loan Takeout Release Amount
    // 29: AD - Cash Advance
    // 14: O - Company Net Comm (₱)
    // 11: L - Company Gross Comm (₱)
    // 13: N - Withholding Tax (Company) (₱)
    // 10: K - Developer Comm Rate (%)
    // 15: P - Agent Comm Rate (%)
    
    const lastRow = sh.getLastRow();
    let targetRow = -1;
    
    // Find the row with matching Sale ID
    for (let i = 2; i <= lastRow; i++) {
      const id = sh.getRange(i, 21).getValue(); // Column U (index 20 + 1 for 1-based)
      if (id === saleId) {
        targetRow = i;
        break;
      }
    }
    
    if (targetRow === -1) {
      return { success: false, error: "Sale not found: " + saleId };
    }
    
    // Agent Net Comm from column T (index 19)
    const agentNetComm = parseFloat(sh.getRange(targetRow, 20).getValue()) || 0;
    
    // Parse ratio
    const [dpPercent, loanPercent] = ratio.split('/').map(p => parseFloat(p));
    const dpTotalAmount = agentNetComm * (dpPercent / 100);
    const loanAmount = agentNetComm * (loanPercent / 100);
    
    // Get cash advance from column AD (index 29)
    let cashAdvance = parseFloat(sh.getRange(targetRow, 30).getValue()) || 0;
    
    // Calculate agent's cash advance portion
    let agentCashAdvance = 0;
    if (cashAdvance > 0) {
      const devCommRate = parseFloat(sh.getRange(targetRow, 11).getValue()) || 0; // Column K
      const agentCommRate = parseFloat(sh.getRange(targetRow, 16).getValue()) || 0; // Column P
      
      if (devCommRate > 0) {
        agentCashAdvance = (cashAdvance / devCommRate) * agentCommRate;
        agentCashAdvance = Math.round(agentCashAdvance * 100) / 100;
      } else {
        agentCashAdvance = cashAdvance;
      }
    }
    
    // Calculate DP per installment
    let dpPerInstallment = dpTranche > 0 ? dpTotalAmount / dpTranche : 0;
    let dpReleaseAmounts = Array(dpTranche).fill(dpPerInstallment);
    
    // Deduct agent portion from first DP release
    if (dpTranche > 0 && agentCashAdvance > 0) {
      dpReleaseAmounts[0] = dpPerInstallment - agentCashAdvance;
      if (dpReleaseAmounts[0] < 0) dpReleaseAmounts[0] = 0;
    }
    
    // Round all amounts
    dpReleaseAmounts = dpReleaseAmounts.map(amt => Math.round(amt * 100) / 100);
    const dpReleaseFormatted = dpReleaseAmounts.map(v => v.toFixed(2)).join(', ');
    
    // Update Company Net Comm (deduct company portion only)
    const companyGross = parseFloat(sh.getRange(targetRow, 12).getValue()) || 0; // Column L
    const withholding = parseFloat(sh.getRange(targetRow, 14).getValue()) || 0; // Column N
    const companyCashAdvance = cashAdvance - agentCashAdvance;
    
    let baseCompanyNet = companyGross - withholding;
    let finalCompanyNet = baseCompanyNet - companyCashAdvance;
    finalCompanyNet = Math.max(0, Math.round(finalCompanyNet * 100) / 100);
    sh.getRange(targetRow, 15).setValue(finalCompanyNet); // Column O
    
    // Save milestone configuration
    sh.getRange(targetRow, 22).setValue(ratio); // Column V
    sh.getRange(targetRow, 23).setValue(dpTranche); // Column W
    sh.getRange(targetRow, 24).setValue(loanTranche); // Column X
    sh.getRange(targetRow, 28).setValue(dpReleaseFormatted); // Column AB
    sh.getRange(targetRow, 29).setValue(loanAmount); // Column AC
    
    // Also update the Agent Cash Advance column (AE) if it exists
    try {
      sh.getRange(targetRow, 31).setValue(); // Column AE
    } catch(e) {
      // Column AE might not exist yet
    }
    
    return { success: true };
  } catch (error) {
    console.error("Error in updateMilestoneConfig:", error.toString());
    return { success: false, error: error.toString() };
  }
}

// ── Other preserved functions ──
function initializeDPReleaseAmounts() {
  const currentYear = new Date().getFullYear();
  const sheetName = currentYear + " Sales";
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  
  if (!sh) return { success: false, error: "Sheet not found" };
  
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const agentNetCommIndex = headers.findIndex(h => h === 'Agent Net Comm (₱)');
  const milestoneRatioIndex = headers.findIndex(h => h === 'Milestone Ratio');
  const dpTrancheIndex = headers.findIndex(h => h === 'DP Tranche');
  const dpReleaseAmountIndex = headers.findIndex(h => h === 'DP Release Amount');
  const loanTakeoutAmountIndex = headers.findIndex(h => h === 'Loan Takeout Release Amount');
  
  if (dpReleaseAmountIndex === -1 || loanTakeoutAmountIndex === -1) {
    return { success: false, error: "Required columns not found" };
  }
  
  const lastRow = sh.getLastRow();
  let updatedCount = 0;
  
  for (let i = 2; i <= lastRow; i++) {
    const agentNetComm = parseFloat(sh.getRange(i, agentNetCommIndex + 1).getValue()) || 0;
    const milestoneRatio = sh.getRange(i, milestoneRatioIndex + 1).getValue() || '80/20';
    const dpTranche = parseInt(sh.getRange(i, dpTrancheIndex + 1).getValue()) || 4;
    
    let dpPercent = 80;
    let loanPercent = 20;
    if (milestoneRatio !== '--' && milestoneRatio.includes('/')) {
      const parts = milestoneRatio.split('/');
      dpPercent = parseInt(parts[0]) || 80;
      loanPercent = parseInt(parts[1]) || 20;
    }
    
    const dpTotalAmount = agentNetComm * (dpPercent / 100);
    const loanTakeoutAmount = agentNetComm * (loanPercent / 100);
    const perInstallment = dpTotalAmount / dpTranche;
    
    const dpReleaseValues = Array(dpTranche).fill(perInstallment).map(v => v.toFixed(2)).join(', ');
    
    sh.getRange(i, dpReleaseAmountIndex + 1).setValue(dpReleaseValues);
    sh.getRange(i, loanTakeoutAmountIndex + 1).setValue(loanTakeoutAmount);
    updatedCount++;
  }
  
  return { success: true, updatedCount };
}

function getDPReleaseAmounts(saleId) {
  try {
    const currentYear = new Date().getFullYear();
    const sheetName = currentYear + " Sales";
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sh = ss.getSheetByName(sheetName);
    
    if (!sh) {
      return { success: false, error: "Sales sheet not found" };
    }
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const saleIdColumnIndex = headers.findIndex(h => h === 'Sale ID');
    const dpReleaseAmountIndex = headers.findIndex(h => h === 'DP Release Amount');
    
    if (saleIdColumnIndex === -1 || dpReleaseAmountIndex === -1) {
      return { success: false, error: "Required columns not found" };
    }
    
    const lastRow = sh.getLastRow();
    let targetRow = -1;
    
    for (let i = 2; i <= lastRow; i++) {
      const id = sh.getRange(i, saleIdColumnIndex + 1).getValue();
      if (id === saleId) {
        targetRow = i;
        break;
      }
    }
    
    if (targetRow === -1) {
      return { success: false, error: "Sale not found" };
    }
    
    let releaseAmounts = [];
    const existingValue = sh.getRange(targetRow, dpReleaseAmountIndex + 1).getValue();
    if (existingValue) {
      releaseAmounts = existingValue.toString().split(',').map(v => parseFloat(v.trim()));
    }
    
    return { success: true, amounts: releaseAmounts };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function updateDPReleaseAmounts(saleId, amounts) {
  try {
    const currentYear = new Date().getFullYear();
    const sheetName = currentYear + " Sales";
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sh = ss.getSheetByName(sheetName);
    
    if (!sh) {
      return { success: false, error: "Sales sheet not found" };
    }
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const saleIdColumnIndex = headers.findIndex(h => h === 'Sale ID');
    const dpReleaseAmountIndex = headers.findIndex(h => h === 'DP Release Amount');
    
    if (saleIdColumnIndex === -1 || dpReleaseAmountIndex === -1) {
      return { success: false, error: "Required columns not found" };
    }
    
    const lastRow = sh.getLastRow();
    let targetRow = -1;
    
    for (let i = 2; i <= lastRow; i++) {
      const id = sh.getRange(i, saleIdColumnIndex + 1).getValue();
      if (id === saleId) {
        targetRow = i;
        break;
      }
    }
    
    if (targetRow === -1) {
      return { success: false, error: "Sale not found" };
    }
    
    const formattedValues = amounts.map(v => v.toFixed(2)).join(', ');
    sh.getRange(targetRow, dpReleaseAmountIndex + 1).setValue(formattedValues);
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function updateLoanTakeoutAmount(saleId, amount) {
  try {
    const currentYear = new Date().getFullYear();
    const sheetName = currentYear + " Sales";
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sh = ss.getSheetByName(sheetName);
    
    if (!sh) {
      return { success: false, error: "Sales sheet not found" };
    }
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const saleIdColumnIndex = headers.findIndex(h => h === 'Sale ID');
    const loanTakeoutAmountIndex = headers.findIndex(h => h === 'Loan Takeout Release Amount');
    
    if (saleIdColumnIndex === -1 || loanTakeoutAmountIndex === -1) {
      return { success: false, error: "Required columns not found" };
    }
    
    const lastRow = sh.getLastRow();
    let targetRow = -1;
    
    for (let i = 2; i <= lastRow; i++) {
      const id = sh.getRange(i, saleIdColumnIndex + 1).getValue();
      if (id === saleId) {
        targetRow = i;
        break;
      }
    }
    
    if (targetRow === -1) {
      return { success: false, error: "Sale not found" };
    }
    
    sh.getRange(targetRow, loanTakeoutAmountIndex + 1).setValue(amount);
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function initializeReleasedAndReceivables() {
  const currentYear = new Date().getFullYear();
  const sheetName = currentYear + " Sales";
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  
  if (!sh) return { success: false, error: "Sheet not found" };
  
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const releasedColIndex = headers.findIndex(h => h === 'Released Commission');
  const receivablesColIndex = headers.findIndex(h => h === 'Receivables');
  const statusColIndex = headers.findIndex(h => h === 'Commission Status');
  const agentNetCommIndex = headers.findIndex(h => h === 'Agent Net Comm (₱)');
  
  if (releasedColIndex === -1 || receivablesColIndex === -1 || statusColIndex === -1) {
    return { success: false, error: "Columns not found" };
  }
  
  const lastRow = sh.getLastRow();
  let updatedCount = 0;
  
  for (let i = 2; i <= lastRow; i++) {
    const agentNetComm = parseFloat(sh.getRange(i, agentNetCommIndex + 1).getValue()) || 0;
    const releasedAmount = parseFloat(sh.getRange(i, releasedColIndex + 1).getValue()) || 0;
    const receivablesAmount = agentNetComm - releasedAmount;
    
    let percentage = '0%';
    if (agentNetComm > 0) {
      const percentValue = (releasedAmount / agentNetComm) * 100;
      percentage = percentValue.toFixed(1) + '%';
      if (percentValue === 100) percentage = '100%';
    }
    
    sh.getRange(i, receivablesColIndex + 1).setValue(receivablesAmount);
    sh.getRange(i, statusColIndex + 1).setValue(percentage);
    updatedCount++;
  }
  
  return { success: true, updatedCount };
}

function updateCommissionRelease(saleId, releaseType, amount, isReleased) {
  try {
    const currentYear = new Date().getFullYear();
    const sheetName = currentYear + " Sales";
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sh = ss.getSheetByName(sheetName);
    
    if (!sh) {
      return { success: false, error: "Sales sheet not found" };
    }
    
    const lastRow = sh.getLastRow();
    let targetRow = -1;
    
    // Find the row with matching Sale ID (column U, index 20)
    for (let i = 2; i <= lastRow; i++) {
      const id = sh.getRange(i, 21).getValue();
      if (id === saleId) {
        targetRow = i;
        break;
      }
    }
    
    if (targetRow === -1) {
      return { success: false, error: "Sale not found" };
    }
    
    // Agent Net Comm from column T (index 19)
    const agentNetComm = parseFloat(sh.getRange(targetRow, 20).getValue()) || 0;
    
    // Handle 'full' release type - release all commissions at once
    if (releaseType === 'full') {
      const currentReleased = Math.max(0, Math.min(amount, agentNetComm));
      const receivables = agentNetComm - currentReleased;
      const statusText = '100%';
      
      sh.getRange(targetRow, 25).setValue(currentReleased); // Column Y
      sh.getRange(targetRow, 26).setValue(receivables); // Column Z
      sh.getRange(targetRow, 27).setValue(statusText); // Column AA
      
      return { 
        success: true, 
        released: currentReleased, 
        receivables: receivables,
        percentage: statusText
      };
    }
    
    // Handle regular releases (dp, loan, cashadvance)
    // Get current released from column Y (index 24)
    let currentReleased = parseFloat(sh.getRange(targetRow, 25).getValue()) || 0;
    
    if (isReleased) {
      // Add the amount to released commission
      currentReleased += amount;
    } else {
      // Subtract the amount from released commission
      currentReleased -= amount;
    }
    
    // Ensure no negative values and not exceeding total commission
    currentReleased = Math.max(0, Math.min(currentReleased, agentNetComm));
    
    // Calculate receivables
    const receivables = agentNetComm - currentReleased;
    
    // Calculate percentage
    let statusText = '0%';
    if (agentNetComm > 0) {
      const percentage = (currentReleased / agentNetComm) * 100;
      statusText = percentage.toFixed(1) + '%';
      if (percentage >= 99.99) statusText = '100%';
    }
    
    // Update all columns
    sh.getRange(targetRow, 25).setValue(currentReleased); // Column Y
    sh.getRange(targetRow, 26).setValue(receivables); // Column Z
    sh.getRange(targetRow, 27).setValue(statusText); // Column AA
    
    return { 
      success: true, 
      released: currentReleased, 
      receivables: receivables,
      percentage: statusText
    };
  } catch (error) {
    return { success: false, error: error.toString() };
  }
}

function generateCommissionSlip(saleId) {
  try {
    const currentYear = new Date().getFullYear();
    const sheetName = currentYear + " Sales";
    const commissions = getCommissions(currentYear);
    const sale = commissions.find(s => s['Sale ID'] === saleId);
    
    if (!sale) {
      return { success: false, error: "Sale not found: " + saleId };
    }
    
    const slipData = {
      success: true,
      slip: {
        saleId: sale['Sale ID'],
        reservationDate: sale['Reservation Date'],
        salesAgent: sale['Sales Agent'],
        developer: sale['Developer'],
        unitId: sale['Unit ID'],
        computedTSP: sale['Computed TSP'] || 0,
        companyGrossComm: sale['Company Gross Comm (₱)'] || 0,
        vatAmount: sale['Net VAT (₱)'] || 0,
        companyWithholdingTax: sale['Withholding Tax (Company) (₱)'] || 0,
        companyNetComm: sale['Company Net Comm (₱)'] || 0,
        agentGrossComm: sale['Agent Gross Comm (₱)'] || 0,
        vatShareAmount: sale['VAT Share (Agent) (₱)'] || 0,
        agentWithholdingTax: sale['Withholding Tax (Agent) (₱)'] || 0,
        agentNetComm: sale['Agent Net Comm (₱)'] || 0
      }
    };
    
    return slipData;
  } catch (error) {
    console.error("Error in generateCommissionSlip:", error.toString());
    return { success: false, error: error.toString() };
  }
}

// ── Get Dashboard Statistics with CORRECT Tax Formula and Company Net ──
function getDashboardStats(yearFilter) {
  let allSales = [];
  
  // Get sales data based on filter
  if (yearFilter === 'all' || !yearFilter) {
    allSales = getAllSalesAcrossYears();
  } else {
    const year = parseInt(yearFilter);
    if (!isNaN(year)) {
      allSales = getSales(year);
      if (allSales.length === 0) {
        allSales = getCommissionsForYear(year);
      }
    } else {
      allSales = getAllSalesAcrossYears();
    }
  }
  
  Logger.log(`Year filter: ${yearFilter}, Sales found: ${allSales.length}`);
  
  if (allSales.length === 0) {
    return {
      totalSales: 0,
      totalCompanyGross: 0,
      totalTax: 0,
      totalNetCommission: 0,  // Changed from totalAgentNet
      activeAgents: 0,
      totalDevelopers: 0,
      topAgents: [],
      recentSales: []
    };
  }
  
  const agents = getAgents();
  const devs = getDevelopers();
  
  let totalSales = 0;
  let totalCompanyGross = 0;
  let totalTax = 0;
  let totalNetCommission = 0;  // Changed from totalAgentNet
  const agentMap = {};
  const recentSalesList = [];
  
  allSales.forEach(s => {
    totalSales++;
    
    const gross = parseFloat(s['Company Gross Comm (₱)']) || 0;
    const netVat = parseFloat(s['Net VAT (₱)']) || 0;
    const withholding = parseFloat(s['Withholding Tax (Company) (₱)']) || 0;
    
    // ✅ CORRECTED: Use Company Net Comm (Column O, index 14)
    const companyNet = parseFloat(s['Company Net Comm (₱)']) || 0;
    
    // Get agent net for top agents and recent sales
    const agentNet = parseFloat(s['Agent Net Comm (₱)']) || 0;
    
    totalCompanyGross += gross;
    
    // ✅ CORRECT: Total Tax = (Gross - Net VAT) + Withholding
    const actualVat = gross - netVat;
    totalTax += actualVat + withholding;
    
    // ✅ CORRECTED: Net Commission = Company Net Comm (Column O)
    totalNetCommission += companyNet;
    
    const agentName = s['Sales Agent'];
    if (agentName) {
      if (!agentMap[agentName]) {
        agentMap[agentName] = { name: agentName, net: 0, count: 0 };
      }
      // Top agents should still use Agent Net Comm
      agentMap[agentName].net += agentNet;
      agentMap[agentName].count += 1;
    }
    
    if (s['Reservation Date']) {
      recentSalesList.push({
        'Sales Agent': s['Sales Agent'] || 'Unknown',
        'Project': s['Project'] || s['Developer'] || 'Unknown',
        'Agent Net Comm': agentNet,
        'Reservation Date': s['Reservation Date']
      });
    }
  });
  
  const topAgents = Object.values(agentMap)
    .sort((a, b) => b.net - a.net)
    .slice(0, 5);
  
  const recentSales = recentSalesList
    .sort((a, b) => {
      const dateA = new Date(a['Reservation Date']);
      const dateB = new Date(b['Reservation Date']);
      return dateB - dateA;
    })
    .slice(0, 5)
    .map(s => ({
      'Sales Agent': s['Sales Agent'],
      'Project': s['Project'],
      'Agent Net Comm': s['Agent Net Comm']
    }));
  
  const activeAgents = agents.filter(a => a['Status'] === 'Active').length;
  
  // DEBUG: Log the calculated values
  Logger.log(`=== DASHBOARD CALCULATIONS ===`);
  Logger.log(`Total Sales: ${totalSales}`);
  Logger.log(`Total Gross Comm: ${totalCompanyGross}`);
  Logger.log(`Total Tax (Actual VAT + Withholding): ${totalTax}`);
  Logger.log(`Total Net Commission (Company Net - Column O): ${totalNetCommission}`);
  Logger.log(`Verification: Gross - Tax = ${totalCompanyGross - totalTax}`);
  Logger.log(`Expected Net (Column O sum): ${totalNetCommission}`);
  
  return {
    totalSales: totalSales,
    totalCompanyGross: totalCompanyGross,
    totalTax: totalTax,
    totalNetCommission: totalNetCommission,  // Changed from totalAgentNet
    activeAgents: activeAgents,
    totalDevelopers: devs.length,
    topAgents: topAgents,
    recentSales: recentSales
  };
}

// ── Test Function to Verify Tax Calculation using Company Net ──
function testTaxCalculation(year) {
  const data = getDashboardStats(year || '2026');
  
  Logger.log(`=== Tax Calculation Test for ${year || '2026'} ===`);
  Logger.log(`Total Sales: ${data.totalSales}`);
  Logger.log(`Total Gross Comm: ${data.totalCompanyGross}`);
  Logger.log(`Total Tax: ${data.totalTax}`);
  Logger.log(`Total Net Comm (Company - Column O): ${data.totalNetCommission}`);
  
  // Verify: Gross - Tax should equal Company Net
  const calculatedNet = data.totalCompanyGross - data.totalTax;
  const actualNet = data.totalNetCommission;
  const difference = Math.abs(calculatedNet - actualNet);
  
  Logger.log(`Gross - Tax = ${calculatedNet}`);
  Logger.log(`Actual Company Net (Column O sum) = ${actualNet}`);
  Logger.log(`Difference: ${difference}`);
  
  if (difference < 0.01) {
    Logger.log('✅ Tax calculation is CORRECT! (Gross - Tax = Company Net)');
  } else {
    Logger.log('⚠️ Tax calculation may need review');
    Logger.log(`Gross - Tax (${calculatedNet}) does not equal Company Net (${actualNet})`);
  }
  
  return data;
}

// ── Get ALL Sales Across All Years using 0-based Column Indexing ──
function getAllSalesAcrossYears() {
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const allSheets = ss.getSheets();
  const allSales = [];
  const salesSheetRegex = /^\d{4} Sales$/;
  
  for (let i = 0; i < allSheets.length; i++) {
    const sheet = allSheets[i];
    const sheetName = sheet.getName();
    
    if (salesSheetRegex.test(sheetName)) {
      Logger.log(`Reading data from: ${sheetName}`);
      
      if (sheet.getLastRow() > 1) {
        const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 21).getValues();
        
        rows.forEach(row => {
          let reservationDate = row[0];
          if (reservationDate instanceof Date) {
            reservationDate = Utilities.formatDate(reservationDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
          }
          
          if (row[20]) {
            allSales.push({
              'Reservation Date': reservationDate || '',
              'Sales Agent': row[1] || '',
              'Developer': row[2] || '',
              'Project': row[3] || '',
              'Unit ID': row[4] || '',
              "Buyer's Name": row[5] || '',
              'DP Term': row[6] || '',
              'ACTIVE SALES AMOUNT': parseFloat(row[7]) || 0,
              'Computed TSP': parseFloat(row[9]) || 0,
              'Company Gross Comm (₱)': parseFloat(row[11]) || 0,
              'Net VAT (₱)': parseFloat(row[12]) || 0,  // Net AFTER VAT deduction
              'Withholding Tax (Company) (₱)': parseFloat(row[13]) || 0,
              'Company Net Comm (₱)': parseFloat(row[14]) || 0,
              'Agent Gross Comm (₱)': parseFloat(row[16]) || 0,
              'VAT Share (Agent) (₱)': parseFloat(row[17]) || 0,
              'Withholding Tax (Agent) (₱)': parseFloat(row[18]) || 0,
              'Agent Net Comm (₱)': parseFloat(row[19]) || 0,
              'Developer Comm Rate (%)': parseFloat(row[10]) || 0,
              'Agent Comm Rate (%)': parseFloat(row[15]) || 0,
              'Sale ID': row[20] || ''
            });
          }
        });
      }
    }
  }
  
  Logger.log(`Total sales records found: ${allSales.length}`);
  return allSales;
}

/**
 * Gets commissions data from ALL years
 * Returns combined array of commission records
 */
function getAllCommissions() {
  const allSales = getAllSalesAcrossYears();
  
  // Transform sales data into commission format
  return allSales.map(sale => {
    const agentNetComm = parseFloat(sale['Agent Net Comm (₱)']) || 0;
    let releasedAmount = parseFloat(sale['Released Commission']) || 0;
    
    let receivablesAmount = agentNetComm - releasedAmount;
    let commissionPercentage = '0%';
    
    if (agentNetComm > 0) {
      const percentValue = (releasedAmount / agentNetComm) * 100;
      commissionPercentage = percentValue.toFixed(1) + '%';
      if (percentValue === 100) commissionPercentage = '100%';
    }
    
    let dpReleaseAmounts = [];
    const dpReleaseRaw = sale['DP Release Amount'] || '';
    if (dpReleaseRaw) {
      dpReleaseAmounts = dpReleaseRaw.toString().split(',').map(v => parseFloat(v.trim()));
    }
    
    const loanTakeoutAmount = parseFloat(sale['Loan Takeout Release Amount']) || 0;
    const cashAdvance = parseFloat(sale['Cash Advance']) || 0;
    
    return {
      'Sale ID': sale['Sale ID'] || '',
      'Reservation Date': sale['Reservation Date'] || '',
      'Sales Agent': sale['Sales Agent'] || '',
      'Developer': sale['Developer'] || '',
      'Project': sale['Project'] || '',
      'Unit ID': sale['Unit ID'] || '',
      'DP Term': sale['DP Term'] || '',
      'Computed TSP': sale['Computed TSP'] || 0,
      'Developer Comm Rate (%)': sale['Developer Comm Rate (%)'] || 0,
      'Company Gross Comm (₱)': sale['Company Gross Comm (₱)'] || 0,
      'Net VAT (₱)': sale['Net VAT (₱)'] || 0,
      'Withholding Tax (Company) (₱)': sale['Withholding Tax (Company) (₱)'] || 0,
      'Company Net Comm (₱)': sale['Company Net Comm (₱)'] || 0,
      'Agent Comm Rate (%)': sale['Agent Comm Rate (%)'] || 0,
      'Agent Gross Comm (₱)': sale['Agent Gross Comm (₱)'] || 0,
      'VAT Share (Agent) (₱)': sale['VAT Share (Agent) (₱)'] || 0,
      'Withholding Tax (Agent) (₱)': sale['Withholding Tax (Agent) (₱)'] || 0,
      'Agent Net Comm (₱)': agentNetComm,
      'Milestone Ratio': sale['Milestone Ratio'] || '--',
      'DP Tranche': sale['DP Tranche'] || '--',
      'Loan Takeout Tranche': sale['Loan Takeout Tranche'] || '--',
      'Released Commission': releasedAmount,
      'Receivables': receivablesAmount,
      'Commission Status': commissionPercentage,
      'DP Release Amounts': dpReleaseAmounts,
      'Loan Takeout Amount': loanTakeoutAmount,
      'Cash Advance': cashAdvance
    };
  });
}

function getYearRange() {
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const allSheets = ss.getSheets();
  const salesSheetRegex = /^\d{4} Sales$/;
  const years = [];
  
  for (let i = 0; i < allSheets.length; i++) {
    const sheet = allSheets[i];
    const sheetName = sheet.getName();
    if (salesSheetRegex.test(sheetName)) {
      const yearMatch = sheetName.match(/(\d{4})/);
      if (yearMatch) {
        years.push(parseInt(yearMatch[1]));
      }
    }
  }
  
  if (years.length === 0) return "No data";
  if (years.length === 1) return years[0].toString();
  
  years.sort((a, b) => a - b);
  return `${years[0]} - ${years[years.length - 1]}`;
}

// ── Get Available Years from Sales Sheets ──
function getAvailableYears() {
  const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
  const allSheets = ss.getSheets();
  const salesSheetRegex = /^\d{4} Sales$/;
  const years = [];
  
  for (let i = 0; i < allSheets.length; i++) {
    const sheet = allSheets[i];
    const sheetName = sheet.getName();
    if (salesSheetRegex.test(sheetName)) {
      const yearMatch = sheetName.match(/(\d{4})/);
      if (yearMatch) {
        years.push(parseInt(yearMatch[1]));
      }
    }
  }
  
  return years;
}

function initializeSaleIds() {
  const currentYear = new Date().getFullYear();
  const sheetName = currentYear + " Sales";
  const result = ensureSaleIds(sheetName);
  console.log(`Initialized ${result.updatedCount} Sale IDs`);
  return result;
}

function updateSaleFromSlip(data) {
  try {
    const currentYear = new Date().getFullYear();
    const sheetName = currentYear + " Sales";
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    const sh = ss.getSheetByName(sheetName);
    
    if (!sh) {
      return { success: false, error: "Sales sheet not found" };
    }
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const saleIdColumnIndex = headers.findIndex(h => h === 'Sale ID');
    
    if (saleIdColumnIndex === -1) {
      return { success: false, error: "Sale ID column not found" };
    }
    
    const lastRow = sh.getLastRow();
    let targetRow = -1;
    
    for (let i = 2; i <= lastRow; i++) {
      const saleId = sh.getRange(i, saleIdColumnIndex + 1).getValue();
      if (saleId === data.saleId) {
        targetRow = i;
        break;
      }
    }
    
    if (targetRow === -1) {
      return { success: false, error: "Sale not found: " + data.saleId };
    }
    
    const updates = [
      { header: 'Reservation Date', value: data.reservationDate },
      { header: 'Sales Agent', value: data.salesAgent },
      { header: 'Developer', value: data.developer },
      { header: 'Unit ID', value: data.unitId },
      { header: 'Computed TSP', value: data.computedTSP },
      { header: 'Company Gross Comm (₱)', value: data.companyGrossComm },
      { header: 'Net VAT (₱)', value: data.vatAmount },
      { header: 'Withholding Tax (Company) (₱)', value: data.companyWithholdingTax },
      { header: 'Company Net Comm (₱)', value: data.companyNetComm },
      { header: 'Agent Gross Comm (₱)', value: data.agentGrossComm },
      { header: 'VAT Share (Agent) (₱)', value: data.vatShareAmount },
      { header: 'Withholding Tax (Agent) (₱)', value: data.agentWithholdingTax },
      { header: 'Agent Net Comm (₱)', value: data.agentNetComm }
    ];
    
    for (const update of updates) {
      const colIndex = headers.findIndex(h => h === update.header);
      if (colIndex !== -1) {
        sh.getRange(targetRow, colIndex + 1).setValue(update.value);
      }
    }
    
    return { success: true };
    
  } catch (error) {
    console.error("Error in updateSaleFromSlip:", error.toString());
    return { success: false, error: error.toString() };
  }
}

// ── Include SalesDashboard.html ──
function includeSalesDashboard() {
  return HtmlService.createHtmlOutputFromFile('SalesDashboard').getContent();
}

// ── Sales Dashboard Backend Functions ──
// Add these to your CMS code.gs file

const SALES_SHEET_ID = "1A-oSVCbEdDjwSREOo4jP2sVaby9Wsun-D_YbKQu9cgI";
const ALLOWED_TEAMS = [
  "Blazing SPARCS - Inhouse",
  "Flameborn Champions",
  "Shining Phoenix",
  "Feisty Heroine",
  "Fiery Achievers"
];

// Matches a signed-in user's display name against the Agent (H) column of a
// sales row, so "My ..." views (My Recent Sales, My Performance, My Top
// Projects/Developers) only ever surface the signed-in agent's own deals.
//
// This USED to also match the Supervisor (J) / Manager (K) columns, on the
// theory that a supervisor's own "My Performance" should include their
// team's sales too — but that meant ANY agent whose name ever ended up in
// those columns (autofill, a past assignment, a data-entry slip) would see
// other agents' sales mixed into what's branded as their own, which is the
// scoping bug reported. Company-wide roles (Admin/Supervisor/Manager, see
// pa_isCompanyWideRole_ in PortalData.gs) already get full company-wide
// visibility elsewhere (Dashboard/Admin), so "My Performance" can safely
// stay Agent-only for everyone. Case/whitespace-insensitive.
function isPersonMatch_(r, person) {
  if (!person) return false;

  const salesAgent = String(r[8] || '').trim();

  if (!salesAgent) return false;

  // Use the same flexible agent-name matching
  // used by Report Card / Commission.
  return agentNameMatches_(salesAgent, person);
}

// Robustly parses a sales-amount cell into a number. Google Sheets cells
// entered by hand can end up as text — with a currency symbol ("₱150,000.00"),
// thousands separators ("150,000"), or stray spaces — instead of a plain
// number. parseFloat() alone chokes on a leading "₱" (returns NaN) and stops
// at the first comma (parseFloat("150,000") is 150, not 150000), and a
// strict `typeof v === 'number'` check drops the row entirely. Both of
// those were silently zeroing out or excluding real sales, which is why
// Report Card and My Performance totals were undercounting for some
// accounts. This strips everything but digits/decimal point/minus sign
// before parsing, so it handles numbers and stringified currency alike.
function toAmount_(v) {
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  if (v === null || v === undefined || v === '') return 0;
  var n = parseFloat(String(v).replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function getSalesData(month, year, team, person) {
  var sheetName = "Form Responses " + year;
  var ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  
  if (!sh) {
    return { 
      reservationSales: 0, 
      activeSales: 0, 
      cancelledSales: 0, 
      activePercent: "0", 
      cancelledPercent: "0" 
    };
  }
  
  var rows = sh.getDataRange().getValues();
  rows.shift(); // Remove header
  
  var reservation = 0, active = 0, cancelled = 0;
  
  rows.forEach(function(r) {
    var dateVal = r[2];    // Column C - Reservation Date
    var rowTeam = r[11];   // Column L - Team
    var resVal = toAmount_(r[5]);  // Column F - Reservation Sales
    var cancelledVal = toAmount_(r[15]); // Column P - Cancelled
    var activeVal = toAmount_(r[16]); // Column Q - Active Sales
    
    var rowMatch = person
      ? isPersonMatch_(r, person)
      : (team ? rowTeam === team : ALLOWED_TEAMS.indexOf(rowTeam) !== -1);
    
    if (dateVal && rowMatch) {
      var date = new Date(dateVal);
      var rowYear = date.getFullYear();
      
      if (rowYear === year && isMonthMatch(month, date)) {
        reservation += resVal;
        active += activeVal;
        cancelled += cancelledVal;
      }
    }
  });

  var total = active + cancelled;
  var activePct = total ? (active / total) * 100 : 0;
  var cancelledPct = total ? (cancelled / total) * 100 : 0;

  return {
    reservationSales: reservation,
    activeSales: active,
    cancelledSales: cancelled,
    activePercent: activePct.toFixed(1),
    cancelledPercent: cancelledPct.toFixed(1)
  };
}

function getRecentSales(month, year, team, person) {
  var sheetName = "Form Responses " + year;
  var ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  
  if (!sh) return [];
  
  var rows = sh.getDataRange().getValues();
  rows.shift(); // Remove header
  
  var results = [];
  
  rows.forEach(function(r) {
    if (!r[2] || !r[16]) return;
    var activeAmount = toAmount_(r[16]);
    if (activeAmount <= 0) return;
    
    var rowTeam = r[11];
    var rowMatch = person
      ? isPersonMatch_(r, person)
      : (team ? rowTeam === team : ALLOWED_TEAMS.indexOf(rowTeam) !== -1);
    if (!rowMatch) return;
    
    var date = new Date(r[2]);
    var rowYear = date.getFullYear();
    if (rowYear === year && isMonthMatch(month, date)) {
      results.push({
        date: Utilities.formatDate(new Date(r[2]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        amount: toAmount_(r[16]),
        project: r[6] || '',
        developer: r[7] || '',
        agent: r[8] || '',
        team: r[11] || '',
        agentPic: r[23] || '',
        projectLogo: r[22] || ''
      });
    }
  });
  
  results.sort(function(a, b) {
    return new Date(b.date) - new Date(a.date);
  });
  
  return results.slice(0, 5);
}

function getTopAgents(month, year, team) {
  var sheetName = "Form Responses " + year;
  var ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  
  if (!sh) return [];
  
  var rows = sh.getDataRange().getValues();
  rows.shift();
  
  var agentMap = {};
  
  rows.forEach(function(r) {
    if (!r[2]) return;
    var rowTeam = r[11];
    var teamMatch = team ? rowTeam === team : ALLOWED_TEAMS.indexOf(rowTeam) !== -1;
    if (!teamMatch) return;
    
    var date = new Date(r[2]);
    if (date.getFullYear() !== year || !isMonthMatch(month, date)) return;
    
    var agent = r[8];
    var amount = toAmount_(r[16]);
    var picture = r[23];
    
    if (!agent || amount === 0) return;
    if (!agentMap[agent]) agentMap[agent] = { name: agent, amount: 0, picture: picture };
    agentMap[agent].amount += amount;
  });
  
  var result = [];
  for (var key in agentMap) {
    result.push(agentMap[key]);
  }
  result.sort(function(a, b) { return b.amount - a.amount; });
  return result.slice(0, 5);
}

function getTopSupervisors(month, year, team) {
  var sheetName = "Form Responses " + year;
  var ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  
  if (!sh) return [];
  
  var rows = sh.getDataRange().getValues();
  rows.shift();
  
  var supervisorMap = {};
  
  rows.forEach(function(r) {
    if (!r[2]) return;
    var rowTeam = r[11];
    var teamMatch = team ? rowTeam === team : ALLOWED_TEAMS.indexOf(rowTeam) !== -1;
    if (!teamMatch) return;
    
    var date = new Date(r[2]);
    if (date.getFullYear() !== year || !isMonthMatch(month, date)) return;
    
    var supervisor = r[9];
    var amount = toAmount_(r[16]);
    var picture = r[24];
    
    if (!supervisor || supervisor.trim() === "" || supervisor === "N/A" || amount === 0) return;
    if (!supervisorMap[supervisor]) supervisorMap[supervisor] = { name: supervisor, amount: 0, picture: picture };
    supervisorMap[supervisor].amount += amount;
  });
  
  var result = [];
  for (var key in supervisorMap) {
    result.push(supervisorMap[key]);
  }
  result.sort(function(a, b) { return b.amount - a.amount; });
  return result.slice(0, 5);
}

function getTopManagers(month, year, team) {
  var sheetName = "Form Responses " + year;
  var ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  
  if (!sh) return [];
  
  var rows = sh.getDataRange().getValues();
  rows.shift();
  
  var managerMap = {};
  
  rows.forEach(function(r) {
    if (!r[2]) return;
    var rowTeam = r[11];
    var teamMatch = team ? rowTeam === team : ALLOWED_TEAMS.indexOf(rowTeam) !== -1;
    if (!teamMatch) return;
    
    var date = new Date(r[2]);
    if (date.getFullYear() !== year || !isMonthMatch(month, date)) return;
    
    var manager = r[10];
    var amount = toAmount_(r[16]);
    var picture = r[25];
    
    if (!manager || amount === 0) return;
    if (!managerMap[manager]) managerMap[manager] = { name: manager, amount: 0, picture: picture };
    managerMap[manager].amount += amount;
  });
  
  var result = [];
  for (var key in managerMap) {
    result.push(managerMap[key]);
  }
  result.sort(function(a, b) { return b.amount - a.amount; });
  return result.slice(0, 5);
}

function getTopProjects(month, year, team, person) {
  var sheetName = "Form Responses " + year;
  var ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  
  if (!sh) return [];
  
  var rows = sh.getDataRange().getValues();
  rows.shift();
  
  var projectMap = {};
  
  rows.forEach(function(r) {
    if (!r[2]) return;
    var rowTeam = r[11];
    var rowMatch = person
      ? isPersonMatch_(r, person)
      : (team ? rowTeam === team : ALLOWED_TEAMS.indexOf(rowTeam) !== -1);
    if (!rowMatch) return;
    
    var date = new Date(r[2]);
    if (date.getFullYear() !== year || !isMonthMatch(month, date)) return;
    
    var project = r[6];
    var amount = toAmount_(r[16]);
    var picture = r[22];
    
    if (!project || amount === 0) return;
    if (!projectMap[project]) projectMap[project] = { name: project, amount: 0, picture: picture, homes: 0 };
    projectMap[project].amount += amount;
    projectMap[project].homes += 1;
  });
  
  var result = [];
  for (var key in projectMap) {
    result.push(projectMap[key]);
  }
  result.sort(function(a, b) { return b.amount - a.amount; });
  return result.slice(0, 5);
}

function getTopDevelopers(month, year, team, person) {
  var sheetName = "Form Responses " + year;
  var ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  var sh = ss.getSheetByName(sheetName);
  
  if (!sh) return [];
  
  var rows = sh.getDataRange().getValues();
  rows.shift();
  
  var developerMap = {};
  
  rows.forEach(function(r) {
    if (!r[2]) return;
    var rowTeam = r[11];
    var rowMatch = person
      ? isPersonMatch_(r, person)
      : (team ? rowTeam === team : ALLOWED_TEAMS.indexOf(rowTeam) !== -1);
    if (!rowMatch) return;
    
    var date = new Date(r[2]);
    if (date.getFullYear() !== year || !isMonthMatch(month, date)) return;
    
    var developer = r[7];
    var amount = toAmount_(r[16]);
    var picture = r[21];
    
    if (!developer || amount === 0) return;
    if (!developerMap[developer]) developerMap[developer] = { name: developer, amount: 0, picture: picture, homes: 0 };
    developerMap[developer].amount += amount;
    developerMap[developer].homes += 1;
  });
  
  var result = [];
  for (var key in developerMap) {
    result.push(developerMap[key]);
  }
  result.sort(function(a, b) { return b.amount - a.amount; });
  return result.slice(0, 5);
}

function isMonthMatch(month, date) {
  var rowMonth = date.getMonth() + 1;
  
  if (month === 0) return true; // All Year
  if (month > 0) return rowMonth === month; // Specific month
  
  if (month === -5) return rowMonth >= 1 && rowMonth <= 6; // January-June
  if (month === -6) return rowMonth >= 7 && rowMonth <= 12; // July-December
  
  var quarter = -month;
  var startMonth = (quarter - 1) * 3 + 1;
  var endMonth = startMonth + 2;
  
  return rowMonth >= startMonth && rowMonth <= endMonth;
}

function testDashboardData(year) {
  Logger.log(`=== Testing Dashboard for ${year || 'all'} ===`);
  const data = getDashboardStats(year || 'all');
  
  Logger.log(`Total Sales: ${data.totalSales}`);
  Logger.log(`Total Gross Comm: ${data.totalCompanyGross}`);
  Logger.log(`Total Tax: ${data.totalTax}`);
  Logger.log(`Total Net Comm: ${data.totalAgentNet}`);
  Logger.log(`Active Agents: ${data.activeAgents}`);
  Logger.log(`Total Developers: ${data.totalDevelopers}`);
  Logger.log(`Top Agents: ${JSON.stringify(data.topAgents)}`);
  Logger.log(`Recent Sales: ${JSON.stringify(data.recentSales)}`);
  
  return data;
}

function testSalesDashboardData() {
  const year = 2026; // Change to a year with data
  const month = 0; // All Year
  const team = ""; // All Teams
  
  const data = getSalesData(month, year, team);
  Logger.log("Sales Data: " + JSON.stringify(data));
  
  const recent = getRecentSales(month, year, team);
  Logger.log("Recent Sales: " + JSON.stringify(recent));
  
  const agents = getTopAgents(month, year, team);
  Logger.log("Top Agents: " + JSON.stringify(agents));
}

// ══════════════════════════════════════════════════════════
// ── Report Card Backend Functions ──
// ══════════════════════════════════════════════════════════

// Define the full list of teams for the Inner SPARC group report card.
const INNER_SPARC_TEAMS = [
  'Blazing SPARCS - Inhouse',
  'Feisty Heroine',
  'Fiery Achievers',
  'Flameborn Champions',
  'Shining Phoenix'
];

/**
 * Gets the spreadsheet ID and sheet name for a given year (Report Card).
 * @param {string} year The year (2020-2026).
 * @return {object} An object containing spreadsheetId and sheetName.
 */
function getDataSource(year) {
  // For 2025 and 2026, use the original spreadsheet ID
  if (year === '2025' || year === '2026') {
    return {
      spreadsheetId: '1A-oSVCbEdDjwSREOo4jP2sVaby9Wsun-D_YbKQu9cgI',
      sheetName: `Form Responses ${year}`
    };
  }

  // For 2020-2024, use the archive spreadsheet ID
  return {
    spreadsheetId: '1Qo2twILiv-jrggqZoCljhyqG4XP8hzJdfzw7glCrN2s',
    sheetName: `Form Responses ${year}`
  };
}

// Normalizes a name for comparison/grouping: trims, collapses internal
// whitespace runs to a single space, and lowercases. The sales sheet is
// hand-entered, so the same person or project can appear as "Juan Dela
// Cruz", " Juan Dela Cruz", or "Juan  Dela Cruz" across different rows —
// a strict === match (or map key) would silently drop or fragment those
// rows instead of counting them together, which is what was causing
// Report Card to undercount units/sales for some accounts.
function normName_(s) {
  return String(s || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

/**
 * Retrieves a sorted, unique list of names based on position and team filters (Report Card).
 *
 * @param {string} position The position to filter by ('Agent', 'Supervisor', 'Manager', 'Inner SPARC').
 * @param {string} year The year to filter by ('2020' through '2026').
 * @return {string} A JSON string containing the list of names.
 */
function getNames(position, year) {
  if (position === 'Inner SPARC') {
    return JSON.stringify([]);
  }

  const { spreadsheetId, sheetName } = getDataSource(year);
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

  if (!sheet) {
    return JSON.stringify([]);
  }

  const data = sheet.getDataRange().getValues();

  const allowedTeams = [
    'Blazing SPARCS - Inhouse',
    'Feisty Heroine',
    'Fiery Achievers',
    'Flameborn Champions'
  ];

  let nameColumnIndex;
  if (position === 'Agent') {
    nameColumnIndex = 8;
  } else if (position === 'Supervisor') {
    nameColumnIndex = 9;
  } else if (position === 'Manager') {
    nameColumnIndex = 10;
  } else {
    return JSON.stringify([]);
  }

  const names = data.slice(1)
    .filter(row => {
      const team = row[11];
      const isTeamAllowed = allowedTeams.includes(team);
      const nameValue = row[nameColumnIndex];
      const isNameValid = nameValue && nameValue !== 'N/A';
      const isColumnCNotNull = row[2] !== null && row[2] !== '';
      const isSalesAmountValid = toAmount_(row[16]) > 0;
      return isTeamAllowed && isNameValid && isColumnCNotNull && isSalesAmountValid;
    })
    .map(row => row[nameColumnIndex])
    .filter(name => name)
    .reduce((unique, name) => {
      const key = normName_(name);
      // Keep the first-seen spelling as the display value, but dedupe on
      // the normalized form so whitespace/casing variants of the same
      // person collapse into a single entry instead of two.
      if (!unique.some(n => normName_(n) === key)) {
        unique.push(String(name).trim());
      }
      return unique;
    }, []);

  return JSON.stringify(names.sort());
}

/**
 * Calculates the leading price range for a selected person/group based on their sales count (Report Card).
 *
 * @param {string} name The name of the person (ignored if position is 'Inner SPARC').
 * @param {string} position The position of the person/group.
 * @param {string} year The year to filter by ('2020' through '2026').
 * @return {string} A JSON string containing the leading range label and all range counts.
 */
function getLeadingPriceRange(name, position, year) {
  const { spreadsheetId, sheetName } = getDataSource(year);
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

  if (!sheet) {
    return JSON.stringify({ leadingRange: "No Data", rangeCounts: {}, sortedRanges: [] });
  }

  const data = sheet.getDataRange().getValues();

  let dataToProcess;

  if (position === 'Inner SPARC') {
    dataToProcess = data.slice(1).filter(row => {
      const team = row[11];
      const isTeamAllowed = INNER_SPARC_TEAMS.includes(team);
      const isColumnCNotNull = row[2] !== null && row[2] !== '';
      const salesAmount = row[16];
      const isAmountValid = toAmount_(salesAmount) > 0;
      return isTeamAllowed && isColumnCNotNull && isAmountValid;
    });
  } else {
    let nameColumnIndex;
    if (position === 'Agent') {
      nameColumnIndex = 8;
    } else if (position === 'Supervisor') {
      nameColumnIndex = 9;
    } else if (position === 'Manager') {
      nameColumnIndex = 10;
    } else {
      return JSON.stringify({ leadingRange: "Invalid Position", rangeCounts: {}, sortedRanges: [] });
    }

    dataToProcess = data.slice(1)
      .filter(row => {
        const isNameMatch = normName_(row[nameColumnIndex]) === normName_(name);
        const isColumnCNotNull = row[2] !== null && row[2] !== '';
        const salesAmount = row[16];
        const isAmountValid = toAmount_(salesAmount) > 0;
        return isNameMatch && isColumnCNotNull && isAmountValid;
      });
  }

  const salesAmounts = dataToProcess.map(row => toAmount_(row[16]));

  if (salesAmounts.length === 0) {
    return JSON.stringify({ leadingRange: "No Valid Sales Data", rangeCounts: {}, sortedRanges: [] });
  }

  const priceRanges = [
    { label: 'Below 1M', min: 0, max: 999999 },
    { label: '1M-2M', min: 1000000, max: 2000000 },
    { label: '2M-3M', min: 2000001, max: 3000000 },
    { label: '3M-4M', min: 3000001, max: 4000000 },
    { label: '4M-5M', min: 4000001, max: 5000000 },
    { label: 'Above 5M', min: 5000001, max: Infinity }
  ];

  const rangeCounts = {};
  priceRanges.forEach(range => rangeCounts[range.label] = 0);

  salesAmounts.forEach(amount => {
    const saleValue = amount;

    for (let i = 0; i < priceRanges.length; i++) {
      const range = priceRanges[i];

      if (range.label === 'Below 1M' && saleValue < range.max) {
        rangeCounts[range.label]++;
        break;
      }

      if (range.label === 'Above 5M' && saleValue >= range.min) {
        rangeCounts[range.label]++;
        break;
      }

      if (saleValue >= range.min && saleValue < range.max) {
        rangeCounts[range.label]++;
        break;
      }
    }
  });

  let leadingRangeLabel = 'N/A';
  let maxCount = 0;

  const rangeEntries = Object.entries(rangeCounts);

  if (rangeEntries.length > 0) {
    maxCount = rangeEntries.reduce((max, [, count]) => Math.max(max, count), 0);
  }

  if (maxCount > 0) {
    leadingRangeLabel = rangeEntries
      .filter(([, count]) => count === maxCount)
      .map(([label]) => label)[0];
  }

  const sortedRanges = rangeEntries
    .filter(([, count]) => count > 0)
    .sort((a, b) => b[1] - a[1]);

  return JSON.stringify({
    leadingRange: leadingRangeLabel,
    rangeCounts: rangeCounts,
    sortedRanges: sortedRanges
  });
}

/**
 * Retrieves all report data for a selected person/group, including profile, sales, and top metrics (Report Card).
 *
 * @param {string} name The name of the person (ignored if position is 'Inner SPARC').
 * @param {string} position The position of the person/group.
 * @param {string} year The year to filter by ('2020' through '2026').
 * @return {string} A JSON string containing all report data.
 */
function getReportData(name, position, year) {
  const { spreadsheetId, sheetName } = getDataSource(year);
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);

  if (!sheet) {
    return JSON.stringify(null);
  }

  const data = sheet.getDataRange().getValues();

  let filteredData;
  const profile = {};

  if (position === 'Inner SPARC') {
    filteredData = data.slice(1).filter(row => {
      const team = row[11];
      const isTeamAllowed = INNER_SPARC_TEAMS.includes(team);
      const isColumnCNotNull = row[2] !== null && row[2] !== '';
      const salesAmount = row[16];
      const isAmountValid = toAmount_(salesAmount) > 0;
      return isTeamAllowed && isColumnCNotNull && isAmountValid;
    });

    profile.name = 'Inner SPARC Realty Corporation';
    profile.position = '--';
    profile.team = '--';
    profile.picture = 'https://i.imgur.com/hKnqT9Y.png';
  } else {
    let nameColumnIndex;
    if (position === 'Agent') {
      nameColumnIndex = 8;
    } else if (position === 'Supervisor') {
      nameColumnIndex = 9;
    } else if (position === 'Manager') {
      nameColumnIndex = 10;
    } else {
      return JSON.stringify(null);
    }

    filteredData = data.slice(1).filter(row => {
      const isNameMatch = normName_(row[nameColumnIndex]) === normName_(name);
      const isColumnCNotNull = row[2] !== null && row[2] !== '';
      const salesAmount = row[16];
      const isAmountValid = toAmount_(salesAmount) > 0;
      return isNameMatch && isColumnCNotNull && isAmountValid;
    });

    if (filteredData.length > 0) {
      profile.name = filteredData[0][nameColumnIndex];
      profile.position = position;
      profile.team = filteredData[0][11];

      if (position === 'Agent') {
        profile.picture = filteredData[0][23];
      } else if (position === 'Supervisor') {
        profile.picture = filteredData[0][24];
      } else if (position === 'Manager') {
        profile.picture = filteredData[0][25];
      }
    }
  }

  if (filteredData.length === 0) {
    return JSON.stringify(null);
  }

  const monthlySales = {
    Jan: 0, Feb: 0, Mar: 0, Apr: 0, May: 0, Jun: 0, Jul: 0, Aug: 0, Sep: 0, Oct: 0, Nov: 0, Dec: 0
  };
  let totalSales = 0;
  let rawProjects = {};
  let rawDevelopers = {};
  let topSources = {};

  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

  filteredData.forEach(row => {
    const date = new Date(row[2]);
    const month = months[date.getMonth()];
    const salesAmount = toAmount_(row[16]);

    monthlySales[month] += salesAmount;
    totalSales += salesAmount;

    const project = row[6];
    const projectLogo = row[22];
    if (project) {
      // Group by normalized name so e.g. "Avida Towers" and "avida towers "
      // (same project, entered inconsistently) count as one project instead
      // of splitting into two 1-unit entries.
      const projectKey = normName_(project);
      rawProjects[projectKey] = rawProjects[projectKey] || { name: String(project).trim(), units: 0, sales: 0, logo: projectLogo };
      rawProjects[projectKey].units++;
      rawProjects[projectKey].sales += salesAmount;
      if (!rawProjects[projectKey].logo && projectLogo) rawProjects[projectKey].logo = projectLogo;
    }

    const developer = row[7];
    const developerLogo = row[21];
    if (developer) {
      const developerKey = normName_(developer);
      rawDevelopers[developerKey] = rawDevelopers[developerKey] || { name: String(developer).trim(), units: 0, sales: 0, logo: developerLogo };
      rawDevelopers[developerKey].units++;
      rawDevelopers[developerKey].sales += salesAmount;
      if (!rawDevelopers[developerKey].logo && developerLogo) rawDevelopers[developerKey].logo = developerLogo;
    }

    const source = row[12];
    if (source) {
      const sourceKey = normName_(source);
      topSources[sourceKey] = topSources[sourceKey] || { name: String(source).trim(), count: 0 };
      topSources[sourceKey].count++;
    }
  });

  let potentialIncome = 0;
  let potentialIncome2;

  if (position === 'Agent') {
    potentialIncome = totalSales * 0.03;
  } else if (position === 'Supervisor') {
    potentialIncome = totalSales * 0.005;
  } else if (position === 'Manager') {
    potentialIncome = totalSales * 0.005;
    potentialIncome2 = totalSales * 0.01;
  }

  const monthlySalesArray = Object.values(monthlySales);

  const highestSaleAmount = Math.max(...monthlySalesArray);
  const highestSaleMonth = months[monthlySalesArray.indexOf(highestSaleAmount)];

  const nonZeroSales = monthlySalesArray.filter(sale => sale > 0);
  let lowestSaleAmount;
  let lowestSaleMonth = '';

  if (nonZeroSales.length > 0) {
    lowestSaleAmount = Math.min(...nonZeroSales);
    const lowestSaleIndex = monthlySalesArray.indexOf(lowestSaleAmount);
    if (lowestSaleIndex !== -1) {
      lowestSaleMonth = months[lowestSaleIndex];
    }
  } else {
    lowestSaleAmount = 0;
    lowestSaleMonth = '';
  }

  const averageSale = totalSales / 12;

  const sortedProjects = Object.entries(rawProjects).sort((a, b) => b[1].sales - a[1].sales);
  const topProjects = sortedProjects.slice(0, 3).map(entry => ({
    name: entry[1].name,
    units: entry[1].units,
    sales: entry[1].sales,
    logo: entry[1].logo || 'https://placehold.co/100x30/F0F8FF/222689?text=N%2FA'
  }));

  while (topProjects.length < 3) {
    topProjects.push({
      name: 'N/A',
      units: 0,
      sales: 0,
      logo: 'https://placehold.co/100x30/dddddd/666666?text=N%2FA'
    });
  }

  const sortedDevelopers = Object.entries(rawDevelopers).sort((a, b) => b[1].sales - a[1].sales);
  const topDevelopers = sortedDevelopers.slice(0, 3).map(entry => ({
    name: entry[1].name,
    units: entry[1].units,
    sales: entry[1].sales,
    logo: entry[1].logo || 'https://placehold.co/100x30/F0F8FF/1C7D38?text=N%2FA'
  }));

  while (topDevelopers.length < 3) {
    topDevelopers.push({
      name: 'N/A',
      units: 0,
      sales: 0,
      logo: 'https://placehold.co/100x30/dddddd/666666?text=N%2FA'
    });
  }

  const sortedSources = Object.values(topSources)
    .sort((a, b) => b.count - a.count)
    .map(s => [s.name, s.count]);
  const allSources = sortedSources;

  return JSON.stringify({
    profile: profile,
    monthlySales: monthlySales,
    totalSales: totalSales,
    potentialIncome: potentialIncome,
    potentialIncome2: potentialIncome2,
    averageSale: averageSale,
    highestSale: {
      amount: highestSaleAmount,
      month: highestSaleMonth
    },
    lowestSale: {
      amount: lowestSaleAmount,
      month: lowestSaleMonth
    },
    topProjects: topProjects,
    topDevelopers: topDevelopers,
    topSources: allSources
  });
}


// ============================================================
// CMS → AGENT PORTAL
// Agent Commission Backend
// ============================================================

const CMS_COMMISSION_SHEET_ID = CMS_SHEET_ID //"1TJVunI2XHpUEBqqGn0jrJLAOkJpDjq0G4ZOFQwXzxqg";

const START_YEAR = 2020;
const CURRENT_YEAR = new Date().getFullYear();


// ============================================================
// GET CMS SPREADSHEET
// ============================================================

function getCmsCommissionSpreadsheet_() {
  return SpreadsheetApp.openById(CMS_COMMISSION_SHEET_ID);
}


// ============================================================
// NORMALIZE AGENT NAME
//
// Handles:
//
// Juan Dela Cruz
// JUAN DELA CRUZ
// Juan D. Cruz
// DELA CRUZ, JUAN
// dela cruz, juan
//
// All become comparable name parts.
// ============================================================

function normalizeAgentName_(value) {
  let name = String(value || '').trim().toLowerCase();
  if (!name) return [];

  name = name.replace(/\./g, '');
  name = name.replace(/,/g, ' ');
  name = name.replace(/[&/]+/g, ' ');
  name = name.replace(/[^\p{L}\p{N}\s]/gu, ' ');
  name = name.replace(/\s+/g, ' ').trim();

  return name ? name.split(' ') : [];
}

function agentNamePartSimilar_(a, b) {
  a = String(a || '').toLowerCase().trim();
  b = String(b || '').toLowerCase().trim();
  if (!a || !b) return false;
  if (a === b) return true;
  if (a.length < 4 || b.length < 4) return false;

  const matrix = Array.from({length: a.length + 1}, () => new Array(b.length + 1).fill(0));
  for (let i = 0; i <= a.length; i++) matrix[i][0] = i;
  for (let j = 0; j <= b.length; j++) matrix[0][j] = j;

  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost
      );
    }
  }

  const distance = matrix[a.length][b.length];
  const similarity = 1 - distance / Math.max(a.length, b.length);
  return similarity >= 0.80;
}

function agentNameMatches_(nameA, nameB) {
  const a = [...new Set(normalizeAgentName_(nameA))];
  const b = [...new Set(normalizeAgentName_(nameB))];
  if (!a.length || !b.length) return false;

  // Every word in the portal agent name must be represented in the CMS name.
  // Extra words are allowed for joint-agent records.
  let matched = 0;
  a.forEach(function(partA) {
    if (b.some(function(partB) { return agentNamePartSimilar_(partA, partB); })) matched++;
  });

  return matched === a.length;
}


// ============================================================
// GET AGENT COMMISSION FROM CMS
// ============================================================

/**
 * Collect ONLY commission records belonging to one agent.
 *
 * Supports different name formats between years:
 *
 * Juan Dela Cruz
 * DELA CRUZ, JUAN
 * JUAN DELA CRUZ
 * Juan D. Cruz
 *
 * @param {string} agentName Agent name from Agent Portal
 * @return {Object} Commission result
 */

function getAgentCommissionFromCms(agentName) {

  try {

    if (!agentName) {

      return {
        success: false,
        error: "Agent name is required.",
        data: []
      };

    }


    // Keep original name.
    // Matching is handled by agentNameMatches_().
    const targetAgent = String(agentName).trim();


    const ss = getCmsCommissionSpreadsheet_();

    const sheets = ss.getSheets();

    const results = [];


    // ============================================================
    // PROCESS EVERY YEAR
    // ============================================================

    const seenSaleIds = new Set();

    sheets.forEach(function(sh) {

      const sheetName = sh.getName();


      // Only sheets such as:
      //
      // 2020 Sales
      // 2021 Sales
      // 2022 Sales
      //
      const yearMatch =
        sheetName.match(/^(\d{4})\s+Sales$/i);


      if (!yearMatch) {
        return;
      }


      const year = Number(yearMatch[1]);


      if (
        year < START_YEAR ||
        year > CURRENT_YEAR
      ) {
        return;
      }


      const lastRow = sh.getLastRow();

      const lastColumn = sh.getLastColumn();


      if (
        lastRow < 2 ||
        lastColumn < 20
      ) {
        return;
      }


      // ==========================================================
      // CMS SALES STRUCTURE
      //
      // A  Reservation Date
      // B  Sales Agent
      // C  Developer
      // D  Project
      // E  Unit ID
      // F  Buyer's Name
      // G  DP Term
      // H  Active Sales Amount
      // I  Agent ID
      // J  Computed TSP
      // K  Developer Comm Rate
      // L  Company Gross Comm
      // M  Net VAT
      // N  Withholding Tax Company
      // O  Company Net Comm
      // P  Agent Comm Rate
      // Q  Agent Gross Comm
      // R  VAT Share Agent
      // S  Withholding Tax Agent
      // T  Agent Net Comm
      // U  Sale ID
      // V  Milestone Ratio
      // W  DP Tranche
      // X  Loan Takeout Tranche
      // Y  Released Commission
      // Z  Receivables
      // AA Commission Status
      // AB DP Release Amount
      // AC Loan Takeout Release Amount
      // AD Cash Advance
      // ==========================================================


      const rows = sh
        .getRange(
          2,
          1,
          lastRow - 1,
          lastColumn
        )
        .getValues();


      rows.forEach(function(row) {

        const salesAgent =
          String(row[1] || "").trim();


        if (!salesAgent) {
          return;
        }


        // ========================================================
        // FLEXIBLE AGENT NAME MATCH
        //
        // This replaces exact-name matching.
        //
        // Example:
        //
        // Portal:
        // Juan Dela Cruz
        //
        // CMS:
        // DELA CRUZ, JUAN
        //
        // => MATCH
        // ========================================================

        if (
          !agentNameMatches_(
            targetAgent,
            salesAgent
          )
        ) {
          return;
        }

        // ========================================================
        // EXCLUDE CANCELLED COMMISSIONS
        // ========================================================
        const commissionStatusRaw = String(row[30] || "").trim();

        if (commissionStatusRaw.toLowerCase() === "cancelled") {
          return;
        }

        
        // ========================================================
        // RESERVATION DATE
        // ========================================================

        let reservationDate = row[0];


        if (
          reservationDate instanceof Date
        ) {

          reservationDate =
            Utilities.formatDate(
              reservationDate,
              Session.getScriptTimeZone(),
              "MM/dd/yyyy"
            );

        }

        
        // Dedupe by Sale ID
        const saleIdKey = String(row[20] || "").trim();
        if (saleIdKey) {
          if (seenSaleIds.has(saleIdKey)) return;
          seenSaleIds.add(saleIdKey);
        }

        // ========================================================
        // ADD COMMISSION RECORD
        // ========================================================

        results.push({

          // ------------------------------------------------------
          // SOURCE
          // ------------------------------------------------------

          year: year,


          // ------------------------------------------------------
          // SALE INFORMATION
          // ------------------------------------------------------

          saleId: row[20] || "",

          reservationDate:
            reservationDate || "",

          salesAgent:
            salesAgent,

          agentId:
            row[8] || "",

          developer:
            row[2] || "",

          project:
            row[3] || "",

          unitId:
            row[4] || "",

          buyerName:
            row[5] || "",

          dpTerm:
            row[6] || "",


          // ------------------------------------------------------
          // SALES
          // ------------------------------------------------------

          activeSalesAmount:
            Number(row[7]) || 0,

          computedTsp:
            Number(row[9]) || 0,


          // ------------------------------------------------------
          // DEVELOPER COMMISSION
          // ------------------------------------------------------

          developerCommRate:
            Number(row[10]) || 0,


          // ------------------------------------------------------
          // COMPANY COMMISSION
          // ------------------------------------------------------

          companyGrossComm:
            Number(row[11]) || 0,

          netVat:
            Number(row[12]) || 0,

          companyWithholdingTax:
            Number(row[13]) || 0,

          companyNetComm:
            Number(row[14]) || 0,


          // ------------------------------------------------------
          // AGENT COMMISSION
          // ------------------------------------------------------

          agentCommRate:
            Number(row[15]) || 0,

          agentGrossComm:
            Number(row[16]) || 0,

          agentVatShare:
            Number(row[17]) || 0,

          agentWithholdingTax:
            Number(row[18]) || 0,

          // MAIN AGENT COMMISSION AMOUNT
          agentNetComm:
            Number(row[19]) || 0,


          // ------------------------------------------------------
          // MILESTONE / RELEASE
          // ------------------------------------------------------

          milestoneRatio:
            row[21] || "",

          dpTranche:
            row[22] || "",

          loanTakeoutTranche:
            row[23] || "",

          releasedCommission:
            Number(row[24]) || 0,

          receivables:
            Number(row[25]) || 0,

          commissionStatus:
            row[26] || "",

          dpReleaseAmount:
            Number(row[27]) || 0,

          loanTakeoutReleaseAmount:
            Number(row[28]) || 0,

          cashAdvance:
            Number(row[29]) || 0

        });

      });

    });


    // ============================================================
    // COMPUTE TOTALS
    // ============================================================

    let totalAgentNetComm = 0;

    let totalReleasedCommission = 0;

    let totalReceivables = 0;


    results.forEach(function(row) {

      totalAgentNetComm +=
        Number(row.agentNetComm) || 0;

      totalReleasedCommission +=
        Number(row.releasedCommission) || 0;

      totalReceivables +=
        Number(row.receivables) || 0;

    });


    // ============================================================
    // RETURN RESULT
    // ============================================================

    return {

      success: true,

      agentName: agentName,

      count: results.length,

      totals: {

        agentNetComm:
          totalAgentNetComm,

        releasedCommission:
          totalReleasedCommission,

        receivables:
          totalReceivables

      },

      data: results

    };


  } catch (error) {

    console.error(
      "getAgentCommissionFromCms error:",
      error.toString()
    );


    return {

      success: false,

      error: error.toString(),

      data: []

    };

  }

}






// ── Commission Slips tab (Agent Portal) — reads Commission Release History ──
const COMMISSION_RELEASE_SHEET_ID =   CMS_SHEET_ID;//"1TJVunI2XHpUEBqqGn0jrJLAOkJpDjq0G4ZOFQwXzxqg";
const COMMISSION_RELEASE_SHEET_NAME = "Commission Release History";

/** Fallback for pa_getMyCommissionSlips: if "Commission Release History"
 *  has no "Reservation Date" column of its own — or the column exists but
 *  the cell is blank — look the date up from the underlying sale record
 *  via Sale ID (searches current + historical years, same merged source
 *  pa_getSalesForAliases_ uses). Cached per execution so it only builds
 *  the lookup map once even with many slips. */
let PA_RESV_DATE_CACHE_ = null;
function pa_lookupReservationDateBySaleId_(saleId) {
  if (!saleId) return '';
  if (!PA_RESV_DATE_CACHE_) {
    PA_RESV_DATE_CACHE_ = {};
    pa_getSalesForAliases_(null).forEach(function(s) {
      if (s.saleId) PA_RESV_DATE_CACHE_[s.saleId] = s.reservationDate || '';
    });
  }
  return PA_RESV_DATE_CACHE_[saleId] || '';
}

/** Returns only the signed-in agent's RELEASED commission slips.
 *  Blank/"Pending"/anything other than "Released" is excluded server-side,
 *  so the client never even sees those rows. Uses the same flexible
 *  agentNameMatches_() already used for CMS commission matching, so name
 *  formatting differences (comma order, middle initial, etc.) still match. */
// function pa_getMyCommissionSlips(sessionId) {
//   try {
//     const session = pa_validateSession(sessionId);
//     if (!session || !session.valid) {
//       return { ok: false, error: "Session expired. Please log in again." };
//     }
//     const agentName = session.displayName || '';
//     if (!agentName) return { ok: true, slips: [] };

//     const ss = SpreadsheetApp.openById(COMMISSION_RELEASE_SHEET_ID);
//     const sh = ss.getSheetByName(COMMISSION_RELEASE_SHEET_NAME);
//     if (!sh || sh.getLastRow() < 2) return { ok: true, slips: [] };

//     const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
//       .map(function(h){ return String(h || '').trim(); });
//     const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

//     const col = {};
//     headers.forEach(function(h, i){ col[h] = i; });

//     function fmtDate_(v) {
//       if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "MM/dd/yyyy");
//       return v || '';
//     }

//     const slips = [];
//     rows.forEach(function(row) {
//       const status = String(col['Status'] !== undefined ? (row[col['Status']] || '') : '').trim();
//       if (status.toLowerCase() !== 'released') return; // blank/Pending/anything else stays hidden

//       const rowAgent = String(col['Sales Agent'] !== undefined ? (row[col['Sales Agent']] || '') : '').trim();
//       if (!rowAgent || !agentNameMatches_(agentName, rowAgent)) return;

//       let breakdown = [];
//       try {
//         const raw = col['Breakdown JSON'] !== undefined ? row[col['Breakdown JSON']] : '';
//         breakdown = raw ? JSON.parse(raw) : [];
//       } catch (e) { breakdown = []; }

//       slips.push({
//         timestamp:          fmtDate_(row[col['Timestamp']]),
//         voucherNo:          row[col['Voucher No.']] || '',
//         expDate:            fmtDate_(row[col['Exp. Date']]),
//         releaseDate:        fmtDate_(row[col['Release Date']]),
//         reservationDate:    (col['Reservation Date'] !== undefined && row[col['Reservation Date']])
//                                ? fmtDate_(row[col['Reservation Date']])
//                                : pa_lookupReservationDateBySaleId_(row[col['Sale ID']]),
//         saleId:             row[col['Sale ID']] || '',
//         salesAgent:         rowAgent,
//         developer:          row[col['Developer']] || '',
//         buyerName:          row[col["Buyer's Name"]] || '',
//         project:            row[col['Project']] || '',
//         unitId:             row[col['Unit ID']] || '',
//         tsp:                Number(row[col['TSP']]) || 0,
//         overallRate:        Number(row[col['Overall Rate (%)']]) || 0,
//         saNetReleased:      Number(row[col['SA Net Released (₱)']]) || 0,
//         othersNetReleased:  Number(row[col['Others Net Released (₱)']]) || 0,
//         totalNetReleased:   Number(row[col['Total Net Released (₱)']]) || 0,
//         breakdown:          breakdown,
//         override:           String(row[col['Override']] || ''),
//         saleYear:           row[col['Sale Year']] || '',
//         status:             status
//       });
//     });

//     slips.sort(function(a, b) {
//       return new Date(b.releaseDate || b.timestamp) - new Date(a.releaseDate || a.timestamp);
//     });

//     return { ok: true, slips: slips };
//   } catch (error) {
//     console.error("pa_getMyCommissionSlips error:", error.toString());
//     return { ok: false, error: error.toString() };
//   }
// }







function pa_getMyCommissionSlips(sessionId) {
  try {
    const session = pa_validateSession(sessionId);
    if (!session || !session.valid) {
      return { ok: false, error: "Session expired. Please log in again." };
    }
    const agentName = session.displayName || '';
    if (!agentName) return { ok: true, slips: [] };

    const ss = SpreadsheetApp.openById(COMMISSION_RELEASE_SHEET_ID);
    const sh = ss.getSheetByName(COMMISSION_RELEASE_SHEET_NAME);
    if (!sh || sh.getLastRow() < 2) return { ok: true, slips: [] };

    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
      .map(function(h){ return String(h || '').trim(); });
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

    const col = {};
    headers.forEach(function(h, i){ col[h] = i; });

    function fmtDate_(v) {
      if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "MM/dd/yyyy");
      return v || '';
    }

    const slips = [];
    rows.forEach(function(row) {
      const status = String(col['Status'] !== undefined ? (row[col['Status']] || '') : '').trim();
      if (status.toLowerCase() !== 'released') return; // blank/Pending/anything else stays hidden

      const rowAgent = String(col['Sales Agent'] !== undefined ? (row[col['Sales Agent']] || '') : '').trim();
      const isMainAgent = !!(rowAgent && agentNameMatches_(agentName, rowAgent));

      let breakdown = [];
      try {
        const raw = col['Breakdown JSON'] !== undefined ? row[col['Breakdown JSON']] : '';
        breakdown = raw ? JSON.parse(raw) : [];
      } catch (e) { breakdown = []; }

      // Lines belonging to THIS agent — matches both the sale's own agent
      // (the SA line, e.g. "Annalyn Violenta — DP Installment 2/4 (partial...)")
      // and any other named recipient sharing the same release
      // (SM/OM/FOUNDER/COFOUNDER/FUNDS), so someone whose name appears in
      // ANOTHER agent's release also sees their own commission slip for it.
      // let myLines = breakdown.filter(function(line){
      //   return line && line.particulars && agentNameMatches_(agentName, String(line.particulars));
      // });

      // let myLines = breakdown.filter(function(line){
      //   if (!line) return false;

      //   const recipientName =
      //     line.name ||
      //     line.particulars ||
      //     '';

      //   return recipientName &&
      //         agentNameMatches_(agentName, String(recipientName));
      // });
      let myLines = breakdown.filter(function(line){
        if (!line) return false;
        const recipientName = line.name || line.particulars || '';
        if (!recipientName) return false;
        // Expand shorthand/team-credit names (e.g. "Kurt/Gerald") to every
        // canonical agent they could refer to before matching, using the same
        // Agent Aliases sheet the CMS's email/Drive flow already relies on.
        return pa_resolveCanonicalNames_(recipientName).some(function(candidate){
          return agentNameMatches_(agentName, candidate);
        });
      });


      // Not the sale's own agent AND no line in the breakdown names this
      // agent — this release has nothing to do with them, skip it.
      if (!isMainAgent && !myLines.length) return;

      // Legacy rows saved before the per-recipient Breakdown JSON existed
      // (or where it failed to parse) — fall back to a single synthetic
      // SA line built from the row's own totals so old releases still
      // render a slip for the sale's own agent.
      if (!myLines.length && isMainAgent) {
        myLines = [{
          position: 'SA',
          particulars: rowAgent,
          rate: Number(col['Overall Rate (%)'] !== undefined ? row[col['Overall Rate (%)']] : 0) || 0,
          gross: 0,
          wht: 0,
          net: Number(col['SA Net Released (₱)'] !== undefined ? row[col['SA Net Released (₱)']] : 0) || 0
        }];
      }

      const myNetReleased = myLines.reduce(function(sum, l){ return sum + (Number(l.net) || 0); }, 0);

      // One slip per position — the same person can hold more than one
      // role on the same release (e.g. SA and OM), and each role gets
      // its own commission slip rather than being merged into one.
      const byPosition = {};
      myLines.forEach(function(l){
        const pos = String(l.position || 'SA').toUpperCase();
        if (!byPosition[pos]) byPosition[pos] = [];
        byPosition[pos].push(l);
      });

      Object.keys(byPosition).forEach(function(pos){
        const posLines = byPosition[pos];
        const posNet = posLines.reduce(function(sum, l){ return sum + (Number(l.net) || 0); }, 0);

        slips.push({
          timestamp:          fmtDate_(row[col['Timestamp']]),
          voucherNo:          row[col['Voucher No.']] || '',
          expDate:            fmtDate_(row[col['Exp. Date']]),
          releaseDate:        fmtDate_(row[col['Release Date']]),
          reservationDate:    (col['Reservation Date'] !== undefined && row[col['Reservation Date']])
                                 ? fmtDate_(row[col['Reservation Date']])
                                 : pa_lookupReservationDateBySaleId_(row[col['Sale ID']]),
          saleId:             row[col['Sale ID']] || '',
          salesAgent:         rowAgent,
          developer:          row[col['Developer']] || '',
          buyerName:          row[col["Buyer's Name"]] || '',
          project:            row[col['Project']] || '',
          unitId:             row[col['Unit ID']] || '',
          tsp:                Number(row[col['TSP']]) || 0,
          overallRate:        Number(row[col['Overall Rate (%)']]) || 0,
          saNetReleased:      Number(row[col['SA Net Released (₱)']]) || 0,
          othersNetReleased:  Number(row[col['Others Net Released (₱)']]) || 0,
          totalNetReleased:   Number(row[col['Total Net Released (₱)']]) || 0,
          myPosition:         pos,
          myNetReleased:      posNet,
          breakdown:          posLines,
          override:           String(row[col['Override']] || ''),
          saleYear:           row[col['Sale Year']] || '',
          status:             status
        });
      });
    });

    slips.sort(function(a, b) {
      return new Date(b.releaseDate || b.timestamp) - new Date(a.releaseDate || a.timestamp);
    });

    return { ok: true, slips: slips };
  } catch (error) {
    console.error("pa_getMyCommissionSlips error:", error.toString());
    return { ok: false, error: error.toString() };
  }
}



// ── Agent Aliases (shared with CMS) ──
// Resolves shorthand/team-credit names in Commission Release History
// breakdown lines (e.g. "Kurt/Gerald") to every canonical agent name they
// refer to — reusing the SAME "Agent Aliases" sheet the CMS's own
// email/Drive-saving flow (groupSlipLinesByRecipient_ / resolveCanonicalAgentNames_
// in the CMS's Code.gs) already reads. Both scripts point CMS_SHEET_ID /
// TARGET_SHEET_ID at the same spreadsheet, so this is a read-only consumer
// of the same source of truth — no separate alias list to maintain.
function pa_normalizeAliasKey_(name) {
  return String(name || '')
    .toUpperCase()
    .replace(/[.,\/&\-'’]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function pa_getAgentAliasMap_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('pa_agent_alias_map_v1');
  if (cached) {
    try { return JSON.parse(cached); } catch (e) { /* fall through and rebuild */ }
  }

  const map = {}; // normalized alias key -> array of canonical names
  function addEntry(key, canonical) {
    if (!key || !canonical) return;
    if (!map[key]) map[key] = [];
    if (map[key].indexOf(canonical) === -1) map[key].push(canonical);
  }

  try {
    const ss = getCmsSpreadsheet(); // already defined in Code.gs, points at 17woAy...
    const sh = ss.getSheetByName('Agent Aliases');
    if (sh && sh.getLastRow() >= 2) {
      const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
      data.forEach(function(row) {
        const canonical = String(row[0] || '').trim();
        if (!canonical) return;
        addEntry(pa_normalizeAliasKey_(canonical), canonical);
        // Semicolon-delimited, same convention as the CMS sheet — some
        // aliases legitimately contain commas ("VIOLENTA, MANNY & ANNALYN").
        String(row[1] || '').split(';').forEach(function(a) {
          const alias = a.trim();
          if (alias) addEntry(pa_normalizeAliasKey_(alias), canonical);
        });
      });
    }
  } catch (e) {
    console.error('pa_getAgentAliasMap_ failed:', e.toString());
  }

  try { cache.put('pa_agent_alias_map_v1', JSON.stringify(map), 3600); } catch (e) {}
  return map;
}

/** Resolves a raw breakdown-line name/particulars string to every
 *  canonical agent it could refer to. "Kurt/Gerald" -> both "Kurt Arevalo"
 *  and "Gerald Carvajal", since both hold that alias. Falls back to the
 *  name as-is (wrapped in an array) when nothing in Agent Aliases matches,
 *  so ordinary full names keep working exactly as before. */
function pa_resolveCanonicalNames_(rawName) {
  const raw = String(rawName || '').trim();
  if (!raw) return [];
  const matches = pa_getAgentAliasMap_()[pa_normalizeAliasKey_(raw)];
  return (matches && matches.length) ? matches.slice() : [raw];
}
