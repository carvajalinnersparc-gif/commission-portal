// ============================================================
// GoalPlanner.gs — per-agent, per-year goal (image + caption + notes)
// Stored in an "Agent Goals" tab in TARGET_SHEET_ID (auto-created).
// ============================================================

const SHEET_AGENT_GOALS = "Agent Goals";
const GOAL_HEADERS_ = ["Email", "Year", "Goal Title", "Goal Note", "Goal Image", "Goal Amount", "Updated At", "Goals JSON", "Images JSON"];
const GOAL_MAX_TITLE_ = 80;
const GOAL_MAX_NOTE_ = 1000;
const GOAL_MAX_IMAGE_CHARS_ = 48000; // Sheets cell limit is 50,000

// ── Structured goal list (Special Goal / Amount rows) ──
const GOAL_MAX_ITEMS_ = 20;
const GOAL_MAX_ITEM_NAME_ = 60;

// ── Multiple goal images, each with its own optional caption ──
const GOAL_MAX_IMAGES_ = 6;
const GOAL_MAX_IMAGE_CHARS_EACH_ = 12000;
const GOAL_MAX_IMAGE_CAPTION_ = 80;

function pa_goalSheet_() {
  const ss = getTargetSpreadsheet(); // Code.gs
  let sh = ss.getSheetByName(SHEET_AGENT_GOALS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_AGENT_GOALS);
    sh.getRange(1, 1, 1, GOAL_HEADERS_.length).setValues([GOAL_HEADERS_])
      .setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    return sh;
  }
  // Self-heal: a sheet created before "Goals JSON" / "Images JSON" existed
  // is widened in place so existing rows keep their columns as-is.
  const need = GOAL_HEADERS_.length;
  if (sh.getLastColumn() < need) {
    sh.insertColumnsAfter(sh.getLastColumn(), need - sh.getLastColumn());
    sh.getRange(1, 1, 1, need).setValues([GOAL_HEADERS_])
      .setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
  }
  return sh;
}

function pa_goalYear_(year) {
  const y = parseInt(year, 10);
  return (y >= 2020 && y <= new Date().getFullYear() + 1) ? y : 0;
}

function pa_goalFindRow_(sh, email, year) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const vals = sh.getRange(2, 1, last - 1, 2).getValues();
  const target = pa_normEmail_(email); // Auth.gs
  for (let i = 0; i < vals.length; i++) {
    if (pa_normEmail_(vals[i][0]) === target && String(vals[i][1]).trim() === String(year)) return i + 2;
  }
  return -1;
}

// Text is forced to plain-text format so a note starting with "=" is never treated as a formula.
function pa_goalWrite_(sh, row, col, text) {
  const cell = sh.getRange(row, col);
  cell.setNumberFormat("@");
  cell.setValue(text);
}

/** Returns the signed-in user's goal for a year. The identity always comes from the session, never the client. */
function pa_getMyGoal(sessionId, year) {
  const auth = pa_authorize_(sessionId);          // PortalData.gs
  if (!auth.ok) return auth;
  const gate = pa_requireActiveAgent_(auth);      // PortalData.gs
  if (gate) return gate;

  const y = pa_goalYear_(year);
  if (!y) return { ok: false, error: "Invalid year." };

  const sh = pa_goalSheet_();
  const row = pa_goalFindRow_(sh, auth.email, y);
  if (row === -1) return { ok: true, goal: { goals: [], images: [] } };

  const v = sh.getRange(row, 1, 1, GOAL_HEADERS_.length).getValues()[0];

  let goals = [];
  try { goals = JSON.parse(String(v[7] || "[]")); } catch (e) { goals = []; }
  if (!Array.isArray(goals)) goals = [];

  let images = [];
  try { images = JSON.parse(String(v[8] || "[]")); } catch (e) { images = []; }
  if (!Array.isArray(images)) images = [];

  // Legacy fallback: a goal saved before this migration only has a single
  // image (column E) + a single caption (column C) — surface it as the
  // first gallery image so nothing already saved appears to vanish.
  if (!images.length && String(v[4] || "").trim()) {
    images = [{ image: String(v[4]), title: String(v[2] || "") }];
  }

  return { ok: true, goal: { goals: goals, images: images } };
}

/** Partial update: only the keys present in `changes` (title / note / image) are written. */
function pa_saveMyGoal(sessionId, year, changes) {
  const auth = pa_authorize_(sessionId);
  if (!auth.ok) return auth;
  const gate = pa_requireActiveAgent_(auth);
  if (gate) return gate;

  const y = pa_goalYear_(year);
  if (!y) return { ok: false, error: "Invalid year." };
  changes = changes || {};

  // Validate everything before touching the sheet.
  const writes = {};

  if (changes.goals !== undefined) {
    const list = changes.goals;
    if (!Array.isArray(list)) return { ok: false, error: "Invalid goal list." };
    if (list.length > GOAL_MAX_ITEMS_) return { ok: false, error: "You can only add up to " + GOAL_MAX_ITEMS_ + " goals." };
    const cleaned = [];
    for (let i = 0; i < list.length; i++) {
      const item = list[i] || {};
      const name = String(item.name || "").trim().slice(0, GOAL_MAX_ITEM_NAME_);
      const amtRaw = String(item.amount || "").trim();
      let amount = 0;
      if (amtRaw !== "") {
        amount = parseFloat(amtRaw);
        if (isNaN(amount) || amount < 0) return { ok: false, error: "Each goal amount must be a valid positive number." };
      }
      if (!name && !amount) continue; // skip a fully blank row
      cleaned.push({ name: name, amount: amount });
    }
    writes[8] = JSON.stringify(cleaned);
  }

  if (changes.images !== undefined) {
    const list = changes.images;
    if (!Array.isArray(list)) return { ok: false, error: "Invalid image list." };
    if (list.length > GOAL_MAX_IMAGES_) return { ok: false, error: "You can only add up to " + GOAL_MAX_IMAGES_ + " photos." };
    const cleaned = [];
    for (let i = 0; i < list.length; i++) {
      const item = list[i] || {};
      const img = String(item.image || "");
      if (img && (img.length > GOAL_MAX_IMAGE_CHARS_EACH_ || !/^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/.test(img))) {
        return { ok: false, error: "One of those photos is too large to save. Try a smaller image." };
      }
      const title = String(item.title || "").trim().slice(0, GOAL_MAX_IMAGE_CAPTION_);
      if (!img) continue; // a caption alone with no image isn't kept
      cleaned.push({ image: img, title: title });
    }
    const combined = JSON.stringify(cleaned);
    if (combined.length > GOAL_MAX_IMAGE_CHARS_) {
      return { ok: false, error: "Those photos are too large together. Try removing one or using smaller images." };
    }
    writes[9] = combined;
  }

  if (!Object.keys(writes).length) return { ok: true };

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { ok: false, error: "Server is busy. Please try again." }; }
  try {
    const sh = pa_goalSheet_();
    let row = pa_goalFindRow_(sh, auth.email, y);
    if (row === -1) row = pa_appendRowAfterData_(sh, 1, [auth.email, y, "", "", "", "", new Date(), "[]", "[]"]); // Auth.gs
    Object.keys(writes).forEach(function (col) { pa_goalWrite_(sh, row, Number(col), writes[col]); });
    sh.getRange(row, 7).setValue(new Date());
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}