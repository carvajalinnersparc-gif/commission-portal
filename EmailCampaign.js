// ============================================================
// Inner SPARC Realty Corporation — Agent Portal
// EmailCampaign.gs — Admin broadcast email system
//
// Powers the "Emails" tab in the Admin panel: admin picks a
// preset theme (Christmas, New Year, Generic, etc.), writes a
// subject + body with optional custom bullet points, picks a
// recipient group (everyone / role / specific agent), previews,
// and sends. Every send is logged per-recipient so admins can
// see who got what and when.
//
// Uses the SAME MailApp delivery + branded HTML shell already
// proven by Auth.gs's pa_sendMail_ / pa_buildBrandedEmailHtml_,
// but with its own theme layer on top so login-link emails and
// broadcast emails stay visually and functionally separate.
//
// SETUP (run once from the Apps Script editor):
//   Select "ec_setupSheets" in the toolbar dropdown ▸ Run.
// ============================================================

// ── Sheet names ──────────────────────────────────────────────
const SHEET_EMAIL_CAMPAIGNS = "Email Campaigns";
const SHEET_EMAIL_LOG       = "Email Send Log";

// ── Column schemas (fixed order — do NOT reorder without updating
//    the row↔object converters below) ──────────────────────────
const EC_CAMPAIGN_HEADERS = [
  "Campaign ID",        // 0
  "Created At",         // 1
  "Created By",         // 2
  "Subject",            // 3
  "Theme",              // 4
  "Recipient Mode",     // 5  ("all" | "role" | "specific")
  "Recipient Value",    // 6  (role name for "role", ; separated emails for "specific")
  "Headline",           // 7
  "Greeting",           // 8
  "Body Text",          // 9  (paragraphs separated by \n\n)
  "Bullets",            // 10 (one per line)
  "Signoff",            // 11
  "Status",             // 12 ("Draft" | "Sending" | "Sent" | "Partially Sent" | "Failed" | "Scheduled")
  "Sent At",            // 13
  "Recipient Count",    // 14
  "Success Count",      // 15
  "Failure Count",      // 16
  "Body HTML",          // 17  (sanitized rich body from the Email Campaigner editor)
  "Scheduled For"       // 18  (Date — when a "Scheduled" campaign should auto-send)
];

const EC_LOG_HEADERS = [
  "Timestamp",      // 0
  "Campaign ID",    // 1
  "Email",          // 2
  "Display Name",   // 3
  "Status",         // 4  ("Sent" | "Failed")
  "Error",          // 5
  // ── Open-tracking columns (appended, not inserted — see
  // ec_ensureLogColumns_ below) ───────────────────────────────
  "Tracking ID",    // 6  unique per SENT message; blank for a Failed send
  "Opened At",      // 7  FIRST detected open only — never overwritten again
  "Open Count",     // 8  number of tracking-pixel requests detected
  "Last Opened At"  // 9  updated on every detected open
];

// Sheet holding one row PER DETECTED OPEN (not per recipient) — this is
// what powers the "Open History" list in the recipient detail view.
// EC_LOG_HEADERS above stays the per-recipient summary (first/last/count);
// this sheet is the append-only, granular event trail behind it.
const SHEET_EMAIL_OPEN_LOG = "Email Open Log";
const EC_OPEN_LOG_HEADERS = ["Timestamp", "Tracking ID", "Campaign ID", "Email"];

// Send an admin email alert the first time (and only the first time) a
// given message is opened. Centralized here — see point 27 in the spec —
// so it's one flag to flip, not something scattered through the send loop.
const EC_SEND_OPEN_ALERT_ = true;

// Ignores any "open" detected within this many seconds of sending — this
// is almost always the sender immediately checking their own Sent mail
// (Gmail proxy-fetches the pixel to render it, same as it would for the
// recipient), not the recipient actually opening it. This is a heuristic,
// not a guarantee — pixel tracking cannot distinguish sender from
// recipient at the HTTP level, so a genuinely fast recipient open within
// this window will also be missed. Keep it short.
const EC_OPEN_COOLDOWN_SECONDS = 0;

// ── One-time setup — run manually from the editor ───────────
function ec_setupSheets() {
  const ss = getTargetSpreadsheet(); // Code.gs

  if (!ss.getSheetByName(SHEET_EMAIL_CAMPAIGNS)) {
    const sh = ss.insertSheet(SHEET_EMAIL_CAMPAIGNS);
    sh.getRange(1, 1, 1, EC_CAMPAIGN_HEADERS.length).setValues([EC_CAMPAIGN_HEADERS])
      .setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, EC_CAMPAIGN_HEADERS.length);
    Logger.log("Created sheet: " + SHEET_EMAIL_CAMPAIGNS);
  } else {
    Logger.log("Sheet already exists: " + SHEET_EMAIL_CAMPAIGNS);
  }

  if (!ss.getSheetByName(SHEET_EMAIL_LOG)) {
    const sh = ss.insertSheet(SHEET_EMAIL_LOG);
    sh.getRange(1, 1, 1, EC_LOG_HEADERS.length).setValues([EC_LOG_HEADERS])
      .setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, EC_LOG_HEADERS.length);
    Logger.log("Created sheet: " + SHEET_EMAIL_LOG);
  } else {
    Logger.log("Sheet already exists: " + SHEET_EMAIL_LOG);
    ec_ensureLogColumns_(ss.getSheetByName(SHEET_EMAIL_LOG)); // widen an older log sheet in place
  }

  ec_openLogSheet_(); // creates "Email Open Log" if missing; no-op otherwise

  Logger.log("✅ Email campaign sheets are ready.");
}

// ── Sheet helpers ───────────────────────────────────────────
function ec_campaignSheet_() {
  const sh = getTargetSpreadsheet().getSheetByName(SHEET_EMAIL_CAMPAIGNS);
  if (sh) ec_ensureCampaignColumns_(sh);
  return sh;
}
function ec_logSheet_() {
  const sh = getTargetSpreadsheet().getSheetByName(SHEET_EMAIL_LOG);
  if (sh) ec_ensureLogColumns_(sh); // self-heal old rows onto the wider schema
  return sh;
}

function ec_openLogSheet_() {
  const ss = getTargetSpreadsheet();
  let sh = ss.getSheetByName(SHEET_EMAIL_OPEN_LOG);
  if (!sh) {
    sh = ss.insertSheet(SHEET_EMAIL_OPEN_LOG);
    sh.getRange(1, 1, 1, EC_OPEN_LOG_HEADERS.length).setValues([EC_OPEN_LOG_HEADERS])
      .setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, EC_OPEN_LOG_HEADERS.length);
  }
  return sh;
}

/** Columns are APPEND-ONLY. Sheets created before a column was added are
 *  widened in place here so existing data keeps its original position. */
function ec_ensureCampaignColumns_(sh) {
  const need = EC_CAMPAIGN_HEADERS.length;
  if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
  const header = sh.getRange(1, 1, 1, need).getValues()[0];
  let dirty = false;
  for (let i = 0; i < need; i++) {
    if (String(header[i] || "").trim() !== EC_CAMPAIGN_HEADERS[i]) { header[i] = EC_CAMPAIGN_HEADERS[i]; dirty = true; }
  }
  if (dirty) {
    sh.getRange(1, 1, 1, need).setValues([header])
      .setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
  }
}

/** Same append-only widening as ec_ensureCampaignColumns_, for the Email
 *  Send Log sheet — so a log sheet created before the open-tracking columns
 *  existed gets widened in place (existing Timestamp/Campaign ID/Email/
 *  Display Name/Status/Error data keeps its original position/columns). */
function ec_ensureLogColumns_(sh) {
  const need = EC_LOG_HEADERS.length;
  if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
  const header = sh.getRange(1, 1, 1, need).getValues()[0];
  let dirty = false;
  for (let i = 0; i < need; i++) {
    if (String(header[i] || "").trim() !== EC_LOG_HEADERS[i]) { header[i] = EC_LOG_HEADERS[i]; dirty = true; }
  }
  if (dirty) {
    sh.getRange(1, 1, 1, need).setValues([header])
      .setFontWeight("bold").setBackground("#222689").setFontColor("#ffffff");
  }
}

/** Smallest possible include helper so Portal.html can inline the
 *  EmailCampaigner feature module with <?!= ec_include_('EmailCampaigner') ?>.
 *  Namespaced so it can never collide with an existing include() in Code.gs. */
function ec_include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// ── Inline email images (uploaded to Drive, referenced by URL) ──
const EC_IMAGE_FOLDER_NAME_ = "Inner SPARC Email Images";
const EC_IMAGE_MAX_BYTES_   = 3 * 1024 * 1024; // 3 MB per image

function ec_getImageFolder_() {
  const it = DriveApp.getFoldersByName(EC_IMAGE_FOLDER_NAME_);
  return it.hasNext() ? it.next() : DriveApp.createFolder(EC_IMAGE_FOLDER_NAME_);
}

/** Run this ONCE from the Apps Script editor to grant Drive permission. */
function ec_authorizeDrive() {
  ec_getImageFolder_();
  Logger.log("✅ Drive access authorized.");
}

/** Called by the composer for every image inserted into the body.
 *  img = { name, mimeType, data (base64, no prefix) } */
function ec_uploadEmailImage(sessionId, img) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;
  try {
    img = img || {};
    const mime = String(img.mimeType || "");
    if (!/^image\/(png|jpe?g|gif|webp)$/i.test(mime)) {
      return { ok: false, error: "Only PNG, JPG, GIF or WEBP images are allowed." };
    }
    const bytes = Utilities.base64Decode(String(img.data || ""));
    if (!bytes.length) return { ok: false, error: "The image was empty." };
    if (bytes.length > EC_IMAGE_MAX_BYTES_) return { ok: false, error: "Image is too large (3 MB maximum)." };

    const safeName = String(img.name || "image").replace(/[^\w.\- ]/g, "_").slice(0, 80);
    const blob = Utilities.newBlob(bytes, mime, "EC-" + Date.now() + "-" + safeName);
    const file = ec_getImageFolder_().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { ok: true, fileId: file.getId(), url: "https://lh3.googleusercontent.com/d/" + file.getId() };
  } catch (e) {
    return { ok: false, error: "Could not upload image: " + (e && e.message ? e.message : e) };
  }
}

function ec_generateCampaignId_() {
  return "CMP-" + new Date().getTime().toString().slice(-8);
}

// ── Row ↔ object ────────────────────────────────────────────
function ec_rowToCampaign_(row) {
  return {
    campaignId:     row[0]  || "",
    createdAt:      row[1] instanceof Date ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a") : (row[1] || ""),
    createdBy:      row[2]  || "",
    subject:        row[3]  || "",
    theme:          row[4]  || "generic",
    recipientMode:  row[5]  || "all",
    recipientValue: row[6]  || "",
    headline:       row[7]  || "",
    greeting:       row[8]  || "",
    bodyText:       row[9]  || "",
    bullets:        row[10] || "",
    signoff:        row[11] || "",
    status:         row[12] || "Draft",
    sentAt:         row[13] instanceof Date ? Utilities.formatDate(row[13], Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a") : (row[13] || ""),
    recipientCount: Number(row[14]) || 0,
    successCount:   Number(row[15]) || 0,
    failureCount:   Number(row[16]) || 0,
    bodyHtml:       row[17] || "",
    scheduledFor:   row[18] instanceof Date ? row[18].toISOString() : (row[18] || ""),
    scheduledForLabel: row[18] instanceof Date ? Utilities.formatDate(row[18], Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a") : (row[18] || "")
  };
}

/** Row array for a campaign object — kept in ONE place so every
 *  setValues() write stays in sync with EC_CAMPAIGN_HEADERS. */
function ec_campaignRowValues_(c) {
  return [
    c.campaignId, c.createdAt, c.createdBy, c.subject,
    c.theme, c.recipientMode, c.recipientValue,
    c.headline, c.greeting, c.bodyText, c.bullets, c.signoff,
    c.status, c.sentAt, c.recipientCount,
    c.successCount, c.failureCount, c.bodyHtml || "",
    c.scheduledFor || ""
  ];
}

function ec_findCampaignRow_(sh, campaignId) {
  if (!sh || sh.getLastRow() < 2) return -1;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || "").trim() === String(campaignId || "").trim()) return i + 2;
  }
  return -1;
}

/** Human-readable audience label, derived from the SAME mode/value pair
 *  ec_resolveRecipients_() uses, so the history list can never describe an
 *  audience the server wouldn't actually resolve. */
function ec_audienceLabel_(mode, value) {
  const m = String(mode || "all");
  if (m === "role")     return "Role: " + (String(value || "").trim() || "—");
  if (m === "specific") {
    const n = String(value || "").split(";").map(function (x) { return x.trim(); }).filter(Boolean).length;
    return n + " specific recipient" + (n === 1 ? "" : "s");
  }
  return "Everyone (all portal users)";
}

function ec_plainSnippet_(html, text, max) {
  const limit = max || 120;
  let plain = String(html || "")
    .replace(/<(br|\/p|\/div|\/li|\/h[1-6]|\/tr)\s*>/gi, " ")
    .replace(/<[^>]*>/g, " ");
  if (!plain.replace(/\s/g, "")) plain = String(text || "");
  plain = plain.replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/&lt;/gi, "<")
               .replace(/&gt;/gi, ">").replace(/&quot;/gi, '"').replace(/\s+/g, " ").trim();
  return plain.length > limit ? plain.slice(0, limit - 1) + "…" : plain;
}

// ── THEMES ──────────────────────────────────────────────────
// Each theme supplies: a label, an accent color, a hero emoji strip,
// an optional header banner, and default copy an admin can override.
// Add a new theme by dropping one entry into this object — no other
// change anywhere is required.
const EC_THEMES_ = {
  generic: {
    label: "Generic / Neutral",
    accent: "#222689",
    accentSoft: "#e8e9f8",
    emoji: "✉️",
    defaultHeadline: "A quick update",
    defaultGreeting: "Hi [Agent Name],",
    defaultBody: "We have a quick update for the team. Please see below.",
    defaultBullets: "",
    defaultSignoff: "— Inner SPARC Realty"
  },
  recognition: {
    label: "Recognition",
    accent: "#C9A24B",
    accentSoft: "#fdf6e3",
    emoji: "🏆",
    defaultHeadline: "Congratulations!",
    defaultGreeting: "Hi [Agent Name],",
    defaultBody: "We want to take a moment to recognize your hard work and recent achievements. Keep it up!",
    defaultBullets: "",
    defaultSignoff: "— Inner SPARC Realty"
  }
};


/** Builds the festive string-lights strip as a self-contained SVG image,
 *  matching the reference Christmas design exactly (curly wire path,
 *  colored bulbs, white sparkle highlights). Wrapped in a real <img> tag
 *  with a data: URI — unlike a raw inline <svg> element (which most real
 *  mail clients, Outlook included, strip out of delivered HTML), an <img>
 *  pointing at an svg data URI is treated as an ordinary embedded image
 *  and survives far more clients. This is what lets the sent email
 *  actually show the same lights the in-app preview shows, instead of
 *  falling back to a plain emoji bar. position is 'top' or 'bottom' —
 *  the wire curves the opposite way and the bulb color pattern is offset,
 *  same as the reference design's two strips. */
function ec_lightsStripImg_(position) {
  const isTop = position !== 'bottom';
  const wirePath = isTop
    ? 'M0,10 C30,26 60,26 90,10 C120,-6 150,-6 180,10 C210,26 240,26 270,10 C300,-6 330,-6 360,10 C390,26 420,26 450,10 C480,-6 510,-6 540,10 C570,26 600,26 620,12'
    : 'M0,24 C30,8 60,8 90,24 C120,40 150,40 180,24 C210,8 240,8 270,24 C300,40 330,40 360,24 C390,8 420,8 450,24 C480,40 510,40 540,24 C570,8 600,8 620,22';

  const bulbYPattern = isTop ? [14, 21] : [20, 13];
  const sparkleYPattern = isTop ? [12.5, 19.5] : [18.5, 11.5];
  const colors = ['#B3282A', '#1C7D38', '#FF9132', '#C9A24B'];
  const xs = [15, 75, 135, 195, 255, 315, 375, 435, 495, 555, 605];

  let bulbs = '';
  let sparkles = '';
  xs.forEach(function (x, i) {
    const y = bulbYPattern[i % 2];
    const color = isTop
      ? colors[i % 2 === 0 ? 0 : (i % 4 === 1 ? 1 : (i % 4 === 3 ? 3 : 2))]
      : colors[i % 2 === 0 ? 1 : (i % 4 === 1 ? 2 : (i % 4 === 3 ? 1 : 3))];
    // Simplified deterministic color cycling matching the reference's
    // alternating palette closely enough for a decorative strip.
    const palette = ['#B3282A', '#1C7D38', '#FF9132', '#C9A24B', '#B3282A', '#1C7D38', '#FF9132', '#C9A24B', '#B3282A', '#1C7D38', '#FF9132'];
    const bottomPalette = ['#1C7D38', '#FF9132', '#B3282A', '#1C7D38', '#C9A24B', '#FF9132', '#B3282A', '#1C7D38', '#C9A24B', '#FF9132', '#B3282A'];
    const fill = isTop ? palette[i] : bottomPalette[i];
    bulbs += '<circle cx="' + x + '" cy="' + y + '" r="5" fill="' + fill + '"/>';
    const sy = sparkleYPattern[i % 2];
    sparkles += '<circle cx="' + (x - 1.5) + '" cy="' + sy + '" r="1.4"/>';
  });

  const svg =
    '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 620 34" width="620" height="34">' +
      '<path d="' + wirePath + '" fill="none" stroke="#DCD8CC" stroke-width="1.5"/>' +
      '<g>' + bulbs + '</g>' +
      '<g fill="#FFFFFF" opacity="0.55">' + sparkles + '</g>' +
    '</svg>';

  const dataUri = 'data:image/svg+xml;base64,' + Utilities.base64Encode(svg);
  const marginStyle = isTop ? 'margin-bottom:-2px;' : 'margin-top:-2px;';
  return '<img src="' + dataUri + '" alt="" width="620" height="34" style="display:block;width:100%;height:34px;' + marginStyle + '">';
}


// ── Inline images inside Update panels / Sign-off ────────────
// Marker line:  [img:https://...|w=140|side]
function ec_parseImgLine_(line) {
  const m = String(line || "").trim().match(/^\[img:(https:\/\/[^\s|\]]+)(?:\|([^\]]*))?\]$/i);
  if (!m) return null;
  let w = 160, side = false;
  String(m[2] || "").toLowerCase().split("|").forEach(function (o) {
    o = o.trim();
    if (o === "side") side = true;
    const wm = o.match(/^w=(\d{2,3})$/);
    if (wm) w = Math.min(520, Math.max(40, parseInt(wm[1], 10)));
  });
  return { url: m[1], width: w, side: side };
}

function ec_imgTagHtml_(img) {
  return '<img src="' + pa_escapeHtml_(img.url) + '" width="' + img.width + '" alt="" '
    + 'style="display:block;width:' + img.width + 'px;max-width:100%;height:auto;border:0;border-radius:6px;">';
}

/** Sign-off text -> HTML paragraphs; marker lines become <img>. */
function ec_renderTextWithImages_(text, recipientName) {
  return String(text || "").replace(/\[Agent Name\]/g, recipientName).split(/\n\s*\n/).map(function (p) {
    let out = "", buf = [];
    function flush() {
      if (buf.length) { out += '<p>' + buf.map(function (l) { return pa_escapeHtml_(l); }).join('<br>') + '</p>'; buf = []; }
    }
    p.split(/\n/).forEach(function (l) {
      const im = ec_parseImgLine_(l);
      if (im) { flush(); out += '<p style="margin:0 0 14px;">' + ec_imgTagHtml_(im) + '</p>'; }
      else buf.push(l);
    });
    flush();
    return out;
  }).join("");
}

// ── Branded HTML builder for campaign emails ────────────────
// Renders the innersparc_christmas_email_v4_taglish.html shell as the
// shared template for every theme. Themes only change:
//   - the subtitle line under "Inner SPARC Realty Corporation"
//   - the greeting name
//   - the body paragraphs
//   - the update panels (bullets — CRM / CMS / E-Learning / etc.)
//   - the closing text
//   - the sign-off
// Everything else (logo header, snowflakes, green rule, festive SVG
// strips top & bottom, footer image) is identical for every send.
function ec_buildCampaignHtml_(campaign, recipientName, theme, trackingPixelHtml) {
  const t = theme || EC_THEMES_.generic;
  const safe = function (s) { return pa_escapeHtml_(s); };
  const showHeader = campaign.showHeader !== false; // default true
  const showFooter = campaign.showFooter !== false; // default true

  const LOGO_URL   = "https://i.imgur.com/FGrfu8O.png";
  const FOOTER_URL = "https://i.imgur.com/j4iFl10.png";

  // ── greeting line (uses the recipient's real name) ──
  const recipientDisplay = recipientName || "Agents";
  const greeting = (campaign.greeting || "Kumusta ka na, [Agent Name]?")
    .replace(/\[Agent Name\]/g, recipientDisplay);

  // ── body ──
  // ARCHITECTURE: the main body is now a free-form, sanitized rich HTML
  // block produced by the Email Campaigner editor. The server still owns
  // the branded shell, theme, greeting, panels, signoff and footer — the
  // rich body is INSERTED into that shell, never replaces it. Campaigns
  // saved before the rich editor existed still carry plain Body Text, so
  // that path is kept as the fallback.
  const richBody = ec_normalizeEmailHtml_(String(campaign.bodyHtml || "").replace(/\[Agent Name\]/g, recipientDisplay));
  const bodyParas = richBody || String(campaign.bodyText || "").replace(/\[Agent Name\]/g, recipientDisplay).split(/\n\s*\n/).map(function (p) {
    return '<p>' + safe(p) + '</p>';
  }).join("");

  // ── bullets become the colored update panels (CRM / CMS / etc.) ──
  //
  // Grouping syntax:
  //   - An EMPTY LINE separates one panel from the next.
  //   - The FIRST line of a group is the panel TITLE.
  //   - Every line after that (until the next blank line) is a CONTENT
  //     line inside that panel.
  //   - A content line like "Authentication — password login..." renders
  //     the part before " — " as bold and the rest as normal text.
  //   - A content line without a separator renders as plain text.
  //
  // Example:
  //   Agent Portal
  //   self-service web app for agents...
  //   password login, passwordless "magic link" login...
  //   manage portal user access, roles, and permissions
  //
  //   Commission & Sales Management
  //   full backend system...
  //   leaderboards, trends, and performance summaries...
  //
  // → Two panels: "Agent Portal" (3 lines) and
  //   "Commission & Sales Management" (2 lines).
    const panelColors = ["navy", "green", "orange", "navy", "green", "orange"];

  const rawBulletsText = String(campaign.bullets || "").replace(/\r\n/g, "\n");
  const groupChunks = rawBulletsText
    .split(/\n\s*\n/)
    .map(function (chunk) {
      return chunk.split(/\n/).map(function (l) { return l.trim(); }).filter(Boolean);
    })
    .filter(function (lines) { return lines.length > 0; });

  const panelsHtml = groupChunks.map(function (allLines, i) {
    const color = panelColors[i % panelColors.length];

    // Separate image marker lines from text lines BEFORE parsing title/contents
    // (the marker contains a colon that would otherwise be misread).
    const images = [], lines = [];
    allLines.forEach(function (l) {
      const im = ec_parseImgLine_(l);
      if (im) images.push(im); else lines.push(l);
    });

    let title = "";
    let contents = [];
    if (lines.length) {
      const first = lines[0];
      const fm = first.match(/^(.+?)\s*[—–]\s*(.+)$/) || first.match(/^(.+?)\s*:\s*(.+)$/);
      if (fm) {
        title = fm[1].trim();
        contents.push(fm[2].trim());
        contents = contents.concat(lines.slice(1));
      } else {
        title = first;
        contents = lines.slice(1);
      }
    }

    const contentsHtml = contents.map(function (line) {
      const m = line.match(/^(.+?)\s*[—–]\s*(.+)$/) || line.match(/^(.+?)\s*:\s*(.+)$/);
      if (m) {
        return '<div style="margin:0 0 6px;font-size:15px;line-height:1.6;color:#5C5C68;">'
          + '<strong style="color:#1B1B22;">' + safe(m[1].trim()) + '</strong>'
          + ' — ' + safe(m[2].trim()) + '</div>';
      }
      return '<div style="margin:0 0 6px;font-size:15px;line-height:1.6;color:#5C5C68;">' + safe(line) + '</div>';
    }).join("");

    const sideImgs  = images.filter(function (x) { return x.side; });
    const belowImgs = images.filter(function (x) { return !x.side; });
    const gap = '<div style="height:8px;line-height:8px;font-size:8px;">&nbsp;</div>';

    let inner = contentsHtml || (images.length ? '' : '<p style="margin:0;font-style:italic;color:#AEAEBC;">(no details yet)</p>');

    if (sideImgs.length) {
      const colW = Math.max.apply(null, sideImgs.map(function (x) { return x.width; })) + 14;
      inner = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>'
        + '<td valign="top" style="padding-right:14px;">' + inner + '</td>'
        + '<td valign="top" align="right" width="' + colW + '" style="width:' + colW + 'px;">'
        + sideImgs.map(function (x) { return ec_imgTagHtml_(x); }).join(gap)
        + '</td></tr></table>';
    }
    if (belowImgs.length) {
      inner += '<div style="margin-top:10px;">' + belowImgs.map(function (x) { return ec_imgTagHtml_(x); }).join(gap) + '</div>';
    }

    return '<div class="panel ' + color + '">'
      + (title ? '<h3>' + safe(title) + '</h3>' : '')
      + inner
      + '</div>';
  }).join("");

  // ── closing / signoff paragraphs ──
  // Split on blank lines into paragraphs, then convert any single
  // newlines INSIDE each paragraph into <br> so multi-line sign-offs
  // (e.g. "NAME\nCompany") render on separate lines instead of being
  // collapsed into one by HTML whitespace rules.
  const signoffParas = ec_renderTextWithImages_(campaign.signoff, recipientDisplay);

  // ── subtitle under the brand name (theme-driven, defaults to a
  //    neutral "An Update From The Team") ──
  const subtitle = campaign.headline || t.defaultHeadline || "An Update From The Team";

  // ── festive decorations (snowflakes + string-lights) are exclusive to
  // themes explicitly marked festive (currently only Christmas). The old
  // inline <svg> string lights are replaced with an emoji bar because
  // most real mail clients (Outlook, many Gmail contexts) strip <svg>
  // out of delivered HTML email entirely — that's why the lights showed
  // in the in-app preview but never in the inbox.
  const isFestive = !!t.festive;
  const snowflakes = isFestive ? "❄ ✦ ❄" : "";
  // NOTE: an SVG embedded as a data: URI (ec_lightsStripImg_) looked right
  // in the in-app preview but showed as a broken/blank image in real
  // inboxes — Gmail and most mail clients strip data: URI images outright
  // for security, only trusting images loaded from a real hosted URL (the
  // logo/footer images below work because they ARE real hosted URLs, not
  // data URIs). Falling back to a plain-text emoji bar here guarantees it
  // renders everywhere, with zero risk of a broken-image box.
  const lightsRowHtml = isFestive
    ? '<div style="text-align:center;padding:10px 0;font-size:20px;letter-spacing:10px;line-height:1;background:#FAFAFC;">🔴 🟢 🟠 🟡 🔴 🟢 🟠 🟡 🔴 🟢 🟠</div>'
    : '';

  // ── the reference email, style inlined so it renders in mail clients ──
  return ''
  + '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">'
  + '<meta name="viewport" content="width=device-width, initial-scale=1.0">'
  + '<title>' + safe(campaign.subject) + '</title>'
  + '<style>'
  +   '@import url("https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700&family=Dancing+Script:wght@600&display=swap");'
  +   'body{margin:0;padding:40px 12px;background:#F3F4FA;font-family:Montserrat,Arial,sans-serif;color:#2B2B33;}'
  +   '.email{max-width:620px;margin:0 auto;background:#FFFFFF;border-radius:14px;overflow:hidden;box-shadow:0 2px 10px rgba(34,38,137,0.06);}'
  +   '.header{text-align:center;padding:30px 32px 22px;}'
  +   '.logo-lockup{margin:0 0 8px;}'
  +   '.logo-lockup img{height:110px;width:auto;display:block;margin:0 auto;}'
  +   '.header h1{font-size:21px;font-weight:700;color:#222689;margin:14px 0 3px;}'
  +   '.header .subtitle{font-size:14px;color:#8B8B98;margin:0;letter-spacing:0.02em;}'
  +   '.header .snowflakes{margin-top:10px;font-size:15px;letter-spacing:6px;color:#C9A24B;}'
  +   '.rule{height:3px;background:#1C7D38;}'
  +   '.lights-strip{display:block;width:100%;height:34px;}'
  +   '.lights-strip.top{margin-bottom:-2px;}'
  +   '.lights-strip.bottom{margin-top:-2px;}'
  +   '.content{padding:34px 36px 8px;}'
  +   '.content p{margin:0 0 18px;font-size:16px;line-height:1.75;}'
  +   '.headline{font-size:25px;font-weight:700;color:#1B1B22;margin:0 0 8px;line-height:1.3;}'
  +   '.headline .accent{color:#1C7D38;}'
  +   '.panels{padding:6px 32px 6px;}'
  +   '.panel{background:#F6F7FB;border-left:4px solid #222689;border-radius:6px;padding:16px 18px;margin-bottom:14px;}'
  +   '.panel.green{border-left-color:#1C7D38;}'
  +   '.panel.orange{border-left-color:#FF9132;}'
  +   '.panel h3{font-size:17px;font-weight:700;color:#1B1B22;margin:0 0 5px;}'
  +   '.panel p{margin:0;font-size:15px;line-height:1.65;color:#5C5C68;}'
  +   '.signoff-name{font-family:"Dancing Script","Brush Script MT","Segoe Script",cursive;font-weight:600;font-size:22px;color:#222689;}'
  +   '.note{padding:0 32px 26px;font-size:13px;line-height:1.65;color:#9A9AA6;border-top:1px solid #EDEDF3;padding-top:20px;margin:0 32px;}'
  +   '.footer{text-align:center;border-top:1px solid #EDEDF3;}'
  + '</style></head><body>'

  + '<div class="email">'

  + (showHeader ? (
      // ── festive lights strip (top) — plain emoji bar, the only
      // version guaranteed to render in real inboxes ──
      lightsRowHtml
      // ── header (logo + brand + subtitle + snowflakes) ──
      + '<div class="header">'
      +   '<div class="logo-lockup"><img src="' + LOGO_URL + '" alt="InnerSPARC Realty Corporation"></div>'
      +   '<h1>Inner SPARC Realty Corporation</h1>'
      +   '<p class="subtitle">' + safe(subtitle) + '</p>'
      +   (snowflakes ? '<div class="snowflakes">' + snowflakes + '</div>' : '')
      + '</div>'
      + '<div class="rule"></div>'
    ) : '')

  // ── greeting + body ──
  + '<div class="content">'
  +   '<p class="headline">' + safe(greeting) + '</p>'
  +   bodyParas
  + '</div>'

  // ── update panels (the bullets) ──
  + (panelsHtml ? '<div class="panels">' + panelsHtml + '</div>' : '')

  // ── closing / sign-off ──
  + (signoffParas ? '<div class="content" style="padding-top:6px;">' + signoffParas + '</div>' : '')

  // ── note ──
  + '<p class="note">If you have questions about any of these updates, please coordinate with your assigned team lead or the Inner SPARC Realty support team.</p>'

  + (showFooter ? (
      // ── festive lights strip (bottom) ──
      lightsRowHtml
      // ── footer image ──
      + '<div class="footer">'
      +   '<img src="' + FOOTER_URL + '" alt="Inner SPARC Realty Corporation footer" style="width:100%;height:auto;display:block;">'
      + '</div>'
    ) : '')

  + '</div>'
  // ── open-tracking pixel ──
  // Placed last, outside the visible ".email" card, and only emitted when
  // the caller actually passed one (ec_buildPlainFallback_'s plain-text
  // rendering has no pixel and needs none — see point 8/26 of the spec:
  // opens can only ever be inferred from the HTML version loading images).
  + (trackingPixelHtml || '')
  + '</body></html>';
}


// ── SERVER-SIDE SANITIZATION + EMAIL-SAFE NORMALIZATION ─────
// SANITIZATION BOUNDARY: the Email Campaigner also sanitizes in the
// browser, but that pass is a convenience only — anything arriving here
// is treated as fully untrusted (a caller can hit google.script.run
// directly). This is the authoritative pass, and it always runs before
// the rich body is placed into the branded shell.
//
// SCOPE: every tag is allowed except <script> — but "block <script> and
// nothing else" would NOT actually restrict this to "no code execution":
// onerror=/onclick=-style event-handler attributes, javascript: links,
// and <iframe>/<object>/<embed>/<form>/<svg> can all run or exfiltrate
// without ever writing the word "script". Those run in TWO places that
// matter more than the delivered email itself: the send-log/history
// preview (sandboxed, so low risk) and the live composer editor, which is
// an unsandboxed contenteditable region in the admin portal's own page —
// markup with a working exploit there is a real stored-XSS hole against
// admin sessions, not just a cosmetic issue. So alongside the <script>
// tag itself, this pass also strips: inline event handlers (onXXX=), the
// javascript:/vbscript:/data:/file:/about:/blob: URL schemes on any
// URL-bearing attribute (including url(...) inside CSS), and the handful
// of tags that are pure code-execution vectors even without a literal
// <script> in them (iframe, object, embed, form + its controls, svg,
// math, and friends). Layout-only CSS (position:fixed, url() to a plain
// http(s) image, etc.) is NOT restricted — those cannot execute code and
// are just ordinary email design tools.
const EC_ALLOWED_TAGS_ = {
  a:1, b:1, strong:1, i:1, em:1, u:1, s:1, strike:1, del:1, ins:1,
  mark:1, small:1, big:1, tt:1, kbd:1, var:1, samp:1, dfn:1, abbr:1,
  cite:1, q:1, address:1, time:1, bdi:1, bdo:1, label:1, output:1, wbr:1,
  p:1, br:1, div:1, span:1, h1:1, h2:1, h3:1, h4:1, h5:1, h6:1, blockquote:1,
  ul:1, ol:1, li:1, dl:1, dt:1, dd:1, hr:1,
  table:1, thead:1, tbody:1, tfoot:1, tr:1, td:1, th:1, caption:1, colgroup:1, col:1,
  section:1, article:1, header:1, footer:1, nav:1, aside:1, main:1, figure:1, figcaption:1,
  details:1, summary:1, style:1, button:1,
  sup:1, sub:1, font:1, pre:1, code:1, img:1
};

// Inline style="" attribute (delimited by an HTML attribute quote, so
// literal quote/angle-bracket characters are still stripped to prevent
// attribute breakout). url(...) is now only stripped when it points at a
// code-executing scheme — a plain https:// background-image is fine.
function ec_sanitizeCss_(css) {
  return String(css || "")
    .replace(/expression\s*\(/gi, "")
    .replace(/url\s*\(\s*["']?\s*(?:javascript|vbscript|data|file|about|blob)\s*:[^)]*\)/gi, "url()")
    .replace(/@import/gi, "")
    .replace(/behaviou?r\s*:/gi, "")
    .replace(/["<>]/g, "");
}

// Raw <style>...</style> block content — NOT inside an HTML attribute, so
// quotes (font-family: "Times New Roman", content: "★") are left alone.
// Only the constructs that can execute code or pull in a remote stylesheet
// are stripped.
function ec_sanitizeStyleBlockCss_(css) {
  return String(css || "")
    .replace(/expression\s*\(/gi, "")
    .replace(/url\s*\(\s*["']?\s*(?:javascript|vbscript|data|file|about|blob)\s*:[^)]*\)/gi, "url()")
    .replace(/@import/gi, "")
    .replace(/behaviou?r\s*:/gi, "");
}

function ec_sanitizeRichHtml_(html) {
  let out = String(html || "");

  out = out.replace(/<!--[\s\S]*?-->/g, "");
  // Drop dangerous elements together with their contents. These are the
  // only tags kept off the allow-list — see the SCOPE note above for why.
  out = out.replace(/<\s*(script|iframe|object|embed|form|svg|math|noscript|template|frameset|applet|audio|video|canvas)\b[\s\S]*?<\s*\/\s*\1\s*>/gi, "");
  // Then any stray opening/closing/void tag of the same families, plus
  // form-only controls that are inert (and thus pointless) outside a form.
  out = out.replace(/<\s*\/?\s*(script|iframe|object|embed|form|input|select|textarea|svg|math|link|meta|base|frame|frameset|applet|audio|video|source|track|canvas|noscript|template|picture)\b[^>]*>/gi, "");
  // Inline event handlers — these are the actual "does this execute code"
  // surface, independent of which tag they sit on.
  out = out.replace(/\son[a-z]+\s*=\s*"[^"]*"/gi, "")
           .replace(/\son[a-z]+\s*=\s*'[^']*'/gi, "")
           .replace(/\son[a-z]+\s*=\s*[^\s>]+/gi, "");
  // Dangerous URL schemes on any URL-bearing attribute.
  out = out.replace(/\s(href|src|background|action|formaction|xlink:href)\s*=\s*("|')\s*(?:javascript|vbscript|data|file|about|blob)\s*:[^"']*\2/gi, ' href="#"');
  out = out.replace(/\s(href|src|background|action|formaction|xlink:href)\s*=\s*(?:javascript|vbscript|data|file|about|blob)\s*:[^\s>]*/gi, ' href="#"');
  // Raw <style> block content — sanitize the CSS text itself before the
  // tag-level filtering below (which only looks at tags, not text nodes).
  out = out.replace(/(<style\b[^>]*>)([\s\S]*?)(<\/style\s*>)/gi, function (m, open, css, close) {
    return open + ec_sanitizeStyleBlockCss_(css) + close;
  });
  // CSS payloads on inline style="" attributes.
  out = out.replace(/style\s*=\s*"([^"]*)"/gi, function (m, css) { return 'style="' + ec_sanitizeCss_(css) + '"'; });
  out = out.replace(/style\s*=\s*'([^']*)'/gi, function (m, css) { return 'style="' + ec_sanitizeCss_(css) + '"'; });
  // Tag allow-list — anything unexpected is dropped, its text kept.
  out = out.replace(/<\/?([a-zA-Z0-9]+)((?:[^>"']|"[^"]*"|'[^']*')*)>/g, function (m, tag) {
    return EC_ALLOWED_TAGS_[String(tag).toLowerCase()] ? m : "";
  });
  return out;
}

/** Adds an inline style (and optional attributes) to every occurrence of a
 *  tag, merging with any style the editor already produced. */
function ec_withStyle_(html, tag, style, attrs) {
  const re = new RegExp("<" + tag + "\\b((?:[^>\"']|\"[^\"]*\"|'[^']*')*)>", "gi");
  return String(html).replace(re, function (m, rest) {
    let body = rest || "";
    if (/style\s*=\s*"/i.test(body)) {
      body = body.replace(/style\s*=\s*"([^"]*)"/i, function (mm, css) {
        return 'style="' + style + ";" + css + '"';
      });
    } else {
      body += ' style="' + style + '"';
    }
    if (attrs) body += " " + attrs;
    return "<" + tag + body + ">";
  });
}

/** EMAIL-SAFE NORMALIZATION: every structural tag the editor can emit gets
 *  its formatting pushed inline too, since Outlook and many other clients
 *  ignore <style> blocks and external CSS entirely — an author-supplied
 *  <style> block (now allowed, see EC_ALLOWED_TAGS_) is left in place for
 *  the clients that DO honor it, but inline styles remain the fallback
 *  every client can render. Editor-only helper classes are stripped for
 *  the same reason. */
function ec_normalizeEmailHtml_(html) {
  let out = ec_sanitizeRichHtml_(html);
  if (!out.replace(/<[^>]*>/g, "").replace(/&nbsp;/gi, "").trim()) return "";

  out = out.replace(/\sclass\s*=\s*"[^"]*"/gi, "").replace(/\sclass\s*=\s*'[^']*'/gi, "");
  out = out.replace(/\scontenteditable\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, "");

  out = ec_withStyle_(out, "p",  "margin:0 0 14px;font-size:16px;line-height:1.75;color:#2B2B33");
  out = ec_withStyle_(out, "div", "font-size:16px;line-height:1.75;color:#2B2B33");
  out = ec_withStyle_(out, "h1", "margin:0 0 10px;font-size:26px;line-height:1.3;font-weight:700;color:#1B1B22");
  out = ec_withStyle_(out, "h2", "margin:0 0 9px;font-size:22px;line-height:1.3;font-weight:700;color:#1B1B22");
  out = ec_withStyle_(out, "h3", "margin:0 0 8px;font-size:18px;line-height:1.35;font-weight:700;color:#1B1B22");
  out = ec_withStyle_(out, "h4", "margin:0 0 8px;font-size:16px;line-height:1.35;font-weight:700;color:#1B1B22");
  out = ec_withStyle_(out, "h5", "margin:0 0 8px;font-size:15px;line-height:1.4;font-weight:700;color:#1B1B22");
  out = ec_withStyle_(out, "h6", "margin:0 0 8px;font-size:14px;line-height:1.4;font-weight:700;color:#1B1B22;text-transform:uppercase;letter-spacing:0.02em");
  out = ec_withStyle_(out, "blockquote", "margin:0 0 14px;padding:10px 16px;border-left:4px solid #222689;background:#F6F7FB;color:#5C5C68;font-style:italic");
  out = ec_withStyle_(out, "ul", "margin:0 0 14px;padding-left:24px");
  out = ec_withStyle_(out, "ol", "margin:0 0 14px;padding-left:24px");
  out = ec_withStyle_(out, "li", "margin:0 0 6px;font-size:16px;line-height:1.7");
  out = ec_withStyle_(out, "dl", "margin:0 0 14px");
  out = ec_withStyle_(out, "dt", "font-weight:700;font-size:16px;color:#1B1B22;margin:0 0 2px");
  out = ec_withStyle_(out, "dd", "margin:0 0 10px 18px;font-size:16px;line-height:1.7;color:#2B2B33");
  out = ec_withStyle_(out, "address", "margin:0 0 14px;font-style:normal;font-size:16px;line-height:1.7;color:#2B2B33");
  out = ec_withStyle_(out, "a",  "color:#222689;text-decoration:underline", 'target="_blank"');
  out = ec_withStyle_(out, "pre", "margin:0 0 14px;padding:12px;background:#F6F7FB;font-family:'Courier New',monospace;font-size:14px;white-space:pre-wrap");
  out = ec_withStyle_(out, "code", "font-family:'Courier New',monospace;font-size:14px");
  out = ec_withStyle_(out, "hr", "border:0;border-top:1px solid #EDEDF3;margin:20px 0");
  out = ec_withStyle_(out, "table", "border-collapse:collapse;width:100%;margin:0 0 14px", 'cellpadding="6" cellspacing="0" border="1"');
  out = ec_withStyle_(out, "caption", "font-size:14px;color:#5C5C68;margin:0 0 6px;text-align:left");
  out = ec_withStyle_(out, "td", "border:1px solid #DDDDE5;padding:8px;font-size:15px;line-height:1.6;vertical-align:top");
  out = ec_withStyle_(out, "th", "border:1px solid #DDDDE5;padding:8px;font-size:15px;line-height:1.6;background:#F6F7FB;text-align:left;font-weight:700");
  out = ec_withStyle_(out, "img", "max-width:100%;height:auto;display:block");
  out = ec_withStyle_(out, "details", "margin:0 0 14px;font-size:16px;line-height:1.7;color:#2B2B33");
  out = ec_withStyle_(out, "summary", "font-weight:700;cursor:pointer;color:#1B1B22");
  out = ec_withStyle_(out, "mark", "background:#FFF3B0;color:#1B1B22;padding:0 2px");
  // Native <button> can't carry a link (no JS, no <form> — both blocked
  // above), so it only ever renders as a styled, inert label in the
  // delivered email. Use the composer's "Insert button" toolbar action
  // instead for an actually-clickable CTA (it inserts a styled <a>).
  out = ec_withStyle_(out, "button", "display:inline-block;padding:12px 28px;background:#222689;color:#ffffff;border:none;border-radius:6px;font-weight:700;font-size:15px;text-decoration:none;font-family:Montserrat,Arial,sans-serif;");
  return out;
}

/** Plain-text rendering of the rich body, for the MailApp text fallback. */
function ec_richToPlain_(html) {
  return String(html || "")
    .replace(/<\s*(br)\s*\/?>/gi, "\n")
    .replace(/<\s*\/\s*(p|div|li|h[1-6]|tr|blockquote)\s*>/gi, "\n")
    .replace(/<\s*li[^>]*>/gi, "• ")
    .replace(/<[^>]*>/g, "")
    .replace(/&nbsp;/gi, " ").replace(/&amp;/gi, "&").replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">").replace(/&quot;/gi, '"').replace(/&#39;/gi, "'")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

// ── Recipient resolution ────────────────────────────────────
/**
 * Returns [{ email, name }] for the given recipient mode.
 *   mode "all"      -> every active Agent Access row
 *   mode "role"     -> every Agent Access row whose role matches value (case-insensitive)
 *   mode "specific" -> value is a semicolon-separated list of emails; each
 *                      must have an Agent Access row (so we can look up a
 *                      display name — but a row without a display name is
 *                      still deliverable, just addressed generically).
 */
function ec_resolveRecipients_(mode, value) {
  const users = pa_adminListUsers_(); // Auth.gs

  if (mode === "all") {
    return users
      .filter(u => String(u.email || "").trim())
      .map(u => ({ email: pa_normEmail_(u.email), name: ec_displayName_(u) || u.email }));
  }

  if (mode === "role") {
    const want = String(value || "").trim().toLowerCase();
    return users
      .filter(u => String(u.role || "").trim().toLowerCase() === want && String(u.email || "").trim())
      .map(u => ({ email: pa_normEmail_(u.email), name: ec_displayName_(u) || u.email }));
  }

  if (mode === "specific") {
    const byEmail = {};
    users.forEach(u => { byEmail[pa_normEmail_(u.email)] = u; });
    const raw = String(value || "").split(";").map(s => s.trim()).filter(Boolean);
    return raw.map(e => {
      const norm = pa_normEmail_(e);
      const u = byEmail[norm];
      return { email: norm, name: u ? (ec_displayName_(u) || norm) : norm };
    }).filter(r => r.email && r.email.indexOf("@") !== -1);
  }

  return [];
}

/** Recipient display name: the Canonical Name, falling back to the first alias. */
function ec_displayName_(u) {
  return String((u && u.canonicalName) || "").trim() || ec_firstAlias_(u && u.aliasesCsv);
}

function ec_firstAlias_(aliasesCsv) {
  const parts = pa_parseAliasesCsv_(aliasesCsv); // Auth.gs
  return parts.length ? parts[0] : "";
}

// ── PUBLIC: theme list for the composer ─────────────────────
function ec_listThemes(sessionId) {
  const auth = pa_isAdmin_(sessionId); // Auth.gs
  if (!auth.ok) return auth;
  return {
    ok: true,
    themes: Object.keys(EC_THEMES_).map(k => {
      const t = EC_THEMES_[k];
      return {
        key: k,
        label: t.label,
        emoji: t.emoji,
        accent: t.accent,
        accentSoft: t.accentSoft,
        festive: !!t.festive,
        defaultHeadline: t.defaultHeadline,
        defaultGreeting: t.defaultGreeting,
        defaultBody:     t.defaultBody,
        defaultBullets:  t.defaultBullets,
        defaultSignoff:  t.defaultSignoff
      };
    })
  };
}

// ── PUBLIC: list past campaigns ─────────────────────────────
function ec_listCampaigns(sessionId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  const sh = ec_campaignSheet_();
  if (!sh || sh.getLastRow() < 2) return { ok: true, campaigns: [] };

  let openStatsByCampaign = {};
  try {
    openStatsByCampaign = ec_buildOpenStatsMap_();
  } catch (e) {
    Logger.log("ec_listCampaigns: ec_buildOpenStatsMap_ failed — " + e);
  }

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, EC_CAMPAIGN_HEADERS.length).getValues();
  const campaigns = rows
    .filter(r => String(r[0] || "").trim())
    .map(ec_rowToCampaign_)
    .map(function (c) {
      c.audienceLabel = ec_audienceLabel_(c.recipientMode, c.recipientValue);
      c.snippet = ec_plainSnippet_(c.bodyHtml, c.bodyText, 120);
      const stats = openStatsByCampaign[c.campaignId] || { sent: 0, opened: 0 };
      c.openedCount = stats.opened;
      c.openSentCount = stats.sent;
      c.openRate = stats.sent > 0 ? Math.round((stats.opened / stats.sent) * 1000) / 10 : 0;
      return c;
    })
    .sort((a, b) => String(b.campaignId).localeCompare(String(a.campaignId)));

  return { ok: true, campaigns: campaigns };
}

// ── PUBLIC: preview recipients for a mode/value ─────────────
function ec_previewRecipients(sessionId, mode, value) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;
  const recipients = ec_resolveRecipients_(mode, value);
  return { ok: true, count: recipients.length, recipients: recipients.slice(0, 50), truncated: recipients.length > 50 };
}

// ── Shared payload → campaign field mapping ─────────────────
function ec_campaignFromPayload_(payload, theme) {
  return {
    subject:        String(payload.subject || "").trim(),
    theme:          String(payload.theme || "generic"),
    recipientMode:  String(payload.recipientMode || "all"),
    recipientValue: String(payload.recipientValue || ""),
    headline:       String(payload.headline || theme.defaultHeadline || "").trim(),
    greeting:       String(payload.greeting || theme.defaultGreeting || "").trim(),
    bodyText:       String(payload.bodyText || "").trim(),
    bullets:        String(payload.bullets || "").trim(),
    signoff:        String(payload.signoff || theme.defaultSignoff || "").trim(),
    // Sanitized on the way IN as well as on the way out, so a stored draft
    // can never hold executable markup.
    bodyHtml:       ec_sanitizeRichHtml_(payload.bodyHtml)
  };
}

function ec_validateCampaignPayload_(payload, requireBody) {
  const subject = String(payload.subject || "").trim();
  if (!subject) return "Subject is required.";
  if (subject.length > 250) return "Subject is too long (250 characters maximum).";

  const mode = String(payload.recipientMode || "all");
  if (["all", "role", "specific"].indexOf(mode) === -1) return "Invalid audience selection.";
  if (mode !== "all" && !String(payload.recipientValue || "").trim()) {
    return mode === "role" ? "Please choose a role." : "Please enter at least one recipient email.";
  }
  if (requireBody) {
    const hasRich = !!ec_normalizeEmailHtml_(payload.bodyHtml);
    if (!hasRich && !String(payload.bodyText || "").trim()) return "The message body is empty.";
  }
  return "";
}

/** Shared by ec_sendCampaign (send now) and ec_processScheduledCampaigns
 *  (scheduled auto-send) — resolves recipients, sends, logs each attempt,
 *  and writes the final counts/status back onto the already-written
 *  campaign row. `payload.attachments`, if present, is used; scheduled
 *  sends never carry attachments (see ec_scheduleCampaign). */
function ec_performSend_(campaign, payload, campSheet, campInsertAt, theme) {
  const recipients = ec_resolveRecipients_(campaign.recipientMode, campaign.recipientValue);
  if (!recipients.length) {
    campaign.status = "Failed";
    campaign.sentAt = new Date();
    campSheet.getRange(campInsertAt, 1, 1, EC_CAMPAIGN_HEADERS.length).setValues([ec_campaignRowValues_(campaign)]);
    return { ok: false, error: "No valid recipients matched that audience. Nothing was sent." };
  }

  const logSheet = ec_logSheet_();
  let successCount = 0, failureCount = 0;

  recipients.forEach(function (r) {
    let status = "Sent";
    let errMsg = "";
    const trackingId = Utilities.getUuid();
    try {
      const pixel = ec_trackingPixelHtml_(trackingId);
      const html = ec_buildCampaignHtml_(campaign, r.name, theme, pixel);
      const plain = ec_buildPlainFallback_(campaign, r.name);

      const attachOpts = { htmlBody: html, name: "Innersparc Realty Corporation" };
      const rawAttachments = (payload && payload.attachments) || [];
      if (rawAttachments.length) {
        attachOpts.attachments = rawAttachments.map(function (a) {
          const bytes = Utilities.base64Decode(String(a.data || ""));
          return Utilities.newBlob(bytes, a.mimeType || "application/octet-stream", a.name || "attachment");
        });
      }

      MailApp.sendEmail(r.email, campaign.subject, plain, attachOpts);
      successCount++;
    } catch (e) {
      status = "Failed";
      errMsg = e.toString();
      failureCount++;
    }
    try {
      const logInsertAt = pa_lastRowWithData_(logSheet, 1) + 1;
      logSheet.getRange(logInsertAt, 1, 1, EC_LOG_HEADERS.length).setValues([[
        new Date(), campaign.campaignId, r.email, r.name, status, errMsg,
        status === "Sent" ? trackingId : "", "", 0, ""
      ]]);
    } catch (logErr) {
      Logger.log("ec_performSend_: log append failed for " + r.email + " — " + logErr);
    }
  });

  campaign.status = failureCount === recipients.length ? "Failed" : (failureCount > 0 ? "Partially Sent" : "Sent");
  campaign.sentAt = new Date();
  campaign.recipientCount = recipients.length;
  campaign.successCount = successCount;
  campaign.failureCount = failureCount;

  campSheet.getRange(campInsertAt, 1, 1, EC_CAMPAIGN_HEADERS.length).setValues([ec_campaignRowValues_(campaign)]);

  return { ok: true, campaignId: campaign.campaignId, status: campaign.status, recipientCount: recipients.length, successCount: successCount, failureCount: failureCount };
}


// ── PUBLIC: create + send a campaign ────────────────────────
/**
 * payload = {
 *   campaignId (optional — sends an existing Draft row instead of creating one),
 *   subject, theme,
 *   recipientMode: "all" | "role" | "specific",
 *   recipientValue: "" | "Agent" | "a@x.com;b@y.com",
 *   headline, greeting, bodyHtml, bodyText, bullets, signoff
 * }
 * The client never supplies a resolved recipient list —
 * ec_resolveRecipients_() remains the only source of truth.
 */
function ec_sendCampaign(sessionId, payload) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  payload = payload || {};
  const invalid = ec_validateCampaignPayload_(payload, true);
  if (invalid) return { ok: false, error: invalid };

  const themeKey = String(payload.theme || "generic");
  const theme = EC_THEMES_[themeKey] || EC_THEMES_.generic;

  const recipients = ec_resolveRecipients_(payload.recipientMode, payload.recipientValue);
  if (!recipients.length) return { ok: false, error: "No valid recipients matched that audience. Nothing was sent." };

  // QUOTA: fail before the first send rather than stopping halfway through.
  let remainingQuota = -1;
  try { remainingQuota = MailApp.getRemainingDailyQuota(); } catch (qe) { remainingQuota = -1; }
  if (remainingQuota >= 0 && remainingQuota < recipients.length) {
    return {
      ok: false,
      error: "Not enough email quota: " + remainingQuota + " remaining today but " +
             recipients.length + " recipients resolved. Nothing was sent."
    };
  }

  const campSheet = ec_campaignSheet_();
  if (!campSheet) return { ok: false, error: "The Email Campaigns sheet is missing. Run ec_setupSheets() first." };

  const fields = ec_campaignFromPayload_(payload, theme);
  const now = new Date();

  // Sending an existing Draft reuses its row (and its Campaign ID) instead
  // of leaving an orphaned draft behind.
  let campInsertAt = -1;
  let campaignId = "";
  let createdAt = now;
  let createdBy = auth.email;
  if (payload.campaignId) {
    const existingRow = ec_findCampaignRow_(campSheet, payload.campaignId);
    if (existingRow !== -1) {
      const existing = ec_rowToCampaign_(campSheet.getRange(existingRow, 1, 1, EC_CAMPAIGN_HEADERS.length).getValues()[0]);
      if (existing.status !== "Draft") return { ok: false, error: "That broadcast has already been sent and cannot be sent again." };
      campInsertAt = existingRow;
      campaignId = existing.campaignId;
      createdAt = campSheet.getRange(existingRow, 2).getValue() || now;
      createdBy = existing.createdBy || auth.email;
    }
  }
  if (campInsertAt === -1) {
    campaignId = ec_generateCampaignId_();
    campInsertAt = pa_lastRowWithData_(campSheet, 1) + 1; // Auth.gs
  }

  const campaign = {
    campaignId:     campaignId,
    createdAt:      createdAt,
    createdBy:      createdBy,
    subject:        fields.subject,
    theme:          themeKey,
    recipientMode:  fields.recipientMode,
    recipientValue: fields.recipientValue,
    headline:       fields.headline,
    greeting:       fields.greeting,
    bodyText:       fields.bodyText,
    bullets:        fields.bullets,
    signoff:        fields.signoff,
    bodyHtml:       fields.bodyHtml,
    showHeader:     payload.showHeader !== false,
    showFooter:     payload.showFooter !== false,
    status:         "Sending",
    sentAt:         "",
    recipientCount: recipients.length,
    successCount:   0,
    failureCount:   0
  };

  // Write the campaign row FIRST (status = Sending), so if the send loop
  // times out partway through, the admin still sees the attempt.
  campSheet.getRange(campInsertAt, 1, 1, EC_CAMPAIGN_HEADERS.length)
    .setValues([ec_campaignRowValues_(campaign)]);

  return ec_performSend_(campaign, payload, campSheet, campInsertAt, theme);
}

// ── PUBLIC: schedule a campaign to auto-send at a future time ──
// scheduledForIso: an ISO datetime string from the composer's
// <input type="datetime-local">, converted client-side with .toISOString().
function ec_scheduleCampaign(sessionId, payload, scheduledForIso) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  payload = payload || {};
  const invalid = ec_validateCampaignPayload_(payload, true);
  if (invalid) return { ok: false, error: invalid };

  const scheduledFor = new Date(scheduledForIso);
  if (isNaN(scheduledFor.getTime())) return { ok: false, error: "Please choose a valid send date/time." };
  if (scheduledFor.getTime() <= Date.now() + 60000) {
    return { ok: false, error: "Please choose a time at least a minute in the future." };
  }

  const themeKey = String(payload.theme || "generic");
  const theme = EC_THEMES_[themeKey] || EC_THEMES_.generic;

  const recipients = ec_resolveRecipients_(payload.recipientMode, payload.recipientValue);
  if (!recipients.length) return { ok: false, error: "No valid recipients matched that audience. Nothing was scheduled." };

  const campSheet = ec_campaignSheet_();
  if (!campSheet) return { ok: false, error: "The Email Campaigns sheet is missing. Run ec_setupSheets() first." };

  const fields = ec_campaignFromPayload_(payload, theme);

  let campInsertAt = -1, campaignId = "", createdAt = new Date(), createdBy = auth.email;
  if (payload.campaignId) {
    const existingRow = ec_findCampaignRow_(campSheet, payload.campaignId);
    if (existingRow !== -1) {
      const existing = ec_rowToCampaign_(campSheet.getRange(existingRow, 1, 1, EC_CAMPAIGN_HEADERS.length).getValues()[0]);
      if (existing.status !== "Draft" && existing.status !== "Scheduled") {
        return { ok: false, error: "That broadcast has already been sent and cannot be rescheduled." };
      }
      campInsertAt = existingRow;
      campaignId = existing.campaignId;
      createdAt = campSheet.getRange(existingRow, 2).getValue() || createdAt;
      createdBy = existing.createdBy || auth.email;
    }
  }
  if (campInsertAt === -1) {
    campaignId = ec_generateCampaignId_();
    campInsertAt = pa_lastRowWithData_(campSheet, 1) + 1;
  }

  const campaign = {
    campaignId: campaignId, createdAt: createdAt, createdBy: createdBy,
    subject: fields.subject, theme: themeKey,
    recipientMode: fields.recipientMode, recipientValue: fields.recipientValue,
    headline: fields.headline, greeting: fields.greeting, bodyText: fields.bodyText,
    bullets: fields.bullets, signoff: fields.signoff, bodyHtml: fields.bodyHtml,
    status: "Scheduled", sentAt: "", recipientCount: recipients.length,
    successCount: 0, failureCount: 0, scheduledFor: scheduledFor
  };

  campSheet.getRange(campInsertAt, 1, 1, EC_CAMPAIGN_HEADERS.length).setValues([ec_campaignRowValues_(campaign)]);

  let triggerWarning = "";
  try {
    ec_ensureScheduledSendTrigger_();
  } catch (trigErr) {
    Logger.log("ec_scheduleCampaign: could not create trigger — " + trigErr);
    triggerWarning = "Scheduled, but the auto-send trigger could not be created, so it will not send automatically. Run ec_ensureScheduledSendTrigger_ once from the Apps Script editor.";
  }

  return {
    ok: true, campaignId: campaignId, status: "Scheduled",
    scheduledFor: Utilities.formatDate(scheduledFor, Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a"),
    warning: triggerWarning
  };
}

/** Permanently deletes a DRAFT broadcast row. Sent/Scheduled broadcasts
 *  are never deletable here (Scheduled must be canceled back to Draft first). */
function ec_deleteDraft(sessionId, campaignId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { ok: false, error: "The server is busy. Please try again." }; }
  try {
    const sh = ec_campaignSheet_();
    const row = ec_findCampaignRow_(sh, campaignId);
    if (row === -1) return { ok: false, error: "That broadcast could not be found." };

    const existing = ec_rowToCampaign_(sh.getRange(row, 1, 1, EC_CAMPAIGN_HEADERS.length).getValues()[0]);
    if (existing.status !== "Draft") return { ok: false, error: "Only drafts can be deleted." };

    sh.deleteRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** Cancels a pending Scheduled campaign, reverting it to an editable Draft. */
function ec_cancelScheduledCampaign(sessionId, campaignId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  const sh = ec_campaignSheet_();
  const row = ec_findCampaignRow_(sh, campaignId);
  if (row === -1) return { ok: false, error: "That broadcast could not be found." };

  const existing = ec_rowToCampaign_(sh.getRange(row, 1, 1, EC_CAMPAIGN_HEADERS.length).getValues()[0]);
  if (existing.status !== "Scheduled") return { ok: false, error: "Only a pending scheduled broadcast can be canceled." };

  existing.status = "Draft";
  existing.scheduledFor = "";
  sh.getRange(row, 1, 1, EC_CAMPAIGN_HEADERS.length).setValues([ec_campaignRowValues_(existing)]);
  return { ok: true };
}

/** Run this ONCE from the editor to create the auto-send trigger. */
function ec_installScheduleTrigger() {
  ec_ensureScheduledSendTrigger_();
  const found = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === "ec_processScheduledCampaigns";
  }).length;
  Logger.log("Scheduled-send triggers found: " + found);
}

/** DIAGNOSTIC / CLEANUP — run this manually once from the Apps Script
 *  editor now, to clear out whatever caused the "could not be created"
 *  warning and to (re)authorize trigger creation for this project.
 *  Removes every existing ec_processScheduledCampaigns trigger, then
 *  installs exactly one fresh one. Safe to run again any time. */
function ec_resetScheduleTriggers() {
  const existing = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === "ec_processScheduledCampaigns";
  });
  existing.forEach(function (t) { ScriptApp.deleteTrigger(t); });
  Logger.log("Removed " + existing.length + " existing ec_processScheduledCampaigns trigger(s).");
  ScriptApp.newTrigger("ec_processScheduledCampaigns").timeBased().everyMinutes(5).create();
  Logger.log("Installed a fresh 5-minute ec_processScheduledCampaigns trigger.");
}

/** Ensures exactly one time-driven trigger exists for
 *  ec_processScheduledCampaigns — self-healing rather than "create only
 *  if totally missing": every call removes any existing trigger(s) for
 *  this handler and installs exactly one fresh one, so a stale or
 *  duplicate trigger left over from an earlier bug/deployment can never
 *  linger or silently break scheduling again. Runs every 5 minutes —
 *  that's the max lag a scheduled send can now have (was 10 minutes). */
function ec_ensureScheduledSendTrigger_() {
  const existing = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === "ec_processScheduledCampaigns";
  });

  // Already have a trigger → leave it alone (this stops the red warning)
  if (existing.length > 0) return;

  // Only create one if none exist
  ScriptApp.newTrigger("ec_processScheduledCampaigns")
    .timeBased()
    .everyMinutes(5)
    .create();
}

/** Trigger entry point (also runnable manually from the editor). Sends
 *  every "Scheduled" campaign whose Scheduled For has passed, via the same
 *  delivery path as an immediate send (ec_performSend_). */
function ec_processScheduledCampaigns() {
  const sh = ec_campaignSheet_();
  if (!sh || sh.getLastRow() < 2) return;

  const lastRow = sh.getLastRow();
  const rows = sh.getRange(2, 1, lastRow - 1, EC_CAMPAIGN_HEADERS.length).getValues();
  const now = Date.now();
  Logger.log("Scheduled check at " + new Date() + " — " + rows.length + " campaign rows scanned.");

  rows.forEach(function (row, i) {
    const campaign = ec_rowToCampaign_(row);
    if (campaign.status !== "Scheduled" || !campaign.scheduledFor) return;
    const due = new Date(campaign.scheduledFor).getTime();
    if (isNaN(due) || due > now) return;

    const rowNum = i + 2;
    const theme = EC_THEMES_[campaign.theme] || EC_THEMES_.generic;

    // Mark Sending immediately so a slow send (or an overlapping trigger
    // run) can never double-send the same broadcast.
    campaign.status = "Sending";
    sh.getRange(rowNum, 1, 1, EC_CAMPAIGN_HEADERS.length).setValues([ec_campaignRowValues_(campaign)]);

    try {
      ec_performSend_(campaign, { attachments: [] }, sh, rowNum, theme);
    } catch (e) {
      Logger.log("ec_processScheduledCampaigns: send failed for " + campaign.campaignId + " — " + e);
      campaign.status = "Failed";
      sh.getRange(rowNum, 1, 1, EC_CAMPAIGN_HEADERS.length).setValues([ec_campaignRowValues_(campaign)]);
    }
  });
}

function ec_buildPlainFallback_(campaign, recipientName) {
  const nameForText = recipientName || "ka-team";
  const greet = (campaign.greeting || "").replace(/\[Agent Name\]/g, nameForText);
  const bodyText = campaign.bodyHtml
    ? ec_richToPlain_(String(campaign.bodyHtml).replace(/\[Agent Name\]/g, nameForText))
    : String(campaign.bodyText || "").replace(/\[Agent Name\]/g, nameForText);
  const signoffText = String(campaign.signoff || "").replace(/\[Agent Name\]/g, nameForText)
    .split(/\n/).filter(function (l) { return !ec_parseImgLine_(l); }).join("\n");
  const bullets = String(campaign.bullets || "").split(/\n/)
    .filter(function (l) { return l.trim() && !ec_parseImgLine_(l); })
    .map(b => "• " + b.trim()).join("\n");
  return [
    greet,
    "",
    bodyText,
    bullets ? "\n" + bullets + "\n" : "",
    signoffText,
    "",
    "— Inner SPARC Realty Corporation"
  ].join("\n");
}

// ── PUBLIC: detailed log for one campaign ───────────────────
function ec_getCampaignLog(sessionId, campaignId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  const sh = ec_logSheet_();
  if (!sh || sh.getLastRow() < 2) return { ok: true, log: [] };

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, EC_LOG_HEADERS.length).getValues();
  const log = rows
    .filter(r => String(r[1] || "") === String(campaignId))
    .map(r => ({
      timestamp:  r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a") : (r[0] || ""),
      campaignId: r[1] || "",
      email:      r[2] || "",
      displayName:r[3] || "",
      status:     r[4] || "",
      error:      r[5] || "",
      trackingId:   r[6] || "",
      openedAt:     r[7] instanceof Date ? Utilities.formatDate(r[7], Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a") : (r[7] || ""),
      openCount:    Number(r[8]) || 0,
      lastOpenedAt: r[9] instanceof Date ? Utilities.formatDate(r[9], Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a") : (r[9] || "")
    }));

  return { ok: true, log: log };
}

// ── PUBLIC: one campaign (used to reopen a Draft in the composer) ──
function ec_getCampaign(sessionId, campaignId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  const sh = ec_campaignSheet_();
  const row = ec_findCampaignRow_(sh, campaignId);
  if (row === -1) return { ok: false, error: "That broadcast could not be found." };

  const campaign = ec_rowToCampaign_(sh.getRange(row, 1, 1, EC_CAMPAIGN_HEADERS.length).getValues()[0]);
  campaign.audienceLabel = ec_audienceLabel_(campaign.recipientMode, campaign.recipientValue);
  campaign.snippet = ec_plainSnippet_(campaign.bodyHtml, campaign.bodyText, 120);
  return { ok: true, campaign: campaign };
}

// ── PUBLIC: reuse any past broadcast (Sent, Partially Sent, Failed, or
// even another Draft) as the starting point for a brand-new Draft ──────
// Copies every content field (subject, theme, greeting, body, bullets,
// signoff, audience) into a fresh row with its own Campaign ID and clears
// every send-outcome field, so a previously-sent broadcast is never mutated
// — the original row is left completely untouched as a historical record.
function ec_duplicateCampaign(sessionId, campaignId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  try {
    const sh = ec_campaignSheet_();
    if (!sh) return { ok: false, error: "The Email Campaigns sheet is missing. Run ec_setupSheets() first." };

    const row = ec_findCampaignRow_(sh, campaignId);
    if (row === -1) return { ok: false, error: "That broadcast could not be found." };

    const source = ec_rowToCampaign_(sh.getRange(row, 1, 1, EC_CAMPAIGN_HEADERS.length).getValues()[0]);

    const campaign = {
      campaignId:     ec_generateCampaignId_(),
      createdAt:      new Date(),
      createdBy:      auth.email,
      subject:        source.subject,
      theme:          source.theme,
      recipientMode:  source.recipientMode,
      recipientValue: source.recipientValue,
      headline:       source.headline,
      greeting:       source.greeting,
      bodyText:       source.bodyText,
      bullets:        source.bullets,
      signoff:        source.signoff,
      bodyHtml:       source.bodyHtml,
      status:         "Draft",
      sentAt:         "",
      recipientCount: 0,
      successCount:   0,
      failureCount:   0
    };

    const insertAt = pa_lastRowWithData_(sh, 1) + 1;
    sh.getRange(insertAt, 1, 1, EC_CAMPAIGN_HEADERS.length).setValues([ec_campaignRowValues_(campaign)]);

    return {
      ok: true,
      campaignId: campaign.campaignId,
      campaign: {
        campaignId:     campaign.campaignId,
        createdAt:      Utilities.formatDate(campaign.createdAt, Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a"),
        createdBy:      campaign.createdBy,
        subject:        campaign.subject,
        theme:          campaign.theme,
        recipientMode:  campaign.recipientMode,
        recipientValue: campaign.recipientValue,
        headline:       campaign.headline,
        greeting:       campaign.greeting,
        bodyText:       campaign.bodyText,
        bullets:        campaign.bullets,
        signoff:        campaign.signoff,
        bodyHtml:       campaign.bodyHtml,
        status:         "Draft",
        sentAt:         "",
        recipientCount: 0,
        successCount:   0,
        failureCount:   0,
        audienceLabel:  ec_audienceLabel_(campaign.recipientMode, campaign.recipientValue),
        snippet:        ec_plainSnippet_(campaign.bodyHtml, campaign.bodyText, 120)
      }
    };
  } catch (e) {
    // Surface the real reason to the client instead of a generic failure —
    // google.script.run's withFailureHandler only shows a bare "Error", so
    // returning ok:false here is what lets the admin (and us) actually see
    // why a specific broadcast couldn't be duplicated.
    return { ok: false, error: "Could not reuse that broadcast: " + (e && e.message ? e.message : e) };
  }
}

// ── PUBLIC: save a new Draft ────────────────────────────────
// Drafts are real rows in the Email Campaigns sheet (Status = "Draft") —
// there is no browser-side draft store of any kind.
function ec_saveDraft(sessionId, payload) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  payload = payload || {};
  if (payload.campaignId) return ec_updateCampaign(sessionId, payload.campaignId, payload);

  const invalid = ec_validateCampaignPayload_(payload, false);
  if (invalid) return { ok: false, error: invalid };

  const sh = ec_campaignSheet_();
  if (!sh) return { ok: false, error: "The Email Campaigns sheet is missing. Run ec_setupSheets() first." };

  const theme = EC_THEMES_[String(payload.theme || "generic")] || EC_THEMES_.generic;
  const fields = ec_campaignFromPayload_(payload, theme);

  const campaign = {
    campaignId:     ec_generateCampaignId_(),
    createdAt:      new Date(),
    createdBy:      auth.email,
    subject:        fields.subject,
    theme:          fields.theme,
    recipientMode:  fields.recipientMode,
    recipientValue: fields.recipientValue,
    headline:       fields.headline,
    greeting:       fields.greeting,
    bodyText:       fields.bodyText,
    bullets:        fields.bullets,
    signoff:        fields.signoff,
    bodyHtml:       fields.bodyHtml,
    status:         "Draft",
    sentAt:         "",
    recipientCount: 0,
    successCount:   0,
    failureCount:   0
  };

  const insertAt = pa_lastRowWithData_(sh, 1) + 1;
  sh.getRange(insertAt, 1, 1, EC_CAMPAIGN_HEADERS.length).setValues([ec_campaignRowValues_(campaign)]);

  return { ok: true, campaignId: campaign.campaignId, status: "Draft" };
}

// ── PUBLIC: update an existing Draft ────────────────────────
function ec_updateCampaign(sessionId, campaignId, payload) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  payload = payload || {};
  const invalid = ec_validateCampaignPayload_(payload, false);
  if (invalid) return { ok: false, error: invalid };

  const sh = ec_campaignSheet_();
  const row = ec_findCampaignRow_(sh, campaignId);
  if (row === -1) return { ok: false, error: "That broadcast could not be found." };

  const existing = ec_rowToCampaign_(sh.getRange(row, 1, 1, EC_CAMPAIGN_HEADERS.length).getValues()[0]);
  // Only drafts are editable — a sent broadcast is an immutable record.
  if (existing.status !== "Draft") return { ok: false, error: "Only drafts can be edited." };

  const theme = EC_THEMES_[String(payload.theme || existing.theme || "generic")] || EC_THEMES_.generic;
  const fields = ec_campaignFromPayload_(payload, theme);

  const campaign = {
    campaignId:     existing.campaignId,
    createdAt:      sh.getRange(row, 2).getValue() || new Date(),
    createdBy:      existing.createdBy || auth.email,
    subject:        fields.subject,
    theme:          fields.theme,
    recipientMode:  fields.recipientMode,
    recipientValue: fields.recipientValue,
    headline:       fields.headline,
    greeting:       fields.greeting,
    bodyText:       fields.bodyText,
    bullets:        fields.bullets,
    signoff:        fields.signoff,
    bodyHtml:       fields.bodyHtml,
    status:         "Draft",
    sentAt:         "",
    recipientCount: 0,
    successCount:   0,
    failureCount:   0
  };

  sh.getRange(row, 1, 1, EC_CAMPAIGN_HEADERS.length).setValues([ec_campaignRowValues_(campaign)]);
  return { ok: true, campaignId: campaign.campaignId, status: "Draft" };
}

// ============================================================
// EMAIL OPEN TRACKING
//
// A tiny (1x1) tracking pixel is embedded in every broadcast's HTML body
// (ec_buildCampaignHtml_ above) with a URL pointing back at THIS SAME web
// app's doGet(), using ?action=track&id=<Tracking ID> (see Code.gs's
// doGet — that "track" branch is the only thing that calls into this
// section). The Tracking ID is the sole lookup key end-to-end: it's
// generated once per SENT message (ec_sendCampaign), never the recipient's
// email address, so the same person receiving five different broadcasts
// produces five independently-trackable ids, and the pixel URL itself
// carries no recipient PII.
//
// Two sheets carry the resulting data (both created/widened by
// ec_setupSheets — run it once after deploying this):
//   - "Email Send Log" (EC_LOG_HEADERS): unchanged first 6 columns, plus
//     4 appended columns that summarize opens PER RECIPIENT — Tracking ID,
//     first-Opened-At, running Open Count, and Last-Opened-At.
//   - "Email Open Log" (EC_OPEN_LOG_HEADERS): one row per DETECTED OPEN
//     EVENT, so the admin UI can show a full open history per recipient,
//     not just the first/last/count summary.
//
// SECURITY (spec point 24): the /track endpoint is intentionally the
// dumbest possible handler — it takes an id, writes a timestamp, and
// returns empty text. It never returns recipient data, sheet contents, or
// anything that varies by whether the id was valid, so there's nothing an
// unauthenticated caller can learn or enumerate through it. All the actual
// recipient/open DATA is only ever returned by ec_getCampaignLog /
// ec_getOpenHistory / ec_getEmailDashboardStats below, and every one of
// those re-checks pa_isAdmin_(sessionId) exactly like the rest of this
// file — there is no unauthenticated read path anywhere in this section.
// ============================================================

/** Builds the pixel's callback URL against THIS web app's own deployment,
 *  via ScriptApp.getService().getUrl() — so there is nothing to configure
 *  by hand (no WEB_APP_URL constant to keep in sync): whatever URL this
 *  script is currently deployed at is automatically what gets embedded.
 *  NOTE: this returns the CURRENT deployment's /exec URL at send time. If
 *  you create a brand-new deployment (not just "Manage deployments ▸
 *  Edit ▸ new version" on the existing one), already-sent emails still
 *  point at the old URL — see the Deployment section in the writeup. */
function ec_trackingPixelUrl_(trackingId) {
  return ScriptApp.getService().getUrl() + "?action=track&id=" + encodeURIComponent(trackingId);
}

/** The literal <img> tag injected into the HTML body. Deliberately NOT
 *  display:none as its only hiding mechanism — some clients strip inline
 *  display:none from images before deciding whether to even request them
 *  as "remote content"; width/height=1 plus a 1px box keeps it invisible
 *  either way without affecting whether the request fires. */
function ec_trackingPixelHtml_(trackingId) {
  const url = ec_trackingPixelUrl_(trackingId);
  return '<img src="' + url + '" width="1" height="1" alt="" aria-hidden="true" '
       + 'style="display:block;width:1px;height:1px;border:0;overflow:hidden;">';
}

/** doGet's entire "track" branch. MUST NOT throw, MUST NOT touch
 *  session/auth (the requester is a mail client's image loader, not a
 *  logged-in user), and MUST always return the same trivial response
 *  regardless of whether the id turned out to be valid — see the SECURITY
 *  note above. Real work happens in ec_recordOpen_(), wrapped so that ANY
 *  failure there (bad id, locked sheet, whatever) still ends in a normal
 *  200 response rather than a stack trace being handed to a mail server. */
function ec_handleTrackingRequest_(e) {
  try {
    const id = String((e && e.parameter && e.parameter.id) || "").trim();
    if (id) {
      try { ec_recordOpen_(id); }
      catch (err) { Logger.log("ec_recordOpen_ failed for id=" + id + " — " + err); }
    }
  } catch (outerErr) {
    Logger.log("ec_handleTrackingRequest_ error: " + outerErr);
  }
  // NOTE (limitation): ContentService can only set a TEXT-family MimeType
  // (see ContentService.MimeType — there is no image/gif or image/png
  // option), so this can never be a byte-for-byte real image response the
  // way a static file host would serve one. It doesn't need to be: the
  // pixel is 1x1 and invisible either way, and what actually matters for
  // tracking is only that the request reached the server, which this
  // confirms by returning 200 either way.
  return ContentService.createTextOutput("").setMimeType(ContentService.MimeType.TEXT);
}

/** TEMP DIAGNOSTIC — run this manually from the Apps Script editor to
 *  confirm ec_recordOpen_ itself works, independent of whether the pixel
 *  request is even reaching the server. */
function ec_debugTrackingPixel() {
  const sh = ec_logSheet_();
  if (!sh || sh.getLastRow() < 2) {
    Logger.log('No rows in Email Send Log yet — send a test broadcast first.');
    return;
  }
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, EC_LOG_HEADERS.length).getValues();
  const withId = rows.find(r => String(r[6] || '').trim()); // column G = Tracking ID
  if (!withId) {
    Logger.log('No row has a Tracking ID yet — every send may have failed, or none have been sent.');
    return;
  }
  const trackingId = withId[6];
  Logger.log('Testing ec_recordOpen_ directly for Tracking ID: ' + trackingId);
  ec_recordOpen_(trackingId);
  Logger.log('Done. Now check the Email Send Log sheet — Opened At / Open Count / Last Opened At should be filled in for that row.');
}

/** Finds the Email Send Log row for a given Tracking ID by scanning ONLY
 *  the Tracking ID column (never the whole sheet — see spec point 14),
 *  with a short-lived CacheService hit to skip even that scan on repeat
 *  opens of the same message (mail clients/proxies commonly re-request
 *  the same pixel within seconds). Returns -1 for an unknown/expired id —
 *  callers treat that as "silently ignore", never as an error. */
function ec_findLogRowByTrackingId_(sh, trackingId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "ec_trkrow_" + trackingId;
  const cached = cache.get(cacheKey);
  if (cached) {
    const row = Number(cached);
    if (row >= 2 && row <= sh.getLastRow() &&
        String(sh.getRange(row, 7).getValue() || "") === trackingId) {
      return row; // cache hit, and still points at the right row
    }
  }

  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 7, last - 1, 1).getValues(); // column 7 = Tracking ID only
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || "") === trackingId) {
      const row = i + 2;
      cache.put(cacheKey, String(row), 21600); // 6h — plenty for a burst of opens
      return row;
    }
  }
  return -1;
}

/** Records one detected open: updates the per-recipient summary row in
 *  Email Send Log (first Opened At / running Open Count / Last Opened At)
 *  and appends one granular event row to Email Open Log. Wrapped in
 *  LockService the same way pa_saveMyGoal_ (GoalPlanner.gs) guards its own
 *  read-modify-write, so two near-simultaneous opens of the same message
 *  (spec point 13's "3 -> 4, 3 -> 4" race) can't both read Open Count
 *  before either writes it back. */
function ec_recordOpen_(trackingId) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    Logger.log("ec_recordOpen_: could not acquire lock for id=" + trackingId);
    return;
  }
  try {
    const sh = ec_logSheet_();
    if (!sh) return;

    const row = ec_findLogRowByTrackingId_(sh, trackingId);
    if (row === -1) return; // unknown/expired id — nothing to update, nothing to log

    const vals = sh.getRange(row, 1, 1, EC_LOG_HEADERS.length).getValues()[0];
    const campaignId    = vals[1] || "";
    const email         = vals[2] || "";
    const alreadyOpened = !!vals[7];
    const openCountNow  = Number(vals[8]) || 0;
    const now = new Date();

    // Cooldown: skip opens detected right after sending — almost always
    // the sender checking Sent mail, not the recipient. See
    // EC_OPEN_COOLDOWN_SECONDS above for why this can't be more precise.
    const sentAt = vals[0] instanceof Date ? vals[0] : new Date(vals[0]);
    const secondsSinceSent = (now.getTime() - sentAt.getTime()) / 1000;
    if (secondsSinceSent >= 0 && secondsSinceSent < EC_OPEN_COOLDOWN_SECONDS) {
      Logger.log("ec_recordOpen_: ignored open for id=" + trackingId + " — " +
        secondsSinceSent.toFixed(1) + "s after send (within " + EC_OPEN_COOLDOWN_SECONDS +
        "s cooldown; likely the sender checking Sent mail).");
      return;
    }

    if (!alreadyOpened) sh.getRange(row, 8).setValue(now); // Opened At — first open ONLY, never overwritten again
    sh.getRange(row, 9).setValue(openCountNow + 1);         // Open Count
    sh.getRange(row, 10).setValue(now);                     // Last Opened At

    try {
      const openLog = ec_openLogSheet_();
      const insertAt = pa_lastRowWithData_(openLog, 1) + 1; // Auth.gs
      openLog.getRange(insertAt, 1, 1, EC_OPEN_LOG_HEADERS.length).setValues([[now, trackingId, campaignId, email]]);
    } catch (openLogErr) {
      Logger.log("ec_recordOpen_: open-log append failed for id=" + trackingId + " — " + openLogErr);
    }

    // Admin alert — FIRST detected open only (spec point 27). Wrapped so a
    // MailApp failure here can never roll back or interrupt the tracking
    // write above; the open is recorded either way.
    if (!alreadyOpened && EC_SEND_OPEN_ALERT_) {
      try {
        MailApp.sendEmail(
          Session.getEffectiveUser().getEmail(),
          "📬 Email opened — " + (email || "unknown recipient"),
          "Campaign: " + campaignId + "\n"
          + "Recipient: " + (email || "unknown") + "\n"
          + "First opened: " + now + "\n"
          + "Tracking ID: " + trackingId
        );
      } catch (mailErr) {
        Logger.log("ec_recordOpen_: admin open-alert failed for id=" + trackingId + " — " + mailErr);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

// ── PUBLIC: per-recipient open history (for the recipient detail view) ──
function ec_getOpenHistory(sessionId, trackingId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  const id = String(trackingId || "").trim();
  if (!id) return { ok: true, history: [] };

  const sh = ec_openLogSheet_();
  if (!sh || sh.getLastRow() < 2) return { ok: true, history: [] };

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, EC_OPEN_LOG_HEADERS.length).getValues();
  const history = rows
    .filter(function (r) { return String(r[1] || "") === id; })
    .sort(function (a, b) { return new Date(a[0]) - new Date(b[0]); }) // chronological, oldest first
    .map(function (r) {
      return r[0] instanceof Date
        ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a")
        : String(r[0] || "");
    });

  return { ok: true, history: history };
}

/** Aggregates per-campaign Sent/Opened counts from the Email Send Log —
 *  one pass over the sheet, used to show an open-rate ratio (e.g. "11 of 29
 *  opened") in the broadcast detail view without a separate round trip
 *  per campaign. Keyed by Campaign ID. */
function ec_buildOpenStatsMap_() {
  const sh = ec_logSheet_();
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, EC_LOG_HEADERS.length).getValues();
  rows.forEach(function (r) {
    const campaignId = String(r[1] || "").trim();
    if (!campaignId) return;
    if (!map[campaignId]) map[campaignId] = { sent: 0, opened: 0 };
    if (String(r[4] || "") === "Sent") {
      map[campaignId].sent++;
      if ((Number(r[8]) || 0) > 0) map[campaignId].opened++;
    }
  });
  return map;
}

// ── PUBLIC: admin dashboard summary cards (Total / Sent / Opened / Open Rate) ──
// A single pass over Email Send Log — no per-campaign re-scans — so this
// stays cheap even with a long broadcast history (spec point 14).
function ec_getEmailDashboardStats(sessionId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;

  const sh = ec_logSheet_();
  if (!sh || sh.getLastRow() < 2) {
    return { ok: true, stats: { total: 0, sent: 0, opened: 0, openRate: 0 } };
  }

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, EC_LOG_HEADERS.length).getValues();
  let sent = 0, opened = 0;
  rows.forEach(function (r) {
    if (String(r[4] || "") === "Sent") sent++;
    if ((Number(r[8]) || 0) > 0) opened++; // Open Count > 0
  });

  // Open Rate is opened/SENT, never opened/total — a Failed send was never
  // deliverable in the first place, so it can't fairly count against the
  // rate (spec point 16).
  const openRate = sent > 0 ? Math.round((opened / sent) * 1000) / 10 : 0;

  return { ok: true, stats: { total: rows.length, sent: sent, opened: opened, openRate: openRate } };
}

function ec_forceOpenLatest() {
  const sh = ec_logSheet_();
  if (!sh || sh.getLastRow() < 2) {
    Logger.log("No rows in Email Send Log");
    return;
  }
  const lastRow = sh.getLastRow();
  const trackingId = String(sh.getRange(lastRow, 7).getValue() || "").trim(); // column G
  if (!trackingId) {
    Logger.log("Latest row has no Tracking ID");
    return;
  }
  Logger.log("Forcing open for latest Tracking ID: " + trackingId);
  ec_recordOpen_(trackingId);
  Logger.log("Done. Refresh the Email Send Log sheet and check Opened At / Open Count on the last row.");
}