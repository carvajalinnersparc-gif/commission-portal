// ============================================================
// Inner SPARC Realty Corporation — Agent Portal
// PortalData.gs — Read-only commission data for authenticated agents
//
// IMPORTANT: This file does NOT duplicate any commission math.
// It reads the same "<Year> Sales" sheets used by Code.gs and
// simply filters/aggregates rows belonging to the logged-in
// agent (matched via the Agent Access alias table).
// ============================================================

const PORTAL_SALES_SHEET_REGEX_ = /^\d{4} Sales$/;

/** Requires a valid session; throws-free — returns {ok, email, aliases, role, active} or {ok:false, error}.
 *  NOTE on the "Active" toggle: an Agent Access row with Active=FALSE (a
 *  resigned/inactive agent) is still a VALID account — it must still be
 *  able to log in. It is intentionally NOT rejected here anymore. What
 *  "inactive" restricts is which portal features that session can reach:
 *  callers that should be limited to a PDF-only commission view for
 *  inactive agents check the returned `active` flag themselves (see
 *  pa_requireActiveAgent_ below) rather than this function refusing the
 *  session outright. A missing Agent Access row entirely (deleted user)
 *  is still rejected below — that's a different case from "inactive". */
function pa_authorize_(sessionId) {
  const session = pa_validateSession_(sessionId);
  if (!session.valid) return { ok: false, error: "Session expired. Please log in again." };
  const access = pa_getAccessRecord_(session.email);
  if (!access) return { ok: false, error: "Access is no longer active for this account." };
  // Sales are matched on the aliases in the CMS "Agent Aliases" tab (looked up by
  // this login's email, else by Canonical Name). Agent Access aliases are only a
  // fallback for someone who has no row in that tab.
  // const cmsAliases = pa_cmsAliasesFor_(session.email, access.canonicalName);
  // return { ok: true, email: session.email, aliases: (cmsAliases && cmsAliases.length) ? cmsAliases : access.aliases, role: access.role, active: access.active };

  const cmsAliases = pa_cmsAliasesFor_(session.email, access.canonicalName) || [];
  const seen = {}, merged = [];
  cmsAliases.concat(access.aliases).forEach(n => {
    const k = pa_aliasKey_(n);
    if (k && !seen[k]) { seen[k] = true; merged.push(n); }
  });
  return { ok: true, email: session.email, aliases: merged, role: access.role, active: access.active };
}

/** Shared gate for portal features that inactive (resigned) agents should
 *  NOT reach — everything except the PDF-only commission view (dashboard)
 *  and profile/password management. Returns {ok:false, error} if the
 *  session belongs to an inactive account, or null if the caller should
 *  proceed normally. This is the server-side enforcement of the same
 *  restriction Portal_JS.html applies client-side (nav hidden, navTo()
 *  redirected to dashboard) — the client behavior is only a convenience,
 *  same as every other auth boundary in this portal. */
function pa_requireActiveAgent_(auth) {
  if (!auth.active) {
    return { ok: false, error: "Your account is inactive. You can still view and download your commission record as a PDF from your dashboard." };
  }
  return null;
}

function pa_norm_(s) {
  return String(s || '').trim().toUpperCase();
}

/** Comparison key for an exact alias match: upper-case, punctuation and extra
 *  spaces ignored ("PASION,LENIZA F." == "Pasion, Leniza F"). Word ORDER still
 *  matters — "LIM, LUZVIMINDA" and "LUZVIMINDA LIM" are separate aliases. */
function pa_aliasKey_(s) {
  return String(s || '').toUpperCase().replace(/[^A-Z0-9\u00C0-\u024F]+/g, ' ').replace(/\s+/g, ' ').trim();
}

/** Rows of the CMS "Agent Aliases" tab (A Canonical Name, B Aliases ';'-separated,
 *  C Email) — cached 5 min so edits in the sheet show up quickly. */
function pa_getCmsAliasRows_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('pa_cms_alias_rows_v1');
  if (hit) { try { return JSON.parse(hit); } catch (e) { /* rebuild */ } }
  const rows = [];
  try {
    const sh = getCmsSpreadsheet().getSheetByName('Agent Aliases'); // Code.gs
    if (sh && sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
        const canonical = String(r[0] || '').replace(/\s+/g, ' ').trim();
        if (!canonical) return;
        rows.push({ canonical: canonical, aliases: pa_parseAliasesCsv_(r[1]), email: pa_normEmail_(r[2]) }); // Auth.gs
      });
    }
  } catch (e) { console.error('pa_getCmsAliasRows_ failed:', e.toString()); }
  try { cache.put('pa_cms_alias_rows_v1', JSON.stringify(rows), 300); } catch (e) {}
  return rows;
}

/** Canonical name + aliases for one agent from the CMS "Agent Aliases" tab, or
 *  null if the tab has no row for them. Matches by email first, then by
 *  Canonical Name (most CMS rows have no email). Never matches on partial names. */
function pa_cmsAliasesFor_(email, canonicalName) {
  const rows = pa_getCmsAliasRows_();
  const em = pa_normEmail_(email);
  let row = em ? rows.find(r => r.email === em) : null;
  if (!row && canonicalName) {
    const k = pa_aliasKey_(canonicalName);
    row = rows.find(r => pa_aliasKey_(r.canonical) === k);
  }
  if (!row) return null;
  const seen = {}, out = [];
  [row.canonical].concat(row.aliases).forEach(n => {
    const k = pa_aliasKey_(n);
    if (k && !seen[k]) { seen[k] = true; out.push(n); }
  });
  return out;
}

/** Spreadsheets whose "<Year> Sales" tabs the portal reads. The Inner SPARC CMS
 *  holds every year. Set the flag below to true to ALSO merge the live sheet's
 *  (TARGET_SHEET_ID) current-year tab, e.g. if sales are recorded there too. */
const PORTAL_SALES_INCLUDE_LIVE_SHEET_ = false;
function pa_salesSpreadsheets_() {
  return PORTAL_SALES_INCLUDE_LIVE_SHEET_ ? [getCmsSpreadsheet(), getTargetSpreadsheet()] : [getCmsSpreadsheet()];
}

/** All years the Sales History year filter should offer: every "<year>
 *  Sales" sheet that actually exists, unioned with every year from 2020
 *  through the current calendar year. The 2020→today half of that union
 *  is computed from Date.getFullYear() at call time, so the filter's top
 *  end rolls forward automatically each January without any code change —
 *  no need to wait for a "2027 Sales" tab to be created before the 2027
 *  option shows up. Returned newest-first. */
function pa_getAllSalesYears_() {
  const years = new Set();
  // Historical years live only in the CMS (see CMS_SHEET_ID note in Code.gs);
  // the current year lives (live/authoritative) in TARGET_SHEET_ID. Union both.
  pa_salesSpreadsheets_().forEach(ss => {
    ss.getSheets().forEach(sh => {
      const name = sh.getName();
      if (PORTAL_SALES_SHEET_REGEX_.test(name)) {
        years.add(parseInt(name.match(/(\d{4})/)[1]));
      }
    });
  });
  const nowYear = new Date().getFullYear();
  for (let y = 2020; y <= nowYear; y++) years.add(y);
  return Array.from(years).sort((a, b) => b - a);
}

/** True if this role should see every sale company-wide rather than only
 *  rows matching its own Agent Access aliases. Admin, Supervisor, and
 *  Manager all get the unscoped, company-wide view; only Agent stays
 *  scoped to its own aliases. */
function pa_isCompanyWideRole_(role) {
  const r = String(role || '').trim().toUpperCase();
  return r === 'ADMIN' || r === 'SUPERVISOR' || r === 'MANAGER';
}

/** Resolves every "<Year> Sales" sheet per year across the two spreadsheets
 *  that can hold them. Returns a Map<year, Sheet[]>.
 *
 *  IMPORTANT (bug fix): this used to keep only ONE sheet per year — whichever
 *  was read last (TARGET_SHEET_ID, since it was processed second in the
 *  merge loop) — on the theory that TARGET_SHEET_ID is always the live,
 *  actively-edited copy and the CMS's copy of that same year is always a
 *  stale, no-longer-edited snapshot. That assumption doesn't hold in
 *  practice: both spreadsheets currently have a "2026 Sales" tab, and rows
 *  keep landing in BOTH of them (agents/admins still editing directly in
 *  the CMS out of old habit, alongside whatever the portal's own
 *  sale-recording flow writes into TARGET_SHEET_ID). Picking only
 *  TARGET_SHEET_ID's sheet silently discarded every row that only existed
 *  in the CMS's "2026 Sales" tab — which is exactly why 2026 sales could
 *  vanish from the Sales History view even though they're really present in
 *  the underlying data. Returning every matching sheet (instead of just the
 *  "winning" one) lets each caller merge rows across both copies by Sale ID
 *  instead of dropping one wholesale — see pa_getSalesForAliases_() and
 *  pa_pivotSalesAgentNames_() below, which both do this merge themselves.
 *  The CMS (Inner SPARC CMS, CMS_SHEET_ID in Code.gs) is the only place
 *  2020 Sales .. 2025 Sales still exist, since that history was never
 *  copied into TARGET_SHEET_ID. */
function pa_getYearSalesSheets_() {
  const byYear = new Map();
  pa_salesSpreadsheets_().forEach(ss => {
    ss.getSheets().forEach(sh => {
      const name = sh.getName();
      if (!PORTAL_SALES_SHEET_REGEX_.test(name)) return;
      const year = parseInt(name.match(/(\d{4})/)[1]);
      if (!byYear.has(year)) byYear.set(year, []);
      byYear.get(year).push(sh);
    });
  });
  return byYear;
}

/** Reads every "<Year> Sales" sheet (full 31 cols), one per year (see
 *  pa_getYearSalesSheets_), and returns rows matching any of the given
 *  aliases — or, when aliases is null, every row company-wide. Passing
 *  null (rather than an empty array) is what triggers the unscoped path,
 *  so an Admin with a genuinely empty alias list can never be confused
 *  with "give me everything." */
function pa_getSalesForAliases_(aliases) {

  const matchAll = aliases === null;

  const aliasList = matchAll
    ? []
    : aliases
        .map(a => String(a || '').trim())
        .filter(Boolean);

  const yearSheets = pa_getYearSalesSheets_();
  const results = [];

  // ============================================================
  // ALIAS MATCH — exact (case/punctuation/spacing-insensitive) match of the
  // row's "Sales Agent" text against the agent's aliases. No fuzzy or
  // word-overlap matching: a name that isn't an alias is not that agent's sale,
  // so one agent can never pick up another agent's rows.
  // ============================================================
  const aliasKeys = new Set(aliasList.map(pa_aliasKey_).filter(Boolean));

  function salesAgentMatchesAliases(agentName) {
    if (matchAll) return true;
    if (!agentName) return false;
    return aliasKeys.has(pa_aliasKey_(agentName));
  }


  // ============================================================
  // READ ALL YEAR SHEETS
  // ============================================================

  yearSheets.forEach((sheets, year) => {

    const bySaleId = new Map();


    sheets.forEach(sheet => {

      if (sheet.getLastRow() < 2) {
        return;
      }


      const numCols =
        Math.min(40, sheet.getLastColumn());

      const headerRow =
        sheet
          .getRange(1, 1, 1, numCols)
          .getValues()[0];


      const dataRows =
        sheet
          .getRange(
            2,
            1,
            sheet.getLastRow() - 1,
            numCols
          )
          .getValues();


      // ========================================================
      // HEADER MAP
      // ========================================================

      const colIdx = {};

      headerRow.forEach((h, i) => {

        const normalized =
          String(h || '')
            .trim()
            .replace(/\*+$/, '');

        colIdx[normalized] = i;

      });


      // Only "Sales Agent" is required. Older year tabs may have no Sale ID
      // column, or blank Sale IDs — those are still real sales and must not
      // be dropped, so they get a stable synthetic key (year + tab + row).
      if (!('Sales Agent' in colIdx)) {
        return;
      }


      // ========================================================
      // STORE SALES BY SALE ID
      // ========================================================

      dataRows.forEach((row, rowIdx) => {

        let saleId =
          row[colIdx['Sale ID']];

        if (!saleId) {
          const hasAgent = String(row[colIdx['Sales Agent']] || '').trim() !== '';
          const hasUnit = colIdx['Unit ID'] !== undefined && String(row[colIdx['Unit ID']] || '').trim() !== '';
          if (!hasAgent && !hasUnit) {
            return; // genuinely empty row
          }
          saleId = 'ROW-' + year + '-' + sheet.getSheetId() + '-' + (rowIdx + 2);
        }

        bySaleId.set(
          saleId,
          {
            row: row,
            colIdx: colIdx,
            key: saleId
          }
        );

      });

    });


    // ==========================================================
    // PROCESS SALES
    // ==========================================================

    bySaleId.forEach(({ row, colIdx, key }) => {

      const agentName =
        row[colIdx['Sales Agent']];


      // Skip cancelled sales
      // AE = column 31 = array index 30
      const saleStatus =
        String(row[30] || '')
          .trim()
          .toLowerCase();

      if (
        saleStatus === 'cancelled' ||
        saleStatus === 'canceled'
      ) {
        return;
      }
      // ========================================================
      // THIS IS THE IMPORTANT FIX
      // ========================================================

      if (
        !salesAgentMatchesAliases(agentName)
      ) {
        return;
      }


      const saleId =
        row[colIdx['Sale ID']] || key;


      // ========================================================
      // RESERVATION DATE
      // ========================================================

      let reservationDate =
        row[colIdx['Reservation Date']];

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


      // ========================================================
      // COMMISSION
      // ========================================================

      const agentNetComm =
        parseFloat(
          row[colIdx['Agent Net Comm (₱)']]
        ) || 0;


      const releasedAmount =
        parseFloat(
          row[colIdx['Released Commission']]
        ) || 0;


      // Remaining is always derived from the two live figures (Agent Net
      // Comm and Released Commission), never read off the sheet's
      // "Receivables" cell — that cell is filled in by hand and drifts out
      // of sync the moment a partial release happens (it was found stuck
      // at the full commission amount on ~49 of 76 units in one report,
      // which is why per-row Remaining values didn't foot to the grand
      // total). Deriving it here guarantees every row's Remaining always
      // equals Commission − Released, so the column always sums correctly.
      const receivablesAmount =
        agentNetComm > 0
          ? Math.max(0, agentNetComm - releasedAmount)
          : 0;


      // ========================================================
      // COMMISSION STATUS
      // ========================================================

      let status = '0%';


      if (agentNetComm > 0) {

        const pct =
          (releasedAmount / agentNetComm) *
          100;

        status =
          pct >= 99.99
            ? '100%'
            : pct.toFixed(1) + '%';

      }


      // ========================================================
      // DP RELEASE AMOUNTS
      // ========================================================

      let dpReleaseAmounts = [];

      const dpRaw =
        row[colIdx['DP Release Amount']];


      if (dpRaw) {

        dpReleaseAmounts =
          String(dpRaw)
            .split(',')
            .map(v =>
              parseFloat(v.trim())
            )
            .filter(n =>
              !isNaN(n)
            );

      }


      // ========================================================
      // PUSH RESULT
      // ========================================================

      results.push({

        year: year,

        saleId:
          saleId || '',

        reservationDate:
          reservationDate || '',

        salesAgent:
          agentName || '',

        developer:
          row[colIdx['Developer']] || '',

        project:
          row[colIdx['Project']] || '',

        unitId:
          row[colIdx['Unit ID']] || '',

        buyerName:
          row[colIdx["Buyer's Name"]] || '',

        dpTerm:
          row[colIdx['DP Term']] || '',

        activeSalesAmount:
          parseFloat(
            row[colIdx['ACTIVE SALES AMOUNT']]
          ) || 0,

        computedTSP:
          parseFloat(
            row[colIdx['Computed TSP']]
          ) || 0,

        developerCommRate:
          parseFloat(
            row[colIdx['Developer Comm Rate (%)']]
          ) || 0,

        companyGrossComm:
          parseFloat(
            row[colIdx['Company Gross Comm (₱)']]
          ) || 0,

        netVat:
          parseFloat(
            row[colIdx['Net VAT (₱)']]
          ) || 0,

        companyWithholdingTax:
          parseFloat(
            row[colIdx['Withholding Tax (₱)']]
          ) || 0,

        companyNetComm:
          parseFloat(
            row[colIdx['Company Net Comm (₱)']]
          ) || 0,

        agentCommRate:
          parseFloat(
            row[colIdx['Agent Comm Rate (%)']]
          ) || 0,

        agentGrossComm:
          parseFloat(
            row[colIdx['Agent Gross Comm (₱)']]
          ) || 0,

        vatShareAgent:
          parseFloat(
            row[colIdx['VAT Share (%)']]
          ) || 0,

        withholdingTaxAgent:
          parseFloat(
            row[colIdx['Withholding Tax (%)']]
          ) || 0,

        agentNetComm:
          agentNetComm,

        milestoneRatio:
          row[colIdx['Milestone Ratio']] || '--',

        dpTranche:
          row[colIdx['DP Tranche']] || '--',

        loanTakeoutTranche:
          row[colIdx['Loan Takeout Tranche']] || '--',

        releasedCommission:
          releasedAmount,

        receivables:
          receivablesAmount,

        commissionStatus:
          String(status || '0%'),

        dpReleaseAmounts:
          dpReleaseAmounts,

        loanTakeoutAmount:
          parseFloat(
            row[colIdx['Loan Takeout Release Amount']]
          ) || 0,

        cashAdvance:
          parseFloat(
            row[colIdx['Cash Advance']]
          ) || 0

      });

    });

  });


  // ============================================================
  // EXACT-DUPLICATE REMOVAL
  //
  // Two rows for the same Unit ID + Subdivision that also share the
  // same buyer, reservation date, and commission figures are the same
  // transaction encoded twice under two different Sale IDs (a
  // copy/paste error) — not a second legitimate tranche, since a real
  // tranche always differs in date and/or amount. Dropping the repeat
  // here, at the single place every sale list is built from, means
  // Dashboard "Units Sold", the Commission by Unit total, and every
  // peso total downstream all agree automatically — none of them can
  // double-count a duplicated row again.
  //
  // Anything with the same unit but a DIFFERENT buyer, date, or amount
  // is left untouched — the backend has no way to know which of two
  // different buyers on the same unit is correct, so that case is only
  // ever flagged (see pa_findDuplicateSaleRecords_ below), never
  // auto-removed.
  // ============================================================
  const seenExactSignatures_ = {};
  const deduped_ = results.filter(function(sale) {
    const key =
      String(sale.unitId || '').trim().toUpperCase() + '||' +
      String(sale.project || '').trim().toUpperCase() + '||' +
      pa_saleDuplicateSignature_(sale);
    if (seenExactSignatures_[key]) return false; // exact repeat — drop it
    seenExactSignatures_[key] = true;
    return true;
  });

  return deduped_;
}

function pa_statusBucket_(sale) {
  const pctStr = String(sale.commissionStatus || '0%').replace('%', '');
  const pct = parseFloat(pctStr) || 0;
  if (pct >= 99.99) return 'Fully Released';
  if (pct > 0) return 'Partially Released';
  return 'Pending';
}

// ============================================================
// CANONICAL UNIT GROUPING
//
// "Sales Count" on the dashboard and the unit total on Commission by
// Unit used to be computed two different ways in two different places
// (one counted raw sale rows, the other grouped by unit client-side) —
// which is why they could disagree. Both now go through this single
// function, so for ANY account the two numbers are the same by
// construction, not by coincidence.
//
// A unit is only considered "the same" if BOTH its Unit ID and its
// Subdivision/Project match (normalized: trim + uppercase). The same
// block/lot text legitimately gets reused across different
// subdivisions, so matching on Unit ID alone would wrongly merge
// unrelated sales.
//
// A sale with a blank Unit ID or Subdivision is never dropped — it's
// bucketed under a clearly-labeled fallback so it still counts toward
// every total.
// ============================================================

const PA_UNASSIGNED_UNIT_LABEL_ = '(No Unit ID)';
const PA_UNKNOWN_SUBDIVISION_LABEL_ = '(No Subdivision)';

/** Normalized identity for one sale row: which unit it belongs to. */
function pa_unitGroupKey_(sale) {
  const unit = String(sale.unitId || '').trim() || PA_UNASSIGNED_UNIT_LABEL_;
  const subdivision = String(sale.project || '').trim() || PA_UNKNOWN_SUBDIVISION_LABEL_;
  return {
    key: unit.toUpperCase() + '||' + subdivision.toUpperCase(),
    unit: unit,
    subdivision: subdivision
  };
}

/** Groups sale records into per-unit buckets. THE single source of
 *  truth for "how many units" a list of sales represents — every
 *  caller that needs a unit count or a per-unit commission total
 *  should go through this, not re-implement grouping locally. */
function pa_groupSalesByUnit_(sales) {
  const map = {};
  const order = [];

  (sales || []).forEach(sale => {
    const g = pa_unitGroupKey_(sale);

    if (!map[g.key]) {
      map[g.key] = {
        unit: g.unit,
        subdivision: g.subdivision,
        groupKey: g.key,
        reservationDate: sale.reservationDate || '',
        totalAgentCommission: 0,
        received: 0,
        remaining: 0,
        sales: []
      };
      order.push(g.key);
    }

    const bucket = map[g.key];
    if (!bucket.reservationDate && sale.reservationDate) {
      bucket.reservationDate = sale.reservationDate;
    }
    bucket.totalAgentCommission += Number(sale.agentNetComm) || 0;
    bucket.received += Number(sale.releasedCommission) || 0;
    bucket.remaining += Number(sale.receivables) || 0;
    bucket.sales.push(sale);
  });

  return order.map(k => map[k]);
}

// ============================================================
// DUPLICATE / CONFLICT DETECTION
//
// pa_groupSalesByUnit_ merges same-unit sales into one bucket, which is
// correct for totals — but it also silently hides the fact that a unit
// bucket can contain rows that shouldn't both exist. Two ways that
// happens in practice:
//
//   - exactDuplicate: the same transaction was encoded twice under two
//     different Sale IDs — same buyer, same reservation date, same
//     commission figures. This double-counts real money in every total
//     and should have one of the two rows removed at the source.
//
//   - buyerConflict: the same unit shows two *different* buyers. That's
//     only legitimate if the unit was genuinely resold (e.g. first
//     buyer cancelled) — otherwise it's almost always a mistyped Unit
//     ID on one of the rows. This needs a human to check against the
//     source sheet; we only flag it, we never guess which row is wrong.
//
// Units where the same buyer appears more than once with genuinely
// different dates/amounts (real multi-tranche commissions) are left
// unflagged — that's expected and not a data problem.
// ============================================================

/** Signature used to detect an exact-duplicate row within a unit group. */
function pa_saleDuplicateSignature_(sale) {
  return [
    String(sale.buyerName || '').trim().toUpperCase(),
    String(sale.reservationDate || '').trim(),
    Number(sale.agentNetComm) || 0,
    Number(sale.releasedCommission) || 0,
    Number(sale.receivables) || 0
  ].join('||');
}

/** Scans a sales list for units whose grouped sale rows look like an
 *  accidental duplicate entry or a same-unit/different-buyer conflict.
 *  Returns only the units worth a human look — not every multi-sale
 *  unit, since most of those are legitimate (multiple commission
 *  tranches for the same buyer). */
function pa_findDuplicateSaleRecords_(sales) {
  const units = pa_groupSalesByUnit_(sales);
  const flagged = [];

  units.forEach(u => {
    if (u.sales.length < 2) return;

    const buyers = new Set(
      u.sales.map(s => String(s.buyerName || '').trim().toUpperCase())
    );

    const seenSignatures = {};
    let exactDuplicate = false;
    u.sales.forEach(s => {
      const sig = pa_saleDuplicateSignature_(s);
      if (seenSignatures[sig]) exactDuplicate = true;
      seenSignatures[sig] = true;
    });

    const buyerConflict = buyers.size > 1;

    if (exactDuplicate || buyerConflict) {
      flagged.push({
        unit: u.unit,
        subdivision: u.subdivision,
        reason: exactDuplicate
          ? (buyerConflict ? 'exactDuplicate+buyerConflict' : 'exactDuplicate')
          : 'buyerConflict',
        sales: u.sales.map(s => ({
          saleId: s.saleId,
          buyerName: s.buyerName,
          reservationDate: s.reservationDate,
          agentGrossComm: s.agentGrossComm,
          agentNetComm: s.agentNetComm,
          releasedCommission: s.releasedCommission,
          receivables: s.receivables
        }))
      });
    }
  });

  return flagged;
}

/** Company-wide duplicate/conflict scan for the Admin panel. */
function pa_adminFindDuplicateSaleRecords_ALL_() {
  const sales = pa_getSalesForAliases_(null); // null = company-wide, all agents
  return pa_findDuplicateSaleRecords_(sales);
}

/** Dashboard summary: lifetime commission, released, receivables, release %, sales count. */
function pa_getDashboard(sessionId) {
  const auth = pa_authorize_(sessionId);
  if (!auth.ok) return auth;

  const sales = pa_getSalesForAliases_(pa_isCompanyWideRole_(auth.role) ? null : auth.aliases);
  let lifetime = 0, released = 0, receivables = 0;
  let totalSalesAmount = 0, totalGrossComm = 0;
  sales.forEach(s => {
    lifetime += s.agentNetComm;
    released += s.releasedCommission;
    receivables += s.receivables;
    totalSalesAmount += s.computedTSP;
    totalGrossComm += s.agentGrossComm;
  });
  const releasePct = lifetime > 0 ? (released / lifetime) * 100 : 0;

  // "Total Tax" is ground truth: the real Gross -> Net gap, so
  // Net Commission = Gross - Total Tax always holds exactly.
  const totalTax = totalGrossComm - lifetime;

  // VAT / Withholding are shown as a *breakdown of totalTax*, not a second,
  // independent read of the sheet — this guarantees VAT + Withholding ==
  // Total Tax on every load, without touching any sheet data.
  //
  // Why not just sum the "VAT Share (%)" / "Withholding Tax (%)" columns
  // directly (as before)? Because on this data, ~68% of rows have those
  // columns populated with a computed reference value that was never
  // actually subtracted from Agent Net Comm — so a direct sum overstates
  // VAT/Withholding versus what was really withheld (see the mismatch
  // analysis: summed columns ≈ ₱6.09M vs. real Gross-Net gap ≈ ₱3.43M
  // across the full sheet).
  //
  // Instead: for each row, take the REAL amount withheld (its own signed
  // Gross - Net gap — a handful of rows have Net > Gross, e.g. manual
  // adjustments, and must stay signed rather than floored at 0, or the
  // totals below stop reconciling exactly) and split it using that row's
  // own recorded VAT:Withholding ratio. Rows with no recorded split at all
  // borrow the portfolio-wide ratio computed from the rows that do have
  // one, so every peso of real tax still lands in VAT or Withholding —
  // none of it is dropped, and none of it is invented.
  let knownVat = 0, knownWht = 0, knownTax = 0, unratedTax = 0;
  sales.forEach(s => {
    const actualTax = s.agentGrossComm - s.agentNetComm;
    const breakdownSum = s.vatShareAgent + s.withholdingTaxAgent;
    if (breakdownSum > 0) {
      knownVat += actualTax * (s.vatShareAgent / breakdownSum);
      knownWht += actualTax * (s.withholdingTaxAgent / breakdownSum);
      knownTax += actualTax;
    } else {
      unratedTax += actualTax;
    }
  });
  const fallbackVatRatio = knownTax !== 0 ? (knownVat / knownTax) : 0.5;
  const totalVat = knownVat + unratedTax * fallbackVatRatio;
  const totalWithholding = knownWht + unratedTax * (1 - fallbackVatRatio);

  const recent = sales
    .slice()
    .sort((a, b) => new Date(b.reservationDate) - new Date(a.reservationDate))
    .slice(0, 5)
    .map(s => ({
      saleId: s.saleId, project: s.project, developer: s.developer,
      reservationDate: s.reservationDate, agentNetComm: s.agentNetComm,
      status: pa_statusBucket_(s)
    }));

  return {
    ok: true,
    // Inactive (resigned) agents still get real numbers here — the data
    // itself isn't restricted, only the UI it's rendered into (PDF-only,
    // no report-card/recent-activity view). See pa_requireActiveAgent_.
    isInactive: !auth.active,
    lifetimeCommission: lifetime,
    releasedCommission: released,
    outstandingReceivables: receivables,
    releasePercentage: releasePct,
    // "Sales Count" means "number of units sold" — the same canonical
    // grouping (pa_groupSalesByUnit_) that Commission by Unit uses, so
    // this can never disagree with that report's unit total, for any
    // account. A unit with multiple commission tranches recorded as
    // separate rows still counts once here, same as there.
    salesCount: pa_groupSalesByUnit_(sales).length,
    // Raw number of underlying sale/commission-tranche rows, kept
    // separately for anywhere that specifically needs the transaction
    // count rather than the unit count.
    saleRecordCount: sales.length,
    totalSalesAmount: totalSalesAmount,
    totalGrossComm: totalGrossComm,
    totalVat: totalVat,
    totalWithholding: totalWithholding,
    totalTax: totalTax,
    recentSales: recent
  };
}

/** Sales history, grouped by year, with optional filters.
 *  filters.year:  '' (all years, from 2020 through the latest year) or a specific year, e.g. '2024'
 *  filters.month: '' (all months) or a month number as a string, '1'..'12'
 */

/** Same sales list as dashboard — use for Report Card / commission by unit.
 *  units is the canonical per-unit grouping (see pa_groupSalesByUnit_) —
 *  the client should render Commission by Unit directly from this array
 *  rather than re-grouping res.sales itself, so there is only ever one
 *  implementation of "what counts as a unit" for any account. */
function pa_getMySalesForReport(sessionId) {
  const auth = pa_authorize_(sessionId);
  if (!auth.ok) return auth;

  const sales = pa_getSalesForAliases_(
    pa_isCompanyWideRole_(auth.role) ? null : auth.aliases
  );

  const units = pa_groupSalesByUnit_(sales);

  return {
    ok: true,
    // Same definition as pa_getDashboard's salesCount: number of units,
    // not number of sale rows — guaranteed equal to units.length below
    // since both come from pa_groupSalesByUnit_.
    salesCount: units.length,
    saleRecordCount: sales.length,
    units: units,
    sales: sales
  };
}

function pa_getSalesHistory(sessionId, filters) {
  const auth = pa_authorize_(sessionId);
  if (!auth.ok) return auth;
  const inactiveGate = pa_requireActiveAgent_(auth);
  if (inactiveGate) return inactiveGate;
  filters = filters || {};

  const allSales = pa_getSalesForAliases_(pa_isCompanyWideRole_(auth.role) ? null : auth.aliases);
  let sales = allSales;

  if (filters.project) sales = sales.filter(s => s.project === filters.project);
  if (filters.developer) sales = sales.filter(s => s.developer === filters.developer);
  if (filters.status) sales = sales.filter(s => pa_statusBucket_(s) === filters.status);
  if (filters.year) sales = sales.filter(s => String(s.year) === String(filters.year));
  if (filters.month) {
    const wantMonth = parseInt(filters.month, 10); // 1-12
    sales = sales.filter(s => {
      if (!s.reservationDate) return false;
      const d = new Date(s.reservationDate);
      return !isNaN(d.getTime()) && (d.getMonth() + 1) === wantMonth;
    });
  }

  const grouped = {};
  sales.forEach(s => {
    if (!grouped[s.year]) grouped[s.year] = [];
    grouped[s.year].push(s);
  });

  const years = Object.keys(grouped).map(Number).sort((a, b) => b - a);
  const groups = years.map(y => ({
    year: y,
    sales: grouped[y].sort((a, b) => new Date(b.reservationDate) - new Date(a.reservationDate))
  }));

  const projects = Array.from(new Set(allSales.map(s => s.project).filter(Boolean))).sort();
  const developers = Array.from(new Set(allSales.map(s => s.developer).filter(Boolean))).sort();

  return {
    ok: true,
    groups: groups,
    projects: projects,
    developers: developers,
    // Full 2020 → latest-year range (see pa_getAllSalesYears_) rather than
    // just the years the agent happens to have sales in, so the dropdown
    // always offers every year and keeps extending itself over time.
    availableYears: pa_getAllSalesYears_(),
    total: sales.length
  };
}

/** Full detail for a single sale, scoped to the logged-in agent's own records only. */
function pa_getSaleDetail(sessionId, saleId) {
  const auth = pa_authorize_(sessionId);
  if (!auth.ok) return auth;
  const inactiveGate = pa_requireActiveAgent_(auth);
  if (inactiveGate) return inactiveGate;

  const sales = pa_getSalesForAliases_(pa_isCompanyWideRole_(auth.role) ? null : auth.aliases);
  const sale = sales.find(s => s.saleId === saleId);
  if (!sale) return { ok: false, error: "Sale not found or not accessible." };

  return { ok: true, sale: sale, statusBucket: pa_statusBucket_(sale) };
}

/** Sales Team directory for the dashboard: agent names (+ status) pulled
 *  straight from the "Agents" sheet via getAgents() (defined in Code.gs —
 *  Apps Script shares one global scope across .gs files, so it's directly
 *  callable here). Any authenticated portal user can see the roster; this
 *  is names/status only, never commission figures, so it doesn't leak
 *  anything an agent's own dashboard shouldn't show. */
function pa_getAgentDirectory(sessionId) {
  const auth = pa_authorize_(sessionId);
  if (!auth.ok) return auth;

  try {
    const agents = getAgents()
      .map(a => ({
        name: String(a['Full Name'] || '').trim(),
        status: String(a['Status'] || '').trim()
      }))
      .filter(a => a.name)
      .sort((a, b) => a.name.localeCompare(b.name));

    return { ok: true, agents: agents, count: agents.length };
  } catch (error) {
    return { ok: false, error: 'Could not load the agent directory.' };
  }
}

/** Profile info for the logged-in agent. Every field here comes from the
 *  trusted server-side session (pa_authorize_) or the "Credentials" sheet
 *  row for that session's email — never from anything the client sends,
 *  so a tampered client request can't make this return someone else's
 *  data or a fabricated role. */
function pa_getProfile(sessionId) {
  const auth = pa_authorize_(sessionId);
  if (!auth.ok) return auth;

  // "Credentials" sheet (Auth.gs, SHEET_CREDENTIALS) columns, 0-indexed
  // as getValues() returns them: [0]=Email,[1]=Password Hash,
  // [2]=Password Set,[3]=Last Login,[4]=Created Date. (Fixed: this used
  // to read a 9-column layout — [1] as email, [6]/[7]/[8] for the rest —
  // that doesn't match this sheet, which is why Last Login/Member Since
  // never actually populated and passwordSet was always false.)
  const credSh = pa_credentialsSheet_();
  const rows = pa_sheetRows_(credSh);
  let passwordSet = false, lastLogin = '', createdDate = '';
  rows.forEach(r => {
    if (pa_normEmail_(r[0]) === auth.email) {
      passwordSet = String(r[2]).toUpperCase() === 'TRUE';
      lastLogin = r[3] instanceof Date ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a") : (r[3] || '');
      createdDate = r[4] instanceof Date ? Utilities.formatDate(r[4], Session.getScriptTimeZone(), "MM/dd/yyyy") : (r[4] || '');
    }
  });

  return {
    ok: true,
    email: auth.email,
    role: auth.role,
    aliases: auth.aliases,
    active: auth.active, // Agent Access "Active" toggle — already surfaced elsewhere (CURRENT_IS_ACTIVE), safe to show here too
    passwordSet: passwordSet,
    lastLogin: lastLogin,
    createdDate: createdDate
  };
}

// ══════════════════════════════════════════════════════════
// DIAGNOSTICS — agent-name matching ("pivoting agent names")
//
// pa_getSalesForAliases_() only keeps a Sales row if its "Sales Agent"
// text (column B of the <Year> Sales sheets) exactly matches one of an
// agent's Agent Access aliases (case/whitespace-insensitive, via
// pa_norm_). If a name in the Sales Agent column doesn't match ANY
// alias on file — a typo, a maiden/married name change, a nickname used
// one year and a full name the next, extra whitespace introduced during
// the CMS migration, etc. — those rows just silently vanish from every
// agent's portal view. No error, no warning; the row is simply skipped.
//
// This is the automated version of the "pivot column B and eyeball it"
// check: it pivots every "Sales Agent" value across every merged Sales
// sheet (both CMS + TARGET_SHEET_ID, current + historical years — see
// pa_getYearSalesSheets_ above), normalizes it the exact same way
// pa_getSalesForAliases_ does, and reports every name that has NO
// matching Agent Access alias — together with how many rows and how
// much Agent Net Comm are affected, so you can prioritize which
// mismatches actually matter.
// ══════════════════════════════════════════════════════════

/** Pivots the "Sales Agent" column across every merged Sales sheet.
 *  Returns a Map<normalizedName, {displayName, rowCount, totalAgentNetComm, years:Set}>. */
function pa_pivotSalesAgentNames_() {
  const yearSheets = pa_getYearSalesSheets_();
  const pivot = new Map();

  yearSheets.forEach((sheets, year) => {
    // Same merge-by-Sale-ID fix as pa_getSalesForAliases_: a year can now
    // resolve to more than one sheet (CMS + TARGET_SHEET_ID both having a
    // "2026 Sales" tab, for instance), so dedupe across them by Sale ID
    // rather than double-counting/ overcounting or reading only one.
    const bySaleId = new Map();
    sheets.forEach(sheet => {
      if (sheet.getLastRow() < 2) return;
      const numCols = Math.min(31, sheet.getLastColumn());
      const headerRow = sheet.getRange(1, 1, 1, numCols).getValues()[0];
      const dataRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, numCols).getValues();

      const colIdx = {};
      headerRow.forEach((h, i) => {
        colIdx[String(h || '').trim().replace(/\*+$/, '')] = i;
      });
      if (!('Sales Agent' in colIdx)) return; // malformed/unexpected sheet — skip rather than throw

      dataRows.forEach((row, rowIdx) => {
        let saleId = row[colIdx['Sale ID']];
        if (!saleId) {
          // mirror pa_getSalesForAliases_: blank Sale ID rows are kept under a synthetic key
          if (!String(row[colIdx['Sales Agent']] || '').trim()) return;
          saleId = 'ROW-' + year + '-' + sheet.getSheetId() + '-' + (rowIdx + 2);
        }
        bySaleId.set(saleId, { row: row, colIdx: colIdx });
      });
    });

    bySaleId.forEach(({ row, colIdx }) => {
      const raw = row[colIdx['Sales Agent']];
      const norm = pa_aliasKey_(raw);
      if (!norm) return; // blank Sales Agent cell — not a matching problem, just an incomplete row

      const agentNetComm = colIdx['Agent Net Comm (₱)'] !== undefined
        ? (parseFloat(row[colIdx['Agent Net Comm (₱)']]) || 0) : 0;

      if (!pivot.has(norm)) {
        pivot.set(norm, { displayName: String(raw).trim(), rowCount: 0, totalAgentNetComm: 0, years: new Set() });
      }
      const entry = pivot.get(norm);
      entry.rowCount++;
      entry.totalAgentNetComm += agentNetComm;
      entry.years.add(year);
    });
  });

  return pivot;
}

/** Every normalized alias currently on file in Agent Access, across every user. */
function pa_allKnownAliases_() {
  const known = new Set();
  pa_adminListUsers_().forEach(u => {
    // pa_parseAliasesCsv_ (Auth.gs) splits on ';', not ',' — agent names
    // like "FLORES, JULIE FE" contain a literal comma, so comma-splitting
    // here would silently fragment them the same way it did in
    // pa_getAccessRecord_, making the diagnostic falsely report a
    // "mismatch" for a name that's actually fine.
    // Canonical Name + Aliases together — both are valid Sales Agent spellings.
    pa_buildAccessNames_(u.canonicalName, u.aliasesCsv).aliases.forEach(a => {
      const norm = pa_aliasKey_(a);
      if (norm) known.add(norm);
    });
  });
  // Plus every name in the CMS "Agent Aliases" tab — that is what matching uses.
  pa_getCmsAliasRows_().forEach(r => {
    [r.canonical].concat(r.aliases).forEach(a => { const norm = pa_aliasKey_(a); if (norm) known.add(norm); });
  });
  return known;
}

/** The actual diagnostic: every distinct "Sales Agent" name that appears in
 *  the Sales sheets but matches NO alias on any Agent Access row. Sorted by
 *  totalAgentNetComm descending, so the mismatches with the biggest real
 *  impact (money an agent can't currently see) show up first. */
function pa_findUnmatchedSalesAgentNames_() {
  const pivot = pa_pivotSalesAgentNames_();
  const known = pa_allKnownAliases_();
  const unmatched = [];

  pivot.forEach((entry, norm) => {
    if (!known.has(norm)) {
      unmatched.push({
        salesAgentName: entry.displayName,
        rowCount: entry.rowCount,
        totalAgentNetComm: entry.totalAgentNetComm,
        years: Array.from(entry.years).sort()
      });
    }
  });

  unmatched.sort((a, b) => b.totalAgentNetComm - a.totalAgentNetComm);
  return unmatched;
}

/** Admin-facing entry point — call from the Admin panel (or wire up a
 *  button later) to get the same list pa_findUnmatchedSalesAgentNames_
 *  computes, gated behind a real Admin session. */
function pa_adminFindUnmatchedSalesAgentNames(sessionId) {
  const auth = pa_isAdmin_(sessionId);
  if (!auth.ok) return auth;
  return { ok: true, unmatched: pa_findUnmatchedSalesAgentNames_() };
}

/** DIAGNOSTIC — run this manually from the Apps Script editor
 *  (Select function > pa_debugUnmatchedSalesAgentNames > Run), then check
 *  View > Logs / Executions. Prints every Sales Agent name with no
 *  matching Agent Access alias, worst-impact first. */
function pa_debugUnmatchedSalesAgentNames() {
  const unmatched = pa_findUnmatchedSalesAgentNames_();
  if (!unmatched.length) {
    Logger.log('No unmatched Sales Agent names found — every name in the Sales sheets matches a known Agent Access alias.');
    return;
  }
  Logger.log(unmatched.length + ' Sales Agent name(s) with NO matching Agent Access alias:');
  unmatched.forEach(u => {
    Logger.log(
      '  "' + u.salesAgentName + '" — ' + u.rowCount + ' row(s), ' +
      '\u20b1' + u.totalAgentNetComm.toLocaleString() + ' in Agent Net Comm invisible to any agent, ' +
      'years: ' + u.years.join(', ')
    );
  });
}