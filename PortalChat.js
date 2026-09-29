// ============================================================
// Inner SPARC Realty Corporation — Agent Portal
// PortalChat.gs — Commission Assistant Chatbot (session-scoped)
//
// Ported from the standalone "Commission Monitoring System" chatbot
// (chatWithBot / callGroqForIntent / generate* formatters). The intent
// classification (Groq) and response formatting are functionally the
// same as the original — what's added here is a data-scoping layer so
// the bot can never show anyone commission data pa_getDashboard() /
// pa_getSalesHistory() wouldn't already show them:
//
//   - Every call re-validates the session via pa_authorize_ (PortalData.gs)
//     — no valid session, no chatbot, same as every other pa_* function.
//   - Admin / Supervisor / Manager (pa_isCompanyWideRole_, PortalData.gs)
//     get the full company-wide bot: every intent, every agent/developer,
//     same as the original standalone version.
//   - Agent gets a bot whose underlying data is ALWAYS pre-filtered to its
//     own Agent Access aliases (via the existing pa_getSalesForAliases_),
//     and whose "about a specific person" intents (get_agent_summary,
//     get_pending_commissions with a named agent) are only honored when
//     the resolved name is the caller's own — otherwise it replies with a
//     clear "you can only ask about your own data" message rather than
//     silently redirecting or quietly answering about someone else.
//     Company-wide-only views (compare_agents, get_top_performers,
//     get_trends, get_overall_summary, get_data_issues, get_week_comparison,
//     get_recent_activity) are blocked outright for this role, since by
//     definition they surface other people's numbers.
// ============================================================

const PA_CHAT_RESTRICTED_MSG =
  "🔒 I can only answer questions about your own commission record. " +
  "Try asking about your own sales, pending commissions, or a specific developer's rates.";

/** Intents that inherently surface company-wide / other-people's data.
 *  Only ever run for a company-wide role (Admin/Supervisor/Manager). */
const PA_CHAT_COMPANY_WIDE_ONLY = [
  'compare_agents', 'get_top_performers', 'get_trends',
  'get_overall_summary', 'get_data_issues', 'get_week_comparison',
  'get_recent_activity'
];

/**
 * Main entry point, called from the client as
 *   google.script.run...pa_chatWithBot(sessionId, query, year, history)
 * Returns { ok:true, reply } or { ok:false, error }.
 */
function pa_chatWithBot(sessionId, query, year, history) {
  const auth = pa_authorize_(sessionId); // PortalData.gs — same auth boundary as every other portal call
  if (!auth.ok) return { ok: false, error: auth.error };

  // Inactive (resigned) agents are limited to the PDF-only commission view
  // — the chatbot is part of the active-agent portal experience, same as
  // Sales History, so it's gated the same way (pa_requireActiveAgent_).
  const inactiveGate = pa_requireActiveAgent_(auth); // PortalData.gs
  if (inactiveGate) {
    return { ok: true, reply: "🔒 The assistant isn't available on an inactive account. You can still view and download your commission record as a PDF from your dashboard." };
  }

  try {
    const companyWide = pa_isCompanyWideRole_(auth.role); // PortalData.gs
    const normalizedQuery = String(query || '').trim();
    const dropdownYear = year || String(new Date().getFullYear());

    // Full name rosters are used only for name-recognition (Groq / local /
    // fuzzy matching) — never as a bypass around scoping. Any authenticated
    // portal user already sees this same agent-name list via
    // pa_getAgentDirectory (PortalData.gs), so exposing just the names here
    // isn't a new leak; the underlying numbers are what stay gated below.
    const agents = getAgents();       // Code.gs
    const developers = getDevelopers(); // Code.gs
    const agentNames = agents.map(a => a['Full Name']).filter(Boolean);
    const developerNames = developers.map(d => d['Developer Name']).filter(Boolean);

    if (mentionsMilestone(normalizedQuery)) {
      return { ok: true, reply: generateMilestoneRedirect() };
    }

    const groqResult = callGroqForIntent(normalizedQuery, dropdownYear, agentNames, developerNames, history);

    let toolName, args;
    if (groqResult.functionCall) {
      toolName = groqResult.functionCall.name;
      args = groqResult.functionCall.args;
    } else {
      const localMatch = findKnownNameInQuery(normalizedQuery, agentNames, developerNames);
      const fuzzyMatch = localMatch ? null : fuzzyFindName(normalizedQuery, agentNames, developerNames);
      if (localMatch || fuzzyMatch) {
        const match = localMatch || fuzzyMatch;
        toolName = match.type === 'agent' ? 'get_agent_summary' : 'get_developer_summary';
        args = match.type === 'agent' ? { agent_name: match.name } : { developer_name: match.name };
      } else {
        return {
          ok: true,
          reply: groqResult.text || generateGenericResponse(pa_chatBuildContext_(auth, dropdownYear, agents, developers))
        };
      }
    }

    const resolvedYear = (args && args.year && /^20\d{2}$/.test(String(args.year))) ? String(args.year) : dropdownYear;

    // ── Security gate #1: company-wide-only views ──
    if (!companyWide && PA_CHAT_COMPANY_WIDE_ONLY.indexOf(toolName) !== -1) {
      return { ok: true, reply: PA_CHAT_RESTRICTED_MSG };
    }

    // ── Security gate #2: "about a specific person" — only about yourself ──
    if (!companyWide && (toolName === 'get_agent_summary' || toolName === 'get_pending_commissions')) {
      const askedAgentName = args && args.agent_name;
      if (askedAgentName) {
        const matched = findMatchingAgent(askedAgentName, agents);
        const isSelf = matched && auth.aliases.some(a => pa_norm_(a) === pa_norm_(matched['Full Name']));
        if (!isSelf) return { ok: true, reply: PA_CHAT_RESTRICTED_MSG };
      }
    }

    const context = pa_chatBuildContext_(auth, resolvedYear, agents, developers);
    let response = '';

    switch (toolName) {
      case 'compare_agents': { // company-wide only (gated above)
        const agent1 = findMatchingAgent(args.agent1_name, agents);
        const agent2 = findMatchingAgent(args.agent2_name, agents);
        if (!agent1 || !agent2) {
          response = `I couldn't find ${!agent1 ? `"${args.agent1_name}"` : `"${args.agent2_name}"`} in the agent list. Could you check the spelling?`;
        } else {
          context.agent1 = agent1;
          context.agent2 = agent2;
          response = generateComparison(context);
        }
        break;
      }
      case 'get_agent_summary': {
        const agent = findMatchingAgent(args.agent_name, agents);
        if (!agent) {
          response = `I couldn't find an agent named "${args.agent_name}". Could you check the spelling?`;
        } else {
          context.agentName = agent;
          response = generateAgentSummary(context);
        }
        break;
      }
      case 'get_developer_summary': {
        const dev = findMatchingDeveloper(args.developer_name, developers);
        if (!dev) {
          response = `I couldn't find a developer named "${args.developer_name}". Could you check the spelling?`;
        } else {
          context.developerName = dev;
          response = generateDeveloperSummary(context); // context.commissions is already scoped for non-company-wide roles
        }
        break;
      }
      case 'get_top_performers': response = generateTopPerformers(context); break; // company-wide only (gated above)
      case 'get_trends': response = generateTrendsSummary(context); break;         // company-wide only (gated above)
      case 'get_overall_summary': response = generateOverallSummary(context); break; // company-wide only (gated above)
      case 'show_help': response = generateHelpResponse(); break;
      case 'get_data_issues': response = generateDataIssuesReport(resolvedYear); break; // company-wide only (gated above) — unmodified, direct sheet read is fine since only reachable when already company-wide
      case 'get_pending_commissions': {
        let focusAgent = args && args.agent_name;
        let focusDeveloper = args && args.developer_name;

        if (!companyWide) {
          // Scoped role: context.commissions is already limited to this
          // person's own sales, so there's nothing left to "focus" onto by
          // name — force company-wide-style rendering off and just show
          // their own totals. (Gate #2 above already rejected any explicit
          // *other* agent's name before we got here.)
          focusAgent = null;
        } else if (!focusAgent && !focusDeveloper) {
          const mentionedThisTurn =
            findKnownNameInQuery(normalizedQuery, agentNames, developerNames) ||
            fuzzyFindName(normalizedQuery, agentNames, developerNames);
          if (!mentionedThisTurn) {
            const activeEntity = getActiveEntityFromHistory(history);
            if (activeEntity) {
              if (activeEntity.type === 'agent') focusAgent = activeEntity.name;
              else focusDeveloper = activeEntity.name;
            }
          }
        }
        context.focusAgent = focusAgent || null;
        context.focusDeveloper = focusDeveloper || null;
        response = generatePendingCommissions(context);
        break;
      }
      case 'get_developer_rate': { // rate card only, not sales data — safe for every role
        const dev = findMatchingDeveloper(args.developer_name, developers);
        if (!dev) {
          response = `I couldn't find a developer named "${args.developer_name}". Could you check the spelling?`;
        } else {
          response = generateDeveloperRateResponse(dev);
        }
        break;
      }
      case 'get_week_comparison': response = generateWeekComparison(resolvedYear); break; // company-wide only (gated above)
      case 'get_recent_activity': response = generateRecentActivity(resolvedYear); break; // company-wide only (gated above)
      default: response = generateGenericResponse(context);
    }

    return { ok: true, reply: response };
  } catch (error) {
    Logger.log(`[pa_chatWithBot] ERROR: ${error.message}`);
    return { ok: false, error: `The assistant hit an error: ${error.message}. Please try again.` };
  }
}

/**
 * Builds the same shape of context object the ported generate* formatters
 * expect, but sourced from pa_getSalesForAliases_ (PortalData.gs) instead of
 * the standalone app's getCommissions() — which is what actually enforces
 * scoping. null aliases (company-wide) reads every sale; an alias list
 * reads only that person's own sales.
 */
function pa_chatBuildContext_(auth, year, agents, developers) {
  const companyWide = pa_isCompanyWideRole_(auth.role);
  const scopedSales = pa_getSalesForAliases_(companyWide ? null : auth.aliases); // every year
  const yearSales = scopedSales.filter(s => String(s.year) === String(year));
  const commissions = (yearSales.length ? yearSales : scopedSales).map(pa_chatAdaptSale_);
  const stats = pa_chatBuildStats_(commissions, agents, developers);

  return { year: year, agents: agents, developers: developers, commissions: commissions, stats: stats };
}

/** Adapts a pa_getSalesForAliases_ sale object (camelCase) into the exact
 *  header-keyed shape the ported generate* functions expect (matching
 *  the standalone app's getCommissions() output), so those functions can
 *  be reused unmodified. */
function pa_chatAdaptSale_(s) {
  return {
    'Sale ID': s.saleId || '',
    'Reservation Date': s.reservationDate || '',
    'Sales Agent': s.salesAgent || '',
    'Developer': s.developer || '',
    'Project': s.project || '',
    'Unit ID': s.unitId || '',
    "Buyer's Name": s.buyerName || '',
    'DP Term': s.dpTerm || '',
    'Computed TSP': s.computedTSP || 0,
    'Developer Comm Rate (%)': s.developerCommRate || 0,
    'Company Gross Comm (₱)': s.companyGrossComm || 0,
    'Net VAT (₱)': s.netVat || 0,
    'Withholding Tax (Company) (₱)': s.companyWithholdingTax || 0,
    'Company Net Comm (₱)': s.companyNetComm || 0,
    'Agent Comm Rate (%)': s.agentCommRate || 0,
    'Agent Gross Comm (₱)': s.agentGrossComm || 0,
    'VAT Share (Agent) (₱)': s.vatShareAgent || 0,
    'Withholding Tax (Agent) (₱)': s.withholdingTaxAgent || 0,
    'Agent Net Comm (₱)': s.agentNetComm || 0,
    'Milestone Ratio': s.milestoneRatio || '--',
    'DP Tranche': s.dpTranche || '--',
    'Loan Takeout Tranche': s.loanTakeoutTranche || '--',
    'Released Commission': s.releasedCommission || 0,
    'Receivables': s.receivables || 0,
    'Commission Status': s.commissionStatus || '0%',
    'DP Release Amounts': s.dpReleaseAmounts || [],
    'Loan Takeout Amount': s.loanTakeoutAmount || 0,
    'Cash Advance': s.cashAdvance || 0
  };
}

/** Port of getDashboardStats()'s aggregation, but taking an already-scoped
 *  `commissions` array directly instead of re-reading the sheet by year —
 *  this is what keeps generateOverallSummary/generateGenericResponse's
 *  "quick stats" scoped to the caller for non-company-wide roles. */
function pa_chatBuildStats_(commissions, agents, devs) {
  if (!commissions || commissions.length === 0) {
    return { totalSales: 0, totalCompanyGross: 0, totalTax: 0, totalNetCommission: 0, activeAgents: 0, totalDevelopers: 0, topAgents: [], recentSales: [] };
  }

  let totalSales = 0, totalCompanyGross = 0, totalTax = 0, totalNetCommission = 0;
  const agentMap = {};
  const recentSalesList = [];

  commissions.forEach(s => {
    totalSales++;
    const gross = parseFloat(s['Company Gross Comm (₱)']) || 0;
    const netVat = parseFloat(s['Net VAT (₱)']) || 0;
    const withholding = parseFloat(s['Withholding Tax (Company) (₱)']) || 0;
    const companyNet = parseFloat(s['Company Net Comm (₱)']) || 0;
    const agentNet = parseFloat(s['Agent Net Comm (₱)']) || 0;

    totalCompanyGross += gross;
    totalTax += (gross - netVat) + withholding;
    totalNetCommission += companyNet;

    const agentName = s['Sales Agent'];
    if (agentName) {
      if (!agentMap[agentName]) agentMap[agentName] = { name: agentName, net: 0, count: 0 };
      agentMap[agentName].net += agentNet;
      agentMap[agentName].count += 1;
    }

    if (s['Reservation Date']) {
      recentSalesList.push({ 'Sales Agent': s['Sales Agent'] || 'Unknown', 'Project': s['Project'] || s['Developer'] || 'Unknown', 'Agent Net Comm': agentNet, 'Reservation Date': s['Reservation Date'] });
    }
  });

  const topAgents = Object.values(agentMap).sort((a, b) => b.net - a.net).slice(0, 5);
  const recentSales = recentSalesList
    .sort((a, b) => new Date(b['Reservation Date']) - new Date(a['Reservation Date']))
    .slice(0, 5)
    .map(s => ({ 'Sales Agent': s['Sales Agent'], 'Project': s['Project'], 'Agent Net Comm': s['Agent Net Comm'] }));

  const activeAgents = agents.filter(a => a['Status'] === 'Active').length;

  return { totalSales, totalCompanyGross, totalTax, totalNetCommission, activeAgents, totalDevelopers: devs.length, topAgents, recentSales };
}