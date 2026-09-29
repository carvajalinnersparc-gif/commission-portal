// ============================================================
// Inner SPARC Realty Corporation — Agent Portal
// ChatEngine.gs — Chatbot intent classification + response formatters
//
// Ported verbatim (unmodified) from the standalone "Commission Monitoring
// System" chatbot's Code.gs. These are pure helpers/formatters — name
// matching, the Groq function-calling call, and the generate* response
// builders — that only ever operate on whatever `commissions`/`agents`/
// `developers` array is handed to them via `context`.
//
// The actual security boundary (session auth + role-based data scoping)
// lives in PortalChat.gs's pa_chatWithBot()/pa_chatBuildContext_(), which
// is what decides WHAT data these functions get to see, and which
// intents (compare_agents, get_top_performers, get_trends,
// get_overall_summary, get_data_issues, get_week_comparison,
// get_recent_activity) are allowed to run at all for a non-company-wide
// role. Nothing in this file should be called directly from the client —
// only pa_chatWithBot (PortalChat.gs) is the entry point.
//
// Setup note: this file calls PropertiesService.getScriptProperties() for
// GROQ_API_KEY / GROQ_API_KEY_1 (see callGroqForIntent below) — add those
// under Project Settings > Script Properties in this project too.
// ============================================================

function getActiveEntityFromHistory(history) {
  if (!history || !history.length) return null;

  let lastBotTurn = null;
  for (let i = history.length - 1; i >= 0; i--) {
    if (history[i] && history[i].role === 'bot' && history[i].content) {
      lastBotTurn = history[i];
      break;
    }
  }
  if (!lastBotTurn) return null;

  const content = String(lastBotTurn.content);

  let m = /Agent Performance Summary:\s*([^\n]+)/.exec(content);
  if (m) return { type: 'agent', name: m[1].trim() };

  m = /Developer Performance Summary:\s*([^\n]+)/.exec(content);
  if (m) return { type: 'developer', name: m[1].trim() };

  m = /🏗️\s*\*\*([^*]+?)\s*—\s*Commission Rates\*\*/.exec(content);
  if (m) return { type: 'developer', name: m[1].trim() };

  return null;
}

/**
 * Typo-tolerant check for whether the query is asking about "milestone(s)".
 * Exact substring match first (fast path), then falls back to a per-word
 * Levenshtein check so small misspellings like "milstone" or "mileston"
 * still route to the milestone redirect instead of falling through to Groq.
 */
function mentionsMilestone(query) {
  const q = (query || '').toLowerCase();
  if (/\bmilestones?\b/.test(q)) return true;

  const words = q.split(/\s+/);
  for (const w of words) {
    if (w.length < 5) continue;
    if (levenshtein(w, 'milestone') <= 2 || levenshtein(w, 'milestones') <= 2) {
      return true;
    }
  }
  return false;
}

function callGroqForIntent(query, dropdownYear, agentNames, developerNames, history) {

  // GROQ_API_KEY_1 is a fallback key — used automatically if the primary key
  // fails with a 400 (e.g. invalid/revoked key on Groq's side).
  const apiKeys = [
    PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY'),
    PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY_1')
  ].filter(Boolean);

  if (apiKeys.length === 0) {
    throw new Error('GROQ_API_KEY is not set in Script Properties. Add it under Project Settings.');
  }

  const model = 'openai/gpt-oss-120b'; // swap to 'openai/gpt-oss-20b' for faster/cheaper
  const url = 'https://api.groq.com/openai/v1/chat/completions';

  const today = new Date();
  const agentListText = (agentNames && agentNames.length) ? agentNames.join(', ') : '(none on file)';
  const developerListText = (developerNames && developerNames.length) ? developerNames.join(', ') : '(none on file)';

  const systemInstruction =
    `You are the intent-detection layer for a real estate commission tracking chatbot. ` +
    `Today's date is ${today.toDateString()}. The currently selected year in the UI is ${dropdownYear}. ` +
    `Valid data years range from 2020 to 2026.\n\n` +
    `KNOWN AGENT NAMES: ${agentListText}\n` +
    `KNOWN DEVELOPER NAMES: ${developerListText}\n\n` +
    `RULES:\n` +
    `1. If the message names or clearly refers to someone (even a nickname, partial name, or surname only) who ` +
    `matches or closely resembles a name from KNOWN AGENT NAMES, you MUST call get_agent_summary (or compare_agents ` +
    `if two agent names are mentioned) — even if the message also contains the word "performance," "summary," or ` +
    `similar. A message about a named individual is NEVER get_overall_summary or get_top_performers, even if it ` +
    `uses the word "performance."\n` +
    `2. Only call get_overall_summary when the message asks about the company/team/business as a whole with NO ` +
    `specific person or developer named (e.g. "how are we doing overall," "company performance").\n` +
    `3. Only call get_developer_summary when a name from KNOWN DEVELOPER NAMES is matched, similarly to rule 1.\n` +
    `4. Extract a year argument ONLY if the user explicitly types one (e.g. "2026"); otherwise omit it entirely — ` +
    `the system will use the UI's selected year.\n` +
    `5. If the message is just a greeting, thanks, or general chit-chat that doesn't match any tool and doesn't ` +
    `reference any known name, do NOT call any tool — instead reply directly with a short, friendly plain-text ` +
    `response.\n` +
    `6. If you are genuinely unsure between two tools, prefer the more specific one (an agent/developer tool over ` +
    `a general one).\n` +
    `7. Use get_data_issues only when the user is asking about data quality/problems/errors, not performance.\n` +
    `8. Use get_pending_commissions when the user asks what's still owed/unreleased/pending, as opposed to totals ` +
    `already paid out.\n` +
    `9. Use get_developer_rate ONLY for a quick rate/percentage lookup on a developer, not a full performance ` +
    `summary — if the user wants performance/transactions/history, use get_developer_summary instead.\n` +
    `10. Use get_week_comparison for week-over-week questions, and get_recent_activity for "who's active lately" ` +
    `questions — these are different from get_trends (which is monthly) and get_top_performers (which is lifetime totals).\n` +
    `11. If the conversation history shows your previous message asked a clarifying question (e.g. "which agent?" ` +
    `or "which developer?"), treat the user's next short reply as the answer to that question, not as a new, ` +
    `unrelated message — match it against KNOWN AGENT NAMES / KNOWN DEVELOPER NAMES even if it's just one word.\n` +
    `12. If your previous reply was a performance summary for a specific agent or developer, and the user's next ` +
    `message is a short, unscoped follow-up (e.g. "pending commissions", "what's still owed", "how about this ` +
    `year") that doesn't name someone else, keep it scoped to that same agent/developer — e.g. pass their name as ` +
    `agent_name/developer_name to get_pending_commissions — rather than answering company-wide.`;

  const messages = [{ role: 'system', content: systemInstruction }];
  // Include recent conversation turns (already trimmed by the caller) so
  // short follow-up replies can be understood in context instead of being
  // treated as standalone, ambiguous messages.
  (history || []).forEach(turn => {
    if (turn && turn.role && turn.content) {
      messages.push({ role: turn.role === 'bot' ? 'assistant' : 'user', content: String(turn.content) });
    }
  });
  messages.push({ role: 'user', content: query });

  const tools = [
    {
      type: 'function',
      function: {
        name: 'get_agent_summary',
        description: 'Get a performance summary for one specific sales agent (transactions, commissions, developers worked with, etc).',
        parameters: {
          type: 'object',
          properties: {
            agent_name: { type: 'string', description: 'The agent\'s name as mentioned by the user' },
            year: { type: 'string', description: 'A 4-digit year, only if explicitly mentioned (e.g. "2026")' }
          },
          required: ['agent_name']
        }
      }
    },
    {
      type: 'function',
      function: {
        name: 'get_developer_summary',
        description: 'Get a performance summary for one specific developer/client (transactions, commissions, agents involved, projects, etc).',
        parameters: {
          type: 'object',
          properties: {
            developer_name: { type: 'string', description: 'The developer\'s name as mentioned by the user' },
            year: { type: 'string', description: 'A 4-digit year, only if explicitly mentioned' }
          },
          required: ['developer_name']
        }
      }
    },
    {
      type: 'function',
      function: {
        name: 'compare_agents',
        description: 'Compare performance between exactly two sales agents.',
        parameters: {
          type: 'object',
          properties: {
            agent1_name: { type: 'string', description: 'First agent\'s name' },
            agent2_name: { type: 'string', description: 'Second agent\'s name' },
            year: { type: 'string', description: 'A 4-digit year, only if explicitly mentioned' }
          },
          required: ['agent1_name', 'agent2_name']
        }
      }
    },
    {
      type: 'function',
      function: {
        name: 'get_top_performers',
        description: 'Get the top-ranked agents and developers by net commission.',
        parameters: {
          type: 'object',
          properties: {
            year: { type: 'string', description: 'A 4-digit year, only if explicitly mentioned' }
          }
        }
      }
    },
    {
      type: 'function',
      function: {
        name: 'get_trends',
        description: 'Get month-by-month commission trends and statistics.',
        parameters: {
          type: 'object',
          properties: {
            year: { type: 'string', description: 'A 4-digit year, only if explicitly mentioned' }
          }
        }
      }
    },
    {
      type: 'function',
      function: {
        name: 'get_overall_summary',
        description: 'Get the overall company-wide performance summary (total sales, commissions, active agents, etc).',
        parameters: {
          type: 'object',
          properties: {
            year: { type: 'string', description: 'A 4-digit year, only if explicitly mentioned' }
          }
        }
      }
    },
    {
      type: 'function',
      function: {
        name: 'show_help',
        description: 'Show the user what kinds of questions this bot can answer.',
        parameters: { type: 'object', properties: {} }
      }
    },
    {
      type: 'function',
      function: {
        name: 'get_data_issues',
        description: 'Scan for data quality problems: sales missing a Sale ID, missing/zero commission rates, and outstanding cash advances that agents still owe against future releases. Use for phrases like "data issues", "anything wrong", "check for errors", "outstanding cash advances".',
        parameters: {
          type: 'object',
          properties: {
            year: { type: 'string', description: 'A 4-digit year, only if explicitly mentioned' }
          }
        }
      }
    },
    {
      type: 'function',
      function: {
        name: 'get_pending_commissions',
        description: 'Get a breakdown of commission still owed/unreleased to agents (pending vs. already released), company-wide or scoped to one agent/developer. Use for phrases like "how much is still unpaid", "pending commissions", "receivables", "what\'s still owed".',
        parameters: {
          type: 'object',
          properties: {
            agent_name: { type: 'string', description: 'An agent\'s name, ONLY if explicitly mentioned this turn or clearly the subject of the ongoing conversation' },
            developer_name: { type: 'string', description: 'A developer\'s name, ONLY if explicitly mentioned this turn or clearly the subject of the ongoing conversation' },
            year: { type: 'string', description: 'A 4-digit year, only if explicitly mentioned' }
          }
        }
      }
    },
    {
      type: 'function',
      function: {
        name: 'get_developer_rate',
        description: 'Quick single-fact lookup of a developer\'s commission rates (developer comm rate, agent comm rate, VAT, withholding tax) without a full performance summary. Use only when the user asks specifically for a rate/percentage, e.g. "what\'s the comm rate for X", "what percentage does X get".',
        parameters: {
          type: 'object',
          properties: {
            developer_name: { type: 'string', description: 'The developer\'s name as mentioned by the user' }
          },
          required: ['developer_name']
        }
      }
    },
    {
      type: 'function',
      function: {
        name: 'get_week_comparison',
        description: 'Compare this week\'s sales/commission activity to last week. Use for phrases like "how are we doing this week", "this week vs last week", "week over week".',
        parameters: { type: 'object', properties: {} }
      }
    },
    {
      type: 'function',
      function: {
        name: 'get_recent_activity',
        description: 'Get the most active agents by transaction count in the last 30 days (momentum/velocity, not lifetime totals). Use for phrases like "who\'s hot right now", "most active agents lately", "recent activity", "who\'s closing deals this month".',
        parameters: {
          type: 'object',
          properties: {
            year: { type: 'string', description: 'A 4-digit year, only if explicitly mentioned' }
          }
        }
      }
    }
  ];

  const payload = {
    model: model,
    messages: messages,
    tools: tools,
    tool_choice: 'auto',
    reasoning_effort: 'high' // improves name/entity matching accuracy at a small latency cost
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let response, statusCode, responseText;
  for (let i = 0; i < apiKeys.length; i++) {
    const attemptOptions = Object.assign({}, options, {
      headers: { Authorization: `Bearer ${apiKeys[i]}` }
    });
    response = UrlFetchApp.fetch(url, attemptOptions);
    statusCode = response.getResponseCode();
    responseText = response.getContentText();

    if (statusCode === 200) break;

    // Only fall through to the next key on a 400 — that's the signature of a
    // bad/invalid/revoked key on Groq's side. Other errors (429 rate limit,
    // 5xx server errors) aren't fixed by swapping keys, so don't waste a
    // second call retrying those with the fallback key.
    if (statusCode === 400 && i < apiKeys.length - 1) {
      Logger.log(`[Groq] Key #${i + 1} returned 400, retrying with fallback key (GROQ_API_KEY_1)`);
      continue;
    }
    break;
  }

  if (statusCode !== 200) {
    Logger.log(`[Groq] ERROR ${statusCode}: ${responseText}`);
    throw new Error(`Groq API error (${statusCode}). Please try again in a moment.`);
  }

  const data = JSON.parse(responseText);
  const message = (data && data.choices && data.choices[0] && data.choices[0].message) || {};
  const toolCalls = message.tool_calls || [];

  if (toolCalls.length > 0) {
    const call = toolCalls[0]; // only ever expect/handle one tool call here
    let args = {};
    try {
      args = call.function && call.function.arguments ? JSON.parse(call.function.arguments) : {};
    } catch (e) {
      Logger.log(`[Groq] Failed to parse tool args: ${call.function && call.function.arguments}`);
    }
    return { functionCall: { name: call.function.name, args: args } };
  }

  // No tool call — treat message content as a direct answer
  const text = (message.content || '').trim();
  return { text: text || "I'm not sure how to help with that. Try asking about an agent, a developer, top performers, or trends!" };
}

// ── MILESTONE REDIRECT ──
/**
 * The chatbot has no data-backed way to answer "milestone" questions — DP/loan
 * release tranches are configured per-sale on the Milestone tab, not part of
 * the commissions data this bot reads. Answering immediately (instead of
 * asking an open-ended clarifying question) avoids a dead-end loop, since no
 * follow-up reply could actually resolve it.
 */
function generateMilestoneRedirect() {
  return `📍 **About Milestones**\n\n` +
    `Milestones here refer to a sale's DP/Loan Takeout release schedule (ratio, tranches, and release amounts) — that's configured per-sale, not something I can look up by agent or developer name.\n\n` +
    `👉 Head to the **Milestone** tab in the sidebar to view or edit a specific sale's release schedule.\n\n` +
    `💬 I *can* help with: an agent's or developer's commission summary, pending/unreleased commissions, or top performers. Just ask!`;
}

// ── FUZZY NAME MATCHING ──
/**
 * Last-resort match for typo'd or partial names (e.g. "violenta" for
 * "Violeta Reyes") that fail both Groq's classification and the exact
 * substring fallback. Compares the query against every known agent/developer
 * name using normalized Levenshtein distance and returns the best match only
 * if it's confident enough to avoid false positives on short/generic queries.
 */
function levenshtein(a, b) {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const dp = Array(n + 1);
  for (let j = 0; j <= n; j++) dp[j] = j;
  for (let i = 1; i <= m; i++) {
    let prev = dp[0];
    dp[0] = i;
    for (let j = 1; j <= n; j++) {
      const temp = dp[j];
      dp[j] = a[i - 1] === b[j - 1]
        ? prev
        : 1 + Math.min(prev, dp[j], dp[j - 1]);
      prev = temp;
    }
  }
  return dp[n];
}

// Generic chatbot vocabulary that should never itself count as a "name" match,
// even though it may literally appear as a word inside a real name (e.g. the
// roster placeholder "Sample Agent" contains the word "agent"). Without this,
// a query like "agent" would fuzzy-match "Sample Agent" at score 1.0 purely
// because "agent" is one of its words — a false positive, not a typo'd name.
const GENERIC_QUERY_WORDS = new Set([
  'agent', 'agents', 'developer', 'developers', 'sales', 'sample',
  'name', 'names', 'commission', 'commissions', 'summary', 'summaries',
  'overall', 'top', 'performer', 'performers', 'trend', 'trends',
  'help', 'pending', 'financial', 'monthly'
]);

function fuzzyFindName(query, agentNames, developerNames) {
  const q = (query || '').toLowerCase().trim();
  // Skip fuzzy matching on very short or multi-word queries — too high a
  // false-positive risk (e.g. "help", "show me trends").
  if (q.length < 4 || q.split(/\s+/).length > 4) return null;
  // Skip generic chatbot vocabulary outright — these aren't typo'd names.
  if (GENERIC_QUERY_WORDS.has(q)) return null;

  let best = null; // { name, type, score }

  const scan = (names, type) => {
    (names || []).forEach(name => {
      if (!name) return;
      const candidate = name.toLowerCase();
      // Compare against the full name and against each individual word (so a
      // single-word query like "violenta" can match a surname or first name).
      // Generic words (e.g. "agent" inside "Sample Agent") are excluded so
      // they can't be the thing that triggers a match.
      const parts = [candidate, ...candidate.split(/\s+/)]
        .filter(p => !GENERIC_QUERY_WORDS.has(p));
      parts.forEach(part => {
        if (part.length < 4) return;
        const dist = levenshtein(q, part);
        const score = 1 - dist / Math.max(q.length, part.length);
        if (score >= 0.72 && (!best || score > best.score)) {
          best = { name: name, type: type, score: score };
        }
      });
    });
  };

  scan(agentNames, 'agent');
  scan(developerNames, 'developer');

  return best;
}


/**
 * Scans the raw user query for a known agent or developer name as a
 * deterministic fallback for when the LLM fails to recognize a name that's
 * actually in our system. Checks agents first, then developers, and prefers
 * the longest matching name (so "Erwin Baguioan" wins over a shorter partial
 * match if both happen to appear). Case-insensitive substring match.
 */
function findKnownNameInQuery(query, agentNames, developerNames) {
  const q = (query || '').toLowerCase();
  let best = null;

  const scan = (names, type) => {
    for (const name of (names || [])) {
      if (!name) continue;
      if (q.includes(name.toLowerCase())) {
        if (!best || name.length > best.name.length) {
          best = { name: name, type: type };
        }
      }
    }
  };

  scan(agentNames, 'agent');
  scan(developerNames, 'developer');

  return best;
}


// ── FIND MATCHING AGENT ──
function findMatchingAgent(name, agents) {
  if (!name) return null;
  
  const searchName = name.toLowerCase().trim();
  
  // Exact match first
  for (const agent of agents) {
    const fullName = (agent['Full Name'] || '').toLowerCase();
    if (fullName === searchName) {
      return agent;
    }
  }
  
  // Partial match
  for (const agent of agents) {
    const fullName = (agent['Full Name'] || '').toLowerCase();
    if (fullName.includes(searchName) || searchName.includes(fullName)) {
      return agent;
    }
  }
  
  // Try by Agent ID
  for (const agent of agents) {
    const agentId = (agent['Agent ID'] || '').toLowerCase();
    if (agentId === searchName || agentId.includes(searchName)) {
      return agent;
    }
  }
  
  return null;
}

// ── FIND MATCHING DEVELOPER ──
function findMatchingDeveloper(name, developers) {
  if (!name) return null;
  
  const searchName = name.toLowerCase().trim();
  
  for (const dev of developers) {
    const devName = (dev['Developer Name'] || '').toLowerCase();
    if (devName === searchName || devName.includes(searchName) || searchName.includes(devName)) {
      return dev;
    }
  }
  
  return null;
}

// ── GENERATE AGENT SUMMARY ──
function generateAgentSummary(context) {
  const { agentName, commissions, year } = context;
  
  if (!agentName) {
    return "I couldn't find that agent. Try asking: 'Show summary for [agent name]'";
  }
  
  const agentFullName = agentName['Full Name'] || 'Unknown';
  
  // Get agent's sales
  const agentSales = commissions.filter(c => 
    (c['Sales Agent'] || '').toLowerCase().includes(agentFullName.toLowerCase())
  );

  // Per policy: sales credited to a team (e.g. "VIOLENTA, MANNY & ANNALYN")
  // are intentionally excluded from an individual's personal summary/totals
  // above — they aren't split or attributed to any one person. Detect them
  // separately purely to surface a transparency note, so a mismatch against
  // a dashboard/leaderboard view (which may show team entries as their own
  // line) doesn't look like missing or broken data.
  const nameTokens = agentFullName.toLowerCase().split(/\s+/).filter(t => t.length >= 3);
  const teamSales = nameTokens.length > 0 ? commissions.filter(c => {
    const salesAgentLC = (c['Sales Agent'] || '').toLowerCase();
    if (!salesAgentLC || salesAgentLC.includes(agentFullName.toLowerCase())) return false; // already counted or blank
    return nameTokens.some(tok => new RegExp(`\\b${tok.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`).test(salesAgentLC));
  }) : [];
  
  const totalTransactions = agentSales.length;
  const totalGross = agentSales.reduce((sum, s) => sum + (parseFloat(s['Agent Gross Comm (₱)']) || 0), 0);
  const totalNet = agentSales.reduce((sum, s) => sum + (parseFloat(s['Agent Net Comm (₱)']) || 0), 0);
  const totalVat = agentSales.reduce((sum, s) => sum + (parseFloat(s['VAT Share (Agent) (₱)']) || 0), 0);
  const totalWT = agentSales.reduce((sum, s) => sum + (parseFloat(s['Withholding Tax (Agent) (₱)']) || 0), 0);
  
  // Get developers worked with
  const developersWorked = [...new Set(agentSales.map(s => s['Developer']).filter(Boolean))];
  
  // Get monthly breakdown
  const monthlyData = {};
  agentSales.forEach(s => {
    const date = s['Reservation Date'];
    if (date) {
      const month = date.split('/')[0] || 'Unknown';
      monthlyData[month] = (monthlyData[month] || 0) + (parseFloat(s['Agent Net Comm (₱)']) || 0);
    }
  });
  
  const status = agentName['Status'] || 'Unknown';
  const team = agentName['Team'] || 'N/A';
  
  let response = `👤 **Agent Performance Summary: ${agentFullName}**\n\n`;
  response += `📋 **Profile:**\n`;
  response += `• Agent ID: ${agentName['Agent ID'] || 'N/A'}\n`;
  response += `• Status: ${status}\n`;
  response += `• Team: ${team}\n`;
  response += `• Phone: ${agentName['Phone'] || 'N/A'}\n`;
  response += `• License: ${agentName['License No.'] || 'N/A'}\n\n`;
  
  response += `💰 **Commission Summary (${year}):**\n`;
  response += `• Total Transactions: ${totalTransactions}\n`;
  response += `• Total Gross Commission: ₱${fmt(totalGross)}\n`;
  response += `• Total VAT Share: ₱${fmt(totalVat)}\n`;
  response += `• Total Withholding Tax: ₱${fmt(totalWT)}\n`;
  response += `• Total Net Commission: ₱${fmt(totalNet)}\n`;
  response += `• Average Net per Transaction: ₱${fmt(totalTransactions > 0 ? totalNet / totalTransactions : 0)}\n\n`;
  
  if (developersWorked.length > 0) {
    response += `🏗️ **Developers Worked With:**\n`;
    developersWorked.slice(0, 5).forEach(dev => {
      response += `• ${dev}\n`;
    });
    if (developersWorked.length > 5) {
      response += `• ... and ${developersWorked.length - 5} more\n`;
    }
    response += '\n';
  }
  
  // Monthly breakdown
  if (Object.keys(monthlyData).length > 0) {
    response += `📅 **Monthly Net Commission:**\n`;
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    months.forEach(month => {
      if (monthlyData[month]) {
        response += `• ${month}: ₱${fmt(monthlyData[month])}\n`;
      }
    });
    response += '\n';
  }
  
  // Performance assessment
  const avgNet = totalTransactions > 0 ? totalNet / totalTransactions : 0;
  if (totalTransactions === 0) {
    response += `📊 **Assessment:** No transactions recorded yet.`;
  } else if (totalTransactions > 10 && avgNet > 50000) {
    response += `🏆 **Assessment:** Excellent performance! Top-tier agent with consistent high-value transactions.`;
  } else if (totalTransactions > 5 && avgNet > 20000) {
    response += `⭐ **Assessment:** Strong performer with good transaction volume and value.`;
  } else if (totalTransactions > 3) {
    response += `📈 **Assessment:** Solid performance with room to grow.`;
  } else {
    response += `🌱 **Assessment:** New or growing agent.`;
  }

  response += `\n\n💬 Want to compare ${agentFullName} to another agent, or see their pending commissions? Just ask.`;

  if (teamSales.length > 0) {
    const teamNet = teamSales.reduce((sum, s) => sum + (parseFloat(s['Agent Net Comm (₱)']) || 0), 0);
    const teamCredits = [...new Set(teamSales.map(s => s['Sales Agent']))];
    response += `\n\nℹ️ **Note:** ${teamSales.length} team-credited sale(s) totaling ₱${fmt(teamNet)} net also ` +
      `mention ${agentFullName.split(' ')[0]} (e.g. "${teamCredits[0]}") but aren't included in the totals above — ` +
      `team-credited sales are kept separate from individual summaries rather than split or double-counted.`;
  }

  return response;
}

// ── GENERATE DEVELOPER SUMMARY ──
function generateDeveloperSummary(context) {
  const { developerName, commissions, year } = context;
  
  if (!developerName) {
    return "I couldn't find that developer. Try asking: 'Show summary for [developer name]'";
  }
  
  const devNameFull = developerName['Developer Name'] || 'Unknown';
  
  const devSales = commissions.filter(c => 
    (c['Developer'] || '').toLowerCase().includes(devNameFull.toLowerCase())
  );
  
  const totalTransactions = devSales.length;
  const totalCompanyGross = devSales.reduce((sum, s) => sum + (parseFloat(s['Company Gross Comm (₱)']) || 0), 0);
  const totalCompanyNet = devSales.reduce((sum, s) => sum + (parseFloat(s['Company Net Comm (₱)']) || 0), 0);
  const totalAgentNet = devSales.reduce((sum, s) => sum + (parseFloat(s['Agent Net Comm (₱)']) || 0), 0);
  
  const agentsWorked = [...new Set(devSales.map(s => s['Sales Agent']).filter(Boolean))];
  const projects = [...new Set(devSales.map(s => s['Project'] || s['Unit ID']).filter(Boolean))];
  
  const devRate = developerName['Developer Comm Rate (%)'] || 0;
  const agentRate = developerName['Agent Comm Rate (%)'] || 0;
  const vatRate = developerName['VAT (%)'] || 0;
  const withholdingRate = developerName['Withholding Tax (%)'] || 0;
  
  let response = `🏗️ **Developer Performance Summary: ${devNameFull}**\n\n`;
  response += `📋 **Profile:**\n`;
  response += `• Developer Comm Rate: ${devRate}%\n`;
  response += `• Agent Comm Rate: ${agentRate}%\n`;
  response += `• VAT Rate: ${vatRate}%\n`;
  response += `• Withholding Tax: ${withholdingRate}%\n`;
  response += `• TIN: ${developerName['TIN'] || 'N/A'}\n`;
  response += `• Address: ${developerName['Developer Address'] || 'N/A'}\n\n`;
  
  response += `💰 **Commission Summary (${year}):**\n`;
  response += `• Total Transactions: ${totalTransactions}\n`;
  response += `• Total Company Gross Commission: ₱${fmt(totalCompanyGross)}\n`;
  response += `• Total Company Net Commission: ₱${fmt(totalCompanyNet)}\n`;
  response += `• Total Agent Net Commission: ₱${fmt(totalAgentNet)}\n`;
  response += `• Avg Company Net per Transaction: ₱${fmt(totalTransactions > 0 ? totalCompanyNet / totalTransactions : 0)}\n\n`;
  
  if (agentsWorked.length > 0) {
    response += `👤 **Agents Working on this Developer:**\n`;
    agentsWorked.slice(0, 5).forEach(agent => {
      response += `• ${agent}\n`;
    });
    if (agentsWorked.length > 5) {
      response += `• ... and ${agentsWorked.length - 5} more\n`;
    }
    response += '\n';
  }
  
  if (projects.length > 0) {
    response += `🏠 **Projects/Units:**\n`;
    projects.slice(0, 5).forEach(project => {
      response += `• ${project}\n`;
    });
    if (projects.length > 5) {
      response += `• ... and ${projects.length - 5} more\n`;
    }
    response += '\n';
  }
  
  // Assessment
  if (totalTransactions === 0) {
    response += `📊 **Assessment:** No transactions recorded yet for this developer.`;
  } else if (totalTransactions > 10 && totalCompanyNet > 1000000) {
    response += `🏆 **Assessment:** High-performing developer with significant commission generation.`;
  } else if (totalTransactions > 5) {
    response += `📈 **Assessment:** Active developer with consistent transaction volume.`;
  } else {
    response += `🌱 **Assessment:** New or developing relationship with limited transactions.`;
  }

  response += `\n\n💬 Want ${devNameFull}'s commission rates at a glance, or to compare it to another developer? Just ask.`;

  return response;
}

// ── GENERATE COMPARISON ──
function generateComparison(context) {
  const { agent1, agent2, commissions, year } = context;
  
  if (!agent1 || !agent2) {
    return "To compare agents, ask: 'Compare [Agent A] and [Agent B]'";
  }
  
  const name1 = agent1['Full Name'] || 'Unknown';
  const name2 = agent2['Full Name'] || 'Unknown';
  
  const sales1 = commissions.filter(c => (c['Sales Agent'] || '').toLowerCase().includes(name1.toLowerCase()));
  const sales2 = commissions.filter(c => (c['Sales Agent'] || '').toLowerCase().includes(name2.toLowerCase()));
  
  const net1 = sales1.reduce((sum, s) => sum + (parseFloat(s['Agent Net Comm (₱)']) || 0), 0);
  const net2 = sales2.reduce((sum, s) => sum + (parseFloat(s['Agent Net Comm (₱)']) || 0), 0);
  const count1 = sales1.length;
  const count2 = sales2.length;
  const avg1 = count1 > 0 ? net1 / count1 : 0;
  const avg2 = count2 > 0 ? net2 / count2 : 0;
  
  const devs1 = [...new Set(sales1.map(s => s['Developer']).filter(Boolean))];
  const devs2 = [...new Set(sales2.map(s => s['Developer']).filter(Boolean))];
  const commonDevs = devs1.filter(d => devs2.includes(d));
  
  let response = `⚖️ **Agent Comparison: ${name1} vs ${name2}**\n\n`;
  response += `📊 **${name1}:**\n`;
  response += `• Transactions: ${count1}\n`;
  response += `• Total Net Commission: ₱${fmt(net1)}\n`;
  response += `• Average per Transaction: ₱${fmt(avg1)}\n\n`;
  
  response += `📊 **${name2}:**\n`;
  response += `• Transactions: ${count2}\n`;
  response += `• Total Net Commission: ₱${fmt(net2)}\n`;
  response += `• Average per Transaction: ₱${fmt(avg2)}\n\n`;
  
  // Comparison
  const diff = net1 - net2;
  const diffPct = net2 > 0 ? (diff / net2) * 100 : 0;
  response += `📈 **Comparison:**\n`;
  response += `• ${name1} is ${Math.abs(diffPct).toFixed(1)}% ${diffPct > 0 ? 'ahead' : 'behind'} ${name2} in total net commission\n`;
  
  if (count1 > count2) {
    response += `• ${name1} has ${count1 - count2} more transactions\n`;
  } else if (count2 > count1) {
    response += `• ${name2} has ${count2 - count1} more transactions\n`;
  } else {
    response += `• Both agents have the same number of transactions\n`;
  }
  
  if (commonDevs.length > 0) {
    response += `• Common developers: ${commonDevs.join(', ')}\n`;
  }
  
  // Verdict
  if (net1 > net2 && avg1 > avg2) {
    response += `\n🏆 **Verdict:** ${name1} outperforms ${name2} in both total and average commission.`;
  } else if (net2 > net1 && avg2 > avg1) {
    response += `\n🏆 **Verdict:** ${name2} outperforms ${name1} in both total and average commission.`;
  } else if (net1 > net2) {
    response += `\n🏆 **Verdict:** ${name1} has higher total commission, but ${name2} has higher average per transaction.`;
  } else {
    response += `\n🏆 **Verdict:** ${name2} has higher total commission, but ${name1} has higher average per transaction.`;
  }
  
  return response;
}

// ── GENERATE TOP PERFORMERS ──
function generateTopPerformers(context) {
  const { commissions, year } = context;
  
  if (!commissions || commissions.length === 0) {
    return `🏆 No data available for ${year}.`;
  }
  
  const agentMap = {};
  commissions.forEach(s => {
    const agent = s['Sales Agent'] || 'Unknown';
    const net = parseFloat(s['Agent Net Comm (₱)']) || 0;
    if (!agentMap[agent]) {
      agentMap[agent] = { name: agent, net: 0, count: 0 };
    }
    agentMap[agent].net += net;
    agentMap[agent].count += 1;
  });
  
  const sortedAgents = Object.values(agentMap)
    .sort((a, b) => b.net - a.net)
    .slice(0, 5);
  
  const devMap = {};
  commissions.forEach(s => {
    const dev = s['Developer'] || 'Unknown';
    const net = parseFloat(s['Company Net Comm (₱)']) || 0;
    if (!devMap[dev]) {
      devMap[dev] = { name: dev, net: 0, count: 0 };
    }
    devMap[dev].net += net;
    devMap[dev].count += 1;
  });
  
  const sortedDevs = Object.values(devMap)
    .sort((a, b) => b.net - a.net)
    .slice(0, 5);
  
  let response = `🏆 **Top Performers for ${year}**\n\n`;
  
  response += `👤 **Top 5 Agents by Net Commission:**\n`;
  sortedAgents.forEach((a, i) => {
    const medal = ['🥇', '🥈', '🥉', '4️⃣', '5️⃣'][i] || '•';
    response += `${medal} ${a.name}: ₱${fmt(a.net)} (${a.count} transactions)\n`;
  });
  
  response += `\n🏗️ **Top 5 Developers by Company Net Commission:**\n`;
  sortedDevs.forEach((d, i) => {
    const medal = ['🥇', '🥈', '🥉', '4️⃣', '5️⃣'][i] || '•';
    response += `${medal} ${d.name}: ₱${fmt(d.net)} (${d.count} transactions)\n`;
  });
  
  return response;
}

// ── GENERATE TRENDS SUMMARY ──
function generateTrendsSummary(context) {
  const { commissions, year } = context;
  
  if (!commissions || commissions.length === 0) {
    return `📈 No data available for ${year}.`;
  }
  
  const monthlyData = {};
  const monthlyCount = {};
  
  commissions.forEach(s => {
    const date = s['Reservation Date'];
    if (date) {
      const parts = date.split('/');
      const month = parts[0] || 'Unknown';
      monthlyData[month] = (monthlyData[month] || 0) + (parseFloat(s['Company Net Comm (₱)']) || 0);
      monthlyCount[month] = (monthlyCount[month] || 0) + 1;
    }
  });
  
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const validMonths = months.filter(m => monthlyData[m] > 0);
  
  if (validMonths.length === 0) {
    return `📈 No monthly data available for ${year}.`;
  }
  
  const totalNet = validMonths.reduce((sum, m) => sum + monthlyData[m], 0);
  const avgMonthly = totalNet / validMonths.length;
  
  let bestMonth = validMonths.reduce((a, b) => monthlyData[a] > monthlyData[b] ? a : b);
  let worstMonth = validMonths.reduce((a, b) => monthlyData[a] < monthlyData[b] ? a : b);
  
  let response = `📈 **Monthly Trends for ${year}**\n\n`;
  validMonths.forEach(m => {
    const bar = '█'.repeat(Math.min(Math.round(monthlyData[m] / (avgMonthly || 1) * 10), 20));
    response += `• ${m}: ₱${fmt(monthlyData[m])} ${bar} (${monthlyCount[m] || 0} transactions)\n`;
  });
  response += '\n';
  
  response += `📊 **Statistics:**\n`;
  response += `• Total Net Commission: ₱${fmt(totalNet)}\n`;
  response += `• Average Monthly Net: ₱${fmt(avgMonthly)}\n`;
  response += `• Best Month: ${bestMonth} (₱${fmt(monthlyData[bestMonth])})\n`;
  response += `• Worst Month: ${worstMonth} (₱${fmt(monthlyData[worstMonth])})\n`;

  response += `\n💬 Want this compared week-over-week, or broken down by top performers for ${year}? Just ask.`;

  return response;
}

// ── GENERATE OVERALL SUMMARY ──
function generateOverallSummary(context) {
  const { stats, agents, developers, year } = context;
  
  const totalSales = stats.totalSales || 0;
  const totalGross = stats.totalCompanyGross || 0;
  const totalTax = stats.totalTax || 0;
  const totalNet = stats.totalNetCommission || 0;
  const activeAgents = stats.activeAgents || 0;
  const totalDevs = stats.totalDevelopers || 0;
  
  let response = `📊 **Performance Summary for ${year}**\n\n`;
  response += `🏢 **Company Overview:**\n`;
  response += `• Total Sales: ${totalSales} transactions\n`;
  response += `• Total Gross Commission: ₱${fmt(totalGross)}\n`;
  response += `• Total Tax (VAT + Withholding): ₱${fmt(totalTax)}\n`;
  response += `• Total Net Commission: ₱${fmt(totalNet)}\n\n`;
  
  response += `👤 **Team Overview:**\n`;
  response += `• Active Agents: ${activeAgents}\n`;
  response += `• Total Developers: ${totalDevs}\n\n`;
  
  const topAgents = stats.topAgents || [];
  if (topAgents.length > 0) {
    response += `🏆 **Top Performer:**\n`;
    const top = topAgents[0];
    response += `• ${top.name}: ₱${fmt(top.net)} (${top.count} transactions)\n\n`;
  }
  
  const recent = stats.recentSales || [];
  if (recent.length > 0) {
    response += `🕐 **Recent Activity:**\n`;
    recent.slice(0, 3).forEach(sale => {
      response += `• ${sale['Sales Agent']}: ₱${fmt(sale['Agent Net Comm'])} (${sale['Project']})\n`;
    });
  }
  
  return response;
}

// ── GENERATE HELP RESPONSE ──
function generateHelpResponse() {
  return `🤖 **Commission Bot Help**\n\n
I can help you with commission and performance data. Here's what you can ask me:

**📊 General:**
• "Show me the overall performance for 2026"
• "What's the company summary?"

**👤 Agent Specific:**
• "Show summary for [agent name]"
• "How is [agent name] performing?"

**🏗️ Developer Specific:**
• "Show summary for [developer name]"
• "How is [developer name] performing?"

**📈 Trends:**
• "Show me monthly trends"

**⚖️ Comparisons:**
• "Compare [Agent A] and [Agent B]"

**🏆 Rankings:**
• "Show me top agents"
• "Who's most active in the last 30 days?"

**💸 Money Owed:**
• "How much commission is still pending?"
• "What's the comm rate for [developer]?"

**📅 Recent Activity:**
• "How are we doing this week vs last week?"

**⚠️ Data Quality:**
• "Check for data issues"

Just type your question naturally! 😊`;
}

// ── GENERIC RESPONSE ──
function generateGenericResponse(context) {
  const { stats, year } = context;
  
  let response = `🤔 I'm not sure I fully understand. Here's what I do know:\n\n`;
  
  response += `📊 **${year} Quick Stats:**\n`;
  response += `• ${stats.totalSales || 0} transactions\n`;
  response += `• ₱${fmt(stats.totalCompanyGross || 0)} gross commission\n`;
  response += `• ₱${fmt(stats.totalNetCommission || 0)} net commission\n`;
  response += `• ${stats.activeAgents || 0} active agents\n\n`;
  
  response += `💡 **Try asking:**\n`;
  response += `• "Show summary for [agent name]"\n`;
  response += `• "Show summary for [developer name]"\n`;
  response += `• "Show me monthly trends"\n`;
  response += `• "Compare Agent A and Agent B"\n`;
  response += `• "Who are the top performers?"\n`;
  response += `• "How much commission is still pending?"\n`;
  response += `• "Check for data issues"\n`;
  response += `• "Help" for more options`;
  
  return response;
}

// ── HELPER: Parse "MM/dd/yyyy" strings into Date objects ──
function parseSlashDate(str) {
  if (!str) return null;
  const parts = String(str).split('/');
  if (parts.length !== 3) return null;
  const month = parseInt(parts[0], 10) - 1;
  const day = parseInt(parts[1], 10);
  const year = parseInt(parts[2], 10);
  if (isNaN(month) || isNaN(day) || isNaN(year)) return null;
  return new Date(year, month, day);
}

// ── GENERATE DATA ISSUES REPORT ──
function generateDataIssuesReport(year) {
  const ss = getTargetSpreadsheet();
  const sheetName = year + ' Sales';
  const sh = ss.getSheetByName(sheetName);

  if (!sh) return `⚠️ No sales sheet found for ${year}.`;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return `⚠️ No data found for ${year}.`;

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const saleIdCol = headers.findIndex(h => h === 'Sale ID');
  const agentCol = headers.findIndex(h => h === 'Sales Agent');
  const devCol = headers.findIndex(h => h === 'Developer');
  const devRateCol = headers.findIndex(h => h === 'Developer Comm Rate (%)');
  const agentRateCol = headers.findIndex(h => h === 'Agent Comm Rate (%)');

  const rows = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();

  const missingSaleId = [];
  const missingRates = [];

  rows.forEach((row, i) => {
    const hasAnyData = row.some(v => v !== '' && v !== null);
    if (!hasAnyData) return;

    const rowNum = i + 2;
    const saleId = saleIdCol !== -1 ? row[saleIdCol] : null;
    const agent = agentCol !== -1 ? row[agentCol] : '';
    const dev = devCol !== -1 ? row[devCol] : '';

    if (saleIdCol !== -1 && (!saleId || saleId === '')) {
      missingSaleId.push(`Row ${rowNum} (${agent || 'Unknown agent'} — ${dev || 'Unknown developer'})`);
    }
    if (devRateCol !== -1 && agentRateCol !== -1) {
      const devRate = parseFloat(row[devRateCol]) || 0;
      const agentRate = parseFloat(row[agentRateCol]) || 0;
      if ((devRate === 0 || agentRate === 0) && (agent || dev)) {
        missingRates.push(`${saleId || 'Row ' + rowNum} — ${agent || 'Unknown agent'} @ ${dev || 'Unknown developer'}`);
      }
    }
  });

  const commissions = getCommissions(year);
  const outstandingCA = commissions.filter(c => (parseFloat(c['Cash Advance']) || 0) > 0);

  let response = `⚠️ **Data Issues Check (${year})**\n\n`;
  let issuesFound = false;

  if (missingSaleId.length > 0) {
    issuesFound = true;
    response += `🔴 **Missing Sale ID (${missingSaleId.length}):**\n`;
    missingSaleId.slice(0, 10).forEach(m => response += `• ${m}\n`);
    if (missingSaleId.length > 10) response += `• ... and ${missingSaleId.length - 10} more\n`;
    response += '\n';
  }

  if (missingRates.length > 0) {
    issuesFound = true;
    response += `🟠 **Missing/Zero Commission Rate (${missingRates.length}):**\n`;
    missingRates.slice(0, 10).forEach(m => response += `• ${m}\n`);
    if (missingRates.length > 10) response += `• ... and ${missingRates.length - 10} more\n`;
    response += '\n';
  }

  if (outstandingCA.length > 0) {
    issuesFound = true;
    const totalCA = outstandingCA.reduce((sum, c) => sum + (parseFloat(c['Cash Advance']) || 0), 0);
    response += `🟡 **Outstanding Cash Advances (${outstandingCA.length}, ₱${fmt(totalCA)} total):**\n`;
    outstandingCA.slice(0, 10).forEach(c => {
      response += `• ${c['Sales Agent']}: ₱${fmt(c['Cash Advance'])} (${c['Sale ID']})\n`;
    });
    if (outstandingCA.length > 10) response += `• ... and ${outstandingCA.length - 10} more\n`;
    response += '\n';
  }

  if (!issuesFound) {
    response += `✅ No data issues found for ${year}. Everything looks clean!`;
  } else {
    response += `💡 Fix these directly in the ${sheetName} sheet, then ask me again to confirm they're cleared.`;
  }

  return response;
}

// ── GENERATE PENDING COMMISSIONS ──
function generatePendingCommissions(context) {
  const { commissions, year, focusAgent, focusDeveloper } = context;

  if (!commissions || commissions.length === 0) {
    return `💸 No commission data available for ${year}.`;
  }

  // Scope to a single agent or developer when the conversation calls for it
  // (either named explicitly this turn, or carried over from the agent/
  // developer just discussed).
  let scoped = commissions;
  let scopeLabel = '';
  if (focusAgent) {
    const needle = focusAgent.toLowerCase();
    scoped = commissions.filter(s => (s['Sales Agent'] || '').toLowerCase().includes(needle));
    scopeLabel = ` for ${focusAgent}`;
  } else if (focusDeveloper) {
    const needle = focusDeveloper.toLowerCase();
    scoped = commissions.filter(s => (s['Developer'] || '').toLowerCase().includes(needle));
    scopeLabel = ` with ${focusDeveloper}`;
  }

  if (scoped.length === 0) {
    return `💸 No commission transactions found${scopeLabel} for ${year}.`;
  }

  const totalNet = scoped.reduce((sum, s) => sum + (parseFloat(s['Agent Net Comm (₱)']) || 0), 0);
  const totalReleased = scoped.reduce((sum, s) => sum + (parseFloat(s['Released Commission']) || 0), 0);
  const totalReceivable = scoped.reduce((sum, s) => sum + (parseFloat(s['Receivables']) || 0), 0);

  let response = `💸 **Pending Commissions${scopeLabel} — ${year}**\n\n`;

  if (focusAgent || focusDeveloper) {
    response += `📊 **Summary:**\n`;
    response += `• Total Net Commission Earned: ₱${fmt(totalNet)}\n`;
    response += `• Already Released: ₱${fmt(totalReleased)}\n`;
    response += `• Still Pending: ₱${fmt(totalReceivable)}\n\n`;

    const pendingSales = scoped.filter(s => (parseFloat(s['Receivables']) || 0) > 0);
    if (pendingSales.length > 0) {
      response += `📋 **Pending Transactions:**\n`;
      pendingSales.slice(0, 10).forEach(s => {
        const counterpart = focusAgent ? (s['Developer'] || 'Unknown developer') : (s['Sales Agent'] || 'Unknown agent');
        response += `• ${s['Sale ID'] || 'N/A'} — ${counterpart}: ₱${fmt(s['Receivables'])} pending\n`;
      });
      if (pendingSales.length > 10) response += `• ... and ${pendingSales.length - 10} more\n`;
    } else {
      response += `✅ No pending commissions — everything's released!`;
    }
    return response;
  }

  // Company-wide view
  const byAgent = {};
  scoped.forEach(s => {
    const receivable = parseFloat(s['Receivables']) || 0;
    if (receivable > 0) {
      const agent = s['Sales Agent'] || 'Unknown';
      byAgent[agent] = (byAgent[agent] || 0) + receivable;
    }
  });

  const sortedAgents = Object.entries(byAgent).sort((a, b) => b[1] - a[1]).slice(0, 5);

  response += `📊 **Company-wide:**\n`;
  response += `• Total Net Commission Earned: ₱${fmt(totalNet)}\n`;
  response += `• Already Released: ₱${fmt(totalReleased)}\n`;
  response += `• Still Pending: ₱${fmt(totalReceivable)}\n\n`;

  if (sortedAgents.length > 0) {
    response += `👤 **Top Agents Awaiting Release:**\n`;
    sortedAgents.forEach(([agent, amt]) => {
      response += `• ${agent}: ₱${fmt(amt)} pending\n`;
    });
  } else {
    response += `✅ No pending commissions — everything's released!`;
  }

  return response;
}

// ── GENERATE DEVELOPER RATE LOOKUP ──
function generateDeveloperRateResponse(dev) {
  const name = dev['Developer Name'] || 'Unknown';

  let response = `🏗️ **${name} — Commission Rates**\n\n`;
  response += `• Developer Comm Rate: ${dev['Developer Comm Rate (%)'] || 0}%\n`;
  response += `• Agent Comm Rate: ${dev['Agent Comm Rate (%)'] || 0}%\n`;
  response += `• VAT: ${dev['VAT (%)'] || 0}%\n`;
  response += `• Withholding Tax: ${dev['Withholding Tax (%)'] || 0}%\n`;
  if (dev['LMF (%)']) response += `• LMF: ${dev['LMF (%)']}%\n`;
  response += `\n💬 Want the full performance summary for ${name} instead? Just ask.`;

  return response;
}

// ── GENERATE WEEK-OVER-WEEK COMPARISON ──
function generateWeekComparison(year) {
  const commissions = getCommissions(year);

  if (!commissions || commissions.length === 0) {
    return `📅 No sales data available for ${year} to compare weeks.`;
  }

  const now = new Date();
  const startOfThisWeek = new Date(now);
  startOfThisWeek.setDate(now.getDate() - 6);
  startOfThisWeek.setHours(0, 0, 0, 0);
  const startOfLastWeek = new Date(startOfThisWeek);
  startOfLastWeek.setDate(startOfThisWeek.getDate() - 7);
  const endOfLastWeek = new Date(startOfThisWeek.getTime() - 1);

  let thisWeekCount = 0, thisWeekNet = 0, lastWeekCount = 0, lastWeekNet = 0;

  commissions.forEach(s => {
    const d = parseSlashDate(s['Reservation Date']);
    if (!d) return;
    const net = parseFloat(s['Company Net Comm (₱)']) || 0;
    if (d >= startOfThisWeek && d <= now) {
      thisWeekCount++;
      thisWeekNet += net;
    } else if (d >= startOfLastWeek && d <= endOfLastWeek) {
      lastWeekCount++;
      lastWeekNet += net;
    }
  });

  const netDelta = thisWeekNet - lastWeekNet;
  const pctDelta = lastWeekNet > 0 ? (netDelta / lastWeekNet * 100) : (thisWeekNet > 0 ? 100 : 0);
  const trendIcon = netDelta > 0 ? '📈' : netDelta < 0 ? '📉' : '➡️';

  let response = `📅 **This Week vs Last Week (${year})**\n\n`;
  response += `**This Week (last 7 days):**\n`;
  response += `• Transactions: ${thisWeekCount}\n`;
  response += `• Company Net Commission: ₱${fmt(thisWeekNet)}\n\n`;
  response += `**Last Week:**\n`;
  response += `• Transactions: ${lastWeekCount}\n`;
  response += `• Company Net Commission: ₱${fmt(lastWeekNet)}\n\n`;
  response += `${trendIcon} **Change:** ${netDelta >= 0 ? '+' : ''}₱${fmt(Math.abs(netDelta))} (${pctDelta >= 0 ? '+' : ''}${pctDelta.toFixed(1)}%)`;

  return response;
}

// ── GENERATE RECENT ACTIVITY (LAST 30 DAYS) ──
function generateRecentActivity(year) {
  const commissions = getCommissions(year);

  if (!commissions || commissions.length === 0) {
    return `🔥 No sales data available for ${year}.`;
  }

  const now = new Date();
  const cutoff = new Date(now);
  cutoff.setDate(now.getDate() - 30);
  cutoff.setHours(0, 0, 0, 0);

  const recentByAgent = {};
  commissions.forEach(s => {
    const d = parseSlashDate(s['Reservation Date']);
    if (!d || d < cutoff || d > now) return;
    const agent = s['Sales Agent'] || 'Unknown';
    if (!recentByAgent[agent]) recentByAgent[agent] = { count: 0, net: 0 };
    recentByAgent[agent].count++;
    recentByAgent[agent].net += parseFloat(s['Agent Net Comm (₱)']) || 0;
  });

  const sorted = Object.entries(recentByAgent).sort((a, b) => b[1].count - a[1].count).slice(0, 5);

  if (sorted.length === 0) {
    return `🔥 No transactions recorded in the last 30 days for ${year}.`;
  }

  let response = `🔥 **Most Active Agents (Last 30 Days)**\n\n`;
  sorted.forEach(([agent, data], i) => {
    const medal = ['🥇', '🥈', '🥉', '4️⃣', '5️⃣'][i] || '•';
    response += `${medal} ${agent}: ${data.count} transactions (₱${fmt(data.net)} net)\n`;
  });

  return response;
}

// ── HELPER: Format Currency ──
function fmt(v) {
  const n = parseFloat(v) || 0;
  const rounded = Math.round((n + Number.EPSILON) * 100) / 100;
  return rounded.toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}