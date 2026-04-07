// Rules are initialized lazily (not at top level) because they reference constants
// from Core_Library.js (e.g. NO_CHANGE) which may not yet be defined when this
// file is first evaluated by Apps Script.
let AUTO_DECISION_RULES;
function getAutoDecisionRules() {
  if (!AUTO_DECISION_RULES) {
    AUTO_DECISION_RULES = [
      {
        name: 'Retention agreement, no damage',
        decision: NO_CHANGE,
        params: {},
        evaluate: (item, _params) => hasRetentionAgreement(item) && !parseDamage(item),
      },
      {
        name: 'Low worldwide holdings',
        decision: NO_CHANGE,
        params: { maxHoldings: 25 },
        evaluate: (item, params) => {
          const h = parseOclcHoldings(item);
          return h !== null && h !== undefined && h <= params.maxHoldings;
        },
      },
      {
        name: 'High circulation',
        decision: NO_CHANGE,
        params: { minCircCount: 1 },
        evaluate: (item, params) => parseFolioCircCount(item) >= params.minCircCount,
      },
    ];
  }
  return AUTO_DECISION_RULES;
}

function applyAutoDecisions(sheet, row, item) {
  const decisionCol = getColumn(DECISION);
  if (!decisionCol) return;

  const existing = sheet.getRange(row, decisionCol).getValue();
  if (existing) return;

  const prefs = getAutoDecisionPreferences();
  for (const rule of getAutoDecisionRules()) {
    const pref = prefs[rule.name];
    if (!pref?.enabled) continue;
    const params = { ...rule.params, ...pref.params };
    if (rule.evaluate(item, params)) {
      const decision = pref.decision ?? rule.decision;
      sheet.getRange(row, decisionCol).setValue(decision);
      const noteCol = getColumn(DECISION_ADDENDUM);
      if (noteCol) sheet.getRange(row, noteCol).setValue(`Auto-decision by rule: ${rule.name}`);
      return;
    }
  }
}

function getAutoDecisionPreferences() {
  const stored = properties.getProperty('auto_decision_preferences');
  const defaults = Object.fromEntries(
    getAutoDecisionRules().map(r => [r.name, { enabled: false, decision: r.decision, params: r.params }])
  );
  return stored ? { ...defaults, ...JSON.parse(stored) } : defaults;
}

function saveAutoDecisionPreferences(prefs) {
  properties.setProperty('auto_decision_preferences', JSON.stringify(prefs));
}

function getAutoDecisionRulesData() {
  const prefs = getAutoDecisionPreferences();
  return {
    decisions: DECISIONS,
    rules: getAutoDecisionRules().map(r => ({
      name: r.name,
      defaultDecision: r.decision,
      decision: prefs[r.name]?.decision ?? r.decision,
      params: Object.fromEntries(
        Object.keys(r.params).map(k => [k, prefs[r.name]?.params?.[k] ?? r.params[k]])
      ),
      enabled: prefs[r.name]?.enabled ?? false,
    })),
  };
}
