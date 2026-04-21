// Rules are initialized lazily (not at top level) because they reference constants
// from Core_Library.js (e.g. NO_CHANGE) which may not yet be defined when this
// file is first evaluated by Apps Script.
let AUTO_DECISION_RULES;
function getAutoDecisionRules() {
  if (!AUTO_DECISION_RULES) {
    AUTO_DECISION_RULES = [
      {
        name: 'Retention agreement',
        decision: NO_CHANGE,
        params: {},
        evaluate: (item, _params) => hasRetentionAgreement(item),
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
      {
        name: 'Damage note',
        decision: WITHDRAW,
        params: { damageContains: '' },
        evaluate: (item, params) => { 
          const damage = parseDamage(item);
          return damage && (!params.damageContains || damage.toLowerCase().includes(params.damageContains.toLowerCase()));
        },
      },
      {
        name: 'Electronic holdings',
        decision: WITHDRAW,
        params: { accessMethod: 'U' },
        evaluate: (item, params) => { 
          const electronicHoldings = JSON.parse(item['electronic_holdings'] || '[]');
          return electronicHoldings.some(
            eHolding => eHolding && (eHolding['access_method'] ?? '') == params.accessMethod
          );
        },
      },
    ];
  }
  return AUTO_DECISION_RULES;
}

function loadExistingDecisions(sheet, startRow, count) {
  const decisionCol = getColumn(DECISION);
  if (!decisionCol) return new Array(count).fill('');
  return sheet.getRange(startRow, decisionCol, count, 1).getValues().map(r => r[0]);
}

function applyAutoDecisions(sheet, row, item, existingDecision) {
  const decisionCol = getColumn(DECISION);
  if (!decisionCol) return;

  if (existingDecision) return;

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
