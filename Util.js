// Performance testing
let lastTime = false;
function logTime(text) {
  const now = Date.now();
  if (lastTime) {
    const elapsed = now - lastTime;
    console.log(`time: ${elapsed / 1000}s, at ${text}`);
  }
  lastTime = now;
}

// Caching
function getOrCreate(cacheKey, creationFunction, cacheTime) {
  let cache = CacheService.getScriptCache();
  let value = JSON.parse(cache.get(cacheKey));
  if (value == null) {
    console.log(`No cached ${cacheKey}; creating.`);
    value = creationFunction();
    cache.put(cacheKey, JSON.stringify(value), cacheTime);
  }
  else {
    console.log(`Fetched ${cacheKey} from cache.`);
  }
  return value;
}
function clearCache() {
  CacheService.getScriptCache().removeAll([
    'authenticate', 
    'loadLocations', 
    'loadStatisticalCodes', 
    'loadDecisionNoteTypeId', 
    'loadInstanceStatusWithdrawnId',
  ]);
}

// Kill switch for long-running tasks
const KILL_SWITCH_KEY = "killSwitch";
const KILL_SWITCH_VALUE = "Delete this to kill job.";
function initKillSwitch() {
  PropertiesService.getScriptProperties().setProperty(KILL_SWITCH_KEY, KILL_SWITCH_VALUE);
}
function flipKillSwitch() {
  PropertiesService.getScriptProperties().deleteProperty(KILL_SWITCH_KEY);
}
function killSwitchFlipped() {
  const value = PropertiesService.getScriptProperties().getProperty(KILL_SWITCH_KEY);
  const flipped = (value == null);
  if (flipped) {
    console.log("kill switch flipped");
  }
  return flipped;
}

// Email
function email(subject, body) {
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: subject,
    htmlBody: body,
  });
}
