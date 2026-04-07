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
    'loadLegacyCircNoteTypeId',
    'loadOclcIdentifierTypeId',
  ]);
}

// Kill switch for long-running tasks
const KILL_SWITCH_KEY = "killSwitch";
const KILL_SWITCH_VALUE = "Delete this to kill job.";
function initKillSwitch() {
  properties.setProperty(KILL_SWITCH_KEY, KILL_SWITCH_VALUE);
}
function flipKillSwitch() {
  properties.deleteProperty(KILL_SWITCH_KEY);
}
function killSwitchFlipped() {
  const value = properties.getProperty(KILL_SWITCH_KEY);
  const flipped = (value == null);
  if (flipped) {
    console.log("kill switch flipped");
  }
  return flipped;
}

// UptimeRobot monitoring
const UPTIME_ROBOT_API_KEY = "uptimeRobotApiKey";
const UPTIME_ROBOT_MONITOR_ID = "uptimeRobotMonitorId";
const UPTIME_ROBOT_EDIT_MONITOR_URL = "https://api.uptimerobot.com/v2/editMonitor";
const UPTIME_ROBOT_HEARTBEAT_URL = "https://heartbeat.uptimerobot.com/";
const UPTIME_ROBOT_HEARTBEAT_KEY = "uptimeRobotHeartbeatKey";
function isMonitoringEnabled() {
  return !!(
    properties.getProperty(UPTIME_ROBOT_API_KEY) &&
    properties.getProperty(UPTIME_ROBOT_MONITOR_ID) &&
    properties.getProperty(UPTIME_ROBOT_HEARTBEAT_KEY)
  );
}
function startMonitoring() {
  if (isMonitoringEnabled()) {
    sendHeartbeat();
    changeMonitoring(1);
  }
}
function stopMonitoring() {
  if (isMonitoringEnabled()) {
    changeMonitoring(0);
  }
}
function sendHeartbeat() {
  const heartbeatKey = properties.getProperty(UPTIME_ROBOT_HEARTBEAT_KEY);
  const response = UrlFetchApp.fetch(UPTIME_ROBOT_HEARTBEAT_URL + heartbeatKey, {
    'method': 'post'
  });
  const responseText = response.getContentText();
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 400) {
    console.error(`UptimeRobot heartbeat: Error response: ${response.getResponseCode()}, ${responseText}`);
    return null;
  }
  else {
    const responseData = JSON.parse(responseText);
    console.log("UptimeRobot heartbeat sent.");
    return responseData;
  }

}
function changeMonitoring(newStatus) {
  const apiKey = properties.getProperty(UPTIME_ROBOT_API_KEY);
  const monitorId = properties.getProperty(UPTIME_ROBOT_MONITOR_ID);
  const formData = {
    'api_key': apiKey,
    'id': monitorId,
    'status': newStatus
  }
  const response = UrlFetchApp.fetch(UPTIME_ROBOT_EDIT_MONITOR_URL, {
    'method': 'post',
    'payload': formData
  });
  const responseText = response.getContentText();
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 400) {
    console.error(`UptimeRobot: Error response: ${response.getResponseCode()}, ${responseText}`);
    return null;
  }
  else {
    const responseData = JSON.parse(responseText);
    console.log("UptimeRobot monitoring status is now " + newStatus);
    return responseData;
  }
}

// Sheet developer metadata
function getSheetMetadata(sheet, key) {
  const found = sheet.createDeveloperMetadataFinder().withKey(key).find();
  return found.length > 0 ? found[0].getValue() : null;
}

function setSheetMetadata(sheet, key, value) {
  const found = sheet.createDeveloperMetadataFinder().withKey(key).find();
  if (found.length > 0) {
    found[0].setValue(value);
  } else {
    sheet.addDeveloperMetadata(key, value);
  }
}

function deleteSheetMetadata(sheet, key) {
  sheet.createDeveloperMetadataFinder().withKey(key).find().forEach(m => m.remove());
}

// Email
function email(subject, body) {
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: subject,
    htmlBody: body,
  });
}
