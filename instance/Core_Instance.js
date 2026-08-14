// Wrap library functions, pass in properties as needed
function onOpen() {
  initProperties();
  checkProperties();
  ProjectPluck.onOpen();
}
function getLoadingMode() {
  initProperties();
  return ProjectPluck.getLoadingMode();
}
function showSidebar() {
  initProperties();
  ProjectPluck.showSidebar();
}
function reloadFolioMetadata() {
  initProperties();
  ProjectPluck.reloadFolioMetadata();
}
function showDeveloperInfo() {
  ProjectPluck.showDeveloperInfo();
}
function getLocationsAndSheetState(environment) {
  PropertiesService.getScriptProperties().setProperty("environment", environment);
  initProperties();
  const locations = ProjectPluck.getLocations();
  const sheetState = getActiveSheetState();
  return { locations, sheetState };
}
function initSheetForLocation(environment, location_id, start_call_number_prefix, end_call_number_prefix) {
  PropertiesService.getScriptProperties().setProperty("environment", environment);
  const sheet = SpreadsheetApp.getActiveSheet();
  setSheetMetadata(sheet, 'location_id', location_id);
  setSheetMetadata(sheet, 'start_call_number_prefix', start_call_number_prefix);
  setSheetMetadata(sheet, 'end_call_number_prefix', end_call_number_prefix);
  setSheetMetadata(sheet, 'sheet_locked', 'true');
  setSheetMetadata(sheet, 'loading_active', 'true');
  initProperties();
  ProjectPluck.initSheetForLocation();
}

function setSheetMetadata(sheet, key, value) {
  const found = sheet.createDeveloperMetadataFinder().withKey(key).find();
  if (found.length > 0) {
    found[0].setValue(value);
  } else {
    sheet.addDeveloperMetadata(key, value);
  }
}
function loadMoreItems() {
  initProperties();
  ProjectPluck.loadMoreItems();
}
function validateCallNumberBoundaries(environment, location_code, start_call_number_prefix, end_call_number_prefix) {
  PropertiesService.getScriptProperties().setProperty("environment", environment);
  initProperties();
  return ProjectPluck.validateCallNumberBoundaries(location_code, start_call_number_prefix, end_call_number_prefix);
}
function stopLoading() {
  initProperties();
  ProjectPluck.stopLoading();
}
function addDecisions() {
  initProperties();
  ProjectPluck.addDecisions();
}
function processFinalStates() {
  initProperties();
  ProjectPluck.processFinalStates();
}

// Pass properties to library, for use during this script execution
function initProperties() {
  ProjectPluck.initProperties(PropertiesService.getScriptProperties());
}

function getEnvironment() {
  return PropertiesService.getScriptProperties().getProperty("environment");
}

function getSidebarInitData() {
  initProperties();
  const loadingMode = ProjectPluck.getLoadingMode();
  const environment = PropertiesService.getScriptProperties().getProperty('environment');
  if (!environment) {
    return { loadingMode, environment: null, locations: null, sheetState: null };
  }
  const locations = ProjectPluck.getLocations();
  const sheetState = getActiveSheetState();
  return { loadingMode, environment, locations, sheetState };
}

function getActiveSheetState() {
  initProperties();
  const sheet = SpreadsheetApp.getActiveSheet();
  const metadata = Object.fromEntries(
    sheet.getDeveloperMetadata().map(m => [m.getKey(), m.getValue()])
  );
  return {
    sheetName: sheet.getName(),
    sheetLocked: metadata['sheet_locked'] === 'true',
    locationId: metadata['location_id'] ?? null,
    startCallNumberPrefix: metadata['start_call_number_prefix'] ?? null,
    endCallNumberPrefix: metadata['end_call_number_prefix'] ?? null,
  };
}

function showAlert(message) {
  SpreadsheetApp.getUi().alert(message);
}

function showColumnPreferences() {
  ProjectPluck.showColumnPreferences();
}

function getColumnPreferencesData() {
  initProperties();
  return ProjectPluck.getColumnPreferencesData();
}

function saveColumnPreferences(prefs) {
  initProperties();
  ProjectPluck.saveColumnPreferences(prefs);
}

function showAutoDecisionRules() {
  ProjectPluck.showAutoDecisionRules();
}

function getAutoDecisionRulesData() {
  initProperties();
  return ProjectPluck.getAutoDecisionRulesData();
}

function saveAutoDecisionPreferences(prefs) {
  initProperties();
  ProjectPluck.saveAutoDecisionPreferences(prefs);
}

function checkProperties() {
  getLoadingMode();
}
