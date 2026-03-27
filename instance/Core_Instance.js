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
function getLocations(environment) {
  PropertiesService.getScriptProperties().setProperty("environment", environment);
  initProperties();
  return ProjectPluck.getLocations();
}
function initSheetForLocation(environment, location_id, start_call_number_prefix, end_call_number_prefix) {
  PropertiesService.getScriptProperties().setProperty("environment", environment);
  PropertiesService.getScriptProperties().setProperty('lastSheetName', SpreadsheetApp.getActiveSheet().getSheetName());
  const sheet = SpreadsheetApp.getActiveSheet();
  setSheetMetadata(sheet, 'location_id', location_id);
  setSheetMetadata(sheet, 'start_call_number_prefix', start_call_number_prefix);
  setSheetMetadata(sheet, 'end_call_number_prefix', end_call_number_prefix);
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

function getActiveSheetState() {
  initProperties();
  const sheet = SpreadsheetApp.getActiveSheet();
  const metadata = Object.fromEntries(
    sheet.getDeveloperMetadata().map(m => [m.getKey(), m.getValue()])
  );
  return {
    sheetName: sheet.getName(),
    hasItems: sheet.getLastRow() > 1,
    locationId: metadata['location_id'] ?? null,
    startCallNumberPrefix: metadata['start_call_number_prefix'] ?? null,
    endCallNumberPrefix: metadata['end_call_number_prefix'] ?? null,
  };
}

function checkProperties() {
  getLoadingMode();
}
