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
  PropertiesService.getScriptProperties().setProperty("location_id", location_id);
  PropertiesService.getScriptProperties().setProperty("start_call_number_prefix", start_call_number_prefix);
  PropertiesService.getScriptProperties().setProperty("end_call_number_prefix", end_call_number_prefix);
  PropertiesService.getScriptProperties().setProperty('lastSheetName', SpreadsheetApp.getActiveSheet().getSheetName());
  initProperties();
  ProjectPluck.initSheetForLocation();
}
function loadMoreItems() {
  initProperties();
  ProjectPluck.loadMoreItems();
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

function hasItems() {
  return SpreadsheetApp.getActiveSheet().getLastRow() > 1;
}

function getCallNumberPrefixes() {
  const props = PropertiesService.getScriptProperties();
  return {
    start: props.getProperty('start_call_number_prefix'),
    end: props.getProperty('end_call_number_prefix'),
  };
}

function checkProperties() {
  getLoadingMode();
}
