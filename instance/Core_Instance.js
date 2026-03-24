// Wrap library functions, pass in properties as needed
function onOpen() {
  initProperties();
  ProjectPluck.onOpen();
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
function initSheetForLocation(environment, location_id) {
  PropertiesService.getScriptProperties().setProperty("environment", environment);
  PropertiesService.getScriptProperties().setProperty("location_id", location_id);
  PropertiesService.getScriptProperties().setProperty('lastSheetName', SpreadsheetApp.getActiveSheet().getSheetName());
  initProperties();
  ProjectPluck.initSheetForLocation();
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
