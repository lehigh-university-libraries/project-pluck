// Wrap library functions that can't be called directly
function onOpen() {
  ProjectPluck.onOpen();
}
function showSidebar() {
  ProjectPluck.showSidebar();
}
function getLocations(config) {
  return ProjectPluck.getLocations(config);
}
function initSheetForLocation(config) {
  ProjectPluck.initSheetForLocation(config);
}
function stopLoading() {
  ProjectPluck.stopLoading();
}
function addDecisions() {
  ProjectPluck.addDecisions();
}
function processFinalStates() {
  ProjectPluck.processFinalStates();
}
