
// Performance tuning
const FOLIO_LOAD_COUNT = 50;
const METADB_LOAD_COUNT = 1000;
const FOLIO_ENRICH_COUNT = 50;
const METADB_ENRICH_COUNT = 50;
const FLUSH_RATE = 5;
const PAUSE_TIME = 5000;

// Spreadsheet UI
const TAB_COMPLETE_COLOR = 'green';
const MAX_ADDENDUM_LENGTH = 500;
const ADD_SUCCESS_MESSAGE = 'Added Note';
const FINAL_STATE_SUCCESS_MESSAGE = 'Final State Processed';
const SUCCESS_BACKGROUND = 'lightgreen';
const FAILURE_BACKGROUND = 'lightcoral';


var DECISION_CODE_TO_ID;
var LOCATIONS;

// Properties are limited to the current run of the script
let properties = null;

function test() {
  // testGetLocations();
  // testInitSheetForLocation();
  // testAddDecision();
  // testProcessFinalStates();
}

function testGetLocations() {
  getLocations({
    'environment': 'test',
  });
}

function testInitSheetForLocation() {
  initSheetForLocation({
    'environment': 'test',
    'location_id': '460df2a6-6146-4749-9ff0-a0d0730e0214',
  });
}

function testAddDecision() {
  initFolio();
  addDecision(SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getRow());
}

function testProcessFinalStates() {
  initFolio();
  processFinalStates(SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getRow());
}

function initProperties(instanceProperties) {
  properties = instanceProperties;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Project Pluck')
    .addItem('Show Sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Select columns', 'showColumnPreferences')
    .addItem('Configure auto-decision rules', 'showAutoDecisionRules')
    .addItem('Reload FOLIO metadata', 'reloadFolioMetadata')
    .addItem('Show developer info', 'showDeveloperInfo')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Project Pluck')
    .setWidth(500);
  SpreadsheetApp.getUi()
    .showSidebar(html);
}


function reloadFolioMetadata() {
  clearCache();
  initFolio();
  SpreadsheetApp.getUi().alert('FOLIO metadata reloaded.');
}

function showDeveloperInfo() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const metadata = Object.fromEntries(
    sheet.getDeveloperMetadata().map(m => [m.getKey(), m.getValue()])
  );
  SpreadsheetApp.getUi().alert(
    `Sheet: ${sheet.getName()}\n` +
    `Sheet locked: ${metadata['sheet_locked'] === 'true'}\n` +
    `Loading active: ${metadata['loading_active'] === 'true'}\n` +
    `Location ID: ${metadata['location_id'] ?? null}\n` +
    `Start call number prefix: ${metadata['start_call_number_prefix'] ?? null}\n` +
    `End call number prefix: ${metadata['end_call_number_prefix'] ?? null}\n` +
    `Columns: ${metadata['headers'] ? JSON.parse(metadata['headers']).join(', ') : 'all (default)'}`
  );
}

function getLocations() {
  initFolio();
  return Object.entries(LOCATIONS).sort((a, b) => {return a['code'] < b['code']});
}

function initSheetForLocation() {
  console.log("initSheetForLocation. config: ", properties);

  // logTime("start initSheetForLocation");
  initKillSwitch();
  loadMoreItems();
}

function loadMoreItems() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    .find(s => s.createDeveloperMetadataFinder().withKey('loading_active').find().length > 0);
  if (!sheet) {
    console.log('No sheet with loading_active found');
    return;
  }
  try {
    tryLoadMoreItems(sheet);
  }
  catch (error) {
    console.log('Error loading items: ', error);
    email(`Error loading items to ${sheet.getName()}`, `${error}`);
  }
}

function tryLoadMoreItems(sheet) {
  if (killSwitchFlipped()) {
    deleteSheetMetadata(sheet, 'loading_active');
    stopMonitoring();
    return;
  }

  console.log(`UptimeRobot monitoring ${isMonitoringEnabled() ? 'enabled' : 'disabled'}.`);
  startMonitoring();

  headers = getSheetHeaders(sheet);
  const loadOclc = headers.some(h => ALL_HEADERS.get(h) === OCLC_SOURCE);
  const loadHathi = headers.some(h => ALL_HEADERS.get(h) === HATHI_SOURCE);
  loadFolioNotes = headers.some(h => ALL_HEADERS.get(h) === FOLIO_NOTES_SOURCE);

  initFolio(loadFolioNotes);
  if (loadOclc) initOclc();
  if (loadHathi) initHathi();
  writeHeaders(sheet);

  const locationId            = getSheetMetadata(sheet, 'location_id');
  const startCallNumberPrefix = getSheetMetadata(sheet, 'start_call_number_prefix');
  const endCallNumberPrefix   = getSheetMetadata(sheet, 'end_call_number_prefix');
  writeTabName(sheet, locationId, startCallNumberPrefix, endCallNumberPrefix);

  const loadingMode = getLoadingMode();
  let offset = sheet.getLastRow() - 1;
  const loadCount = loadingMode === 'folio' ? FOLIO_LOAD_COUNT : METADB_LOAD_COUNT;
  const enrichCount = loadingMode === 'folio' ? FOLIO_ENRICH_COUNT : METADB_ENRICH_COUNT;
  const needed = sheet.getLastRow() + loadCount - sheet.getMaxRows();
  if (needed > 0) sheet.insertRowsAfter(sheet.getMaxRows(), needed);

  let items;
  if (loadingMode === 'folio') {
    console.log("Loading items in 'folio' mode.");
    items = loadItemsFolio(locationId, offset, loadCount);
  } else {
    console.log("Loading items in 'metadb' mode.");
    let locationCode = LOCATIONS[locationId]?.['code'];
    items = loadItemsMetadb(locationCode, startCallNumberPrefix, endCallNumberPrefix, offset, loadCount);
  }
  console.log(`writing items to sheet with offset ${offset} and count ${loadCount}`);
  if (items.length == 0) {
    console.log("Loaded all items for this sheet");
    const lastRow = sheet.getLastRow();
    const blankRows = sheet.getMaxRows() - lastRow;
    if (blankRows > 0) sheet.deleteRows(lastRow + 1, blankRows);
    sheet.setTabColor(TAB_COMPLETE_COLOR);
    deleteSheetMetadata(sheet, 'loading_active');
    email(`${sheet.getName()} load complete`, `Google Sheets is done loading the items ${sheet.getName()}.`);
    stopMonitoring();
    return;
  }

  if (loadingMode === 'folio') {
    for (const item of items) {
      enrichItem(item, true, true, true);
      normalizeFolioItem(item);
      if (killSwitchFlipped()) {
        deleteSheetMetadata(sheet, 'loading_active');
        stopMonitoring();
        return;
      }
    }
  }

  let row = sheet.getLastRow();
  for (let i = 0; i < items.length; i += enrichCount) {
    const batch = items.slice(i, i + enrichCount);
    if (loadOclc) enrichBatchFromOclc(batch);
    if (killSwitchFlipped()) {
      deleteSheetMetadata(sheet, 'loading_active');
      stopMonitoring();
      return;
    }
    if (loadHathi) enrichBatchFromHathi(batch);
    const batchStartRow = row + 1;
    const existingDecisions = loadExistingDecisions(sheet, batchStartRow, batch.length);
    for (let i = 0; i < batch.length; i++) {
      const item = batch[i];
      row++;
      writeItemToSheet(sheet, row, item);
      initDecision(sheet, row);
      applyAutoDecisions(sheet, row, item, existingDecisions[i]);
      if (row % FLUSH_RATE == 0) {
        SpreadsheetApp.flush();
      }
    }
    if (killSwitchFlipped()) {
      break;
    }
  }

  scheduleLoadMoreItems();
  stopMonitoring();
}

function scheduleLoadMoreItems() {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    ScriptApp.deleteTrigger(trigger);
  }
  const trigger = ScriptApp.newTrigger('loadMoreItems')
    .timeBased()
    .after(PAUSE_TIME)
    .create();
  console.log('loadMoreItems trigger scheduled to fire at: ' + new Date(Date.now() + PAUSE_TIME));
}

function stopLoading() {
  flipKillSwitch();
}






function initDecision(sheet, row) {
  sheet.getRange(row, getColumn(DECISION)).setDataValidation(DECISIONS_VALIDATION);
}

function addDecisions() {
  headers = getSheetHeaders(SpreadsheetApp.getActiveSheet());
  initFolio();
  processSelectedRows(addDecision);
}

function processFinalStates() {
  headers = getSheetHeaders(SpreadsheetApp.getActiveSheet());
  initFolio();
  processSelectedRows(processFinalState);
}

function processSelectedRows(callback) {
  const selection = SpreadsheetApp.getActiveSheet().getSelection();
  const ranges = selection.getActiveRangeList().getRanges();
  for (const range of ranges) {
    const start = range.getRow();
    const end = range.getLastRow();
    for (let row = start; row <= end; row ++) {
      const decision = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(DECISION)).getValue();
      if (!decision) continue;
      callback(row);
    }
  }
}

function addDecision(row) {
  console.log("adding decision for row " + row);
  const item = loadItemForRow(row);

  const decision = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(DECISION)).getValue();
  const now = new Date().toString();
  const decisionAddendum = String(
    SpreadsheetApp.getActiveSheet().getRange(row, getColumn(DECISION_ADDENDUM)).getValue() || ''
  ).substring(0, MAX_ADDENDUM_LENGTH);
  let decisionNote = `${decision} : ${now}`;
  if (decisionAddendum) {
    decisionNote += `: ${decisionAddendum}`;
  }
  item['notes'].push({
    itemNoteTypeId: DECISION_NOTE_TYPE_ID,
    note: decisionNote,
    staffOnly: true,
  });

  const error = putItem(item);
  const addStatusCell = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(ADD_DECISION_STATUS));
  if (error) {
    addStatusCell.setValue(error);
    addStatusCell.setBackground(FAILURE_BACKGROUND);
  }
  else {
    addStatusCell.setValue(ADD_SUCCESS_MESSAGE);
    addStatusCell.setBackground(SUCCESS_BACKGROUND);
  }
}

function processFinalState(row) {
  console.log("processing final state for row " + row);
  const item = loadItemForRow(row, {holdingsRecord: true, instance: true});

  const decision = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(DECISION)).getValue();
  const finalStateCode = DECISION_TO_FINAL_STATE.get(decision);
  const finalStateCodeId = DECISION_CODE_TO_ID[finalStateCode];
  item['statisticalCodeIds'].push(finalStateCodeId);

  if (finalStateCode == FINAL_STATE_WITHDRAW) {
    item['status']['name'] = 'Withdrawn';
    item['discoverySuppress'] = true;
  }

  if (MISSING == decision) {
    item['circulationNotes'].push({
      noteType: MISSING_CHECK_IN_NOTE_TYPE,
      note: MISSING_CHECK_IN_NOTE_TEXT,
      staffOnly: true,
    });
  }

  let error = putItem(item);
  const processFinalStateCell = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(PROCESS_FINAL_STATE_STATUS));
  if (error) {
    processFinalStateCell.setValue('Error processing final state for item: ' + error);
    processFinalStateCell.setBackground(FAILURE_BACKGROUND);
    return;
  }
  
  if (finalStateCode == FINAL_STATE_WITHDRAW) {
    if (!hasUnsuppressedItems(item.holdingsRecord, item)) {
      console.log("suppress holdings record");
      item.holdingsRecord['discoverySuppress'] = true;
      error = putHoldingsRecord(item.holdingsRecord);
      if (error) {
        processFinalStateCell.setValue('Error withdrawing holdings record: ' + error);
        processFinalStateCell.setBackground(FAILURE_BACKGROUND);
        return;
      }
    
      if (!hasUnsuppressedHoldingsRecords(item.instance, item.holdingsRecord)) {
        console.log("suppress instance");
        item.instance['discoverySuppress'] = true;
        item.instance['statusId'] = INSTANCE_STATUS_WITHDRAWN_ID;
        error = putInstance(item.instance);
        if (error) {
          processFinalStateCell.setValue('Error withdrawing instance: ' + error);
          processFinalStateCell.setBackground(FAILURE_BACKGROUND);
          return;
        }
      }
    }
  }

  processFinalStateCell.setValue(FINAL_STATE_SUCCESS_MESSAGE);
  processFinalStateCell.setBackground(SUCCESS_BACKGROUND);
}

function loadItemForRow(row, {holdingsRecord = false, instance = false, circulations = false} = {}) {
  const barcode = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(BARCODE), 1, 1).getValue();
  return loadItemForBarcode(barcode, holdingsRecord, instance, circulations);
}
