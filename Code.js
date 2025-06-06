// column headers
const BARCODE = 'barcode';
const EFFECTIVE_CALL_NUMBER = 'effective_call_number';
const TITLE = 'title';
const CONTRIBUTOR = 'contributor';
const PUBLICATION_DATE = 'publication_date';
const ITEM_STATUS = 'item_status';
const RETENTION = 'EAST Retention';
const FACULTY_AUTHOR = 'Faculty Author';
const LEGACY_CIRC_COUNT = 'ole_circ_count';
const FOLIO_CIRC_COUNT = 'folio_circ_count';
const OCLC_NUMBER = 'oclc_number';
const OCLC_HOLDINGS = 'OCLC Holdings';
const PALCI_HOLDINGS = 'PALCI Holdings';
const HATHI_EBOOK = 'Hathi e-book';
const INSTANCE_UUID = 'instance_uuid';
const INSTANCE_HRID = 'instance_hrid';
const ITEM_EFFECTIVE_LOCATION_NAME = 'item_effective_location_name';
const HOLDINGS_PERMANENT_LOCATION_NAME = 'holdings_permanent_location_name';
const MATERIAL_TYPE = 'material_type';
const DECISION = 'Decision';
const DECISION_ADDENDUM = 'Decision Note';
const ADD_DECISION_STATUS = 'Add Decision Status';
const PROCESS_FINAL_STATE_STATUS = 'Process Final State Status';

const HEADERS = [
  BARCODE,
  EFFECTIVE_CALL_NUMBER,
  TITLE,
  CONTRIBUTOR,
  PUBLICATION_DATE,
  ITEM_STATUS,
  RETENTION,
  FACULTY_AUTHOR,
  LEGACY_CIRC_COUNT,
  FOLIO_CIRC_COUNT,
  OCLC_NUMBER,
  OCLC_HOLDINGS,
  PALCI_HOLDINGS,
  HATHI_EBOOK,
  INSTANCE_UUID,
  INSTANCE_HRID,
  ITEM_EFFECTIVE_LOCATION_NAME,
  HOLDINGS_PERMANENT_LOCATION_NAME,
  MATERIAL_TYPE,
  DECISION,
  DECISION_ADDENDUM,
  ADD_DECISION_STATUS,
  PROCESS_FINAL_STATE_STATUS,
];

const MAX_COLUMNS = HEADERS.length;

const DEFAULT_COUNT = 50;
const FLUSH_RATE = 5;
const PAUSE_TIME = 5000;
const TAB_COMPLETE_COLOR = 'green';

// Decisions
const WITHDRAW = 'withdrawn';
const REMOTE = 'move to remote storage';
const SC = 'move to special collections';
const MISSING = 'missing';
const NO_CHANGE = 'no change';
const DECISIONS = [ WITHDRAW, REMOTE, SC, MISSING, NO_CHANGE] ;
const DECISIONS_RULE = SpreadsheetApp.newDataValidation().requireValueInList(DECISIONS).build();

// Final States
const FINAL_STATE_KEEP = 'decision-keep-2024';
const FINAL_STATE_WITHDRAW = 'decision-withdraw-2024';

const DECISION_TO_FINAL_STATE = new Map([
  [WITHDRAW, FINAL_STATE_WITHDRAW],
  [REMOTE, FINAL_STATE_KEEP],
  [SC, FINAL_STATE_KEEP],
  [MISSING, FINAL_STATE_WITHDRAW],
  [NO_CHANGE, FINAL_STATE_KEEP],
]);

// status columns
const ADD_SUCCESS_MESSAGE = 'Added Note';
const FINAL_STATE_SUCCESS_MESSAGE = 'Final State Processed';
const SUCCESS_BACKGROUND = 'lightgreen';
const FAILURE_BACKGROUND = 'lightcoral';

// Retention Statistical Codes
const RETENTION_IDS = [
  'ba16cd17-fb83-4a14-ab40-23c7ffa5ccb5',
]

// Decision Note
const DECISION_NOTE_ITEM_TYPE = 'Project Pluck Decision';

const MISSING_CHECK_IN_NOTE_TYPE = 'Check in';
const MISSING_CHECK_IN_NOTE_TEXT = 'Withdrawn.  Route to Cataloging.';

const FACULTY_AUTHOR_NOTE_TEXT = "Lehigh Faculty Author Publication";
const LEGACY_CIRC_COUNT_NOTE_TYPE_ID = '8f26b475-d7e3-4577-8bd0-c3d3bf44f73b';
const OCLC_NUMBER_IDENTIFIER_TYPE_ID = '439bfbae-75bc-4f74-9fc7-b2a2d47ce3ef';

const INSTANCE_STATUS_WITHDRAWN_CODE = 'Withdrawn';

var DECISION_CODE_TO_ID;
var LOCATIONS;

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

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Project Pluck')
    .addItem('Show Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  clearCache();
  var html = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Project Pluck')
    .setWidth(500);
  SpreadsheetApp.getUi()
    .showSidebar(html);
}

function getLocations(config) {
  initFolio();
  return Object.entries(LOCATIONS).sort((a, b) => {return a['code'] < b['code']});
}

function initSheetForLocation(config) {
  console.log("initSheetForLocation: ", config);
  PropertiesService.getScriptProperties().setProperty("config", JSON.stringify(config));
  PropertiesService.getScriptProperties().setProperty('lastSheetName', SpreadsheetApp.getActiveSheet().getSheetName());

  // logTime("start initSheetForLocation");
  initKillSwitch();
  loadMoreItems();
}

function loadMoreItems() {
  const sheetName = PropertiesService.getScriptProperties().getProperty('lastSheetName');
  try {
    tryLoadMoreItems(sheetName);    
  }
  catch (error) {
    console.log('Error loading items: ', error);
    email(`Error loading items to ${sheetName}`, `${error}`);
  }
}

function tryLoadMoreItems(sheetName) {
  if (killSwitchFlipped()) {
    stopMonitoring();
    return;
  }

  startMonitoring();

  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));

  SpreadsheetApp.getActive().getSheetByName(sheetName).activate();

  initFolio();
  initOclc();
  initHathi();
  writeHeaders();

  let locationId = config.location_id;
  writeTabName(locationId);

  let offset = SpreadsheetApp.getActiveSheet().getLastRow() - 1;
  let count = DEFAULT_COUNT;
  console.log(`writing items to sheet with offset ${offset} and count ${count}`);
  const items = loadItems(locationId, offset, count);
  if (items.length == 0) {
    console.log("Loaded all items for this sheet");
    SpreadsheetApp.getActiveSheet().setTabColor(TAB_COMPLETE_COLOR);
    email(`${sheetName} load complete`, `Google Sheets is done loading the items ${sheetName}.`);
    stopMonitoring();
    return;
  }

  let row = SpreadsheetApp.getActiveSheet().getLastRow();
  for (const item of items) {
    row++;
    // logTime('before enrichment');
    enrichItem(item, true, true, true);
    enrichFromOclc(item);
    enrichFromHathi(item);
    writeItemToSheet(row, item);
    initDecision(row);
    if (row % FLUSH_RATE == 0) {
      SpreadsheetApp.flush();
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
  ScriptApp.newTrigger('loadMoreItems')
    .timeBased()
    .after(PAUSE_TIME)
    .create();
}

function stopLoading() {
  flipKillSwitch();
}

function writeHeaders() {
  SpreadsheetApp.getActiveSheet().getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  SpreadsheetApp.getActiveSheet().setFrozenRows(1);

  let column = getColumnLetter(PALCI_HOLDINGS);
  SpreadsheetApp.getActiveSheet().getRange(`${column}1:${column}`).setHorizontalAlignment("right");
}

function writeTabName(locationId) {
  let code = LOCATIONS[locationId]?.['code'];
  SpreadsheetApp.getActiveSheet().setName(code);
  PropertiesService.getScriptProperties().setProperty('lastSheetName', SpreadsheetApp.getActiveSheet().getSheetName());
}

function getColumn(text) {
  let index = HEADERS.findIndex((element) => element == text);
  if (index < 0) {
    return null;
  }
  return index + 1;
}

function getColumnLetter(text) {
  return String.fromCharCode(64 + getColumn(text))
}

function writeItemToSheet(row, item) {
  initWriteToRow();
  writeToRow(getColumn(BARCODE), item['barcode']);
  writeToRow(getColumn(EFFECTIVE_CALL_NUMBER), item['effectiveCallNumberComponents']?.['callNumber']);
  writeToRow(getColumn(TITLE), item['title']);
  writeToRow(getColumn(CONTRIBUTOR), item['contributorNames']?.[0]?.['name']);
  writeToRow(getColumn(PUBLICATION_DATE), item.instance['publication']?.[0]?.['dateOfPublication']);
  writeToRow(getColumn(ITEM_STATUS), item['status']['name']);
  writeToRow(getColumn(RETENTION), hasRetentionAgreement(item));
  writeToRow(getColumn(FACULTY_AUTHOR), isFacultyAuthor(item));
  writeToRow(getColumn(LEGACY_CIRC_COUNT), parseLegacyCircCount(item));
  writeToRow(getColumn(FOLIO_CIRC_COUNT), parseFolioCircCount(item));
  writeToRow(getColumn(OCLC_NUMBER), parseOclcNumber(item));
  writeToRow(getColumn(OCLC_HOLDINGS), parseOclcHoldings(item));
  writeToRow(getColumn(PALCI_HOLDINGS), parsePalciHoldings(item));
  writeToRow(getColumn(HATHI_EBOOK), parseHathiEbook(item));
  writeToRow(getColumn(INSTANCE_UUID), item.instance.id);
  writeToRow(getColumn(INSTANCE_HRID), item.instance.hrid);
  writeToRow(getColumn(ITEM_EFFECTIVE_LOCATION_NAME), item['effectiveLocation']?.['name']);
  writeToRow(getColumn(HOLDINGS_PERMANENT_LOCATION_NAME), parseLocation(item.holdingsRecord['permanentLocationId']));
  writeToRow(getColumn(MATERIAL_TYPE), item['materialType']?.['name']);
  commitWriteToRow(row);
}

let writeBuffer;
function initWriteToRow() {
  writeBuffer = Array(MAX_COLUMNS).fill('');
}
function writeToRow(column, value) {
  writeBuffer[column - 1] = value;
}
function commitWriteToRow(row) {
  SpreadsheetApp.getActiveSheet().getRange(row, 1, 1, MAX_COLUMNS).setValues([writeBuffer]);
}

function initDecision(row) {
  SpreadsheetApp.getActiveSheet().getRange(row, getColumn(DECISION)).setDataValidation(DECISIONS_RULE);
}

function addDecisions() {
  initFolio();
  processSelectedRows(addDecision);
}

function processFinalStates() {
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
      callback(row);
    }
  }
}

function addDecision(row) {
  console.log("adding decision for row " + row);
  const item = loadItemForRow(row);

  const decision = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(DECISION)).getValue();
  const now = new Date().toString();
  const decisionAddendum = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(DECISION_ADDENDUM)).getValue();
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
