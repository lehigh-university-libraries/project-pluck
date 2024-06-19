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
const DECISION_NOTE = 'Decision Note';
const SAVE_STATUS = 'Save Status';
const WITHDRAW_STATUS = 'Withdraw Status';

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
  DECISION_NOTE,
  SAVE_STATUS,
  WITHDRAW_STATUS,
];

const MAX_COLUMNS = HEADERS.length;

const DEFAULT_COUNT = 20;

// decisions
const WITHDRAW = 'to be withdrawn';
const REMOTE = 'remote storage';
const SC = 'possible sc material';
const KEEP = 'keep';
const MISSING = 'missing - book is not on shelf';
const DECISIONS = [ WITHDRAW, REMOTE, SC, KEEP, MISSING] ;
const DECISIONS_RULE = SpreadsheetApp.newDataValidation().requireValueInList(DECISIONS).build();

const WITHDRAW_NOTES = [ 'low circ', 'second copy' ];
const REMOTE_NOTES = [ 'fragile' ];
const SC_NOTES = [ 'date', 'subject', 'scarcely held', 'bookplate', 'signed copy', 'accession number' ];
const KEEP_NOTES = [ 'scarcely held', 'seminal work', 'curricular/interest' ];

const RULES = new Map();
RULES.set(WITHDRAW, SpreadsheetApp.newDataValidation().requireValueInList(WITHDRAW_NOTES).build());
RULES.set(REMOTE, SpreadsheetApp.newDataValidation().requireValueInList(REMOTE_NOTES).build());
RULES.set(SC, SpreadsheetApp.newDataValidation().requireValueInList(SC_NOTES).build());
RULES.set(KEEP, SpreadsheetApp.newDataValidation().requireValueInList(KEEP_NOTES).build());

const DECISION_CODES = new Map([
  [WITHDRAW, 'to-withdraw'],
  [REMOTE, 'for-rm'],
  [SC, 'may-be-sc'],
  [KEEP, 'decision-keep'],
]);

// status columns
const SAVE_SUCCESS_MESSAGE = 'Saved';
const WITHDRAW_PENDING_MESSAGE = 'Pending';
const WITHDRAW_SUCCESS_MESSAGE = 'Withdrawn';
const SUCCESS_BACKGROUND = 'lightgreen';
const FAILURE_BACKGROUND = 'lightcoral';

const RETENTION_IDS = [
  'ba16cd17-fb83-4a14-ab40-23c7ffa5ccb5',
]

const DECISION_NOTE_ITEM_TYPE = 'Project Pluck Decision';

const FACULTY_AUTHOR_NOTE_TEXT = "Lehigh Faculty Author Publication";
const LEGACY_CIRC_COUNT_NOTE_TYPE_ID = '8f26b475-d7e3-4577-8bd0-c3d3bf44f73b';
const OCLC_NUMBER_IDENTIFIER_TYPE_ID = '439bfbae-75bc-4f74-9fc7-b2a2d47ce3ef';

const INSTANCE_STATUS_WITHDRAWN_CODE = 'Withdrawn';

var DECISION_CODE_TO_ID;
var LOCATIONS;

function test() {
  // testInitSheetForLocation();
  // testEdit();
  // testProcessDecision();
  // testProcessWithdrawal();
}

function testInitSheetForLocation() {
  initSheetForLocation({
    'environment': 'test',
    'location_id': '460df2a6-6146-4749-9ff0-a0d0730e0214',
  });
}

function testEdit() {
  initFolio();
  onEdit({
    source : SpreadsheetApp.getActiveSpreadsheet(),
    range : SpreadsheetApp.getActiveSpreadsheet().getActiveCell(),
    value : SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue(),
  });
}

function testProcessDecision() {
  initFolio();
  processDecision(SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getRow());
}

function testProcessWithdrawal() {
  initFolio();
  processWithdrawal(SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getRow());
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Project Pluck')
    .addItem('Show Sidebar', 'showSidebar')
    .addToUi();
}

function showSidebar() {
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
  // logTime("start initSheetForLocation");
  initFolio();
  initOclc();
  initHathi();
  writeHeaders();

  let locationId = config.location_id;
  writeTabName(locationId);

  let offset = config.start_row ? parseInt(config.start_row) : SpreadsheetApp.getActiveSheet().getLastRow() - 1;
  let count = config.row_count ? parseInt(config.row_count) : DEFAULT_COUNT;
  console.log(`writing items to sheet with offset ${offset} and count ${count}`);
  const items = loadItems(locationId, offset, count);
  let row = SpreadsheetApp.getActiveSheet().getLastRow();
  for (const item of items) {
    row++;
    // logTime('before enrichment');
    enrichItem(item, true, true);
    enrichFromOclc(item);
    enrichFromHathi(item);
    writeItemToSheet(row, item);
    initDecision(row);
    if (row % 10 == 1) {
      SpreadsheetApp.flush();
    }
  }
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
}

function onEdit(e) {
  var column = e.range.getColumn();
  if (column == getColumn(BARCODE)) {
    barcodeChanged(e.range.getRow());
  }
  else if (column == getColumn(DECISION)) {
    decisionChanged(e.range.getRow());
  }
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

function barcodeChanged(row) {
  console.log("barcode changed in row " + row);
  const item = loadItemForRow(row, {holdingsRecord: true, instance: true});
  item.circulations = loadCirculationLogs(item, 'Checked out');
  writeItemToSheet(row, item);
  initDecision(row);
}

function writeItemToSheet(row, item) {
  initWriteToRow();
  writeToRow(getColumn(BARCODE), item['barcode']);
  writeToRow(getColumn(EFFECTIVE_CALL_NUMBER), item['effectiveShelvingOrder']);
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

function decisionChanged(row) {
  const noteCell = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(DECISION_NOTE));
  noteCell.clear({contentsOnly: true, validationsOnly: true});
  const decision = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(DECISION)).getValue();
  const rule = RULES.get(decision);
  if (rule) {
    noteCell.setDataValidation(rule);
  }
}

function processDecisions() {
  initFolio();
  processSelectedRows(processDecision);
}

function processWithdrawals() {
  initFolio();
  processSelectedRows(processWithdrawal);
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

function processDecision(row) {
  console.log("processing decision for row " + row);
  const item = loadItemForRow(row);

  const decision = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(DECISION)).getValue();
  const statisticalCode = DECISION_CODES.get(decision);
  const statisticalCodeId = DECISION_CODE_TO_ID[statisticalCode];
  if (statisticalCodeId) {
    item['statisticalCodeIds'].push(statisticalCodeId);
  }

  const decisionNote = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(DECISION_NOTE)).getValue();
  if (decisionNote) {
    item['notes'].push({
      itemNoteTypeId: DECISION_NOTE_TYPE_ID,
      note: decisionNote,
      staffOnly: true,
    });
  }

  const error = putItem(item);
  const saveStatusCell = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(SAVE_STATUS));
  if (error) {
    saveStatusCell.setValue(error);
    saveStatusCell.setBackground(FAILURE_BACKGROUND);
  }
  else {
    saveStatusCell.setValue(SAVE_SUCCESS_MESSAGE);
    saveStatusCell.setBackground(SUCCESS_BACKGROUND);

    if (WITHDRAW == decision) {
      const withdrawalStatusCell = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(WITHDRAW_STATUS));
      withdrawalStatusCell.setValue(WITHDRAW_PENDING_MESSAGE);
    }
  }
}

function processWithdrawal(row) {
  console.log("processing withdrawal for row " + row);
  const item = loadItemForRow(row, {holdingsRecord: true, instance: true});

  item['status']['name'] = 'Withdrawn';
  item['discoverySuppress'] = true;
  let error = putItem(item);
  const withdrawalStatusCell = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(WITHDRAW_STATUS));
  if (error) {
    withdrawalStatusCell.setValue('Error withdrawing item: ' + error);
    withdrawalStatusCell.setBackground(FAILURE_BACKGROUND);
    return;
  }

  if (!hasUnsuppressedItems(item.holdingsRecord, item)) {
    console.log("suppress holdings record");
    item.holdingsRecord['discoverySuppress'] = true;
    error = putHoldingsRecord(item.holdingsRecord);
    if (error) {
      withdrawalStatusCell.setValue('Error withdrawing holdings record: ' + error);
      withdrawalStatusCell.setBackground(FAILURE_BACKGROUND);
      return;
    }
  
    if (!hasUnsuppressedHoldingsRecords(item.instance, item.holdingsRecord)) {
      console.log("suppress instance");
      item.instance['discoverySuppress'] = true;
      item.instance['statusId'] = INSTANCE_STATUS_WITHDRAWN_ID;
      error = putInstance(item.instance);
      if (error) {
        withdrawalStatusCell.setValue('Error withdrawing instance: ' + error);
        withdrawalStatusCell.setBackground(FAILURE_BACKGROUND);
        return;
      }
    }
  }

  withdrawalStatusCell.setValue(WITHDRAW_SUCCESS_MESSAGE);
  withdrawalStatusCell.setBackground(SUCCESS_BACKGROUND);
}

function loadItemForRow(row, {holdingsRecord = false, instance = false} = {}) {
  const barcode = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(BARCODE), 1, 1).getValue();
  return loadItemForBarcode(barcode, holdingsRecord, instance);
}
