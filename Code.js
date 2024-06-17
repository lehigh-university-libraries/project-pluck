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
const INSTANCE_UUID = 'instance_uuid';
const INSTANCE_HRID = 'instance_hrid';
const ITEM_EFFECTIVE_LOCATION_NAME = 'item_effective_location_name';
const HOLDINGS_PERMANENT_LOCATION_NAME = 'holdings_permanent_location_name';
const MATERIAL_TYPE = 'material_type';
const DECISION = 'Decision';
const DECISION_NOTE = 'Decision Note';

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
  INSTANCE_UUID,
  INSTANCE_HRID,
  ITEM_EFFECTIVE_LOCATION_NAME,
  HOLDINGS_PERMANENT_LOCATION_NAME,
  MATERIAL_TYPE,
  DECISION,
  DECISION_NOTE,
];

const MAX_COLUMNS = 100;

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
  // testProcessWithdraw();
}

function testInitSheetForLocation() {
  initSheetForLocation({
    'environment': 'test',
    'location_id': '460df2a6-6146-4749-9ff0-a0d0730e0214',
    'start_row': 0,
    'row_count' : 10,
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

function testProcessWithdraw() {
  initFolio();
  processWithdraw(SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getRow());
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
  initFolio();
  writeHeaders();

  let locationId = config.location_id;
  let offset = parseInt(config.start_row);
  let count = parseInt(config.row_count);
  const items = loadItems(locationId, offset, count);
  let row = SpreadsheetApp.getActiveSheet().getLastRow();
  for (const item of items) {
    row++;
    enrichItem(item, true, true);
    writeItemToSheet(row, item);
    initDecision(row);
  }
}

function writeHeaders() {
  SpreadsheetApp.getActiveSheet().getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
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
  const headers = SpreadsheetApp.getActiveSheet().getRange(1, 1, 1, MAX_COLUMNS).getValues()[0];
  for (var i=0; i < headers.length; i++) {
    if (text == headers[i]) {
      return i + 1;
    }
  }
  return null;
}

function barcodeChanged(row) {
  console.log("barcode changed in row " + row);
  const item = loadItemForRow(row, {holdingsRecord: true, instance: true});
  item.circulations = loadCirculationLogs(item, 'Checked out');
  writeItemToSheet(row, item);
  initDecision(row);
}

function writeItemToSheet(row, item) {
  writeToSheet(row, getColumn(EFFECTIVE_CALL_NUMBER), item['effectiveShelvingOrder']);
  writeToSheet(row, getColumn(TITLE), item['title']);
  writeToSheet(row, getColumn(CONTRIBUTOR), item['contributorNames']?.[0]?.['name']);
  writeToSheet(row, getColumn(PUBLICATION_DATE), item.instance['publication']?.[0]?.['dateOfPublication']);
  writeToSheet(row, getColumn(ITEM_STATUS), item['status']['name']);
  writeToSheet(row, getColumn(RETENTION), hasRetentionAgreement(item));
  writeToSheet(row, getColumn(FACULTY_AUTHOR), isFacultyAuthor(item));
  writeToSheet(row, getColumn(LEGACY_CIRC_COUNT), parseLegacyCircCount(item));
  writeToSheet(row, getColumn(FOLIO_CIRC_COUNT), parseFolioCircCount(item));
  writeToSheet(row, getColumn(OCLC_NUMBER), parseOclcNumber(item));
  writeToSheet(row, getColumn(INSTANCE_UUID), item.instance.id);
  writeToSheet(row, getColumn(INSTANCE_HRID), item.instance.hrid);
  writeToSheet(row, getColumn(ITEM_EFFECTIVE_LOCATION_NAME), item['effectiveLocation']?.['name']);
  writeToSheet(row, getColumn(HOLDINGS_PERMANENT_LOCATION_NAME), parseLocation(item.holdingsRecord['permanentLocationId']));
  writeToSheet(row, getColumn(MATERIAL_TYPE), item['materialType']?.['name']);
}

function writeToSheet(row, column, value) {
  SpreadsheetApp.getActiveSheet().getRange(row, column).setValue(value);
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

  putItem(item);
}

function processWithdraw(row) {
  console.log("processing withdrawal for row " + row);
  const item = loadItemForRow(row, {holdingsRecord: true, instance: true});

  item['status']['name'] = 'Withdrawn';
  item['discoverySuppress'] = true;
  putItem(item);

  if (!hasUnsuppressedItems(item.holdingsRecord, item)) {
    console.log("suppress holdings record");
    item.holdingsRecord['discoverySuppress'] = true;
    putHoldingsRecord(item.holdingsRecord);

    if (!hasUnsuppressedHoldingsRecords(item.instance, item.holdingsRecord)) {
      console.log("suppress instance");
      item.instance['discoverySuppress'] = true;
      item.instance['statusId'] = INSTANCE_STATUS_WITHDRAWN_ID;
      putInstance(item.instance);
    }
  }
}

function loadItemForRow(row, {holdingsRecord = false, instance = false} = {}) {
  const barcode = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(BARCODE), 1, 1).getValue();
  const item = loadItem(barcode);
  if (!item) {
    console.error("No item matching barcode: " + barcode);
    return null;
  }
  enrichItem(item, holdingsRecord, instance);
  return item;
}

function enrichItem(item, holdingsRecord, instance) {
  if (holdingsRecord) {
    item.holdingsRecord = loadHoldingsRecord(item);
  }
  if (instance) {
    item.instance = loadInstance(item);
  }
  item.circulations = loadCirculationLogs(item, 'Checked out');
}

function initFolio(config = null) { 
  if (config) {
    PropertiesService.getScriptProperties().setProperty("config", JSON.stringify(config));
  }
  config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));
  authenticate(config);
  LOCATIONS = loadLocations();
  DECISION_CODE_TO_ID = loadStatisticalCodes();
  DECISION_NOTE_TYPE_ID = loadDecisionNoteTypeId();
  INSTANCE_STATUS_WITHDRAWN_ID = loadInstanceStatusWithdrawnId();
}

function authenticate(config) {
  PropertiesService.getScriptProperties().setProperty("config", JSON.stringify(config));
  config.username = PropertiesService.getScriptProperties().getProperty("username");
  config.password = Utilities.newBlob(Utilities.base64Decode(
      PropertiesService.getScriptProperties().getProperty("password")))
      .getDataAsString();
  FOLIOAUTHLIBRARY.authenticateAndSetHeaders(config);
}

function loadLocations() {
  const url = `/locations?limit=100`;
  const locations = queryFolioGet(url)['locations'];
  return locations.reduce((map, location) => { map[location.id] = location; return map; }, {} );
}

function loadStatisticalCodes() {
  const url = `/statistical-codes?limit=100`;
  const statisticalCodes = queryFolioGet(url)['statisticalCodes'];
  return statisticalCodes.reduce((map, statisticalCode) => { map[statisticalCode.code] = statisticalCode.id; return map; }, {} );
}

function loadDecisionNoteTypeId() {
  const url = `/item-note-types?limit=100&query=${encodeURIComponent(`name=="${DECISION_NOTE_ITEM_TYPE}"`)}`;
  const noteTypes = queryFolioGet(url)['itemNoteTypes'];
  const noteType = noteTypes[0];
  return noteType.id;
}

function loadInstanceStatusWithdrawnId() {
  const url = `/instance-statuses?limit=100&query=${encodeURIComponent(`code=="${INSTANCE_STATUS_WITHDRAWN_CODE}"`)}`;
  const instanceStatuses = queryFolioGet(url)['instanceStatuses'];
  const instanceStatus = instanceStatuses[0];
  return instanceStatus.id;
}

function loadItems(locationId, offset, count) {
  const url = `/inventory/items?query=${encodeURIComponent(`effectiveLocationId=="${locationId}" sortby effectiveCallNumberComponents.callNumber`)}&limit=${count}&offset=${offset}`;
  const items = queryFolioGet(url)['items'];
  return items;
}

function loadItem(barcode) {
  const url = `/inventory/items?query=${encodeURIComponent(`barcode==${barcode}`)}`;
  const items = queryFolioGet(url)['items'];
  const item = items[0] ?? null;
  return item;
}

function loadHoldingsRecord(item) {
  const url = `/holdings-storage/holdings/${item.holdingsRecordId}`;
  const holdingsRecord = queryFolioGet(url);
  return holdingsRecord;
}

function loadInstance(item) {
  const url = `/inventory/instances/${item.holdingsRecord.instanceId}`;
  const instance = queryFolioGet(url);
  return instance;
}

function loadCirculationLogs(item, action) {
  const url = `/audit-data/circulation/logs?query=${encodeURIComponent(`(items=="*${item.id}*" and action=="${action}")`)}`;
  const logs = queryFolioGet(url);
  return logs;
}

function putItem(item) {
  const url = `/inventory/items/${item.id}`;
  return queryFolioPut(url, item);
}

function hasUnsuppressedItems(holdingsRecord, ignoreItem) {
  const url = `/inventory/items-by-holdings-id?query=${encodeURIComponent(`holdingsRecordId==${holdingsRecord.id}`)}`;
  let items = queryFolioGet(url)['items'];
  return hasUnsuppressedRecord(items, ignoreItem);
}

function putHoldingsRecord(holdingsRecord) {
  const url = `/holdings-storage/holdings/${holdingsRecord.id}`;
  return queryFolioPut(url, holdingsRecord);
}

function hasUnsuppressedHoldingsRecords(instance, ignoreHoldingsRecord) {
  const url = `/holdings-storage/holdings?query=${encodeURIComponent(`instanceId==${instance.id}`)}`;
  holdingsRecords = queryFolioGet(url)['holdingsRecords'];
  return hasUnsuppressedRecord(holdingsRecords, ignoreHoldingsRecord);
}

function putInstance(instance) {
  const url = `/inventory/instances/${instance.id}`;
  return queryFolioPut(url, instance);
}

function hasUnsuppressedRecord(recordList, ignoreRecord) {
  const unsuppressedRecords = recordList.filter(
    (record) => (!record['discoverySuppress']) && (record.id != ignoreRecord.id)
  );
  for (record of recordList) {
    if (!record['discoverySuppress']) {
      return true;
    }
  }
  return false;
}

function hasRetentionAgreement(item) {
  const statisticalCodeIds = item['statisticalCodeIds'];
  const retentionIds = statisticalCodeIds.filter(element => RETENTION_IDS.includes(element));
  return retentionIds.length > 0;
}

function isFacultyAuthor(item) {
  const notes = item.instance['notes'];
  for (let note of notes) {
    if (FACULTY_AUTHOR_NOTE_TEXT == note['note']) {
      return true;
    }
  }
  return false;
}

function parseLegacyCircCount(item) {
  const notes = item['notes'];
  for (let note of notes) {
    if (note.itemNoteTypeId == LEGACY_CIRC_COUNT_NOTE_TYPE_ID) {
      return note.note;
    }
  }
  return null;
}

function parseFolioCircCount(item) {
  const circ_count = item.circulations['totalRecords'];
  return circ_count;
}

function parseOclcNumber(item) {
  const identifiers = item.instance['identifiers'];
  for (let identifier of identifiers) {
    if (OCLC_NUMBER_IDENTIFIER_TYPE_ID == identifier['identifierTypeId']) {
      return identifier['value'];
    }
  }
  return null;
}

function parseLocation(locationId) {
  return LOCATIONS[locationId]?.['name'];
}

function queryFolioGet(url) {
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));

  // execute query
  const query = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) + url;
  console.log('Executing GET query: ', query);
  const getOptions = FOLIOAUTHLIBRARY.getHttpGetOptions();
  const response = UrlFetchApp.fetch(query, getOptions);

  // parse response
  const responseText = response.getContentText();
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 400) {
    console.error(`Error response: ${response.getResponseCode()}, ${responseText}`);
    return null;
  }
  else {
    const responseData = JSON.parse(responseText);
    console.log("response data: ", responseData);
    return responseData;
  }  
}

function queryFolioPut(url, payload) {
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));

  // execute query
  const query = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) + url;
  const payloadString = JSON.stringify(payload);
  console.log(`Executing PUT query with url ${url} and payload ${payloadString}`);
  const headers = FOLIOAUTHLIBRARY.getHttpGetHeaders();
  headers['Accept'] = '*/*'
  const options = {
    'method': 'put',
    'contentType': 'application/json',
    'headers': headers,
    'payload': payloadString,
    'muteHttpExceptions': true,
  };
  let response = UrlFetchApp.fetch(query, options);

  // parse response
  let responseContent = response.getContentText()
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 400) {
    console.error(`Error response: ${response.getResponseCode()}, ${responseContent}`);
  }
  else {
    console.log(`Got code ${response.getResponseCode()}, response ${JSON.stringify(responseContent)}`);
  }
}
