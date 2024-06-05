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

const MAX_COLUMNS = 100;

// statistical codes
const RETENTION_IDS = [
  'ba16cd17-fb83-4a14-ab40-23c7ffa5ccb5',
]

const FACULTY_AUTHOR_NOTE_TEXT = "Lehigh Faculty Author Publication";
const LEGACY_CIRC_COUNT_NOTE_TYPE_ID = '8f26b475-d7e3-4577-8bd0-c3d3bf44f73b';
const OCLC_NUMBER_IDENTIFIER_TYPE_ID = '439bfbae-75bc-4f74-9fc7-b2a2d47ce3ef';

function testEdit() {
  onEdit({
    source : SpreadsheetApp.getActiveSpreadsheet(),
    range : SpreadsheetApp.getActiveSpreadsheet().getActiveCell(),
    value : SpreadsheetApp.getActiveSpreadsheet().getActiveCell().getValue(),
  });
}

function onEdit(e) {
  var column = e.range.getColumn();
  if (column == getColumn(BARCODE)) {
    barcodeChanged(e.range.getRow());
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
  initFolio();
  const barcode = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(BARCODE), 1, 1).getValue();
  const item = loadItem(barcode);
  if (!item) {
    console.error("No item matching barcode: " + barcode);
  }
  item.instance = loadInstance(barcode);
  item.holdingsRecord = loadHoldingsRecord(item);
  item.circulations = loadCirculationLogs(item, 'Checked out');
  writeItemToSheet(row, item);
}

function initFolio() {
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));
  authenticate(config);
  LOCATIONS = loadLocations();
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
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));

  // execute query
  const locationsQuery = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) +
    `/locations?limit=100`;
  console.log('Loading locations with query: ', encodeURI(locationsQuery));
  const getOptions = FOLIOAUTHLIBRARY.getHttpGetOptions();
  const locationsResponse = UrlFetchApp.fetch(locationsQuery, getOptions);

  // parse response
  const locationsResponseText = locationsResponse.getContentText();
  const locations = JSON.parse(locationsResponseText)['locations'];
  
  return locations.reduce((map, location) => { map[location.id] = location; return map; }, {} );
}

function loadItem(barcode) {
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));

  // execute query
  const DQ = encodeURIComponent("\"");
  const itemsQuery = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) +
    `/inventory/items?query=barcode==${DQ}${barcode}${DQ}`;
  console.log('Loading item with query: ', encodeURI(itemsQuery));
  const getOptions = FOLIOAUTHLIBRARY.getHttpGetOptions();
  const itemsResponse = UrlFetchApp.fetch(itemsQuery, getOptions);

  // parse response
  const itemsResponseText = itemsResponse.getContentText();
  const items = JSON.parse(itemsResponseText)['items'];
  let item = items[0] ?? null;

  return item;
}

function loadHoldingsRecord(item) {
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));

  // execute query
  const holdingsRecordQuery = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) +
    `/holdings-storage/holdings/${item.holdingsRecordId}`;
  console.log('Loading holdings record with query: ', encodeURI(holdingsRecordQuery));
  const getOptions = FOLIOAUTHLIBRARY.getHttpGetOptions();
  const holdingsRecordResponse = UrlFetchApp.fetch(holdingsRecordQuery, getOptions);

  // parse response
  const holdingsRecordResponseText = holdingsRecordResponse.getContentText();
  const holdingsRecord = JSON.parse(holdingsRecordResponseText);
  
  return holdingsRecord;
}

function loadInstance(barcode) {
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));

  // execute initial query -- to fine the instance ID
  const instancesQuery = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) +
    `/search/instances?query=${encodeURIComponent(`items.barcode=="${barcode}"`)}`;
  console.log('Loading instances with query: ', encodeURI(instancesQuery));
  const getOptions = FOLIOAUTHLIBRARY.getHttpGetOptions();
  const instancesResponse = UrlFetchApp.fetch(instancesQuery, getOptions);

  // parse initial response
  const instancesResponseText = instancesResponse.getContentText();
  const instances = JSON.parse(instancesResponseText)['instances'];
  let instance = instances[0] ?? null;

  // must load the full instance for the full body, including notes
  let instanceQuery = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) +
    `/inventory/instances/${instance.id}`;
  console.log('Loading full instance with query: ', encodeURI(instanceQuery));
  const instanceResponse = UrlFetchApp.fetch(instanceQuery, getOptions);

  // parse second response
  const instanceResponseText = instanceResponse.getContentText();
  instance = JSON.parse(instanceResponseText);

  return instance;
}

function loadCirculationLogs(item, action) {
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));

  // execute query
  const logsQuery = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) +
    `/audit-data/circulation/logs?query=${encodeURIComponent(`(items=="*${item.id}*" and action=="${action}")`)}`;
  console.log('Loading circ logs with query: ', encodeURI(logsQuery));
  const getOptions = FOLIOAUTHLIBRARY.getHttpGetOptions();
  const logsResponse = UrlFetchApp.fetch(logsQuery, getOptions);

  // parse response
  const logsResponseText = logsResponse.getContentText();
  const logs = JSON.parse(logsResponseText);

  return logs;
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
