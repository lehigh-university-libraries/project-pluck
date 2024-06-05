// column headers
const BARCODE = 'barcode';
const EFFECTIVE_CALL_NUMBER = 'effective_call_number';
const TITLE = 'title';
const CONTRIBUTOR = 'contributor';
const PUBLICATION_DATE = 'publication_date';
const ITEM_STATUS = 'item_status';

const MAX_COLUMNS = 100;

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
  const barcode = SpreadsheetApp.getActiveSheet().getRange(row, getColumn(BARCODE), 1, 1).getValue();
  const item = loadItem(barcode);
  item.instance = loadInstance(barcode);
  writeItemToSheet(row, item);
}

function authenticate(config) {
  PropertiesService.getScriptProperties().setProperty("config", JSON.stringify(config));
  config.username = PropertiesService.getScriptProperties().getProperty("username");
  config.password = Utilities.newBlob(Utilities.base64Decode(
      PropertiesService.getScriptProperties().getProperty("password")))
      .getDataAsString();
  FOLIOAUTHLIBRARY.authenticateAndSetHeaders(config);
}

function loadItem(barcode) {
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));
  authenticate(config);

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

function loadInstance(barcode) {
  const config = JSON.parse(PropertiesService.getScriptProperties().getProperty("config"));

  // execute query
  const DQ = encodeURIComponent("\"");
  const instancesQuery = FOLIOAUTHLIBRARY.getBaseOkapi(config.environment) +
    `/search/instances?query=${encodeURIComponent(`items.barcode=="${barcode}"`)}`;
  console.log('Loading instances with query: ', encodeURI(instancesQuery));
  const getOptions = FOLIOAUTHLIBRARY.getHttpGetOptions();
  const instancesResponse = UrlFetchApp.fetch(instancesQuery, getOptions);

  // parse response
  const instancesResponseText = instancesResponse.getContentText();
  const instances = JSON.parse(instancesResponseText)['instances'];
  let instance = instances[0] ?? null;

  return instance;
}

function writeItemToSheet(row, item) {
  writeToSheet(row, getColumn(EFFECTIVE_CALL_NUMBER), item['effectiveShelvingOrder']);
  writeToSheet(row, getColumn(TITLE), item['title']);
  writeToSheet(row, getColumn(CONTRIBUTOR), item['contributorNames']?.[0]?.['name'] ?? '');
  writeToSheet(row, getColumn(PUBLICATION_DATE), item.instance['publication']?.[0]?.['dateOfPublication']);
  writeToSheet(row, getColumn(ITEM_STATUS), item['status']['name']);
}

function writeToSheet(row, column, value) {
  SpreadsheetApp.getActiveSheet().getRange(row, column).setValue(value);
}
