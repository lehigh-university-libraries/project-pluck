// column header string constants
const BARCODE = 'Barcode';
const EFFECTIVE_CALL_NUMBER = 'Effective Call Number';
const TITLE = 'Title';
const CONTRIBUTOR = 'Contributor';
const PUBLICATION_DATE = 'Publication Date';
const ITEM_STATUS = 'Item Status';
const RETENTION = 'EAST Retention';
const INVENTORIED = 'Inventoried';
const FACULTY_AUTHOR = 'Faculty Author';
const LEGACY_CIRC_COUNT = 'OLE Circ Count';
const FOLIO_CIRC_COUNT = 'FOLIO Circ Count';
const DAMAGE = 'Damage';
const OCLC_NUMBER = 'OCLC Number';
const OCLC_HOLDINGS = 'OCLC Holdings';
const PALCI_HOLDINGS = 'PALCI Holdings';
const LVAIC_HOLDINGS = 'LVAIC Holdings';
const HATHI_EBOOK = 'Hathi e-book';
const INSTANCE_UUID = 'Instance UUID';
const INSTANCE_HRID = 'Instance HRID';
const ITEM_EFFECTIVE_LOCATION_NAME = 'Item Effective Location';
const HOLDINGS_PERMANENT_LOCATION_NAME = 'Holdings Permanent Location';
const MATERIAL_TYPE = 'Material Type';
const DECISION = 'Decision';
const DECISION_ADDENDUM = 'Decision Note';
const ADD_DECISION_STATUS = 'Add Decision Status';
const PROCESS_FINAL_STATE_STATUS = 'Process Final State Status';

// data source type constants
const INITIAL_LOAD = 'initialLoad';
const OCLC_SOURCE = 'oclc';
const HATHI_SOURCE = 'hathi';
const DECISION_SOURCE = 'decisions';

const ALL_HEADERS = new Map([
  [BARCODE,                          INITIAL_LOAD],
  [EFFECTIVE_CALL_NUMBER,            INITIAL_LOAD],
  [TITLE,                            INITIAL_LOAD],
  [CONTRIBUTOR,                      INITIAL_LOAD],
  [PUBLICATION_DATE,                 INITIAL_LOAD],
  [ITEM_STATUS,                      INITIAL_LOAD],
  [RETENTION,                        INITIAL_LOAD],
  [INVENTORIED,                      INITIAL_LOAD],
  [FACULTY_AUTHOR,                   INITIAL_LOAD],
  [LEGACY_CIRC_COUNT,                INITIAL_LOAD],
  [FOLIO_CIRC_COUNT,                 INITIAL_LOAD],
  [DAMAGE,                           INITIAL_LOAD],
  [OCLC_NUMBER,                      INITIAL_LOAD],  // comes from FOLIO/metadb, not OCLC API
  [OCLC_HOLDINGS,                    OCLC_SOURCE],
  [PALCI_HOLDINGS,                   OCLC_SOURCE],
  [LVAIC_HOLDINGS,                   OCLC_SOURCE],
  [HATHI_EBOOK,                      HATHI_SOURCE],
  [INSTANCE_UUID,                    INITIAL_LOAD],
  [INSTANCE_HRID,                    INITIAL_LOAD],
  [ITEM_EFFECTIVE_LOCATION_NAME,     INITIAL_LOAD],
  [HOLDINGS_PERMANENT_LOCATION_NAME, INITIAL_LOAD],
  [MATERIAL_TYPE,                    INITIAL_LOAD],
  [DECISION,                         DECISION_SOURCE],
  [DECISION_ADDENDUM,                DECISION_SOURCE],
  [ADD_DECISION_STATUS,              DECISION_SOURCE],
  [PROCESS_FINAL_STATE_STATUS,       DECISION_SOURCE],
]);

const LOCKED_COLUMNS = [BARCODE, EFFECTIVE_CALL_NUMBER, DECISION, DECISION_ADDENDUM, ADD_DECISION_STATUS, PROCESS_FINAL_STATE_STATUS];

// Active headers for the current operation — set from sheet metadata at the start of each operation
let headers = [...ALL_HEADERS.keys()];

function getColumn(text) {
  const index = headers.findIndex(h => h === text);
  return index < 0 ? null : index + 1;
}

function getColumnLetter(text) {
  const col = getColumn(text);
  return col ? String.fromCharCode(64 + col) : null;
}

let writeBuffer;
function initWriteToRow() {
  writeBuffer = Array(headers.length).fill('');
}
function writeToRow(column, value) {
  if (column === null) return;
  writeBuffer[column - 1] = value;
}
function commitWriteToRow(sheet, row) {
  sheet.getRange(row, 1, 1, headers.length).setValues([writeBuffer]);
}

function writeItemToSheet(sheet, row, item) {
  initWriteToRow();
  writeToRow(getColumn(BARCODE), item.barcode);
  writeToRow(getColumn(EFFECTIVE_CALL_NUMBER), item['effective_call_number']);
  writeToRow(getColumn(TITLE), item.title);
  writeToRow(getColumn(CONTRIBUTOR), item.contributor);
  writeToRow(getColumn(PUBLICATION_DATE), item.publication_date);
  writeToRow(getColumn(ITEM_STATUS), item['item_status']);
  writeToRow(getColumn(RETENTION), hasRetentionAgreement(item));
  writeToRow(getColumn(INVENTORIED), isInventoried(item));
  writeToRow(getColumn(FACULTY_AUTHOR), isFacultyAuthor(item));
  writeToRow(getColumn(LEGACY_CIRC_COUNT), parseLegacyCircCount(item));
  writeToRow(getColumn(FOLIO_CIRC_COUNT), parseFolioCircCount(item));
  writeToRow(getColumn(DAMAGE), parseDamage(item));
  writeToRow(getColumn(OCLC_NUMBER), item.oclc_number);
  writeToRow(getColumn(OCLC_HOLDINGS), parseOclcHoldings(item));
  writeToRow(getColumn(PALCI_HOLDINGS), parsePalciHoldings(item));
  writeToRow(getColumn(LVAIC_HOLDINGS), parseLvaicHoldings(item));
  writeToRow(getColumn(HATHI_EBOOK), parseHathiEbook(item));
  writeToRow(getColumn(INSTANCE_UUID), item.instance_uuid);
  writeToRow(getColumn(INSTANCE_HRID), item.instance_hrid);
  writeToRow(getColumn(ITEM_EFFECTIVE_LOCATION_NAME), item['item_effective_location_name']);
  writeToRow(getColumn(HOLDINGS_PERMANENT_LOCATION_NAME), item['holdings_permanent_location_name']);
  writeToRow(getColumn(MATERIAL_TYPE), item['material_type']);
  commitWriteToRow(sheet, row);
}

function writeHeaders(sheet) {
  if (getSheetMetadata(sheet, 'headers')) return;
  const maxRows = sheet.getMaxRows();
  if (maxRows > 2) sheet.deleteRows(3, maxRows - 2);
  const activeHeaders = getActiveHeaders();
  setSheetMetadata(sheet, 'headers', JSON.stringify(activeHeaders));
  headers = activeHeaders;
  sheet.getRange(1, 1, 1, activeHeaders.length).setValues([activeHeaders]);
  sheet.setFrozenRows(1);
  // Text barcode -- allow leading zeroes
  let column = getColumnLetter(BARCODE);
  sheet.getRange(`${column}1:${column}`).setNumberFormat("@");
  for (const countCol of [PALCI_HOLDINGS, LVAIC_HOLDINGS]) {
    column = getColumnLetter(countCol);
    if (column) sheet.getRange(`${column}1:${column}`).setHorizontalAlignment("right");
  }
}

function getSheetHeaders(sheet) {
  const stored = getSheetMetadata(sheet, 'headers');
  return stored ? JSON.parse(stored) : getActiveHeaders();
}

function getColumnPreferences() {
  const defaults = Object.fromEntries([...ALL_HEADERS.keys()].map(h => [h, true]));
  const stored = properties.getProperty('column_preferences');
  return stored ? { ...defaults, ...JSON.parse(stored) } : defaults;
}

function saveColumnPreferences(prefs) {
  properties.setProperty('column_preferences', JSON.stringify(prefs));
}

function getActiveHeaders() {
  const prefs = getColumnPreferences();
  return [...ALL_HEADERS.keys()].filter(h => prefs[h] !== false);
}

function getColumnPreferencesData() {
  const prefs = getColumnPreferences();
  return [...ALL_HEADERS.entries()].map(([name, source]) => ({
    name,
    source,
    locked: LOCKED_COLUMNS.includes(name),
    enabled: prefs[name] !== false,
  }));
}

function showColumnPreferences() {
  const html = HtmlService.createHtmlOutputFromFile('columnPreferences')
    .setWidth(420)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select columns to load on new tabs');
}

function writeTabName(sheet, locationId, startCallNumberPrefix, endCallNumberPrefix) {
  const code = LOCATIONS[locationId]?.['code'];
  const name = (getLoadingMode() === 'metadb')
    ? `${startCallNumberPrefix} - ${endCallNumberPrefix} | ${code}`
    : code;
  sheet.setName(name);
  setSheetMetadata(sheet, 'location_id', locationId);
  setSheetMetadata(sheet, 'start_call_number_prefix', startCallNumberPrefix ?? '');
  setSheetMetadata(sheet, 'end_call_number_prefix', endCallNumberPrefix ?? '');
}
