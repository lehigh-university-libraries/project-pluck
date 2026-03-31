const FOLIO_CACHE_TIME = 60 * 60 * 24 * 7; // 7 days

const ITEM_STATUSES = [
  'Available',
  'In process',
  'In process (non-requestable)',
  'In transit',
  'Missing',
  'Restricted',
  'Unavailable',
];

if (typeof FOLIOAUTHLIBRARY === 'undefined') {
  FOLIOAUTHLIBRARY = {
    getBaseOkapi: getBaseOkapi,
    authenticateAndSetHeaders: authenticateAndSetHeaders,
    getHttpGetOptions: getHttpGetOptions,
    getHttpGetHeaders: getHttpGetHeaders,
  }
}

function getLoadingMode() {
  const mode = properties.getProperty('loadingMode');
  if (mode !== 'folio' && mode !== 'metadb') {
    throw new Error(`loadingMode must be 'folio' or 'metadb', got: ${mode}`);
  }
  return mode;
}

function loadItemForBarcode(barcode, holdingsRecord, instance, circulations) {
  const item = loadItem(barcode);
  if (!item) {
    console.error("No item matching barcode: " + barcode);
    return null;
  }
  if (getLoadingMode() === 'folio') {
    enrichItem(item, holdingsRecord, instance, circulations);
  }
  return item;
}

function initFolio() {
  getOrCreate('authenticate', authenticate, FOLIO_CACHE_TIME);
  LOCATIONS = getOrCreate('loadLocations', loadLocations, FOLIO_CACHE_TIME);
  DECISION_CODE_TO_ID = getOrCreate('loadStatisticalCodes', loadStatisticalCodes, FOLIO_CACHE_TIME);
  DECISION_NOTE_TYPE_ID = getOrCreate('loadDecisionNoteTypeId', loadDecisionNoteTypeId, FOLIO_CACHE_TIME);
  INSTANCE_STATUS_WITHDRAWN_ID = getOrCreate('loadInstanceStatusWithdrawnId', loadInstanceStatusWithdrawnId, FOLIO_CACHE_TIME);
  // logTime('after FOLIO init');
}

function authenticate() {
  const folioConfig = {
    'environment': properties.getProperty("environment"),
    'username': properties.getProperty("username"),
    'password': Utilities.newBlob(Utilities.base64Decode(properties.getProperty("password")))
      .getDataAsString(),
  };
  FOLIOAUTHLIBRARY.authenticateAndSetHeaders(folioConfig);
  return true;
}

function loadLocations() {
  const url = `/locations?limit=1000&query=${encodeURIComponent(`cql.allRecords = 1 sortby code`)}`;
  const locations = queryFolioGet(url)['locations'];

  // identify active sheet
  const activeSheet = SpreadsheetApp.getActiveSheet().getName();
  for (const location of locations) {
    if (location['code'] == activeSheet) {
      location['activeSheet'] = true;
    }
  }

  return locations.reduce((map, location) => { map[location.id] = location; return map; }, {});
}

function loadStatisticalCodes() {
  const url = `/statistical-codes?limit=100`;
  const statisticalCodes = queryFolioGet(url)['statisticalCodes'];
  return statisticalCodes.reduce((map, statisticalCode) => { map[statisticalCode.code] = statisticalCode.id; return map; }, {});
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

function validateCallNumberBoundaries(locationCode, startPrefix, endPrefix) {
  const payload = {
    'url': 'https://raw.githubusercontent.com/lehigh-university-libraries/project-pluck/refs/heads/social-sciences/metadb/validate_call_number_boundaries.sql',
    'params': {
      'location_code': locationCode,
      'start_call_number_prefix': startPrefix,
      'end_call_number_prefix': endPrefix,
    },
    'limit': 1,
  };
  const result = queryFolioPost('/ldp/db/reports', payload);
  return result?.['records']?.[0] ?? { start_found: false, end_found: false };
}

function loadItemsMetadb(locationCode, startCallNumberPrefix, endCallNumberPrefix, offset, count) {
  const payload = {
    'url': 'https://raw.githubusercontent.com/lehigh-university-libraries/project-pluck/refs/heads/social-sciences/metadb/get_items_between_call_number_prefixes.sql',
    'params': {
      'location_code': locationCode,
      'start_call_number_prefix': startCallNumberPrefix,
      'end_call_number_prefix': endCallNumberPrefix,
      'query_limit': String(count),
      'query_offset': String(offset),
    },
    'limit': count,
  };
  return queryFolioPost('/ldp/db/reports', payload)['records'];
}


function loadItem(barcode) {
  const url = `/inventory/items?query=${encodeURIComponent(`barcode==${barcode}`)}`;
  const items = queryFolioGet(url)['items'];
  const item = items[0] ?? null;
  return item;
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
  const holdingsRecords = queryFolioGet(url)['holdingsRecords'];
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
  const codes = JSON.parse(item.statistical_codes || "[]");
  return codes.some(code => RETENTION_IDS.includes(code));
}

function isFacultyAuthor(item) {
  return item.faculty_author;
}

function parseLegacyCircCount(item) {
  return item.legacy_circ_count;
}

function parseFolioCircCount(item) {
  return item.folio_circ_count;
}

function parseOclcNumber(item) {
  return item['oclc_number'];
}

function parseLocation(locationId) {
  return LOCATIONS[locationId]?.['name'];
}

function queryFolioGet(url) {
  // execute query
  const environment = properties.getProperty("environment");
  const query = FOLIOAUTHLIBRARY.getBaseOkapi(environment) + url;
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
    // console.log("response data: ", responseData);
    return responseData;
  }
}

function queryFolioPost(url, payload) {
  const environment = properties.getProperty("environment");
  const query = FOLIOAUTHLIBRARY.getBaseOkapi(environment) + url;
  const payloadString = JSON.stringify(payload);
  console.log(`Executing POST query with url ${url} and payload ${payloadString}`);
  const headers = FOLIOAUTHLIBRARY.getHttpPostHeaders();
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': headers,
    'payload': payloadString,
    'muteHttpExceptions': true,
  };
  const response = UrlFetchApp.fetch(query, options);
  const responseText = response.getContentText();
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 400) {
    console.error(`Error response: ${response.getResponseCode()}, ${responseText}`);
    return null;
  }
  return JSON.parse(responseText);
}

function queryFolioPut(url, payload) {
  // execute query
  const environment = properties.getProperty("environment");
  const query = FOLIOAUTHLIBRARY.getBaseOkapi(environment) + url;
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
    return `Error, see log.  At ${(new Date()).toString()}`;
  }
  else {
    console.log(`Got code ${response.getResponseCode()}, response ${JSON.stringify(responseContent)}`);
    return false;
  }
}
