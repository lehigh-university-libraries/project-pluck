function loadItemForBarcode(barcode, holdingsRecord, instance) {
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

function parseOclcNumber(item, stripPrefix = false) {
  const identifiers = item.instance['identifiers'];
  for (let identifier of identifiers) {
    if (OCLC_NUMBER_IDENTIFIER_TYPE_ID == identifier['identifierTypeId']) {
      let oclcNumber = identifier['value'];
      if (stripPrefix) {
        oclcNumber = oclcNumber.replace("(OCoLC)", "");
      }
      return oclcNumber;
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
    return `Error, see log.  At ${(new Date()).toString()}`;
  }
  else {
    console.log(`Got code ${response.getResponseCode()}, response ${JSON.stringify(responseContent)}`);
    return false;
  }
}
