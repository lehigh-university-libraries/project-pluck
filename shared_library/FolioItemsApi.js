// FOLIO REST API item loading — used only in 'folio' loadingMode.
// In 'metadb' mode these functions are not called; see loadItemsMetadb in Folio.js.

function loadItemsFolio(locationId, offset, count) {
  const statusNamesString = ITEM_STATUSES.map((status) => `"${status}"`).join(' OR ');
  const url = `/inventory/items?query=${encodeURIComponent(`effectiveLocationId=="${locationId}" AND (instance.discoverySuppress=="false") AND (status.name = (${statusNamesString})) sortby effectiveCallNumberComponents.callNumber`)}&limit=${count}&offset=${offset}`;
  return queryFolioGet(url)['items'];
}

function enrichItem(item, holdingsRecord, instance, circulations) {
  if (holdingsRecord) {
    item.holdingsRecord = loadHoldingsRecord(item);
  }
  if (instance) {
    item.instance = loadInstance(item);
  }
  if (circulations) {
    item.circulations = loadCirculationLogs(item, 'Checked out');
  }
}

function loadHoldingsRecord(item) {
  const url = `/holdings-storage/holdings/${item.holdingsRecordId}`;
  return queryFolioGet(url);
}

function loadInstance(item) {
  const url = `/inventory/instances/${item.holdingsRecord.instanceId}`;
  return queryFolioGet(url);
}

function loadCirculationLogs(item, action) {
  const url = `/audit-data/circulation/logs?query=${encodeURIComponent(`(items=="*${item.id}*" and action=="${action}")`)}`;
  return queryFolioGet(url);
}

// Normalize a FOLIO API item's nested fields into the flat schema used by MetaDB,
// so that writeItemToSheet and all helper functions work identically in both modes.
function normalizeFolioItem(item) {
  item.effective_call_number = item.effectiveCallNumberComponents?.callNumber;
  item.contributor = item.contributorNames?.[0]?.name;
  item.publication_date = item.instance?.publication?.[0]?.dateOfPublication;
  item.item_status = item.status?.name;
  item.statistical_codes = (item.statisticalCodeIds || [])
    .map(id => STATISTICAL_CODE_BY_ID[id] ?? id)
    .join('; ');
  item.item_notes = JSON.stringify((item.notes || []).map(n => ({
    type: ITEM_NOTE_TYPE_BY_ID[n.itemNoteTypeId] ?? n.itemNoteTypeId,
    note: n.note,
  })));
  item.faculty_author = parseFacultyAuthorFolio(item);
  item.legacy_circ_count = parseLegacyCircCountFolio(item);
  item.folio_circ_count = item.circulations?.totalRecords;
  item.oclc_number = parseOclcNumberFolio(item);
  item.instance_uuid = item.instance?.id;
  item.instance_hrid = item.instance?.hrid;
  item.item_effective_location_name = item.effectiveLocation?.name;
  item.holdings_permanent_location_name = parseLocation(item.holdingsRecord?.permanentLocationId);
  item.material_type = item.materialType?.name;
}

function parseFacultyAuthorFolio(item) {
  const notes = item.instance?.['notes'] ?? [];
  for (const note of notes) {
    if (FACULTY_AUTHOR_NOTE_TEXT == note['note']) {
      return true;
    }
  }
  return false;
}

function parseLegacyCircCountFolio(item) {
  return parseLegacyCircCount(item);
}

function parseOclcNumberFolio(item) {
  const identifiers = item.instance?.['identifiers'] ?? [];
  for (const identifier of identifiers) {
    if (OCLC_NUMBER_IDENTIFIER_TYPE_ID == identifier['identifierTypeId']) {
      let oclcNumber = identifier['value'].trim();
      oclcNumber = oclcNumber.replace("(OCoLC)", "");
      oclcNumber = oclcNumber.replace("ocn", "");
      return oclcNumber.trim();
    }
  }
  return null;
}
