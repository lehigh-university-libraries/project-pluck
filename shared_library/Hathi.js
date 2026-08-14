const HATHI_BASE_URL = 'https://catalog.hathitrust.org/api/volumes';

function initHathi() {
  // noop
  // logTime('after Hathi init');
}

function buildVolumesBriefRequest(oclcNumber) {
  return {
    url: HATHI_BASE_URL + `/brief/oclc/${oclcNumber}.json`,
    muteHttpExceptions: true,
  };
}

function enrichBatchFromHathi(items) {
  const requestMap = [];
  for (const item of items) {
    const oclcNumber = parseOclcNumber(item);
    if (!oclcNumber) continue;
    requestMap.push({ item, request: buildVolumesBriefRequest(oclcNumber) });
  }
  if (requestMap.length === 0) return;

  console.log(`Fetching Hathi data for ${requestMap.length} items.`);
  const responses = UrlFetchApp.fetchAll(requestMap.map(r => r.request));
  console.log(`Hathi fetch complete.`);
  for (let i = 0; i < responses.length; i++) {
    const response = responses[i];
    const item = requestMap[i].item;
    const code = response.getResponseCode();
    if (code === 404) {
      console.log(`Hathi: no record found for ${item.barcode}`);
      item.hathi = null;
    } else if (code < 200 || code >= 400) {
      console.error(`Hathi batch error for ${item.barcode}: ${code}`);
      item.hathi = null;
    } else {
      item.hathi = JSON.parse(response.getContentText());
    }
  }
}

function parseHathiEbook(item) {
  const items = item?.hathi?.items ?? [];
  const rightsCodes = items
    .map((item) => item['rightsCode'])
    .filter((value, index, array) => {return array.indexOf(value) == index});
  const rightsCodesString = rightsCodes.join(', ');
  return rightsCodesString;
}

