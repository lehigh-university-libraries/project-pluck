const HATHI_BASE_URL = 'https://catalog.hathitrust.org/api/volumes';

function initHathi() {
  // noop
  // logTime('after Hathi init');
}

function enrichFromHathi(item) {
  const oclcNumber = parseOclcNumber(item);
  if (!oclcNumber) {
    console.log("Cannot enrich from Hathi, no OCLC num.")
    return;
  }
  const volumes = loadVolumesBrief(oclcNumber);
  item.hathi = volumes;
  // logTime('after Hathi enrichment');
}

function loadVolumesBrief(oclcNumber) {
  const url = `/brief/oclc/${oclcNumber}.json`;
  const volumes = queryHathiGet(url);
  return volumes;
}

function parseHathiEbook(item) {
  const items = item?.hathi?.items ?? [];
  const rightsCodes = items
    .map((item) => item['rightsCode'])
    .filter((value, index, array) => {return array.indexOf(value) == index});
  const rightsCodesString = rightsCodes.join(', ');
  return rightsCodesString;
}

function queryHathiGet(url) {
  // execute query
  const query = HATHI_BASE_URL + url;
  console.log('Executing GET query: ', query);
  const response = UrlFetchApp.fetch(query);

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
