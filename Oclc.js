TOKEN_URL = 'https://oauth.oclc.org/token';
WORLDCATSEARCH_BASE_URL = 'https://americas.discovery.api.oclc.org/worldcat/search/v2';
WORLDCATSEARCH_SCOPES = 'wcapi';

function initOclc() {
  const id = PropertiesService.getScriptProperties().getProperty('oclcId');
  const secret = PropertiesService.getScriptProperties().getProperty('oclcSecret');

  authenticateOclc(id, secret);
}

// https://github.com/googleworkspace/apps-script-oauth2
function authenticateOclc(id, secret) {
  WORLDCAT_SEARCH_SERVICE = OAuth2.createService('oclc_worldcat_search_2')
    .setGrantType('client_credentials')
    .setTokenUrl(TOKEN_URL)
    .setClientId(id)
    .setClientSecret(secret)
    .setScope(WORLDCATSEARCH_SCOPES)
    .setPropertyStore(PropertiesService.getScriptProperties());
}

function enrichFromOclc(item) {
  const oclcNumber = parseOclcNumber(item, true);
  if (!oclcNumber) {
    console.log("Cannot enrich from OCLC, no OCLC num.")
    return;
  }
  const bibsHoldings = loadBibsHoldings(oclcNumber);
  item.oclcBibsHoldings = bibsHoldings;
}

function loadBibsHoldings(oclcNumber) {
  const url = `/bibs-holdings?oclcNumber=${oclcNumber}&limit=50`;
  return queryWorldCatSearchGet(url);
}

function parseOclcHoldings(item) {
  const briefRecords = item.oclcBibsHoldings['briefRecords'];
  if (!briefRecords.length) {
    console.log("no brief records, cannot parse oclc holdings");
    return null;
  }
  return briefRecords[0]['institutionHolding']['totalHoldingCount'];
}

function queryWorldCatSearchGet(url) {
  const query = WORLDCATSEARCH_BASE_URL + url;
  console.log('Executing GET query: ', query);
  const token = WORLDCAT_SEARCH_SERVICE.getAccessToken();
  let response = UrlFetchApp.fetch(query, {
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + token,
    },
  }); 
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
