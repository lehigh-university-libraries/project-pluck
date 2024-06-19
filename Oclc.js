TOKEN_URL = 'https://oauth.oclc.org/token';
WORLDCATSEARCH_BASE_URL = 'https://americas.discovery.api.oclc.org/worldcat/search/v2';
WORLDCATSEARCH_SCOPES = 'wcapi';

const PALCI_OCLC_SYMBOLS = ['AVL','BEA','BMC','PBU','PBE','CRC','PMC','HHC','PBB','LQS','MAN','XR4','ALL','DKC','DRU','DXU','DUQ','ETS','EAS','ELZ','LFM','PGU','GDC','GBL','HUSAT','HVC','HFC','PZI','PJU','KOL','KZS','LRC','LAS','LAF','VFL','LVC','LYU','LYC','WHV','MRW','QRA','PGM','MVS','CMZ','NJM','MOR','EVI','ZMU','ZYU','UPM','CSC','REC','EIB','PHU','PMN','PTP','ROB','NJG','NJR','PSF','SJD','STH','SQP','SRS','PHA','SUS','PSC','TEU','PCT','PAU','PIT','SRU','URS','PVU','PUG','QWC','WVX','WVU','WFN','UWC','YCP'];
const PALCI_OCLC_SYMBOLS_SET = new Set(PALCI_OCLC_SYMBOLS);

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
  const briefRecords = item.oclcBibsHoldings?.['briefRecords'];
  if (!briefRecords) {
    console.log("no brief records, cannot parse oclc holdings");
    return null;
  }
  return briefRecords?.[0]?.['institutionHolding']?.['totalHoldingCount'];
}

// This is an approximation only, since we only have 50 OCLC results to search through.
function parsePalciHoldings(item) {
  const briefHoldings = item.oclcBibsHoldings?.['briefRecords']?.[0]?.['institutionHolding']?.['briefHoldings'];
  if (!briefHoldings) {
    console.log("no brief holdings, cannot parse PALCI holdings");
    return null;
  }
  let holdingsSymbols = briefHoldings.map((briefHolding) => briefHolding['oclcSymbol']);
  const matches = holdingsSymbols.filter((symbol) => PALCI_OCLC_SYMBOLS_SET.has(symbol));
  return matches.length + '+';
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
