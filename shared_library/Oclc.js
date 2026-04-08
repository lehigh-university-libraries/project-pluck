TOKEN_URL = 'https://oauth.oclc.org/token';
WORLDCATSEARCH_BASE_URL = 'https://americas.discovery.api.oclc.org/worldcat/search/v2';
WORLDCATSEARCH_SCOPES = 'wcapi';

const PALCI_OCLC_SYMBOLS_SET = new Set(PALCI_OCLC_SYMBOLS);
const LVAIC_OCLC_SYMBOLS_SET = new Set(LVAIC_OCLC_SYMBOLS);

function initOclc() {
  const id = properties.getProperty('oclcId');
  const secret = properties.getProperty('oclcSecret');

  authenticateOclc(id, secret);
  // logTime('after OCLC init');
}

// https://github.com/googleworkspace/apps-script-oauth2
function authenticateOclc(id, secret) {
  WORLDCAT_SEARCH_SERVICE = OAuth2.createService('oclc_worldcat_search_2')
    .setGrantType('client_credentials')
    .setTokenUrl(TOKEN_URL)
    .setClientId(id)
    .setClientSecret(secret)
    .setScope(WORLDCATSEARCH_SCOPES)
    .setPropertyStore(properties);
}

function buildBibsHoldingsRequest(oclcNumber, token) {
  return {
    url: WORLDCATSEARCH_BASE_URL + `/bibs-holdings?oclcNumber=${oclcNumber}&limit=50`,
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token },
  };
}

function enrichBatchFromOclc(items) {
  const token = WORLDCAT_SEARCH_SERVICE.getAccessToken();
  const requestMap = [];
  for (const item of items) {
    const oclcNumber = parseOclcNumber(item);
    if (!oclcNumber) continue;
    requestMap.push({ item, request: buildBibsHoldingsRequest(oclcNumber, token) });
  }
  if (requestMap.length === 0) return;

  console.log(`Fetching OCLC data for ${requestMap.length} items.`);
  const responses = UrlFetchApp.fetchAll(requestMap.map(r => r.request));
  console.log(`OCLC fetch complete.`);
  for (let i = 0; i < responses.length; i++) {
    const response = responses[i];
    const item = requestMap[i].item;
    const code = response.getResponseCode();
    if (code < 200 || code >= 400) {
      console.error(`OCLC batch error for ${item.barcode}: ${code}`);
      item.oclcBibsHoldings = null;
    } else {
      item.oclcBibsHoldings = JSON.parse(response.getContentText());
    }
  }
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
function parseConsortiumHoldings(item, name, symbolsSet) {
  const briefHoldings = item.oclcBibsHoldings?.['briefRecords']?.[0]?.['institutionHolding']?.['briefHoldings'];
  if (!briefHoldings) {
    console.log(`no brief holdings, cannot parse ${name} holdings`);
    return null;
  }
  const symbols = briefHoldings.map(h => h['oclcSymbol']);
  return symbols.filter(s => symbolsSet.has(s)).length + '+';
}

function parsePalciHoldings(item) {
  return parseConsortiumHoldings(item, 'PALCI', PALCI_OCLC_SYMBOLS_SET);
}

function parseLvaicHoldings(item) {
  return parseConsortiumHoldings(item, 'LVAIC', LVAIC_OCLC_SYMBOLS_SET);
}

