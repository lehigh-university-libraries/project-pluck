// START OF CONFIGURATION

const FOLIO_API_URL = {   // OKAPI ENDPOINTS
  'prod': 'https://lehigh-okapi.folio.indexdata.com',
  'test': 'https://lehigh-test-okapi.folio.indexdata.com'
 };

const TENANT = "lu";

// END OF CONFIGURATION

function getBaseOkapi(environment) {
  return FOLIO_API_URL[environment];
}

function login(config) {
  console.log("FolioAuth login called");
  if ('username' in config && 'password' in config) {
    console.log("using supplied username [" + config.username + "] and password");
    const token = authenticateWithCredentials(FOLIO_API_URL[config.environment], config.username, config.password);
    properties.setProperty("folioToken", token);
  } else {
    throw new Error("Config must contain username and password attributes.");
  }
}

function getHttpGetHeaders() {
  return {
    "Accept": "application/json",
    "x-okapi-tenant": TENANT,
    "x-okapi-token": properties.getProperty("folioToken"),
  };
}

function getHttpGetOptions() {
  return {
    'headers': getHttpGetHeaders(),
    'muteHttpExceptions': true,
  };
}

function getHttpPostHeaders() {
  return {
    "Accept": "application/json",
    "Content-type": "application/json",
    "x-okapi-tenant": TENANT,
    "x-okapi-token": properties.getProperty("folioToken"),
  };
}

function authenticateWithCredentials(baseOkapi, username, password) {
  const headers = {
    "Accept": "application/json,text/plain",
    "x-okapi-tenant": TENANT,
  };
  const data = {
    'tenant': TENANT,
    'username': username,
    'password': password,
  };
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'headers': headers,
    'muteHttpExceptions': false,
    'payload': JSON.stringify(data),
  };
  const response = UrlFetchApp.fetch(baseOkapi + '/authn/login', options);
  if (response.getResponseCode() != 201) {
    throw new Error("FOLIO login failed: " + response.getResponseCode());
  }
  return response.getHeaders()['x-okapi-token'];
}
