"use strict";
/******************************************************************************/
// Configuration
/******************************************************************************/

// Grab config variables from the Google Sheet
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var USERNAME = sheet.getRange(21, 2).getValue();
var PASSWORD = sheet.getRange(22, 2).getValue();
var CLIENT_ID = sheet.getRange(23, 2).getValue();
var CLIENT_SECRET = sheet.getRange(24, 2).getValue();

// Add Custom Menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("RingRx")
    .addItem("Get Token", "getToken")
    .addItem("Get Billing Info", "getBillingInfo")
    .addItem("Get PBX Callbacks", "getPBXCallbacks")
    .addToUi();
}

/******************************************************************************/
// Helper Functions
/******************************************************************************/

// Get multiple script properties in one call, then log them all.
function logUserProperties() {
  var userProperties = PropertiesService.getUserProperties();
  var data = userProperties.getProperties();
  for (var key in data) {
    Logger.log("Key: %s, Value: %s", key, data[key]);
  }
}

// Communication Monitor
function logUrlFetch(url, opt_params) {
  var params = opt_params || {};
  params.muteHttpExceptions = true;
  var request = UrlFetchApp.getRequest(url, params);
  Logger.log("Request:       >>> " + JSON.stringify(request));
  var response = UrlFetchApp.fetch(url, params);
  Logger.log("Response Code: <<< " + response.getResponseCode());
  Logger.log("Response Text: <<< " + response.getContentText());
  if (response.getResponseCode() >= 400) {
    throw Error("Error in response: " + response);
  }
  return response;
}

/******************************************************************************/
// RingRx Calls
/******************************************************************************/

// Create a token from user credentials
function getToken() {
  var service = getRingRxService_();

  if (service.hasAccess()) {
    Logger.log("App has access.");
    var api =
      "https://portal.ringrx.com/auth/token?username=USERNAME&password=PASSWORD";

    var options = {
      method: "POST",
      muteHttpExceptions: true,
    };

    var response = logUrlFetch(api, options);

    // Parse the JSON reply
    var json = response.getContentText();
    var data = JSON.parse(json);
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.getRange(2, 2).setValue(data["token_type"]);
    sheet.getRange(3, 2).setValue(data["access_token"]);
    sheet.getRange(4, 2).setValue(data["expires_in"]);
    sheet.getRange(5, 2).setValue(data["refresh_token"]);
  } else {
    Logger.log("App has no access yet.");

    // open this url to gain authorization from RingRx
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log(
      "Open the following URL and re-run the script: %s",
      authorizationUrl
    );
  }
}

// Get PBX Callbacks
function getPBXCallbacks() {
  var service = getRingRxService_();

  if (service.hasAccess()) {
    Logger.log("App has access.");
    var api = "https://portal.ringrx.com/pbxcallbacks";

    var headers = {
      Authorization: "Bearer " + sheet.getRange(3, 2).getValue(),
      Accept: "application/json",
    };

    var options = {
      headers: headers,
      method: "GET",
      muteHttpExceptions: true,
    };

    var response = logUrlFetch(api, options);
  } else {
    Logger.log("App has no access yet.");

    // open this url to gain authorization from RingRx
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log(
      "Open the following URL and re-run the script: %s",
      authorizationUrl
    );
  }
}

// Get customer billing information for current account
function getBillingInfo() {
  var service = getRingRxService_();

  if (service.hasAccess()) {
    Logger.log("App has access.");
    var api = "https://portal.ringrx.com/billing";

    var headers = {
      Authorization: "Bearer " + sheet.getRange(3, 2).getValue(), //getRingRxService_().getAccessToken(),
      Accept: "application/json",
    };

    var options = {
      headers: headers,
      method: "GET",
      muteHttpExceptions: true,
    };

    var response = logUrlFetch(api, options);
  } else {
    Logger.log("App has no access yet.");

    // open this url to gain authorization from RingRx
    var authorizationUrl = service.getAuthorizationUrl();
    Logger.log(
      "Open the following URL and re-run the script: %s",
      authorizationUrl
    );
  }
}

/******************************************************************************/
// Oauth2 Library Services
/******************************************************************************/

// Create RingRx service for persisting the authorized token
function getRingRxService_() {
  return (
    OAuth2.createService("RingRx")

      // Set the endpoint URLs, which are the same for all Google services.
      //    .setAuthorizationBaseUrl('https://portal.ringrx.com/auth/token')
      //    .setTokenUrl('https://portal.ringrx.com/auth/token')
      .setAuthorizationBaseUrl("https://accounts.google.com/o/oauth2/auth")
      .setTokenUrl("https://accounts.google.com/o/oauth2/token")

      // Set the client ID and secret, from the Google Developers Console.
      .setClientId(CLIENT_ID)
      .setClientSecret(CLIENT_SECRET)

      // Set the name of the callback function to be invoked to complete
      // the OAuth flow.
      .setCallbackFunction("authCallback")

      // Set the property store where authorized tokens should be persisted.
      .setPropertyStore(PropertiesService.getUserProperties())

      // Set the scopes to request (space-separated for Google services).
      .setScope(
        "https://www.googleapis.com/auth/script.external_request https://www.googleapis.com/auth/spreadsheets"
      )
  );
}

// Handle the callback
function authCallback(request) {
  var ringrxService = getRingRxService_();
  var isAuthorized = ringrxService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput("Success! You can close this tab.");
  } else {
    return HtmlService.createHtmlOutput("Denied. You can close this tab");
  }
}

// Reset the authorization state, so that it can be re-tested.
function reset() {
  getRingRxService_().reset();
}

// Logs the redict URI to register in the Google Developers Console.
function logRedirectUri() {
  Logger.log(OAuth2.getRedirectUri());
}
