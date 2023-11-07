//Config
var config = {
  clientId: "YOUR_CLIENT_ID",
  clientSecret: "YOUR_CLIENT_SECRET"
}
// Main 
function updateData() {

  var accessToken = getAccessToken();
  var header = ['01. AccountName', '02. AccountID', '03. Currency', '04. Spend', '05. Value', '06. Orders', '07. ROAS', '08. Cost/Pur.', '09. Impr.', '10. View Content', '11. Cost/VC', '12. ATC', '13.CPM', '14. LP View', '15. Leads', '16. Cost/Lead', '17. post_engagement', '18. link_click', '19. outbound_clicks', '20. cost_outbound_clicks', '21. Date Start', '22. Date Stop'];
  
  var last_month = [header];
  var yesterday = [header];
  var last_90d = [header];
  var last_30d = [header];
  var last_14d = [header];
  var last_7d = [header];
  var last_3d = [header];

  var accounts = getAccounts(accessToken);
  // Loop through accounts
  for (var i = 0; i < accounts.length; i++) {
    var item = accounts[i];
    if (item.id) {
      var accountDataLastMonth = getData(accessToken, item.id, "last_month");
      var accountDataYesterday = getData(accessToken, item.id, "yesterday");
      var accountData90 = getData(accessToken, item.id, "last_90d");
      var accountData30 = getData(accessToken, item.id, "last_30d");
      var accountData14 = getData(accessToken, item.id, "last_14d");
      var accountData7 = getData(accessToken, item.id, "last_7d");
      var accountData3 = getData(accessToken, item.id, "last_3d");

      function pushIfValidData(data, targetArray) {
        if (data !== undefined && data !== null && data.length > 0) {
          targetArray.push(data);
        }
      }
      pushIfValidData(accountDataLastMonth, last_month);
      pushIfValidData(accountDataYesterday, yesterday);
      pushIfValidData(accountData90, last_90d);
      pushIfValidData(accountData30, last_30d);
      pushIfValidData(accountData14, last_14d);
      pushIfValidData(accountData7, last_7d);
      pushIfValidData(accountData3, last_3d);
    }
  }

  pushToSheet(last_month,"LAST_MONTH");
  pushToSheet(yesterday,"YESTERDAY");
  pushToSheet(last_90d,"LAST_90D");
  pushToSheet(last_30d,"LAST_30D");
  pushToSheet(last_14d,"LAST_14D");
  pushToSheet(last_7d,"LAST_7D");
  pushToSheet(last_3d,"LAST_3D");
}

// Get Account Data
function getData(accessToken, entityId, datePreset){
  var data = [];
  var url = 'https://graph.facebook.com/v18.0/'
    + entityId
    + '/insights?'
    + 'date_preset=' + datePreset
    + '&fields=account_name,account_id,account_currency,spend,cpm,impressions,purchase_roas,actions,action_values,cost_per_action_type,cost_per_outbound_click,outbound_clicks'
    + '&access_token=' + accessToken;

  var response = UrlFetchApp.fetch(url);
  if (response.getResponseCode() == 200) {
    var json = response.getContentText();
    var jsonData = JSON.parse(json);
    var dataArray = jsonData.data;
    var accountData = dataArray[0];

    if (accountData){
      function getValue(obj, property, defaultValue = "-") {
        return obj && obj[property] ? obj[property] : defaultValue;
      }

      function getArrValue(arr, actionType) {
        if (arr) {
          const item = arr.find(item => item.action_type === actionType);
          return item ? item.value : "-";
        }
        return "-";
      }
      var properties = [
        'account_name',
        'account_id',
        'account_currency',
        'spend',
        'action_values|omni_purchase',
        'actions|omni_purchase',
        'purchase_roas|omni_purchase',
        'cost_per_action_type|omni_purchase',
        'impressions',
        'actions|omni_view_content',
        'cost_per_action_type|omni_view_content',
        'actions|omni_add_to_cart',
        'cpm',
        'actions|landing_page_view',
        'actions|lead',
        'cost_per_action_type|lead',
        'actions|post_engagement',
        'actions|link_click',
        'outbound_clicks|outbound_click',
        'cost_per_outbound_click|outbound_click',
        'date_start',
        'date_stop'
      ];

      var data = properties.map(function (property) {
        const [field, actionType] = property.split('|');
        if (actionType) {
          return getArrValue(accountData[field], actionType);
        } else {
          return getValue(accountData, field);
        }
      });
    }
  } 
  if (data !== undefined && data !== null){
    return data;
  }
}

// Get All Accounts
function getAccounts(accessToken){
  var url = 'https://graph.facebook.com/v18.0/me/adaccounts?fields=name,id'
    + '&access_token=' + accessToken;
  var response = UrlFetchApp.fetch(url);
  if (response.getResponseCode() == 200) {
    var json = response.getContentText();
    var data = JSON.parse(json);
    var accounts = data.data; 
  }  
  return accounts;
}

// Get Acces Token
function getAccessToken() {
  var sessionToken = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Token").getRange("A1").getValue();
  
  if (sessionToken) {
    var longLivedToken = getLongLivedToken(sessionToken); // Function to obtain the long-lived token using the session token
    if (longLivedToken) {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Token").getRange("A1").setValue(longLivedToken);
      return longLivedToken;
    }
  }
}

// Get Long Lived Token
function getLongLivedToken(sessionToken) {
  var exchangeToken = sessionToken;
  
  var url = "https://graph.facebook.com/v18.0/oauth/access_token" +
            "?grant_type=fb_exchange_token" +
            "&client_id=" + config.clientId +
            "&client_secret=" + config.clientSecret +
            "&fb_exchange_token=" + exchangeToken;

  var response = UrlFetchApp.fetch(url);

  if (response.getResponseCode() == 200) {
    var json = response.getContentText();
    var data = JSON.parse(json);

    if (data && data.access_token) {
      return data.access_token;
    } else {
      return "Access token not found in JSON response.";
    }
  } else {
    return "Error: " + response.getResponseCode() + " - " + response.getContentText();
  }
}

// Push Data to Spreadsheet
function pushToSheet(dataAll, sheetName) {
  var spread = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spread.getSheetByName(sheetName);
  
  // Calculate the number of rows and columns in the data
  var numRows = dataAll.length;
  var numCols = dataAll[0].length;
  // Clear only the data range, starting from the second row
  sheet.getRange(1, 1, sheet.getLastRow(), numCols).clearContent();
  // Set the values for the entire data range
  sheet.getRange(1, 1, numRows, numCols).setValues(dataAll);
}

