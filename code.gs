// custom menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
 
  ui.createMenu('Strava App')
    .addItem('Get data', 'getStravaActivityData')
    .addToUi();
}
 
// Get athlete activity data
function getStravaActivityData() {
   
  // get the sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Sheet1');

  // This represents ALL the data
var range = sheet.getDataRange();
var values = range.getValues();
var currentData = [];

// This logs the spreadsheet in CSV format with a trailing comma and adds it to an array of current data in the Sheet
for (var i = 1; i < values.length; i++) {
  //var row = [];
  /*for (var j = 1; j < values[i].length; j++) {
    if (values[i][j]) {
      row = row + values[i][j];
    }
    row = row + ",";
  }*/
  currentData.push(values[i][0]);
  Logger.log(currentData);
}
 
  // call the Strava API to retrieve data
  var data = callStravaAPI();
   
  // empty arrays to hold all activity data and new activity data
  var stravaData = [];
     
  // loop over activity data and add to stravaData array for Sheet
  data.forEach(function(activity) {
    var arr = [];
    arr.push(
      activity.id,
      activity.name,
      activity.type,
      activity.distance,
      activity.total_elevation_gain
    );
    stravaData.push(arr);
    Logger.log(stravaData);
  });
   //Try to compare and only pull out unique values
    var newStravaData = [];
    stravaData.forEach( function(each){
      if (!currentData.includes(each[0])) { newStravaData.push(each);
    }});
    /*currentData.forEach(function(each){
      if (!stravaData.includes(each[0])) { newStravaData.push(each);}
    });*/
    Logger.log(newStravaData);

  // paste the values into the Sheet
  sheet.getRange(sheet.getLastRow() + 1, 1, newStravaData.length, newStravaData[0].length).setValues(newStravaData);
}
 
// call the Strava API
function callStravaAPI() {
   
  // set up the service
  var service = getStravaService();
   
  if (service.hasAccess()) {
    Logger.log('App has access.');
     
    var endpoint = 'https://www.strava.com/api/v3/athlete/activities';
    var params = '?after=1677628800&per_page=200';
 
    var headers = {
      Authorization: 'Bearer ' + service.getAccessToken()
    };
     
    var options = {
      headers: headers,
      method : 'GET',
      muteHttpExceptions: true
    };
     
    var response = JSON.parse(UrlFetchApp.fetch(endpoint + params, options));
     
    return response;  
  }
  else {
    Logger.log("App has no access yet.");
     
    // open this url to gain authorization from github
    var authorizationUrl = service.getAuthorizationUrl();
     
    Logger.log("Open the following URL and re-run the script: %s",
        authorizationUrl);
  }
}