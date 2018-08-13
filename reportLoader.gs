// constants
var reportsTabMeta = {
  'name': 'Reports',
  'colSavedName': 2,
  'maxCols': 2,
  'maxRows': 5
};

//-----------------------------------------------------------------------------------------------------------------
//-------------------------------------------------- DEFAULT FUNCTIONS --------------------------------------------
//-----------------------------------------------------------------------------------------------------------------

function onInstall(e) {
  // defaults
  var properties = {
    'host': 'https://myserver.com',
    'company':'myCompany',
    'user':'myUser',
    'psw':'secret',
    'apiKey':'myApiKey'
  };
  PropertiesService.getScriptProperties().setProperties(properties, true);
  //

  onOpen(e);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('WFR Report Export')
      .addItem('Export Report List', 'menuExportReportList')
      .addSeparator()
      .addItem('Export Now', 'menuExportNow')
      .addToUi();
}

//-----------------------------------------------------------------------------------------------------------------
//-------------------------------------------------- MENU -- FUNCTIONS --------------------------------------------
//-----------------------------------------------------------------------------------------------------------------

function menuExportNow() {
  var reportIds = getReportsSetup(reportsTabMeta.name);
  //SpreadsheetApp.getUi().alert('You clicked the Export Now! '+reportIds);

  if (reportIds.length == 0) {
    SpreadsheetApp.getUi().alert(
     'Nothing to load. Make sure '+reportsTabMeta.name+' tab exists and contains list of saved reports.');
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Logging in...');
  var token = loginWfr();

  for (var i = 0; i < reportIds.length; i++) {

    var reportSettingId = reportIds[i][0];
    var reportSavedName = reportIds[i][1];
    var format = 'text/csv';
    //var format = 'application/xml';

    SpreadsheetApp.getActiveSpreadsheet().toast('Loading '+reportSavedName+' ...');
    var reportResult = getReport(token, reportSettingId, format);
    var sheet = getSheet(reportSavedName);

    if (reportResult.code == 200) {
      if(reportResult.format == 'application/xml')
      {
        var document = XmlService.parse(reportResult.data);
        var root = document.getRootElement();
        populateReportXML(root, sheet);
      }
      else if (reportResult.format == 'text/csv')
      {
        var csv = parseCsvResponse_(reportResult.data);
        populateReportCSV(csv,sheet);
      }
      else
      {
        populateExceptionMessage('Format '+ reportResult.format + 'not supported');
      }
    }
    else {
      populateExceptionMessage(reportResult.root, sheet);
    }

    fillReportUpdatedDate(reportSavedName);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Done!');
}

function menuExportReportList() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(reportsTabMeta.name);
  if (sheet != null) {
    var ui = SpreadsheetApp.getUi();

    var result = ui.alert(
      'Please confirm',
      'Are you sure you want to continue?\nIn case you proceed report list would be overridden.',
      ui.ButtonSet.YES_NO);

    if (result != ui.Button.YES) {
      return;
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Logging in...');
  var token = loginWfr();
  SpreadsheetApp.getActiveSpreadsheet().toast('Loading report list...');
  var result = getReportList(token);
  var sheet = getSheet(reportsTabMeta.name);

  if (result.code == 200) {
    populateReportListTab(result.root, sheet);
  } else {
    populateExceptionMessage(result.root, sheet);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Done!');
}

//-----------------------------------------------------------------------------------------------------------------
//---------------------------------------------------- SHEET FUNCTIONS --------------------------------------------
//-----------------------------------------------------------------------------------------------------------------

function getReportsSetup(sheetName) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName(sheetName);

  if (sheet == null) {
    return [];
  }

  var reports = [];
  var colSettingId = 1;

  var range = sheet.getRange(1, 1, reportsTabMeta.maxRows, reportsTabMeta.maxCols);
  for (var i = 1; i <= reportsTabMeta.maxRows; i++) {
    var reportSettingId = range.getCell(i, colSettingId).getValue();
    var reportSavedName = range.getCell(i, reportsTabMeta.colSavedName).getValue();

    if (reportSettingId != '' && reportSavedName != '') {
      reports.push([reportSettingId,reportSavedName]);
    }
  }

  return reports;
}

function populateReportListTab(root, sheet) {

  var reports = root.getChildren('report');
  for (var i = 0; i < reports.length; i++) {
    var settingId = reports[i].getChild('SettingId').getText();
    var savedName = reports[i].getChild('SavedName').getText();

    sheet.appendRow([settingId, savedName]);
  }
}

function populateReportXML(root, sheet) {

  var headerRow = [];
  var headers = root.getChild('header').getChildren('col');
  for (var i = 0; i < headers.length; i++) {
    var label = headers[i].getChild('label').getText();
    headerRow.push(label);

  }
  sheet.appendRow(headerRow);

  // rows
  // TODO group processing (not just body > rows, but body > group > body)
  var rows = root.getChild('body').getChildren('row');
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    Logger.log(row);
    var regularRow = [];
    var cols = row.getChildren('col');
    for (var j = 0; j < cols.length; j++) {
      var val = cols[j].getText();

      regularRow.push(val);
    }
    sheet.appendRow(regularRow);

  }
  // TODO footer
}

function populateReportCSV(csvContent,sheet){

  sheet.clearContents().clearFormats();

  // set the values in the sheet (as efficiently as we know how)
  sheet.getRange(
    1, 1,
    csvContent.length, /* rows */
    csvContent[0].length /* columns */).setValues(csvContent);
}

function fillReportUpdatedDate(reportName) {

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(reportsTabMeta.name);

  // extra col for 'updated date'
  var numCols = reportsTabMeta.maxCols + 1;
  var range = sheet.getRange(1, 1, reportsTabMeta.maxRows, numCols);
  for (var i = 1; i <= reportsTabMeta.maxRows; i++) {
    var reportSavedName = range.getCell(i, reportsTabMeta.colSavedName).getValue();

    if (reportSavedName == reportName) {
      range.getCell(i, numCols).setValue(new Date());
    }
  }
}

function getSheet(sheetName) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName(sheetName);

  if (sheet == null) {
    sheet = activeSpreadsheet.insertSheet();
    sheet.setName(sheetName);
  }
  sheet.clear(); //removing prev contents

  return sheet;
}

function populateExceptionMessage(root, sheet) {
  var errors = root.getChildren('error');
  for (var i = 0; i < errors.length; i++) {
    var errorMsg = errors[i].getChild('message').getText();
    sheet.appendRow([errorMsg]);
  }
}

// This will parse a delimited string into an array of
// arrays. The default delimiter is the comma, but this
// can be overriden in the second argument.
function parseCsvResponse_( strData, strDelimiter ){
    // Check to see if the delimiter is defined. If not,
    // then default to comma.
    strDelimiter = (strDelimiter || ",");
    strData = strData.replace(/^\s+|\s+$/g, '');

    // Create a regular expression to parse the CSV values.
    var objPattern = new RegExp(
        (
            // Delimiters.
            "(\\" + strDelimiter + "|\\r?\\n|\\r|^)" +

            // Quoted fields.
            "(?:\"([^\"]*(?:\"\"[^\"]*)*)\"|" +

            // Standard fields.
            "([^\"\\" + strDelimiter + "\\r\\n]*))"
        ),
        "gi"
        );


    // Create an array to hold our data. Give the array
    // a default empty first row.
    var arrData = [[]];

    // Create an array to hold our individual pattern
    // matching groups.
    var arrMatches = null;


    // Keep looping over the regular expression matches
    // until we can no longer find a match.
    while (arrMatches = objPattern.exec( strData )){

        // Get the delimiter that was found.
        var strMatchedDelimiter = arrMatches[ 1 ], strMatchedValue;

        // Check to see if the given delimiter has a length
        // (is not the start of string) and if it matches
        // field delimiter. If id does not, then we know
        // that this delimiter is a row delimiter.
        if (
            strMatchedDelimiter.length &&
            (strMatchedDelimiter !== strDelimiter)
            ){

            // Since we have reached a new row of data,
            // add an empty row to our data array.
            arrData.push( [] );

        }

        // Now that we have our delimiter out of the way,
        // let's check to see which kind of value we
        // captured (quoted or unquoted).
        if (arrMatches[ 2 ]){

            // We found a quoted value. When we capture
            // this value, unescape any double quotes.
            strMatchedValue = arrMatches[ 2 ].replace(
                new RegExp( "\"\"", "g" ),
                "\""
                );

        } else {

            // We found a non-quoted value.
            strMatchedValue = arrMatches[ 3 ];

        }

        // Now that we have our value string, let's add
        // it to the data array.
      if(strMatchedValue === undefined){
        strMatchedValue = '';
      }
        arrData[ arrData.length - 1 ].push( strMatchedValue );

    }

    // Return the parsed data.
    return arrData;
}

//-----------------------------------------------------------------------------------------------------------------
//---------------------------------------- Workforce Ready API Functions ------------------------------------------
//-----------------------------------------------------------------------------------------------------------------

function loginWfr() {
  var props = PropertiesService.getScriptProperties();

  var host = props.getProperty('host');
  var company = props.getProperty('company');
  var user = props.getProperty('user');
  var psw = props.getProperty('psw');
  var apiKey = props.getProperty('apiKey');

  var options = {
    'method' : 'post',
    headers : {
      "Accept" : "application/json",
      "Content-Type" : "application/json;charset=ISO-8859-1",
      "Api-Key" : apiKey
    },
    'payload':'{"credentials": {"company":"'+company+'","username": "'+user+'","password": "'+psw+'"}}'
  };

  var response = UrlFetchApp.fetch(host+"/ta/rest/v1/login?origin=script", options);
  var strResponse = response.getContentText();

  var json = JSON.parse(strResponse);

  return json.token;
}

function getReportList(token) {
  var host = PropertiesService.getScriptProperties().getProperty('host');

  var options = {
    headers : {
      "Authentication" : "bearer "+token,
      "Accept" : "application/xml"
    },
    'muteHttpExceptions' : true
  };

  var response = UrlFetchApp.fetch(host+"/ta/rest/v1/reports?type=saved&origin=script", options);
  var responseCode = response.getResponseCode();
  var strResponse = response.getContentText();

  var document = XmlService.parse(strResponse);
  var root = document.getRootElement();

  return {'code': responseCode, 'root': root};
}

function getReport(token, reportId, format) {
  var host = PropertiesService.getScriptProperties().getProperty('host');
  var options = {
    headers : {
      "Authentication" : "bearer "+token,
      "Accept" : format
    },
    'muteHttpExceptions' : true
  };
  var response = UrlFetchApp.fetch(host+"/ta/rest/v1/report/saved/"+reportId+"?origin=script", options);
  var responseCode = response.getResponseCode();

  var strResponse = response.getContentText();

  return {'code': responseCode, 'data': strResponse, format:format};
}