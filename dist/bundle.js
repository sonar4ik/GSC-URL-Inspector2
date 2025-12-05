const SHEET_NAME_QUICK_CHECK = 'Quick Check';
const SHEET_NAME_SETTINGS = 'Settings';
const SHEET_NAME_LOGS = 'Logs';

const QUICK_CHECK_HEADERS = ['URL', 'Result', 'IndexingState', 'CoverageState', 'LastCrawlTime', 'RobotsTxtState', 'PageFetchState', 'MobileUsability', 'CanonicalUrl', 'InspectionTime', 'Error'];
const SETTINGS_HEADERS = ['Key', 'Value'];
const LOGS_HEADERS = ['Timestamp', 'Type', 'Message', 'Details'];

const URL_INSPECTION_ENDPOINT = 'https://searchconsole.googleapis.com/v1/urlInspection/index:inspect';

function executeQuickCheck() {
  runInspection(false);
}

function executeFullInspection() {
  runInspection(true);
}

function runInspection(full) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_QUICK_CHECK);
  if (!sheet) {
    throw new Error('Quick Check sheet not found');
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No URLs found in Quick Check sheet.');
    return;
  }
  var siteUrl = getSetting('siteUrl');
  if (!siteUrl) {
    SpreadsheetApp.getUi().alert('Please set siteUrl in Settings sheet (Key=siteUrl, Value=your Search Console property).');
    throw new Error('Missing siteUrl setting');
  }
  var range = sheet.getRange(2, 1, lastRow - 1, 1);
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    var rowIndex = i + 2;
    var url = values[i][0];
    if (!url) {
      continue;
    }
    var existingResult = sheet.getRange(rowIndex, 2).getValue();
    if (!full && existingResult) {
      continue;
    }
    var result;
    try {
      result = inspectSingleUrl(url, siteUrl);
      writeInspectionResult(sheet, rowIndex, url, result);
      appendLog('INFO', 'Inspection', 'OK for ' + url);
    } catch (e) {
      writeErrorResult(sheet, rowIndex, url, e);
      appendLog('ERROR', 'Inspection', 'Failed for ' + url + ': ' + e.message);
    }
    Utilities.sleep(1100);
  }
}

function inspectSingleUrl(inspectionUrl, siteUrl) {
  var payload = {
    inspectionUrl: inspectionUrl,
    siteUrl: siteUrl,
    languageCode: 'en-US'
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };
  var response = UrlFetchApp.fetch(URL_INSPECTION_ENDPOINT, options);
  var code = response.getResponseCode();
  var text = response.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('HTTP ' + code + ' ' + text);
  }
  var data = JSON.parse(text);
  if (!data.inspectionResult) {
    throw new Error('No inspectionResult in response');
  }
  return data.inspectionResult;
}

function writeInspectionResult(sheet, rowIndex, url, inspectionResult) {
  var indexStatus = inspectionResult.indexStatusResult || {};
  var mobile = inspectionResult.mobileUsabilityResult || {};
  var indexingState = indexStatus.indexingState || '';
  var coverageState = indexStatus.coverageState || '';
  var lastCrawlTime = indexStatus.lastCrawlTime || '';
  var robotsTxtState = indexStatus.robotsTxtState || '';
  var pageFetchState = indexStatus.pageFetchState || '';
  var mobileVerdict = mobile.verdict || '';
  var canonicalUrl = indexStatus.googleCanonical || indexStatus.userCanonical || '';
  var inspectionTime = new Date();
  var resultText = indexingState || coverageState || 'OK';
  var rowValues = [
    resultText,
    indexingState,
    coverageState,
    lastCrawlTime,
    robotsTxtState,
    pageFetchState,
    mobileVerdict,
    canonicalUrl,
    inspectionTime,
    ''
  ];
  sheet.getRange(rowIndex, 2, 1, rowValues.length).setValues([rowValues]);
}

function writeErrorResult(sheet, rowIndex, url, error) {
  var existing = sheet.getRange(rowIndex, 1, 1, QUICK_CHECK_HEADERS.length).getValues()[0];
  var rowValues = existing;
  rowValues[1] = 'ERROR';
  rowValues[2] = '';
  rowValues[3] = '';
  rowValues[4] = '';
  rowValues[5] = '';
  rowValues[6] = '';
  rowValues[7] = '';
  rowValues[8] = '';
  rowValues[9] = new Date();
  rowValues[10] = error.message;
  sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
}

function getSetting(key) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_SETTINGS);
  if (!sheet) {
    return null;
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }
  var range = sheet.getRange(2, 1, lastRow - 1, 2);
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === key) {
      return values[i][1];
    }
  }
  return null;
}

function appendLog(type, message, details) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!sheet) {
    return;
  }
  var row = [new Date(), type, message, details || ''];
  sheet.appendRow(row);
}
