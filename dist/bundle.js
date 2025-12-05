const URL_INSPECTION_ENDPOINT = 'https://searchconsole.googleapis.com/v1/urlInspection/index:inspect';

function executeQuickCheck(mode) {
  return runInspection('quick', mode || 'new');
}

function executeFullInspection(mode) {
  return runInspection('full', mode || 'all');
}

function checkForUpdates() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Logs');
  if (sheet) {
    sheet.appendRow([new Date(), 'INFO', 'CheckForUpdates', 'No remote update logic implemented yet']);
  }
  return 'Current code version: 1.0.0';
}

function runInspection(type, mode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var quickSheet = ss.getSheetByName('Quick Check');
  var settingsSheet = ss.getSheetByName('Settings');
  var logSheet = ss.getSheetByName('Logs');
  if (!quickSheet) {
    throw new Error('Quick Check sheet not found');
  }
  if (!settingsSheet) {
    throw new Error('Settings sheet not found');
  }
  var settings = getSettingsMap(settingsSheet);
  var siteUrl = settings.siteUrl;
  if (!siteUrl) {
    throw new Error('Settings: siteUrl is empty');
  }
  var maxRequestsPerRun = parseInt(settings.maxRequestsPerRun || '100', 10);
  if (isNaN(maxRequestsPerRun) || maxRequestsPerRun <= 0) {
    maxRequestsPerRun = 100;
  }
  var delayMs = parseInt(settings.delayMs || '1100', 10);
  if (isNaN(delayMs) || delayMs < 0) {
    delayMs = 1100;
  }
  var lastRow = quickSheet.getLastRow();
  if (lastRow < 2) {
    return 'No URLs in Quick Check sheet';
  }
  var range = quickSheet.getRange(2, 1, lastRow - 1, 11);
  var values = range.getValues();
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('stopRequested');
  var processed = 0;
  var inspectedRows = 0;
  for (var i = 0; i < values.length; i++) {
    var rowIndex = i + 2;
    var row = values[i];
    var url = row[0];
    var resultValue = row[1];
    var errorValue = row[10];
    if (!url) {
      continue;
    }
    if (!shouldProcessRowForMode(mode, resultValue, errorValue)) {
      continue;
    }
    var stopFlag = props.getProperty('stopRequested');
    if (stopFlag === '1') {
      appendLog(logSheet, 'INFO', 'Stop', 'Stop requested by user');
      break;
    }
    if (processed >= maxRequestsPerRun) {
      appendLog(logSheet, 'INFO', 'Limit', 'Reached maxRequestsPerRun: ' + maxRequestsPerRun);
      break;
    }
    inspectedRows++;
    try {
      var inspectionResult = callUrlInspectionApi(url, siteUrl);
      writeSuccessRow(quickSheet, rowIndex, inspectionResult);
      appendLog(logSheet, 'INFO', 'Inspection', 'OK ' + url);
    } catch (e) {
      writeErrorRow(quickSheet, rowIndex, e);
      appendLog(logSheet, 'ERROR', 'Inspection', 'Failed ' + url + ' ' + e.message);
    }
    processed++;
    if (delayMs > 0) {
      Utilities.sleep(delayMs);
    }
  }
  if (inspectedRows === 0) {
    return 'No rows matched mode "' + mode + '"';
  }
  return 'Processed ' + processed + ' URLs in mode "' + mode + '"';
}

function shouldProcessRowForMode(mode, resultValue, errorValue) {
  if (mode === 'all') {
    return true;
  }
  if (mode === 'new' || mode === 'empty') {
    return !resultValue;
  }
  if (mode === 'errors') {
    return resultValue === 'ERROR' || !!errorValue;
  }
  return true;
}

function getSettingsMap(sheet) {
  var map = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return map;
  }
  var range = sheet.getRange(2, 1, lastRow - 1, 3);
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    var key = values[i][0];
    var value = values[i][2];
    if (key) {
      map[key] = value;
    }
  }
  return map;
}

function callUrlInspectionApi(inspectionUrl, siteUrl) {
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

function writeSuccessRow(sheet, rowIndex, inspectionResult) {
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

function writeErrorRow(sheet, rowIndex, error) {
  var inspectionTime = new Date();
  var rowValues = [
    'ERROR',
    '',
    '',
    '',
    '',
    '',
    '',
    '',
    inspectionTime,
    error.message || String(error)
  ];
  sheet.getRange(rowIndex, 2, 1, rowValues.length).setValues([rowValues]);
}

function appendLog(logSheet, type, message, details) {
  if (!logSheet) {
    return;
  }
  var row = [new Date(), type, message, details || ''];
  logSheet.appendRow(row);
}
