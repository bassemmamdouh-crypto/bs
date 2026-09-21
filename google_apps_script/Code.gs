/**
 * Bound Apps Script for the telesales daily refresh.
 *
 * Install: Extensions > Apps Script, paste this file plus SyncLogic.js,
 * then Project Settings > Script properties for METABASE_USERNAME / METABASE_PASSWORD.
 * Run installDailyTrigger once, or use the "Metabase sync" menu.
 */

var TELESALES_CONFIG = {
  METABASE_BASE_URL: "https://bi.marbah.info/api",
  QUESTION_ID: "590",
  SPREADSHEET_ID: "1STzx1zHsztQ1LNAE_0eF9FU0Crfcbes5Q62ZlxaJG6Q",
  WORKSHEET_GID: 292013814,
  WORKSHEET_NAME: "telesales_test",
  TRIGGER_HOUR: 4, // 04:00 in the script timezone (Asia/Baghdad)
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Metabase sync")
    .addItem("Run sync now", "dailySync")
    .addItem("Install daily trigger (04:00)", "installDailyTrigger")
    .addItem("Run self-tests", "runSelfTests")
    .addToUi();
}

function dailySync() {
  var sheet = getTargetSheet_();
  var sheetRows = readSheetRows_(sheet);
  var metabase = fetchMetabaseQuestion_(TELESALES_CONFIG.QUESTION_ID);
  var merged = mergeKeptAndNew(sheetRows, metabase.header, metabase.body);
  writeSheetRows_(sheet, merged);

  var originalData = Math.max(sheetRows.length - (sheetRows.length ? 1 : 0), 0);
  var keptData = Math.max(keepRowsWithColumnP(sheetRows, true).length - 1, 0);
  var summary = {
    original_data_rows: originalData,
    removed_blank_p_rows: originalData - keptData,
    kept_data_rows: keptData,
    appended_rows: metabase.body.length,
    final_data_rows: Math.max(merged.length - 1, 0),
  };
  Logger.log(JSON.stringify(summary));
  toast_(
    "Removed " + summary.removed_blank_p_rows +
      " blank-P rows, appended " + summary.appended_rows + " from Metabase 590."
  );
  return summary;
}

function installDailyTrigger() {
  var handler = "dailySync";
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyDays(1)
    .atHour(TELESALES_CONFIG.TRIGGER_HOUR)
    .create();
  toast_("Daily trigger installed for " + TELESALES_CONFIG.TRIGGER_HOUR + ":00 Asia/Baghdad.");
}

function runSelfTests() {
  var header = [];
  for (var i = 0; i < 16; i++) {
    header.push("h" + i);
  }
  header[15] = "status";
  var keep = header.map(function () { return ""; });
  keep[0] = "keep";
  keep[15] = "done";
  var drop = header.map(function () { return ""; });
  drop[0] = "drop";
  var merged = mergeKeptAndNew(
    [header, keep, drop],
    ["h0", "status"],
    [["new-day", "today"]]
  );
  if (merged.length !== 3) {
    throw new Error("Expected 3 rows after merge, got " + merged.length);
  }
  if (merged[1][0] !== "keep" || merged[2][0] !== "new-day" || merged[2][15] !== "today") {
    throw new Error("Merge did not drop blank column P and append new rows.");
  }
  toast_("Self-tests passed.");
}

function toast_(message) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Metabase sync", 8);
  } catch (error) {
    Logger.log(message);
  }
}

function getTargetSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    ss = SpreadsheetApp.openById(TELESALES_CONFIG.SPREADSHEET_ID);
  }
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() === TELESALES_CONFIG.WORKSHEET_GID) {
      return sheets[i];
    }
  }
  var named = ss.getSheetByName(TELESALES_CONFIG.WORKSHEET_NAME) ||
    ss.getSheetByName("." + TELESALES_CONFIG.WORKSHEET_NAME);
  if (!named) {
    throw new Error(
      "Could not find tab telesales_test (gid " + TELESALES_CONFIG.WORKSHEET_GID + ")."
    );
  }
  return named;
}

function readSheetRows_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) {
    return [];
  }
  return sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
}

function writeSheetRows_(sheet, rows) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 16);
  if (lastRow > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).clearContent();
  }
  if (!rows || !rows.length) {
    return;
  }
  var width = 0;
  for (var i = 0; i < rows.length; i++) {
    width = Math.max(width, rows[i].length);
  }
  var padded = rows.map(function (row) {
    var copy = row.slice();
    while (copy.length < width) {
      copy.push("");
    }
    return copy;
  });
  sheet.getRange(1, 1, padded.length, width).setValues(padded);
}

function fetchMetabaseQuestion_(questionId) {
  var props = PropertiesService.getScriptProperties();
  var username = props.getProperty("METABASE_USERNAME");
  var password = props.getProperty("METABASE_PASSWORD");
  if (!username || !password) {
    throw new Error(
      "Set Script properties METABASE_USERNAME and METABASE_PASSWORD " +
        "(Project Settings in the Apps Script editor)."
    );
  }

  var sessionResponse = UrlFetchApp.fetch(TELESALES_CONFIG.METABASE_BASE_URL + "/session", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ username: username, password: password }),
    muteHttpExceptions: true,
    followRedirects: true,
  });
  if (sessionResponse.getResponseCode() < 200 || sessionResponse.getResponseCode() >= 300) {
    throw new Error(
      "Metabase session failed (" + sessionResponse.getResponseCode() + "): " +
        sessionResponse.getContentText().substring(0, 300)
    );
  }
  var sessionJson = JSON.parse(sessionResponse.getContentText());
  if (!sessionJson || !sessionJson.id) {
    throw new Error("Metabase session response is missing id.");
  }

  var csvResponse = UrlFetchApp.fetch(
    TELESALES_CONFIG.METABASE_BASE_URL + "/card/" + questionId + "/query/csv",
    {
      method: "post",
      contentType: "application/json",
      headers: { "X-Metabase-Session": sessionJson.id },
      payload: JSON.stringify({ parameters: [] }),
      muteHttpExceptions: true,
      followRedirects: true,
    }
  );
  if (csvResponse.getResponseCode() < 200 || csvResponse.getResponseCode() >= 300) {
    throw new Error(
      "Metabase CSV query failed (" + csvResponse.getResponseCode() + "): " +
        csvResponse.getContentText().substring(0, 300)
    );
  }

  var csvText = csvResponse.getContentText();
  if (csvText.charCodeAt(0) === 0xfeff) {
    csvText = csvText.substring(1);
  }
  var parsed = Utilities.parseCsv(csvText);
  if (!parsed || !parsed.length) {
    return { header: [], body: [] };
  }
  return { header: parsed[0], body: parsed.slice(1) };
}
