/**
 * Shared sheet helpers for the two telesales Apps Script entry points.
 */

var TELESALES_CONFIG = {
  METABASE_BASE_URL: "https://bi.marbah.info/api",
  QUESTION_ID: "590",
  SPREADSHEET_ID: "1STzx1zHsztQ1LNAE_0eF9FU0Crfcbes5Q62ZlxaJG6Q",
  WORKSHEET_GID: 292013814,
  WORKSHEET_NAME: "telesales_test",
  CLEANUP_HOUR: 4,
  APPEND_HOUR: 4,
  APPEND_MINUTE: 15,
};

function toast_(message) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Telesales scripts", 8);
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

function padRows_(rows) {
  var width = 0;
  var i;
  for (i = 0; i < rows.length; i++) {
    width = Math.max(width, rows[i].length);
  }
  return rows.map(function (row) {
    var copy = row.slice();
    while (copy.length < width) {
      copy.push("");
    }
    return copy;
  });
}

function replaceSheetRows_(sheet, rows) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 12);
  if (lastRow > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).clearContent();
  }
  if (!rows || !rows.length) {
    return;
  }
  var padded = padRows_(rows);
  sheet.getRange(1, 1, padded.length, padded[0].length).setValues(padded);
}

function appendSheetRows_(sheet, rows) {
  if (!rows || !rows.length) {
    return;
  }
  var padded = padRows_(rows);
  var startRow = sheet.getLastRow() + 1;
  if (startRow < 1) {
    startRow = 1;
  }
  sheet.getRange(startRow, 1, padded.length, padded[0].length).setValues(padded);
}
