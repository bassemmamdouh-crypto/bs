/**
 * Menu and daily triggers for the two independent telesales scripts.
 *
 * 1) removeEmptyColumnL — check column L, then delete empty rows
 * 2) appendMetabaseData — extract Metabase 590 and append
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Telesales scripts")
    .addItem("1. Check empty column L", "checkEmptyColumnL")
    .addItem("1. Remove empty column L", "removeEmptyColumnL")
    .addSeparator()
    .addItem("2. Append Metabase 590", "appendMetabaseData")
    .addSeparator()
    .addItem("Install daily triggers", "installDailyTriggers")
    .addItem("Run self-tests", "runSelfTests")
    .addToUi();
}

function installDailyTriggers() {
  var handlers = {
    removeEmptyColumnL: true,
    appendMetabaseData: true,
  };
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (handlers[trigger.getHandlerFunction()]) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("removeEmptyColumnL")
    .timeBased()
    .everyDays(1)
    .atHour(TELESALES_CONFIG.CLEANUP_HOUR)
    .create();

  ScriptApp.newTrigger("appendMetabaseData")
    .timeBased()
    .everyDays(1)
    .atHour(TELESALES_CONFIG.APPEND_HOUR)
    .nearMinute(TELESALES_CONFIG.APPEND_MINUTE)
    .create();

  toast_(
    "Triggers installed: remove empty L at " + TELESALES_CONFIG.CLEANUP_HOUR +
      ":00, append Metabase at " + TELESALES_CONFIG.APPEND_HOUR + ":" +
      String(TELESALES_CONFIG.APPEND_MINUTE).padStart(2, "0") + " Asia/Baghdad."
  );
}

function runSelfTests() {
  var header = [];
  var i;
  for (i = 0; i < 12; i++) {
    header.push("h" + i);
  }
  header[11] = "col_l";
  var keep = header.map(function () { return ""; });
  keep[0] = "keep";
  keep[11] = "filled";
  var drop = header.map(function () { return ""; });
  drop[0] = "drop";

  var report = inspectEmptyColumnRows([header, keep, drop], COLUMN_L_INDEX, true);
  if (report.blank_count !== 1 || report.blank_sheet_rows[0] !== 3) {
    throw new Error("Column L check did not flag the empty row.");
  }
  var kept = keepRowsWithFilledColumn([header, keep, drop], COLUMN_L_INDEX, true);
  if (kept.length !== 2 || kept[1][0] !== "keep") {
    throw new Error("Column L cleanup kept the wrong rows.");
  }

  var appended = appendAlignedRows(kept, ["h0", "col_l"], [["new-day", "today"]]);
  if (appended.length !== 3 || appended[2][0] !== "new-day" || appended[2][11] !== "today") {
    throw new Error("Metabase append did not add aligned rows under existing data.");
  }
  toast_("Self-tests passed.");
}
