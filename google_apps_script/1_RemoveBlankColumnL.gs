/**
 * Script 1 — check telesales_test, then delete rows whose column L is empty.
 *
 * Run checkEmptyColumnL to inspect only.
 * Run removeEmptyColumnL to inspect, log the blank rows, then delete them.
 * The header row is never deleted.
 */

function checkEmptyColumnL() {
  var sheet = getTargetSheet_();
  var rows = readSheetRows_(sheet);
  var report = inspectEmptyColumnRows(rows, COLUMN_L_INDEX, true);
  Logger.log(JSON.stringify(report));
  toast_(
    "Column L check: " + report.blank_count + " empty row(s) of " +
      report.data_rows + " data rows. Row numbers: " +
      (report.blank_sheet_rows.join(", ") || "none")
  );
  return report;
}

function removeEmptyColumnL() {
  var sheet = getTargetSheet_();
  var rows = readSheetRows_(sheet);
  var report = inspectEmptyColumnRows(rows, COLUMN_L_INDEX, true);
  Logger.log("Column L check before delete: " + JSON.stringify(report));

  if (!report.blank_count) {
    toast_("Column L check: no empty rows to remove.");
    return report;
  }

  var kept = keepRowsWithFilledColumn(rows, COLUMN_L_INDEX, true);
  replaceSheetRows_(sheet, kept);

  report.removed_count = report.blank_count;
  report.kept_data_rows = report.keep_count;
  Logger.log("Column L cleanup done: " + JSON.stringify(report));
  toast_(
    "Removed " + report.removed_count + " row(s) with empty column L. Kept " +
      report.keep_count + " data row(s)."
  );
  return report;
}
