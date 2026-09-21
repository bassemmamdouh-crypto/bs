const assert = require("assert");
const {
  COLUMN_L_INDEX,
  isBlankCell,
  inspectEmptyColumnRows,
  keepRowsWithFilledColumn,
  alignNewRows,
  appendAlignedRows,
} = require("./google_apps_script/SyncLogic.js");

function row() {
  const cells = Array.from(arguments);
  while (cells.length < 12) {
    cells.push("");
  }
  return cells;
}

assert.strictEqual(COLUMN_L_INDEX, 11);
assert.strictEqual(isBlankCell(""), true);
assert.strictEqual(isBlankCell("  "), true);
assert.strictEqual(isBlankCell("complete"), false);
assert.strictEqual(isBlankCell(0), false);

const header = row(...Array.from({ length: 12 }, (_, i) => "h" + i));
header[11] = "col_l";
const keep = row("keep-a");
keep[11] = "done";
const drop = row("drop-me");

const report = inspectEmptyColumnRows([header, keep, drop]);
assert.strictEqual(report.blank_count, 1);
assert.deepStrictEqual(report.blank_sheet_rows, [3]);
assert.strictEqual(report.keep_count, 1);

const kept = keepRowsWithFilledColumn([header, keep, drop]);
assert.strictEqual(kept.length, 2);
assert.strictEqual(kept[1][0], "keep-a");

const aligned = alignNewRows(
  [["2026-09-21", "Ali", "today"]],
  ["date", "agent", "col_l"],
  ["date", "agent"].concat(Array(9).fill("")).concat(["col_l"])
);
assert.strictEqual(aligned[0][0], "2026-09-21");
assert.strictEqual(aligned[0][11], "today");

const existing = [header, keep];
const appended = appendAlignedRows(existing, ["h0", "col_l"], [["new-day", "today"]]);
assert.strictEqual(appended.length, 3, "append must not drop existing rows");
assert.strictEqual(appended[1][0], "keep-a");
assert.strictEqual(appended[2][0], "new-day");
assert.strictEqual(appended[2][11], "today");

console.log("ok: two-script column L check/remove + Metabase append tests passed");
