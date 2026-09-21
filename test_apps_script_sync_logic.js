const assert = require("assert");
const {
  isBlankCell,
  keepRowsWithColumnP,
  alignNewRows,
  mergeKeptAndNew,
} = require("./google_apps_script/SyncLogic.js");

function row() {
  const cells = Array.from(arguments);
  while (cells.length < 16) {
    cells.push("");
  }
  return cells;
}

assert.strictEqual(isBlankCell(""), true);
assert.strictEqual(isBlankCell("  "), true);
assert.strictEqual(isBlankCell("complete"), false);
assert.strictEqual(isBlankCell(0), false);

const header = row(...Array.from({ length: 16 }, (_, i) => "h" + i));
header[15] = "status";
const keep = row("keep-a");
keep[15] = "done";
const drop = row("drop-me");
const kept = keepRowsWithColumnP([header, keep, drop]);
assert.strictEqual(kept.length, 2);
assert.strictEqual(kept[1][0], "keep-a");

const aligned = alignNewRows(
  [["2026-09-21", "Ali", "today"]],
  ["date", "agent", "status"],
  ["date", "agent"].concat(Array(13).fill("")).concat(["status"])
);
assert.strictEqual(aligned[0][0], "2026-09-21");
assert.strictEqual(aligned[0][15], "today");

const merged = mergeKeptAndNew(
  [header, keep, drop],
  ["h0", "status"],
  [["new-day", "today"]]
);
assert.strictEqual(merged.length, 3);
assert.strictEqual(merged[1][0], "keep-a");
assert.strictEqual(merged[2][0], "new-day");
assert.strictEqual(merged[2][15], "today");

console.log("ok: Apps Script sync logic tests passed");
