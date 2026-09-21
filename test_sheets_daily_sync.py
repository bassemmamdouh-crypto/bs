import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

import pandas as pd

from sheets_daily_sync import (
    align_new_rows,
    is_blank,
    keep_rows_with_column_p,
    merge_kept_and_new,
    run_sync,
)


def _row(*values, width=16):
    cells = list(values)
    if len(cells) < width:
        cells.extend([""] * (width - len(cells)))
    return cells


class IsBlankTests(unittest.TestCase):
    def test_none_and_empty(self):
        self.assertTrue(is_blank(None))
        self.assertTrue(is_blank(""))
        self.assertTrue(is_blank("   "))
        self.assertTrue(is_blank(float("nan")))

    def test_null_tokens(self):
        self.assertTrue(is_blank("NULL"))
        self.assertTrue(is_blank("#N/A"))
        self.assertTrue(is_blank("none"))

    def test_real_values(self):
        self.assertFalse(is_blank(0))
        self.assertFalse(is_blank("2026-09-21"))
        self.assertFalse(is_blank(" ليز "))


class KeepRowsTests(unittest.TestCase):
    def test_keeps_header_and_filled_p(self):
        header = _row(*[f"col_{i}" for i in range(16)])
        header[15] = "P"
        filled = _row("keep-a")
        filled[15] = "done"
        blank = _row("drop-me")
        blank[15] = ""
        short = ["only", "a", "few", "cells"]
        whitespace = _row("also-drop")
        whitespace[15] = "  "

        kept = keep_rows_with_column_p([header, filled, blank, short, whitespace])
        self.assertEqual(kept[0][15], "P")
        self.assertEqual(len(kept), 2)
        self.assertEqual(kept[1][0], "keep-a")

    def test_no_header_filters_every_row(self):
        filled = _row("keep")
        filled[15] = "x"
        blank = _row("drop")
        kept = keep_rows_with_column_p([filled, blank], has_header=False)
        self.assertEqual(len(kept), 1)
        self.assertEqual(kept[0][0], "keep")


class AlignNewRowsTests(unittest.TestCase):
    def test_matches_header_names_case_insensitively(self):
        frame = pd.DataFrame(
            {
                "Date": ["2026-09-21"],
                "Agent": ["Ali"],
                "Extra": ["ignored-unless-in-header"],
            }
        )
        header = ["date", "agent"] + [""] * 13 + ["status"]
        rows = align_new_rows(frame, header)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0][0], "2026-09-21")
        self.assertEqual(rows[0][1], "Ali")
        self.assertEqual(rows[0][15], "")

    def test_positional_fallback_when_names_do_not_overlap(self):
        frame = pd.DataFrame([["a", "b", "c"]])
        header = ["one", "two", "three"]
        rows = align_new_rows(frame, header)
        self.assertEqual(rows, [["a", "b", "c"]])


class MergeTests(unittest.TestCase):
    def test_drops_blank_p_then_appends(self):
        header = _row(*[f"h{i}" for i in range(16)])
        keep = _row("old-keep")
        keep[15] = "present"
        drop = _row("old-drop")
        new = pd.DataFrame(
            {
                "h0": ["new-day"],
                "h15": ["today"],
            }
        )
        # Name the 16th column so it maps onto header h15.
        new = new.rename(columns={"h15": "h15"})
        header[15] = "h15"
        merged = merge_kept_and_new([header, keep, drop], new)
        self.assertEqual(len(merged), 3)
        self.assertEqual(merged[1][0], "old-keep")
        self.assertEqual(merged[2][0], "new-day")
        self.assertEqual(merged[2][15], "today")

    def test_empty_sheet_uses_dataframe_header(self):
        new = pd.DataFrame({"date": ["2026-09-21"], "qty": [3]})
        merged = merge_kept_and_new([], new)
        self.assertEqual(merged[0], ["date", "qty"])
        self.assertEqual(merged[1][0], "2026-09-21")
        self.assertEqual(merged[1][1], 3)


class LocalSyncTests(unittest.TestCase):
    def test_local_csv_round_trip(self):
        header = [f"col_{i}" for i in range(16)]
        header[15] = "status"
        keep = _row("keep-row")
        keep[15] = "ok"
        drop = _row("drop-row")
        sheet = pd.DataFrame([keep, drop], columns=header)
        incoming = pd.DataFrame(
            {
                "col_0": ["fresh"],
                "status": ["new"],
            }
        )

        with TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            sheet_path = tmp_path / "sheet.csv"
            new_path = tmp_path / "new.csv"
            out_path = tmp_path / "out.csv"
            sheet.to_csv(sheet_path, index=False)
            incoming.to_csv(new_path, index=False)

            summary = run_sync(
                question_id=None,
                spreadsheet_id=None,
                worksheet_name=None,
                local_input=str(sheet_path),
                local_new=str(new_path),
                local_output=str(out_path),
            )
            result = pd.read_csv(out_path, dtype=object).fillna("")
            self.assertEqual(summary["removed_blank_p_rows"], 1)
            self.assertEqual(summary["kept_data_rows"], 1)
            self.assertEqual(summary["appended_rows"], 1)
            self.assertEqual(list(result["col_0"]), ["keep-row", "fresh"])
            self.assertEqual(list(result["status"]), ["ok", "new"])


if __name__ == "__main__":
    unittest.main()
