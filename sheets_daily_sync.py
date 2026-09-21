#!/usr/bin/env python3
"""Daily Google Sheet refresh from a Metabase question.

Each run:
  1. Drops existing data rows whose column P (16th column) is blank.
  2. Appends the latest Metabase question results as the new day's rows.
  3. Leaves the header row and any rows with a value in column P in place.

Configure via environment variables or CLI flags. See README.md.
"""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
from typing import Any, Sequence

import pandas as pd

from metabase_client import ret_metabase

logger = logging.getLogger(__name__)

# Column P is the 16th column (A=1).
COLUMN_P_INDEX = 15

# Target sheet / Metabase card for the telesales daily refresh.
DEFAULT_SPREADSHEET_ID = "1STzx1zHsztQ1LNAE_0eF9FU0Crfcbes5Q62ZlxaJG6Q"
DEFAULT_WORKSHEET_NAME = "telesales_test"
DEFAULT_WORKSHEET_GID = 292013814
DEFAULT_QUESTION_ID = "590"


def coalesce_config(*values: Any, default: str = "") -> str:
    """Return the first non-empty value, treating blank env vars as unset."""
    for value in values:
        if value is None:
            continue
        text = str(value).strip()
        if text:
            return text
    return default


def is_blank(value: Any) -> bool:
    """True when a cell should be treated as empty for the column-P rule."""
    if value is None:
        return True
    if isinstance(value, float) and pd.isna(value):
        return True
    text = str(value).strip()
    if not text:
        return True
    return text.lower() in {"nan", "none", "null", "#n/a"}


def cell_at(row: Sequence[Any], index: int) -> Any:
    if index < 0 or index >= len(row):
        return ""
    return row[index]


def normalize_header(value: Any) -> str:
    text = "" if value is None else str(value)
    return " ".join(text.replace("\u200f", "").replace("\u200e", "").strip().lower().split())


def keep_rows_with_column_p(
    rows: list[list[Any]],
    *,
    has_header: bool = True,
    column_index: int = COLUMN_P_INDEX,
) -> list[list[Any]]:
    """Keep the header (optional) plus rows whose column P is not blank."""
    if not rows:
        return []
    if has_header:
        header = list(rows[0])
        body = rows[1:]
        kept = [list(row) for row in body if not is_blank(cell_at(row, column_index))]
        return [header] + kept
    return [list(row) for row in rows if not is_blank(cell_at(row, column_index))]


def _stringify_cell(value: Any) -> Any:
    if is_blank(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    if hasattr(value, "item") and not isinstance(value, (bytes, str)):
        try:
            value = value.item()
        except (ValueError, AttributeError):
            pass
    return value


def align_new_rows(frame: pd.DataFrame, header: list[Any]) -> list[list[Any]]:
    """Turn Metabase results into sheet rows aligned to the existing header.

    Prefers name matching (case-insensitive). If no header names overlap the
    DataFrame columns, falls back to positional order.
    """
    if frame is None or frame.empty:
        return []

    header = list(header)
    df_cols = [str(c) for c in frame.columns]
    header_keys = [normalize_header(h) for h in header]
    col_by_key = {normalize_header(c): c for c in df_cols}
    overlap = {key for key in header_keys if key and key in col_by_key}

    rows: list[list[Any]] = []
    if overlap:
        for record in frame.to_dict(orient="records"):
            out: list[Any] = []
            for key in header_keys:
                if key in col_by_key:
                    out.append(_stringify_cell(record.get(col_by_key[key], "")))
                else:
                    out.append("")
            rows.append(out)
        return rows

    for record in frame.itertuples(index=False, name=None):
        values = [_stringify_cell(v) for v in record]
        if len(values) < len(header):
            values.extend([""] * (len(header) - len(values)))
        elif len(values) > len(header):
            values = values[: len(header)]
        rows.append(values)
    return rows


def merge_kept_and_new(
    sheet_rows: list[list[Any]],
    new_frame: pd.DataFrame,
    *,
    has_header: bool = True,
    column_index: int = COLUMN_P_INDEX,
) -> list[list[Any]]:
    """Filter blank-P rows, then append aligned Metabase rows."""
    kept = keep_rows_with_column_p(
        sheet_rows, has_header=has_header, column_index=column_index
    )

    if not kept:
        header = [str(c) for c in new_frame.columns]
        return [header] + align_new_rows(new_frame, header)

    header = list(kept[0]) if has_header else [str(c) for c in new_frame.columns]
    if not has_header:
        kept = [header] + kept
    new_rows = align_new_rows(new_frame, header)
    return kept + new_rows


def load_csv_as_rows(path: str) -> list[list[Any]]:
    frame = pd.read_csv(path, dtype=object, keep_default_na=False)
    header = [str(c) for c in frame.columns]
    body = [
        [_stringify_cell(v) for v in record]
        for record in frame.itertuples(index=False, name=None)
    ]
    return [header] + body


def write_rows_csv(path: str, rows: list[list[Any]]) -> None:
    if not rows:
        pd.DataFrame().to_csv(path, index=False)
        return
    header = [str(c) for c in rows[0]]
    body = rows[1:]
    pd.DataFrame(body, columns=header).to_csv(path, index=False)


def _service_account_info() -> dict:
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip()
    if path:
        with open(path, encoding="utf-8") as handle:
            return json.load(handle)
    if raw:
        return json.loads(raw)
    raise RuntimeError(
        "Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_FILE "
        "to a Google Cloud service-account key."
    )


def resolve_worksheet(spreadsheet, worksheet_name: str | None, worksheet_gid: int | str | None):
    """Open a tab by gid first (stable), then by name, including a leading-dot alias."""
    if worksheet_gid not in (None, ""):
        return spreadsheet.get_worksheet_by_id(int(worksheet_gid))

    names_to_try: list[str] = []
    if worksheet_name:
        names_to_try.append(worksheet_name)
        stripped = worksheet_name.lstrip(".")
        if stripped and stripped != worksheet_name:
            names_to_try.append(stripped)
        dotted = f".{stripped}" if stripped else worksheet_name
        if dotted not in names_to_try:
            names_to_try.append(dotted)

    last_error: Exception | None = None
    for name in names_to_try:
        try:
            return spreadsheet.worksheet(name)
        except Exception as exc:  # gspread.WorksheetNotFound when the lib is installed
            last_error = exc
            continue

    if last_error is not None:
        raise last_error
    return spreadsheet.sheet1


def open_worksheet(
    spreadsheet_id: str,
    worksheet_name: str | None,
    worksheet_gid: int | str | None = None,
):
    try:
        import gspread
    except ImportError as exc:
        raise RuntimeError(
            "gspread is required for Google Sheets access. "
            "Install dependencies with: pip install -r requirements.txt"
        ) from exc

    client = gspread.service_account_from_dict(_service_account_info())
    spreadsheet = client.open_by_key(spreadsheet_id)
    return resolve_worksheet(spreadsheet, worksheet_name, worksheet_gid)


def read_sheet_rows(worksheet) -> list[list[Any]]:
    return worksheet.get_all_values()


def write_sheet_rows(worksheet, rows: list[list[Any]]) -> None:
    worksheet.clear()
    if not rows:
        return
    worksheet.update(range_name="A1", values=rows, value_input_option="USER_ENTERED")


def fetch_metabase_frame(question_id: str) -> pd.DataFrame:
    frame = ret_metabase(question_id)
    if frame is None:
        return pd.DataFrame()
    return frame


def run_sync(
    *,
    question_id: str | None,
    spreadsheet_id: str | None,
    worksheet_name: str | None,
    worksheet_gid: int | str | None = None,
    local_input: str | None = None,
    local_output: str | None = None,
    local_new: str | None = None,
    has_header: bool = True,
    dry_run: bool = False,
) -> dict[str, int]:
    if local_input:
        sheet_rows = load_csv_as_rows(local_input)
    else:
        if not spreadsheet_id:
            raise RuntimeError("spreadsheet_id is required unless --local-input is set.")
        worksheet = open_worksheet(spreadsheet_id, worksheet_name, worksheet_gid)
        sheet_rows = read_sheet_rows(worksheet)

    original_data_rows = max(len(sheet_rows) - (1 if has_header and sheet_rows else 0), 0)

    if local_new:
        new_frame = pd.read_csv(local_new, dtype=object)
    else:
        if not question_id:
            raise RuntimeError("question_id is required unless --local-new is set.")
        new_frame = fetch_metabase_frame(question_id)

    merged = merge_kept_and_new(sheet_rows, new_frame, has_header=has_header)
    kept_data_rows = max(len(keep_rows_with_column_p(sheet_rows, has_header=has_header)) - 1, 0)
    appended_rows = max(len(merged) - 1 - kept_data_rows, 0)
    removed_rows = original_data_rows - kept_data_rows

    summary = {
        "original_data_rows": original_data_rows,
        "removed_blank_p_rows": removed_rows,
        "kept_data_rows": kept_data_rows,
        "appended_rows": appended_rows,
        "final_data_rows": max(len(merged) - 1, 0),
    }
    logger.info("Sheet sync summary: %s", summary)

    if dry_run:
        return summary

    if local_output:
        write_rows_csv(local_output, merged)
        return summary

    if not spreadsheet_id:
        raise RuntimeError("spreadsheet_id is required to write back to Google Sheets.")
    worksheet = open_worksheet(spreadsheet_id, worksheet_name, worksheet_gid)
    write_sheet_rows(worksheet, merged)
    return summary


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Remove sheet rows with a blank column P, then append Metabase data."
    )
    parser.add_argument(
        "--question-id",
        default=None,
        help="Metabase card/question ID. Defaults to 590 (telesales).",
    )
    parser.add_argument(
        "--spreadsheet-id",
        default=None,
        help="Google Sheet ID from the spreadsheet URL.",
    )
    parser.add_argument(
        "--worksheet",
        default=None,
        help="Tab name. Defaults to telesales_test.",
    )
    parser.add_argument(
        "--worksheet-gid",
        default=None,
        help="Google Sheet tab gid. Defaults to 292013814 (telesales_test).",
    )
    parser.add_argument(
        "--local-input",
        help="CSV stand-in for the current sheet (skips Google read).",
    )
    parser.add_argument(
        "--local-new",
        help="CSV stand-in for Metabase results (skips Metabase).",
    )
    parser.add_argument(
        "--local-output",
        help="Write the merged result to this CSV instead of Google Sheets.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Compute the merge but do not write anything.",
    )
    parser.add_argument(
        "--no-header",
        action="store_true",
        help="Treat the first sheet row as data instead of a header.",
    )
    parser.add_argument("--log-level", default="INFO")
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    logging.basicConfig(
        level=getattr(logging, str(args.log_level).upper(), logging.INFO),
        format="%(asctime)s %(levelname)s %(message)s",
    )
    summary = run_sync(
        question_id=coalesce_config(
            args.question_id, os.environ.get("METABASE_QUESTION_ID"), default=DEFAULT_QUESTION_ID
        ),
        spreadsheet_id=coalesce_config(
            args.spreadsheet_id,
            os.environ.get("GOOGLE_SPREADSHEET_ID"),
            default=DEFAULT_SPREADSHEET_ID,
        ),
        worksheet_name=coalesce_config(
            args.worksheet,
            os.environ.get("GOOGLE_WORKSHEET_NAME"),
            default=DEFAULT_WORKSHEET_NAME,
        ),
        worksheet_gid=coalesce_config(
            args.worksheet_gid,
            os.environ.get("GOOGLE_WORKSHEET_GID"),
            default=str(DEFAULT_WORKSHEET_GID),
        ),
        local_input=args.local_input,
        local_output=args.local_output,
        local_new=args.local_new,
        has_header=not args.no_header,
        dry_run=args.dry_run,
    )
    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())
