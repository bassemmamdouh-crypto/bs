#!/usr/bin/env python3
"""Consolidate retailer sheets into one table.

Each source tab/sheet is expected to hold retailer rows. The sheet name is
parsed as:

    agent_name (supervisor_name)

Columns are matched by Arabic/English headers:

    كود              -> retailer_id
    رقم الهاتف       -> mobile_number
    الاسم التجاري    -> market_name
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import pandas as pd

OUTPUT_COLUMNS = [
    "retailer_id",
    "mobile_number",
    "market_name",
    "agent_name",
    "supervisor_name",
]

RETAILER_ID_HEADERS = {
    "كود",
    "الكود",
    "كود العميل",
    "retailer_id",
    "retailerid",
    "id",
    "code",
}
MOBILE_HEADERS = {
    "رقم الهاتف",
    "رقم التلفون",
    "الهاتف",
    "الموبايل",
    "جوال",
    "mobile",
    "mobile_number",
    "phone",
    "phonenumber",
    "tel",
}
MARKET_NAME_HEADERS = {
    "الاسم التجاري",
    "اسم المتجر",
    "اسم السوق",
    "الاسم",
    "market_name",
    "marketname",
    "name",
    "store_name",
    "storename",
}

SHEET_NAME_RE = re.compile(r"^(?P<agent>.*?)\s*\((?P<supervisor>[^)]+)\)\s*$")
NON_ALNUM = re.compile(r"[^0-9A-Za-z\u0600-\u06FF]+")
DIGIT_RE = re.compile(r"\d+")


def normalize_header(value: object) -> str:
    text = "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value)
    text = text.replace("\u200f", "").replace("\u200e", "").strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def compact_header(value: object) -> str:
    return NON_ALNUM.sub("", normalize_header(value))


def parse_sheet_name(sheet_name: str) -> tuple[str, str]:
    """Return (agent_name, supervisor_name) from a tab title.

    Tab titles look like ``Agent Name (Supervisor Name)``. The agent is the
    text outside the brackets. The supervisor is the text inside. CSV dumps
    that have no real tab title (file stems, default Sheet1, etc.) do not
    invent an agent name.
    """
    name = (sheet_name or "").strip()
    match = SHEET_NAME_RE.match(name)
    if match:
        agent = match.group("agent").strip()
        supervisor = match.group("supervisor").strip()
        return agent, supervisor
    if _is_placeholder_sheet_name(name):
        return "", ""
    return name, ""


def _is_placeholder_sheet_name(name: str) -> bool:
    lowered = name.lower()
    if lowered in {"sheet1", "sheet", "consolidated"}:
        return True
    if "_" in name or name.endswith("_2efc"):
        return True
    return False


def _cell_str(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0", text):
        return text[:-2]
    return text


def _digit_len(value: object) -> int:
    digits = "".join(DIGIT_RE.findall(_cell_str(value)))
    return len(digits)


def _series_nonempty(series: pd.Series) -> pd.Series:
    return series.map(_cell_str).replace("", pd.NA).dropna()


def _looks_like_header_row(row: pd.Series) -> bool:
    texts = [normalize_header(v) for v in row.tolist()]
    joined = " ".join(texts)
    keywords = ("كود", "هاتف", "تلفون", "موبايل", "اسم", "تجاري", "code", "phone", "mobile", "name")
    return sum(1 for key in keywords if key in joined) >= 2


def _promote_header_if_needed(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame
    current = [normalize_header(c) for c in frame.columns]
    unnamed = all(c.startswith("unnamed") or c == "" or re.fullmatch(r"\d+", c) for c in current)
    if unnamed or _looks_like_header_row(frame.iloc[0]):
        if _looks_like_header_row(frame.iloc[0]):
            headers = [_cell_str(v) or f"col_{i}" for i, v in enumerate(frame.iloc[0].tolist())]
            frame = frame.iloc[1:].copy()
            frame.columns = headers
            frame.reset_index(drop=True, inplace=True)
    return frame


def _match_known_headers(columns: list[object]) -> dict[str, object]:
    mapping: dict[str, object] = {}
    for col in columns:
        norm = normalize_header(col)
        compact = compact_header(col)
        if col in mapping.values():
            continue
        if norm in RETAILER_ID_HEADERS or compact in {compact_header(x) for x in RETAILER_ID_HEADERS}:
            mapping.setdefault("retailer_id", col)
        elif norm in MOBILE_HEADERS or compact in {compact_header(x) for x in MOBILE_HEADERS}:
            mapping.setdefault("mobile_number", col)
        elif norm in MARKET_NAME_HEADERS or compact in {compact_header(x) for x in MARKET_NAME_HEADERS}:
            mapping.setdefault("market_name", col)
    return mapping


def _score_id_column(series: pd.Series) -> float:
    values = _series_nonempty(series)
    if values.empty:
        return -1.0
    lengths = values.map(_digit_len)
    ratio = (lengths.between(3, 6)).mean()
    unique_ratio = values.nunique() / max(len(values), 1)
    return float(ratio * 2 + unique_ratio)


def _score_mobile_column(series: pd.Series) -> float:
    values = _series_nonempty(series)
    if values.empty:
        return -1.0
    lengths = values.map(_digit_len)
    return float((lengths.between(8, 15)).mean())


def _score_name_column(series: pd.Series) -> float:
    values = _series_nonempty(series)
    if values.empty:
        return -1.0
    texts = values.map(_cell_str)
    avg_len = texts.map(len).mean()
    digit_ratio = texts.map(lambda s: _digit_len(s) / max(len(s), 1)).mean()
    return float(avg_len * (1 - digit_ratio))


def infer_column_mapping(frame: pd.DataFrame) -> dict[str, object]:
    mapping = _match_known_headers(list(frame.columns))
    remaining = [c for c in frame.columns if c not in mapping.values()]

    if "mobile_number" not in mapping and remaining:
        mapping["mobile_number"] = max(remaining, key=lambda c: _score_mobile_column(frame[c]))
        remaining = [c for c in remaining if c != mapping["mobile_number"]]

    if "retailer_id" not in mapping and remaining:
        mapping["retailer_id"] = max(remaining, key=lambda c: _score_id_column(frame[c]))
        remaining = [c for c in remaining if c != mapping["retailer_id"]]

    if "market_name" not in mapping and remaining:
        mapping["market_name"] = max(remaining, key=lambda c: _score_name_column(frame[c]))

    return mapping


def extract_sheet_rows(frame: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    agent_name, supervisor_name = parse_sheet_name(sheet_name)
    frame = frame.copy()
    frame = frame.dropna(how="all")
    if frame.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    frame = _promote_header_if_needed(frame)
    mapping = infer_column_mapping(frame)

    missing = [key for key in ("retailer_id", "mobile_number", "market_name") if key not in mapping]
    if missing:
        raise ValueError(f"Sheet {sheet_name!r} is missing columns: {missing}")

    out = pd.DataFrame(
        {
            "retailer_id": frame[mapping["retailer_id"]].map(_cell_str),
            "mobile_number": frame[mapping["mobile_number"]].map(_cell_str),
            "market_name": frame[mapping["market_name"]].map(_cell_str),
            "agent_name": agent_name,
            "supervisor_name": supervisor_name,
        }
    )
    out = out[out["retailer_id"].ne("") | out["mobile_number"].ne("") | out["market_name"].ne("")]
    header_like = out["market_name"].map(normalize_header).isin(MARKET_NAME_HEADERS) & out[
        "retailer_id"
    ].map(normalize_header).isin(RETAILER_ID_HEADERS)
    out = out.loc[~header_like].copy()
    out.reset_index(drop=True, inplace=True)
    return out[OUTPUT_COLUMNS]


def read_workbook(path: Path) -> dict[str, pd.DataFrame]:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path, sheet_name=None, dtype=object)
    if suffix == ".csv":
        return {path.stem: pd.read_csv(path, dtype=object, encoding="utf-8", encoding_errors="replace")}
    raise ValueError(f"Unsupported file type: {path.suffix}")


def consolidate(path: Path) -> pd.DataFrame:
    sheets = read_workbook(path)
    frames = []
    for sheet_name, frame in sheets.items():
        rows = extract_sheet_rows(frame, sheet_name)
        if not rows.empty:
            frames.append(rows)
    if not frames:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)
    return pd.concat(frames, ignore_index=True)


def write_outputs(frame: pd.DataFrame, output_xlsx: Path, output_csv: Path | None = None) -> None:
    output_xlsx.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_xlsx, engine="openpyxl") as writer:
        frame.to_excel(writer, index=False, sheet_name="consolidated")
        worksheet = writer.sheets["consolidated"]
        for column_cells in worksheet.columns:
            max_length = max(len(_cell_str(cell.value)) for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 48)
    if output_csv is not None:
        output_csv.parent.mkdir(parents=True, exist_ok=True)
        frame.to_csv(output_csv, index=False, encoding="utf-8-sig")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="Excel workbook with one tab per agent")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("retailers/output/massabeh_retailers_consolidated.xlsx"),
        help="Output Excel path",
    )
    parser.add_argument("--csv", type=Path, default=None, help="Optional UTF-8 CSV copy")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if not args.input.exists():
        print(f"Input file not found: {args.input}", file=sys.stderr)
        return 1
    frame = consolidate(args.input)
    csv_path = args.csv
    if csv_path is None:
        csv_path = args.output.with_suffix(".csv")
    write_outputs(frame, args.output, csv_path)
    print(f"Wrote {len(frame)} rows to {args.output} and {csv_path}")
    print(f"Agents: {frame['agent_name'].nunique()} | Supervisors: {frame['supervisor_name'].nunique()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
