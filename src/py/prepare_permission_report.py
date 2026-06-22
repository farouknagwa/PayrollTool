#!/usr/bin/env python3
"""Prepare the Nagwa permission-details report for PayrollTool.

Reads the full HR export ``Nagwa_Permission_Request_Report.xls`` and writes a
filtered ``Nagwa_Permission_Request_permission_details.xls`` in the same folder.

Transformations vs the source report:
  - Drop ``Branch Code`` and ``Workflow ID``.
  - Add ``Total Permission Period`` (duration from Start Time → End Time).
  - Keep only rows with ``Status == "Approved"``.
  - Filter by ``Effective Date`` inside the HR payroll period for the chosen
    month (21st of the previous month through 20th of the target month).
  - Exclude rows whose ``Request Date`` is after the period end plus *N*
    grace days (default ``N = 0``, i.e. request must be on or before period end).

Output workbook presentation (by default):
  - Header row filled with very light gray.
  - Thin black borders on every cell.
  - All rows 0.24 inch tall.
  - First row frozen (header stays visible while scrolling).

Month examples (year 2026):
  --month 4  →  2026-03-21 .. 2026-04-20
  --month 5  →  2026-04-21 .. 2026-05-20

Usage:
    
    # prompts for month
    python3 -B prepare_permission_report.py          


# Default: Request Date must be on or before the period end (20th of the month)
python3 -B prepare_permission_report.py --month 5 --year 2026

# Allow requests up to the period end + 4 days (24th of the month)
python3 -B prepare_permission_report.py --month 5 --year 2026 --request-cutoff-days 4

# No request-date filter (keep all approved in-period rows)
python3 -B prepare_permission_report.py --month 5 --year 2026 --no-request-cutoff


"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import date, datetime, time, timedelta

import pandas as pd
import xlwt

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INPUT = os.path.join(SCRIPT_DIR, "Nagwa_Permission_Request_Report.xls")
DEFAULT_OUTPUT = os.path.join(
    SCRIPT_DIR, "Nagwa_Permission_Request_permission_details.xls"
)

OUTPUT_COLUMNS = [
    "Employee Code",
    "Employee Name",
    "Request Date",
    "Effective Date",
    "Start Time",
    "End Time",
    "Total Permission Period",
    "Time ",  # trailing space matches the HR / PayrollTool column name
    "Transaction Type",
    "Transaction Sub Type",
    "WF Template",
    "Status",
]

# Column widths copied from raw data/Nagwa_Permission_Request_permission_details.xls
# (xlwt / xlrd units: 1/256 of the zero-character width).
COLUMN_WIDTHS = [4522, 10240, 4096, 4181, 3541, 3328, 6186, 2602, 5760, 8576, 6058, 2773]

DURATION_COL = OUTPUT_COLUMNS.index("Total Permission Period")
EMPLOYEE_CODE_COL = OUTPUT_COLUMNS.index("Employee Code")

_THIN_BORDER_XF = (
    "borders: left thin, right thin, top thin, bottom thin, "
    "left_colour black, right_colour black, top_colour black, bottom_colour black"
)

# xlwt row height is in twips (1/20 pt). 1 inch = 72 pt = 1440 twips.
ROW_HEIGHT_TWIPS = round(0.24 * 72 * 20)


def _cell_style() -> xlwt.XFStyle:
    return xlwt.easyxf(_THIN_BORDER_XF)


def _header_style() -> xlwt.XFStyle:
    return xlwt.easyxf(
        _THIN_BORDER_XF + "; pattern: pattern solid, fore_colour gray25"
    )


def _duration_style() -> xlwt.XFStyle:
    return xlwt.easyxf(_THIN_BORDER_XF, num_format_str="h:mm")


def _set_row_height(sheet: xlwt.Worksheet.Worksheet, row_index: int) -> None:
    row = sheet.row(row_index)
    row.height = ROW_HEIGHT_TWIPS
    row.height_mismatch = 1


def period_from_month(month: int, year: int) -> tuple[date, date]:
    """Return (start, end) for HR month *month* in *year*.

    Month *N* means the period ending on the 20th of month *N*:
      start = 21st of month N-1 (December 21 when N == 1)
      end   = 20th of month N
    """
    if not 1 <= month <= 12:
        raise ValueError(f"Month must be between 1 and 12, got {month!r}.")
    if month == 1:
        start = date(year - 1, 12, 21)
    else:
        start = date(year, month - 1, 21)
    end = date(year, month, 20)
    return start, end


def parse_date(value) -> date | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    return datetime.strptime(text, "%d/%m/%Y").date()


def parse_time_value(value) -> datetime | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if isinstance(value, time):
        return datetime.combine(date.today(), value)
    if isinstance(value, datetime):
        return value
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    for fmt in ("%I:%M %p", "%I:%M%p", "%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def format_date(value) -> str:
    parsed = parse_date(value)
    if parsed is None:
        return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()
    return parsed.strftime("%d/%m/%Y")


def format_time(value) -> str:
    parsed = parse_time_value(value)
    if parsed is None:
        return "" if value is None or (isinstance(value, float) and pd.isna(value)) else str(value).strip()
    return parsed.strftime("%I:%M %p")


def permission_period_fraction(start_value, end_value) -> float:
    """Return duration as an Excel time serial (fraction of a day)."""
    start = parse_time_value(start_value)
    end = parse_time_value(end_value)
    if start is None or end is None:
        return 0.0
    seconds = int((end - start).total_seconds())
    if seconds < 0:
        seconds += 24 * 3600
    return seconds / 86400.0


def permission_period(start_value, end_value) -> str:
    """Return duration as HH:MM:SS between Start Time and End Time."""
    start = parse_time_value(start_value)
    end = parse_time_value(end_value)
    if start is None or end is None:
        return ""
    seconds = int((end - start).total_seconds())
    if seconds < 0:
        seconds += 24 * 3600
    hours, remainder = divmod(seconds, 3600)
    minutes, secs = divmod(remainder, 60)
    return f"{hours:02d}:{minutes:02d}:{secs:02d}"


def read_source_report(path: str) -> pd.DataFrame:
    if not os.path.isfile(path):
        raise FileNotFoundError(f"Input report not found: {path}")
    df = pd.read_excel(path, engine="calamine")
    df.columns = [str(c).strip() if c != "Time " else "Time " for c in df.columns]
    required = {
        "Employee Code",
        "Employee Name",
        "Request Date",
        "Effective Date",
        "Start Time",
        "End Time",
        "Time ",
        "Transaction Type",
        "Transaction Sub Type",
        "WF Template",
        "Status",
    }
    missing = required - set(df.columns)
    if missing:
        raise SystemExit(
            f"Missing required column(s) in {path!r}: {sorted(missing)}"
        )
    return df


def transform_report(
    df: pd.DataFrame,
    month: int,
    year: int,
    request_cutoff_days: int | None = 0,
) -> tuple[pd.DataFrame, date, date, int]:
    start_date, end_date = period_from_month(month, year)

    working = df.copy()
    working["_effective_date"] = working["Effective Date"].map(parse_date)
    invalid_effective = working["_effective_date"].isna().sum()
    if invalid_effective:
        print(
            f"  Warning: {invalid_effective} row(s) with unparseable "
            "Effective Date skipped."
        )

    mask = (
        working["_effective_date"].notna()
        & (working["_effective_date"] >= start_date)
        & (working["_effective_date"] <= end_date)
        & (working["Status"].astype(str).str.strip() == "Approved")
    )

    request_excluded = 0
    if request_cutoff_days is not None:
        cutoff_date = end_date + timedelta(days=request_cutoff_days)
        working["_request_date"] = working["Request Date"].map(parse_date)
        invalid_request = working.loc[mask, "_request_date"].isna().sum()
        if invalid_request:
            print(
                f"  Warning: {invalid_request} approved in-period row(s) with "
                "unparseable Request Date skipped."
            )
        before_request_filter = mask.sum()
        request_mask = working["_request_date"].notna() & (
            working["_request_date"] <= cutoff_date
        )
        mask &= request_mask
        request_excluded = int(before_request_filter - mask.sum())
        if request_excluded:
            print(
                f"  Request-date cutoff: excluded {request_excluded} row(s) with "
                f"Request Date after {cutoff_date.isoformat()} "
                f"(period end + {request_cutoff_days} day(s))."
            )

    filtered = working.loc[mask].copy()
    filtered["Total Permission Period"] = filtered.apply(
        lambda row: permission_period(row["Start Time"], row["End Time"]),
        axis=1,
    )

    for col in ("Request Date", "Effective Date"):
        filtered[col] = filtered[col].map(format_date)
    for col in ("Start Time", "End Time", "Time "):
        filtered[col] = filtered[col].map(format_time)

    filtered["Employee Code"] = filtered["Employee Code"].astype(int)
    filtered = filtered[OUTPUT_COLUMNS].reset_index(drop=True)
    return filtered, start_date, end_date, request_excluded


def write_xls(df: pd.DataFrame, path: str) -> None:
    """Write output matching raw data layout plus sheet presentation defaults."""
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("Worksheet")

    text_style = _cell_style()
    header_style = _header_style()
    duration_style = _duration_style()

    for col_idx, width in enumerate(COLUMN_WIDTHS):
        sheet.col(col_idx).width = width

    _set_row_height(sheet, 0)
    for col_idx, header in enumerate(OUTPUT_COLUMNS):
        sheet.write(0, col_idx, header, header_style)

    for row_idx in range(len(df)):
        excel_row = row_idx + 1
        _set_row_height(sheet, excel_row)
        row = df.iloc[row_idx]

        for col_idx, col_name in enumerate(OUTPUT_COLUMNS):
            value = row[col_name]

            if col_name == "Total Permission Period":
                sheet.write(
                    excel_row,
                    col_idx,
                    permission_period_fraction(row["Start Time"], row["End Time"]),
                    duration_style,
                )
            elif col_name == "Employee Code":
                sheet.write(excel_row, col_idx, int(value), text_style)
            elif pd.isna(value):
                sheet.write(excel_row, col_idx, "", text_style)
            else:
                sheet.write(excel_row, col_idx, str(value), text_style)

    sheet.panes_frozen = True
    sheet.horz_split_pos = 1
    sheet.vert_split_pos = 0

    workbook.save(path)


def prompt_month() -> int:
    while True:
        raw = input("Enter payroll month (1-12, e.g. 5 for Apr 21 → May 20): ").strip()
        try:
            month = int(raw)
            if 1 <= month <= 12:
                return month
        except ValueError:
            pass
        print("  Please enter an integer between 1 and 12.")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "-i",
        "--input",
        default=DEFAULT_INPUT,
        help=f"Source HR report (default: {os.path.basename(DEFAULT_INPUT)})",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=DEFAULT_OUTPUT,
        help=(
            "Prepared output file "
            f"(default: {os.path.basename(DEFAULT_OUTPUT)})"
        ),
    )
    parser.add_argument(
        "-m",
        "--month",
        type=int,
        choices=range(1, 13),
        help="Payroll month (period ends on the 20th of this month).",
    )
    parser.add_argument(
        "-y",
        "--year",
        type=int,
        default=2026,
        help="Calendar year for the payroll period (default: 2026).",
    )
    parser.add_argument(
        "-n",
        "--request-cutoff-days",
        type=int,
        default=0,
        metavar="N",
        help=(
            "Drop rows whose Request Date is after period end + N days "
            "(default: 0 = requests must be on or before the 20th)."
        ),
    )
    parser.add_argument(
        "--no-request-cutoff",
        action="store_true",
        help="Do not filter by Request Date (keep all approved in-period rows).",
    )
    args = parser.parse_args()

    month = args.month if args.month is not None else prompt_month()

    print(f"Reading: {args.input}")
    source_df = read_source_report(args.input)
    print(f"  {len(source_df)} row(s) loaded.")

    request_cutoff_days = None if args.no_request_cutoff else args.request_cutoff_days
    prepared_df, start_date, end_date, _ = transform_report(
        source_df,
        month,
        args.year,
        request_cutoff_days=request_cutoff_days,
    )
    print(
        f"Filtered to period {start_date.isoformat()} → {end_date.isoformat()} "
        f"(month {month}, year {args.year})."
    )
    if request_cutoff_days is not None:
        cutoff_date = end_date + timedelta(days=request_cutoff_days)
        print(
            f"  Request-date cutoff: on or before {cutoff_date.isoformat()} "
            f"(period end + {request_cutoff_days} day(s))."
        )
    else:
        print("  Request-date cutoff: disabled.")
    print(f"  {len(prepared_df)} approved row(s) kept.")

    write_xls(prepared_df, args.output)
    print(f"Saved: {args.output}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nCancelled.", file=sys.stderr)
        sys.exit(130)
