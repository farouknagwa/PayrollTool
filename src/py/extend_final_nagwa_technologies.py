#!/usr/bin/env python3
"""Replace the attendance period in the Final Nagwa Technologies report.

The source workbook lives in ``templates/`` and the result is written to
``output/`` using the same file name.

The period is auto-detected from the raw attendance report
(``raw data/Attendance Report.xls``): the first date in the "Attendance
Day" column decides which 21st->20th period to build. For example, a first
date of 28-04-2026 yields the period 21 April -> 20 May.

The script rewrites the date columns (``I3:AM3``) so they cover the detected
period from the 21st of one month to the 20th of the next month, and clears
any existing attendance values in the employee rows for those columns. All
other formatting, formulas and summary columns are preserved.
"""

from __future__ import annotations

import argparse
import calendar
import os
from datetime import date, datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Paths are resolved relative to the project root (the parent of scripts/).
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(SCRIPT_DIR)
DEFAULT_INPUT = os.path.join(
    PROJECT_ROOT, "templates", "Final Nagwa Technologies.xlsx"
)
DEFAULT_OUTPUT_DIR = os.path.join(PROJECT_ROOT, "output")
DEFAULT_REPORT = os.path.join(PROJECT_ROOT, "raw data", "Attendance Report.xls")

DATE_ROW = 3
HEADER_ROW = 4
FIRST_DATA_ROW = 5
DATE_COL_START = 9   # column I
DATE_COL_END = 39    # column AM (inclusive) -> 31 columns total


# ----------------------------------------------------------------------
# Period detection
# ----------------------------------------------------------------------

def compute_period(ref: date) -> tuple[date, date]:
    """Return the (start, end) of the HR period containing ``ref``.

    Periods run from the 21st of one month through the 20th of the next
    month, inclusive. The period chosen is the one that contains ``ref``:
      - day >= 21 -> period starts on the 21st of ``ref``'s month;
      - day <= 20 -> period starts on the 21st of the previous month.
    """
    if ref.day >= 21:
        start = date(ref.year, ref.month, 21)
    elif ref.month == 1:
        start = date(ref.year - 1, 12, 21)
    else:
        start = date(ref.year, ref.month - 1, 21)

    if start.month == 12:
        end = date(start.year + 1, 1, 20)
    else:
        end = date(start.year, start.month + 1, 20)
    return start, end


def detect_first_attendance_date(report_path: str) -> date:
    """Read the first date in the "Attendance Day" column of the report."""
    df = pd.read_excel(report_path, header=0)
    col = "Attendance Day"
    if col not in df.columns:
        raise SystemExit(
            f"Column {col!r} not found in {report_path!r}. "
            f"Available columns: {list(df.columns)}"
        )
    series = df[col].dropna()
    if series.empty:
        raise SystemExit(f"No dates found in column {col!r} of {report_path!r}.")

    first = series.iloc[0]
    if isinstance(first, str):
        return datetime.strptime(first.strip(), "%d/%m/%Y").date()
    return pd.to_datetime(first).date()


def build_period(start: date, end: date) -> list[datetime]:
    """Return the list of dates from ``start`` to ``end`` (inclusive)."""
    dates = []
    cur = start
    while cur <= end:
        dates.append(datetime(cur.year, cur.month, cur.day))
        # next day
        d = cur.day + 1
        m = cur.month
        y = cur.year
        if d > calendar.monthrange(y, m)[1]:
            d = 1
            m += 1
            if m > 12:
                m = 1
                y += 1
        cur = date(y, m, d)
    return dates


def replace_period(input_path: str, output_path: str, ref_date: date) -> None:
    start_date, end_date = compute_period(ref_date)
    dates = build_period(start_date, end_date)
    n_dates = len(dates)
    n_slots = DATE_COL_END - DATE_COL_START + 1
    if n_dates > n_slots:
        raise RuntimeError(
            f"Period has {n_dates} days but the sheet only has "
            f"{n_slots} date columns ({get_column_letter(DATE_COL_START)}:"
            f"{get_column_letter(DATE_COL_END)})."
        )

    wb = load_workbook(input_path)
    ws = wb.active

    # 1) Rewrite the date row (row 3). Leave trailing slots empty if the
    #    period is shorter than the available 31 columns.
    for offset in range(n_slots):
        col = DATE_COL_START + offset
        cell = ws.cell(row=DATE_ROW, column=col)
        if offset < n_dates:
            cell.value = dates[offset]
            # Preserve a date format if the cell did not already have one.
            if not cell.number_format or cell.number_format == "General":
                cell.number_format = "dd-mmm"
        else:
            cell.value = None

    # 2) Clear existing attendance values for every employee row inside the
    #    date column range. Skip the header rows (1-4) and the summary block
    #    after column AM.
    last_row = ws.max_row
    for row in range(FIRST_DATA_ROW, last_row + 1):
        for col in range(DATE_COL_START, DATE_COL_END + 1):
            cell = ws.cell(row=row, column=col)
            if cell.value is not None:
                cell.value = None

    wb.save(output_path)
    wb.close()


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Replace the attendance period in the Final Nagwa report."
    )
    parser.add_argument(
        "-i", "--input",
        default=DEFAULT_INPUT,
        help="Source workbook (default: templates/Final Nagwa Technologies.xlsx).",
    )
    parser.add_argument(
        "-o", "--output",
        default=None,
        help=(
            "Output workbook. Defaults to the output/ folder using the "
            "same file name as the input."
        ),
    )
    parser.add_argument(
        "-r", "--report",
        default=DEFAULT_REPORT,
        help=(
            "Attendance report used to detect the period "
            "(default: raw data/Attendance Report.xls)."
        ),
    )
    args = parser.parse_args()

    inp = os.path.abspath(args.input)
    if not os.path.isfile(inp):
        parser.error(f"Input file not found: {inp}")

    report = os.path.abspath(args.report)
    if not os.path.isfile(report):
        parser.error(f"Attendance report not found: {report}")

    if args.output:
        out = os.path.abspath(args.output)
    else:
        out = os.path.join(DEFAULT_OUTPUT_DIR, os.path.basename(inp))

    os.makedirs(os.path.dirname(out), exist_ok=True)

    first_date = detect_first_attendance_date(report)
    print(f"Detected first attendance date: {first_date.isoformat()}")

    replace_period(inp, out, first_date)

    start_date, end_date = compute_period(first_date)
    dates = build_period(start_date, end_date)
    print(
        f"Wrote {len(dates)} dates "
        f"({dates[0].strftime('%d %b %Y')} - {dates[-1].strftime('%d %b %Y')}) "
        f"to '{out}'."
    )


if __name__ == "__main__":
    main()
