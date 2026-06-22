#!/usr/bin/env python3
"""Create day-wise sums for columns C and D from an Excel or CSV file.

Example:
    python3 daily_sum_by_date.py input.xlsx
    python3 daily_sum_by_date.py input.xlsx --output daily_totals.xlsx
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


DATE_COLUMN = "SLAG_Date"
COLUMN_C = "HM_Temp_Mean_For_Cast"
COLUMN_D = "HM_QTY"


def load_data(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)
    return pd.read_excel(path)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Sum HM_Temp_Mean_For_Cast and HM_QTY day-wise using SLAG_Date."
    )
    parser.add_argument("input_file", type=Path, help="Source .xlsx, .xls, or .csv file")
    parser.add_argument(
        "--output",
        type=Path,
        help="Output Excel file (default: <input>_daily_totals.xlsx)",
    )
    args = parser.parse_args()

    df = load_data(args.input_file)
    # Removes accidental spaces around headers, e.g. " HM_QTY ".
    df.columns = df.columns.astype(str).str.strip()

    needed = [DATE_COLUMN, COLUMN_C, COLUMN_D]
    missing = [column for column in needed if column not in df.columns]
    if missing:
        raise ValueError(
            f"Missing column(s): {', '.join(missing)}. Available: {', '.join(df.columns)}"
        )

    df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN], errors="coerce")
    df[COLUMN_C] = pd.to_numeric(df[COLUMN_C], errors="coerce").fillna(0)
    df[COLUMN_D] = pd.to_numeric(df[COLUMN_D], errors="coerce").fillna(0)

    daily_totals = (
        df.dropna(subset=[DATE_COLUMN])
        .groupby(DATE_COLUMN, as_index=False)[[COLUMN_C, COLUMN_D]]
        .sum()
        .sort_values(DATE_COLUMN)
    )

    output = args.output or args.input_file.with_name(
        f"{args.input_file.stem}_daily_totals.xlsx"
    )
    with pd.ExcelWriter(output, engine="openpyxl", date_format="yyyy-mm-dd") as writer:
        daily_totals.to_excel(writer, index=False, sheet_name="Daily Totals")

    print(f"Created: {output}")


if __name__ == "__main__":
    main()
