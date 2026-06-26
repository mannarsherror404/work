#!/usr/bin/env python3
"""
Aggregate recipe detail rows by date, then merge them with daily totals.

Example:
    python merge_daily_recipe_totals.py \
        --daily "input_daily_totals_no_outliers 2.xlsx" \
        --recipe "recipe details.csv" \
        --output "matched_daily_recipe_output.xlsx"
"""

from __future__ import annotations

import argparse
from pathlib import Path
import sys

import pandas as pd


DEFAULT_DAILY_DATE_COL = "SLAG_Date"
DEFAULT_RECIPE_DATE_COL = "SentDate"


def read_table(path: Path, sheet_name: str | int | None = None) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xls", ".xlsm"}:
        return pd.read_excel(path, sheet_name=0 if sheet_name is None else sheet_name)
    if suffix == ".csv":
        return pd.read_csv(path)
    if suffix == ".tsv":
        return pd.read_csv(path, sep="\t")
    raise ValueError(f"Unsupported file type for {path}. Use .xlsx, .xls, .xlsm, .csv, or .tsv.")


def normalize_date(series: pd.Series, column_name: str) -> pd.Series:
    parsed = pd.to_datetime(series, errors="coerce")
    if parsed.isna().all():
        raise ValueError(f"Could not parse any dates from column '{column_name}'.")
    return parsed.dt.normalize()


def ensure_column(df: pd.DataFrame, column: str, file_label: str) -> None:
    if column not in df.columns:
        available = ", ".join(map(str, df.columns))
        raise KeyError(f"Column '{column}' not found in {file_label}. Available columns: {available}")


def aggregate_recipe_by_date(recipe_df: pd.DataFrame, recipe_date_col: str) -> pd.DataFrame:
    ensure_column(recipe_df, recipe_date_col, "recipe file")

    recipe = recipe_df.copy()
    recipe["_merge_date"] = normalize_date(recipe[recipe_date_col], recipe_date_col)
    recipe = recipe.dropna(subset=["_merge_date"])

    numeric_cols = [
        col
        for col in recipe.select_dtypes(include="number").columns
        if col != "_merge_date"
    ]
    if not numeric_cols:
        raise ValueError("No numeric columns found in the recipe file to sum by date.")

    agg_map = {"Recipe_Row_Count": (recipe_date_col, "size")}
    agg_map.update({col: (col, "sum") for col in numeric_cols})

    daily_totals = (
        recipe.groupby("_merge_date", as_index=False)
        .agg(**agg_map)
        .sort_values("_merge_date")
    )
    daily_totals.insert(0, "Recipe_Date", daily_totals["_merge_date"].dt.date)
    return daily_totals


def merge_daily_with_recipe(
    daily_df: pd.DataFrame,
    recipe_daily_totals: pd.DataFrame,
    daily_date_col: str,
    join_type: str,
) -> pd.DataFrame:
    ensure_column(daily_df, daily_date_col, "daily totals file")

    daily = daily_df.copy()
    daily["_merge_date"] = normalize_date(daily[daily_date_col], daily_date_col)

    merged = daily.merge(
        recipe_daily_totals,
        on="_merge_date",
        how=join_type,
        suffixes=("", "_Recipe"),
    )
    return merged.drop(columns=["_merge_date"])


def unmatched_dates(
    left_dates: pd.DataFrame,
    right_dates: pd.DataFrame,
) -> pd.DataFrame:
    unmatched = left_dates.merge(right_dates, on="_merge_date", how="left", indicator=True)
    unmatched = unmatched[unmatched["_merge"] == "left_only"].drop(columns=["_merge"])
    unmatched = unmatched.copy()
    unmatched.insert(0, "Date", unmatched["_merge_date"].dt.date)
    return unmatched.drop(columns=["_merge_date"])


def build_unmatched_checks(
    daily_df: pd.DataFrame,
    recipe_daily_totals: pd.DataFrame,
    daily_date_col: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    daily_dates = pd.DataFrame(
        {"_merge_date": normalize_date(daily_df[daily_date_col], daily_date_col).dropna().unique()}
    )
    recipe_dates = recipe_daily_totals[["_merge_date"]].drop_duplicates()

    daily_without_recipe = unmatched_dates(daily_dates, recipe_dates)
    recipe_without_daily = unmatched_dates(recipe_dates, daily_dates)
    return daily_without_recipe, recipe_without_daily


def write_output(
    output_path: Path,
    merged: pd.DataFrame,
    recipe_daily_totals: pd.DataFrame,
    daily_without_recipe: pd.DataFrame,
    recipe_without_daily: pd.DataFrame,
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    recipe_sheet = recipe_daily_totals.drop(columns=["_merge_date"])

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="Merged_Output", index=False)
        recipe_sheet.to_excel(writer, sheet_name="Recipe_Daily_Totals", index=False)
        daily_without_recipe.to_excel(writer, sheet_name="Daily_No_Recipe_Match", index=False)
        recipe_without_daily.to_excel(writer, sheet_name="Recipe_No_Daily_Match", index=False)


def parse_sheet_arg(value: str | None) -> str | int | None:
    if value is None:
        return None
    return int(value) if value.isdigit() else value


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Aggregate recipe detail data by date and merge with daily totals.")
    parser.add_argument("--daily", required=True, type=Path, help="Path to input_daily_totals_no_outliers file.")
    parser.add_argument("--recipe", required=True, type=Path, help="Path to recipe details file.")
    parser.add_argument("--output", required=True, type=Path, help="Path for output .xlsx file.")
    parser.add_argument("--daily-date-col", default=DEFAULT_DAILY_DATE_COL, help=f"Daily file date column. Default: {DEFAULT_DAILY_DATE_COL}")
    parser.add_argument("--recipe-date-col", default=DEFAULT_RECIPE_DATE_COL, help=f"Recipe file date column. Default: {DEFAULT_RECIPE_DATE_COL}")
    parser.add_argument("--daily-sheet", default=None, help="Excel sheet name/index for daily file. Default: first sheet.")
    parser.add_argument("--recipe-sheet", default=None, help="Excel sheet name/index for recipe file if Excel. Default: first sheet.")
    parser.add_argument(
        "--join",
        default="inner",
        choices=["inner", "left", "right", "outer"],
        help="Merge type. Use inner for only matching dates; left to keep all daily dates. Default: inner.",
    )
    args = parser.parse_args()
    args.daily_sheet = parse_sheet_arg(args.daily_sheet)
    args.recipe_sheet = parse_sheet_arg(args.recipe_sheet)
    return args


def main() -> int:
    args = parse_args()

    try:
        daily_df = read_table(args.daily, args.daily_sheet)
        recipe_df = read_table(args.recipe, args.recipe_sheet)

        recipe_daily_totals = aggregate_recipe_by_date(recipe_df, args.recipe_date_col)
        merged = merge_daily_with_recipe(daily_df, recipe_daily_totals, args.daily_date_col, args.join)
        daily_without_recipe, recipe_without_daily = build_unmatched_checks(daily_df, recipe_daily_totals, args.daily_date_col)

        write_output(args.output, merged, recipe_daily_totals, daily_without_recipe, recipe_without_daily)

        print(f"Saved: {args.output}")
        print(f"Daily rows: {len(daily_df)}")
        print(f"Recipe rows: {len(recipe_df)}")
        print(f"Recipe daily total rows: {len(recipe_daily_totals)}")
        print(f"Merged output rows: {len(merged)}")
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())


# python merge_daily_recipe_totals.py --daily "input/input_daily_totals_no_outliers 2.xlsx" --recipe "input/recipe details.csv" --output "output/matched_daily_recipe_output.xlsx" 
