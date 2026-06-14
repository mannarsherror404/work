"""
Create PCI and Coke Rate condition analysis from matched HM data.

Purpose:
    - Create one table sorted by high PCI / injection rate with recipe details.
    - Create one table sorted by low coke rate with recipe details.
    - Summarize average recipe and operating conditions by PCI and coke-rate class.

Default input:
    ~/Desktop/HM_Si_Project/input/matched_hm_analysis_output_randomized.xlsx

Default output:
    ~/Desktop/HM_Si_Project/output/PCI_CR_Condition_Analysis.xlsx
"""

from __future__ import annotations

import argparse
import math
import re
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


BASE_DIR = Path.home() / "Desktop" / "HM_Si_Project"
DEFAULT_INPUT = BASE_DIR / "input" / "matched_hm_analysis_output_randomized.xlsx"
DEFAULT_OUTPUT = BASE_DIR / "output" / "PCI_CR_Condition_Analysis.xlsx"
DEFAULT_SHEET = "HM Match Output"

PCI_CANDIDATES = ["InjRate", "PCI", "PCI Rate", "Injection Rate", "PCI / Injection Rate"]
CR_CANDIDATES = ["CR", "Coke Rate", "CR / Coke Rate"]

CONDITION_PRIORITY = [
    "RecipeName",
    "S",
    "P",
    "O",
    "Sinter",
    "Pellet",
    "Ore",
    "NC",
    "Dolo",
    "QTZ",
    "LS",
    "SS",
    "OxideWt",
    "CB",
    "NCR",
    "FuelRate",
    "FluxRate",
    "HMperCh",
    "SR",
    "B2",
    "SlagAl2O3",
    "SlagMgO",
]


def normalize_name(value) -> str:
    value = "" if value is None else str(value)
    return re.sub(r"\s+", " ", value.strip())


def compact_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.lower())


def read_matched_sheet(path: Path, sheet_name: str | None) -> tuple[pd.DataFrame, str]:
    xls = pd.ExcelFile(path)
    sheet = sheet_name or DEFAULT_SHEET
    if sheet not in xls.sheet_names:
        sheet = xls.sheet_names[0]

    raw = pd.read_excel(path, sheet_name=sheet, header=None)
    keywords = {
        "cast_number",
        "sentdate",
        "opening_time",
        "hm_si",
        "recipename",
        "injrate",
        "cr",
    }

    best_row = 0
    best_score = -1
    for idx in range(min(10, len(raw))):
        vals = {normalize_name(v).lower() for v in raw.iloc[idx].tolist()}
        score = len(keywords.intersection(vals))
        if score > best_score:
            best_score = score
            best_row = idx

    if best_score >= 3:
        columns = [
            normalize_name(v) if normalize_name(v) else f"Unnamed_{i}"
            for i, v in enumerate(raw.iloc[best_row].tolist())
        ]
        df = raw.iloc[best_row + 1 :].copy()
        df.columns = columns
    else:
        df = pd.read_excel(path, sheet_name=sheet)
        df.columns = [normalize_name(c) for c in df.columns]

    df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
    return df, sheet


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    exact = {str(c).lower(): c for c in df.columns}
    compact = {compact_name(str(c)): c for c in df.columns}

    for candidate in candidates:
        if candidate.lower() in exact:
            return exact[candidate.lower()]

    for candidate in candidates:
        key = compact_name(candidate)
        if key in compact:
            return compact[key]

    for candidate in candidates:
        key = compact_name(candidate)
        for real_col in df.columns:
            if key and key in compact_name(str(real_col)):
                return real_col
    return None


def ordered_present_columns(df: pd.DataFrame, candidates: list[str]) -> list[str]:
    found = []
    used = set()
    for candidate in candidates:
        col = find_column(df, [candidate])
        if col and col not in used:
            found.append(col)
            used.add(col)
    return found


def numeric_or_blank(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    out = df.copy()
    for col in columns:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce")
    return out


def prepare_analysis_data(df: pd.DataFrame, pci_col: str, cr_col: str) -> tuple[pd.DataFrame, list[str]]:
    condition_cols = ordered_present_columns(df, CONDITION_PRIORITY)

    selected_cols = []
    for col in [pci_col, cr_col] + condition_cols:
        if col in df.columns and col not in selected_cols:
            selected_cols.append(col)

    numeric_cols = [pci_col, cr_col] + [col for col in condition_cols if col != "RecipeName"]
    working = numeric_or_blank(df[selected_cols], numeric_cols)
    working = working.dropna(subset=[pci_col, cr_col], how="any").copy()
    return working.reset_index(drop=True), condition_cols


def build_pci_table(analysis_df: pd.DataFrame, pci_col: str, cr_col: str) -> pd.DataFrame:
    table = analysis_df.copy()
    table["PCI_Rank_Max_to_Min"] = table[pci_col].rank(method="min", ascending=False).astype(int)
    recipe_cols = [col for col in table.columns if col not in {pci_col, cr_col, "PCI_Rank_Max_to_Min"}]
    ordered_cols = ["PCI_Rank_Max_to_Min", pci_col] + recipe_cols
    return (
        table[ordered_cols]
        .sort_values(by=[pci_col], ascending=[False])
        .reset_index(drop=True)
    )


def build_cr_table(analysis_df: pd.DataFrame, pci_col: str, cr_col: str) -> pd.DataFrame:
    table = analysis_df.copy()
    table["CR_Rank_Min_to_Max"] = table[cr_col].rank(method="min", ascending=True).astype(int)
    recipe_cols = [col for col in table.columns if col not in {pci_col, cr_col, "CR_Rank_Min_to_Max"}]
    ordered_cols = ["CR_Rank_Min_to_Max", cr_col] + recipe_cols
    return (
        table[ordered_cols]
        .sort_values(by=[cr_col], ascending=[True])
        .reset_index(drop=True)
    )


def make_bucket_label(value: float, bucket_size: int) -> str:
    lower = math.floor(value / bucket_size) * bucket_size
    upper = lower + bucket_size
    return f"{lower:g}-{upper:g}"


def build_group_table(
    source: pd.DataFrame,
    group_col: str,
    group_name: str,
    pci_col: str,
    cr_col: str,
    condition_cols: list[str],
    bucket_size: int,
    sort_ascending: bool,
) -> pd.DataFrame:
    grouped_source = source.copy()
    class_col = f"{group_name}_Class"
    start_col = f"{group_name}_Class_Start"
    grouped_source[class_col] = grouped_source[group_col].apply(
        lambda value: make_bucket_label(float(value), bucket_size)
        if pd.notna(value)
        else np.nan
    )
    grouped_source[start_col] = grouped_source[group_col].apply(
        lambda value: math.floor(float(value) / bucket_size) * bucket_size
        if pd.notna(value)
        else np.nan
    )

    numeric_condition_cols = [
        col for col in condition_cols
        if col in grouped_source.columns and col != "RecipeName"
    ]
    average_cols = [pci_col, cr_col] + numeric_condition_cols
    for col in average_cols:
        grouped_source[col] = pd.to_numeric(grouped_source[col], errors="coerce")

    agg_map = {
        "Row_Count": (group_col, "count"),
        "Avg_PCI": (pci_col, "mean"),
        "Avg_CR": (cr_col, "mean"),
        "Min_CR": (cr_col, "min"),
        "Max_CR": (cr_col, "max"),
        "Min_PCI": (pci_col, "min"),
        "Max_PCI": (pci_col, "max"),
    }
    for col in numeric_condition_cols:
        if col in grouped_source.columns:
            agg_map[f"Avg_{col}"] = (col, "mean")

    summary = (
        grouped_source.groupby([start_col, class_col], dropna=True)
        .agg(**agg_map)
        .reset_index()
        .sort_values(start_col, ascending=sort_ascending)
    )
    summary = summary.drop(columns=[start_col])

    numeric_cols = [col for col in summary.columns if col not in {class_col, "Row_Count"}]
    summary[numeric_cols] = summary[numeric_cols].round(4)
    return summary


def write_output(
    pci_table: pd.DataFrame,
    cr_table: pd.DataFrame,
    pci_group: pd.DataFrame,
    cr_group: pd.DataFrame,
    output: Path,
    source_file: Path,
    source_sheet: str,
    pci_col: str,
    cr_col: str,
    bucket_size: int,
) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(
            [
                ["Source File", str(source_file)],
                ["Source Sheet", source_sheet],
                ["PCI Column", pci_col],
                ["Coke Rate Column", cr_col],
                ["Class Size", bucket_size],
                ["Purpose", "Separate PCI high-to-low and coke-rate low-to-high condition analysis."],
            ],
            columns=["Field", "Value"],
        ).to_excel(writer, index=False, sheet_name="README")
        pci_table.to_excel(writer, index=False, sheet_name="PCI_Max_to_Min")
        cr_table.to_excel(writer, index=False, sheet_name="CR_Min_to_Max")
        pci_group.to_excel(writer, index=False, sheet_name="PCI_Class_Averages")
        cr_group.to_excel(writer, index=False, sheet_name="CR_Class_Averages")

    style_workbook(output, pci_col, cr_col)


def style_workbook(output: Path, pci_col: str, cr_col: str) -> None:
    wb = load_workbook(output)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    rank_fill = PatternFill("solid", fgColor="DDEBF7")
    metric_fill = PatternFill("solid", fgColor="FCE4D6")
    recipe_fill = PatternFill("solid", fgColor="E2F0D9")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = border
                cell.font = Font(name="Arial", size=10)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 24)

    pci_sheet = wb["PCI_Max_to_Min"]
    cr_sheet = wb["CR_Min_to_Max"]
    for ws in [pci_sheet, cr_sheet]:
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.column == 1:
                    cell.fill = rank_fill
                elif cell.column == 2:
                    cell.fill = metric_fill
                else:
                    cell.fill = recipe_fill
                if cell.row == 1:
                    cell.font = Font(name="Arial", size=10, bold=True, color="000000")

    for ws, metric_col, high_is_good in [
        (pci_sheet, pci_col, True),
        (cr_sheet, cr_col, False),
    ]:
        headers = [cell.value for cell in ws[1]]
        if metric_col not in headers:
            continue
        col_letter = get_column_letter(headers.index(metric_col) + 1)
        if high_is_good:
            start_color, end_color = "F4CCCC", "00B050"
        else:
            start_color, end_color = "00B050", "C00000"
        ws.conditional_formatting.add(
            f"{col_letter}2:{col_letter}{ws.max_row}",
            ColorScaleRule(
                start_type="min",
                start_color=start_color,
                mid_type="percentile",
                mid_value=50,
                mid_color="FFF2CC",
                end_type="max",
                end_color=end_color,
            ),
        )

    for sheet_name in ["PCI_Class_Averages", "CR_Class_Averages"]:
        grouped = wb[sheet_name]
        for row in grouped.iter_rows(min_row=2, max_row=grouped.max_row, min_col=1, max_col=grouped.max_column):
            for cell in row:
                if cell.column == 1:
                    cell.fill = rank_fill
                elif cell.column <= 8:
                    cell.fill = metric_fill
                else:
                    cell.fill = recipe_fill
        for cell in grouped[1]:
            if cell.column == 1:
                cell.fill = rank_fill
            elif cell.column <= 8:
                cell.fill = metric_fill
            else:
                cell.fill = recipe_fill
            cell.font = Font(name="Arial", size=10, bold=True, color="000000")

    for ws in [pci_sheet, cr_sheet, wb["PCI_Class_Averages"], wb["CR_Class_Averages"]]:
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                if isinstance(cell.value, (int, float)):
                    cell.number_format = "0.0000"

    readme = wb["README"]
    readme.column_dimensions["A"].width = 24
    readme.column_dimensions["B"].width = 90

    wb.save(output)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=str(DEFAULT_INPUT), help="Matched HM analysis Excel file")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="Matched data sheet")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="Output Excel file")
    parser.add_argument("--bucket-size", type=int, default=10, help="Coke rate class size")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser()
    output_path = Path(args.output).expanduser()

    df, source_sheet = read_matched_sheet(input_path, args.sheet)
    pci_col = find_column(df, PCI_CANDIDATES)
    cr_col = find_column(df, CR_CANDIDATES)

    if not pci_col:
        raise ValueError(f"Could not find PCI column. Tried: {PCI_CANDIDATES}")
    if not cr_col:
        raise ValueError(f"Could not find coke rate column. Tried: {CR_CANDIDATES}")

    analysis_df, condition_cols = prepare_analysis_data(df, pci_col, cr_col)
    pci_table = build_pci_table(analysis_df, pci_col, cr_col)
    cr_table = build_cr_table(analysis_df, pci_col, cr_col)
    pci_group = build_group_table(
        analysis_df,
        pci_col,
        "PCI",
        pci_col,
        cr_col,
        condition_cols,
        args.bucket_size,
        sort_ascending=False,
    )
    cr_group = build_group_table(
        analysis_df,
        cr_col,
        "CR",
        pci_col,
        cr_col,
        condition_cols,
        args.bucket_size,
        sort_ascending=True,
    )
    write_output(
        pci_table,
        cr_table,
        pci_group,
        cr_group,
        output_path,
        input_path,
        source_sheet,
        pci_col,
        cr_col,
        args.bucket_size,
    )

    print("Done.")
    print(f"Rows in PCI table: {len(pci_table)}")
    print(f"Rows in CR table: {len(cr_table)}")
    print(f"PCI classes created: {len(pci_group)}")
    print(f"CR classes created: {len(cr_group)}")
    print(f"Output saved to: {output_path.resolve()}")


if __name__ == "__main__":
    main()
