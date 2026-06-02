from datetime import datetime
from pathlib import Path

import pandas as pd

from logger import setup_logger
from mapping import MappingRule, generate_mapping_from_dataframe, read_mapping, write_mapping
from validator import ValidationError, ensure_input_file


def transform_forward(
    input_path: str | Path,
    mapping_path: str | Path | None = None,
    output_dir: str | Path = "output",
) -> Path:
    input_file = ensure_input_file(input_path, "Input file")
    df = _read_input_data(input_file)

    if mapping_path:
        rules = read_mapping(mapping_path)
    else:
        rules = generate_mapping_from_dataframe(df)

    _validate_required_columns(df, [rule.original for rule in rules], "input")
    transformed = _apply_forward(df, rules)

    output_path = _build_output_path(output_dir, "VendorData")
    _write_output(transformed, output_path)

    saved_mapping_path = _build_mapping_path(output_path)
    write_mapping(rules, saved_mapping_path)

    logger = setup_logger(Path(output_dir).resolve().parent / "logs")
    logger.info(
        "Forward transformation complete | input=%s | mapping=%s | output=%s | rows=%s | columns=%s",
        input_file,
        saved_mapping_path,
        output_path,
        len(transformed),
        len(transformed.columns),
    )

    return output_path


def transform_reverse(
    input_path: str | Path,
    mapping_path: str | Path | None = None,
    output_dir: str | Path = "output",
) -> Path:
    input_file = ensure_input_file(input_path, "Input file")

    if mapping_path is None:
        mapping_path = find_saved_mapping(input_file)

    rules = read_mapping(mapping_path)
    df = _read_input_data(input_file)

    _validate_required_columns(df, [rule.vendor for rule in rules], "vendor-returned")
    transformed = _apply_reverse(df, rules)

    output_path = _build_output_path(output_dir, "InternalData")
    _write_output(transformed, output_path)

    logger = setup_logger(Path(output_dir).resolve().parent / "logs")
    logger.info(
        "Reverse transformation complete | input=%s | mapping=%s | output=%s | rows=%s | columns=%s",
        input_file,
        mapping_path,
        output_path,
        len(transformed),
        len(transformed.columns),
    )

    return output_path


def find_saved_mapping(vendor_file: str | Path) -> Path:
    vendor_path = Path(vendor_file)
    mapping_path = _build_mapping_path(vendor_path)

    if mapping_path.exists():
        return mapping_path

    raise ValidationError(
        "Could not find the saved mapping file for this vendor file. "
        f"Expected it here: {mapping_path}"
    )


def _apply_forward(df: pd.DataFrame, rules: list[MappingRule]) -> pd.DataFrame:
    transformed = df.copy()

    for rule in rules:
        if rule.data_type == "numeric":
            transformed[rule.original] = _scale_numeric_column(
                transformed[rule.original],
                rule.factor,
                operation="multiply",
                column_name=rule.original,
            )

    rename_map = {rule.original: rule.vendor for rule in rules}
    return transformed.rename(columns=rename_map)


def _apply_reverse(df: pd.DataFrame, rules: list[MappingRule]) -> pd.DataFrame:
    transformed = df.copy()

    for rule in rules:
        if rule.data_type == "numeric":
            transformed[rule.vendor] = _scale_numeric_column(
                transformed[rule.vendor],
                rule.factor,
                operation="divide",
                column_name=rule.vendor,
            )

    rename_map = {rule.vendor: rule.original for rule in rules}
    return transformed.rename(columns=rename_map)


def _read_input_data(input_file: Path) -> pd.DataFrame:
    if input_file.suffix.lower() == ".csv":
        return pd.read_csv(input_file)

    return pd.read_excel(input_file)


def _scale_numeric_column(
    series: pd.Series,
    factor: float,
    operation: str,
    column_name: str,
) -> pd.Series:
    numeric_values = pd.to_numeric(series, errors="coerce")
    invalid_mask = series.notna() & numeric_values.isna()

    if invalid_mask.any():
        first_bad_index = invalid_mask[invalid_mask].index[0]
        raise ValidationError(
            f"Column '{column_name}' contains non-numeric data at row {first_bad_index + 2}."
        )

    if operation == "multiply":
        return numeric_values * factor

    if operation == "divide":
        return numeric_values / factor

    raise ValueError(f"Unsupported numeric operation: {operation}")


def _validate_required_columns(
    df: pd.DataFrame,
    required_columns: list[str],
    file_label: str,
) -> None:
    missing = [column for column in required_columns if column not in df.columns]

    if missing:
        raise ValidationError(
            f"The {file_label} file is missing required columns: {', '.join(missing)}"
        )


def _build_output_path(output_dir: str | Path, prefix: str) -> Path:
    output_path = Path(output_dir).resolve()
    output_path.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return output_path / f"{prefix}_{timestamp}.xlsx"


def _build_mapping_path(output_path: Path) -> Path:
    return output_path.with_name(f"{output_path.stem}_mapping.xlsx")


def _write_output(df: pd.DataFrame, output_path: Path) -> None:
    df.to_excel(output_path, index=False)
