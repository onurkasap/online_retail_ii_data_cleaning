"""Reproducible, audit-friendly cleaning pipeline for Online Retail II."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = PROJECT_ROOT / "data" / "online_retail_II.xlsx"
DEFAULT_CLEANED = PROJECT_ROOT / "outputs" / "cleaned_transactions.csv"
DEFAULT_SALES = PROJECT_ROOT / "outputs" / "sales_analysis_ready.csv"
DEFAULT_REPORT = PROJECT_ROOT / "reports" / "cleaning_report.md"

# The workbook uses these original labels. Aliases also support common UCI variants.
COLUMN_ALIASES = {
    "invoice": "invoice_no",
    "invoiceno": "invoice_no",
    "stockcode": "stock_code",
    "description": "description",
    "quantity": "quantity",
    "invoicedate": "invoice_date",
    "price": "unit_price",
    "unitprice": "unit_price",
    "customerid": "customer_id",
    "country": "country",
}
REQUIRED = ("invoice_no", "stock_code", "quantity", "invoice_date", "unit_price")


def canonical_name(source_name: Any) -> str:
    """Normalize a source header and map known dataset fields to stable names."""
    key = "".join(character for character in str(source_name).casefold() if character.isalnum())
    return COLUMN_ALIASES.get(key, "_".join(str(source_name).strip().casefold().split()))


def _trim_string(series: pd.Series) -> pd.Series:
    """Strip surrounding whitespace without turning missing values into text."""
    return series.astype("string").str.strip().replace("", pd.NA)


def _identifier(series: pd.Series) -> pd.Series:
    """Keep identifiers textual and remove Excel's trailing .0 for integral IDs."""
    values = _trim_string(series)
    numeric = pd.to_numeric(values, errors="coerce")
    integral = numeric.notna() & numeric.mod(1).eq(0)
    values.loc[integral] = numeric.loc[integral].astype("int64").astype("string")
    return values


def process_worksheet(
    source: pd.DataFrame,
    worksheet: str,
    excel_start_row: int = 2,
) -> tuple[pd.DataFrame, pd.DataFrame, dict[str, Any]]:
    """Create auditable canonical records and the eligible positive-sales subset."""
    original_columns = list(source.columns)
    canonical_columns = [canonical_name(column) for column in original_columns]
    if len(set(canonical_columns)) != len(canonical_columns):
        raise ValueError(f"Worksheet {worksheet!r} has source columns that normalize to duplicate names")

    # Retain the source values under source_* columns before any normalization.
    result = pd.DataFrame(index=source.index)
    for column in original_columns:
        result[f"source_{canonical_name(column)}"] = source[column]
    result["source_worksheet"] = worksheet
    result["source_excel_row"] = range(excel_start_row, excel_start_row + len(source))
    result["source_row_id"] = [f"{worksheet}:{row}" for row in result["source_excel_row"]]

    values = source.copy()
    values.columns = canonical_columns
    for column in values.columns:
        if pd.api.types.is_object_dtype(values[column]) or isinstance(values[column].dtype, pd.StringDtype):
            values[column] = _trim_string(values[column])

    # Invoice and stock identifiers retain their exact textual representation,
    # including leading zeros. Customer IDs may arrive as Excel floats (e.g. 42.0).
    for identifier_column in ("invoice_no", "stock_code"):
        if identifier_column in values:
            values[identifier_column] = _trim_string(values[identifier_column])
    if "customer_id" in values:
        values["customer_id"] = _identifier(values["customer_id"])
    for numeric_column in ("quantity", "unit_price"):
        if numeric_column in values:
            values[numeric_column] = pd.to_numeric(values[numeric_column], errors="coerce")
    if "invoice_date" in values:
        values["invoice_date"] = pd.to_datetime(values["invoice_date"], format="mixed", errors="coerce")
    for column in values.columns:
        result[column] = values[column]

    # Duplicate identity uses every original workbook field, before trimming/casting.
    duplicate_members = source.duplicated(keep=False)
    duplicate_later_rows = source.duplicated(keep="first")
    result["exact_duplicate_flag"] = duplicate_members.astype(bool)
    result["description_missing_flag"] = result.get("description", pd.Series(pd.NA, index=result.index)).isna()
    result["customer_id_missing_flag"] = result.get("customer_id", pd.Series(pd.NA, index=result.index)).isna()

    invoice = result.get("invoice_no", pd.Series(pd.NA, index=result.index, dtype="string"))
    quantity = pd.to_numeric(result.get("quantity", pd.Series(pd.NA, index=result.index)), errors="coerce")
    price = pd.to_numeric(result.get("unit_price", pd.Series(pd.NA, index=result.index)), errors="coerce")
    invoice_date = result.get("invoice_date", pd.Series(pd.NaT, index=result.index))

    result["cancellation_coded_flag"] = invoice.astype("string").str.strip().str.upper().str.startswith("C", na=False)
    result["negative_quantity_flag"] = quantity.lt(0)
    result["zero_quantity_flag"] = quantity.eq(0)
    result["zero_unit_price_flag"] = price.eq(0)
    result["negative_unit_price_flag"] = price.lt(0)
    result["invalid_quantity_flag"] = quantity.isna() | ~quantity.map(lambda value: pd.notna(value) and abs(value) != float("inf"))
    result["invalid_unit_price_flag"] = price.isna() | ~price.map(lambda value: pd.notna(value) and abs(value) != float("inf"))
    result["invalid_invoice_date_flag"] = invoice_date.isna()
    missing_required = pd.Series(False, index=result.index)
    for column in REQUIRED:
        missing_required |= result.get(column, pd.Series(pd.NA, index=result.index)).isna()
    result["required_field_review_flag"] = missing_required

    # Cancellation precedence makes a C-coded line identifiable even when quantity < 0.
    # Other negative quantities precede zero-price lines; invalid/missing required data
    # and negative prices take top priority as records requiring review.
    review = result["required_field_review_flag"] | result["negative_unit_price_flag"] | result["zero_quantity_flag"]
    status = pd.Series("normal_sale", index=result.index, dtype="string")
    status.loc[result["zero_unit_price_flag"]] = "zero_price_transaction"
    status.loc[result["negative_quantity_flag"]] = "other_negative_quantity_transaction"
    status.loc[result["cancellation_coded_flag"]] = "cancellation_coded_transaction"
    status.loc[review] = "requires_review"
    result["transaction_status"] = status

    # A signed extended line amount is valid for finite numeric operands, including
    # negative values. Negative quantity or price therefore affects the sign.
    valid_operands = quantity.notna() & price.notna() & quantity.map(lambda x: pd.notna(x) and abs(x) != float("inf")) & price.map(lambda x: pd.notna(x) and abs(x) != float("inf"))
    result["line_total"] = (quantity * price).where(valid_operands)

    eligible = quantity.gt(0) & price.gt(0) & ~result["cancellation_coded_flag"]
    sales = result.loc[eligible].copy()

    flags = [
        "description_missing_flag", "customer_id_missing_flag", "exact_duplicate_flag",
        "cancellation_coded_flag", "negative_quantity_flag", "zero_quantity_flag",
        "zero_unit_price_flag", "negative_unit_price_flag", "invalid_quantity_flag",
        "invalid_unit_price_flag", "invalid_invoice_date_flag", "required_field_review_flag",
    ]
    before_missing = {str(column): int(source[column].isna().sum()) for column in original_columns}
    after_missing = {canonical_name(column): int(result[canonical_name(column)].isna().sum()) for column in original_columns}
    exclusion_masks = {
        "quantity not positive": ~quantity.gt(0),
        "unit price not positive": ~price.gt(0),
        "cancellation-coded invoice": result["cancellation_coded_flag"],
    }
    exclusion_counts = {reason: int(mask.sum()) for reason, mask in exclusion_masks.items()}
    stats: dict[str, Any] = {
        "worksheet": worksheet,
        "input_rows": len(source),
        "cleaned_rows": len(result),
        "sales_rows": len(sales),
        "duplicate_later_rows": int(duplicate_later_rows.sum()),
        "duplicate_flagged_rows": int(duplicate_members.sum()),
        "missing_before": before_missing,
        "missing_after": after_missing,
        "flags": {flag: int(result[flag].sum()) for flag in flags},
        "exclusion_counts_overlap": exclusion_counts,
        "invalid_numeric_rows": int((result["invalid_quantity_flag"] | result["invalid_unit_price_flag"]).sum()),
        "invalid_dates": int(result["invalid_invoice_date_flag"].sum()),
        "line_total_missing": int(result["line_total"].isna().sum()),
        "eligible_sales_rows": int(eligible.sum()),
        "validation": {
            "row_count_preserved": len(result) == len(source),
            "source_row_id_unique": bool(result["source_row_id"].is_unique),
            "source_row_sequence_valid": result["source_excel_row"].tolist() == list(range(excel_start_row, excel_start_row + len(source))),
            "canonical_required_columns_present": all(column in result.columns for column in REQUIRED),
            "line_total_matches_operands": bool((result.loc[valid_operands, "line_total"] == (quantity * price).loc[valid_operands]).all()),
        },
    }
    return result.reset_index(drop=True), sales.reset_index(drop=True), stats


def _table(headers: list[str], rows: list[list[Any]]) -> str:
    if not rows:
        return "_No rows._"
    lines = ["| " + " | ".join(headers) + " |", "| " + " | ".join("---" for _ in headers) + " |"]
    lines.extend("| " + " | ".join(str(cell).replace("|", "\\|") for cell in row) + " |" for row in rows)
    return "\n".join(lines)


def build_report(input_path: Path, stats: list[dict[str, Any]]) -> str:
    """Render worksheet-level pipeline results and assumptions as Markdown."""
    lines = [
        "# Data Cleaning Pipeline Report", "",
        f"- **Input:** `{input_path}` (read-only)",
        f"- **Cleaned output:** `{DEFAULT_CLEANED}`",
        f"- **Sales subset:** `{DEFAULT_SALES}`",
        "- **Processing:** worksheet-by-worksheet; outputs are recreated on each run.", "",
        "## Row counts by worksheet", "",
        _table(["Worksheet", "Input", "Cleaned", "Sales analysis ready", "Excluded from sales subset"], [
            [s["worksheet"], f'{s["input_rows"]:,}', f'{s["cleaned_rows"]:,}', f'{s["sales_rows"]:,}', f'{s["input_rows"] - s["sales_rows"]:,}'] for s in stats
        ]), "",
        "## Missing values before and after standardization", "",
    ]
    for item in stats:
        lines.extend([f'### {item["worksheet"]}', "", _table(["Source column", "Missing before", "Missing after"], [
            [column, f'{count:,}', f'{item["missing_after"].get(canonical_name(column), 0):,}'] for column, count in item["missing_before"].items()
        ]), ""])
    lines.extend(["## Quality flags", "", "Counts are row counts; flags may overlap.", ""])
    flag_names = list(stats[0]["flags"]) if stats else []
    lines.extend([_table(["Flag", *[s["worksheet"] for s in stats]], [[flag, *[f'{s["flags"][flag]:,}' for s in stats]] for flag in flag_names]), ""])
    lines.extend(["## Sales subset exclusion reasons", "", "These reason counts overlap because a row can fail more than one criterion. Eligibility requires positive quantity, positive unit price, and a non-cancellation invoice. The subset is gross positive sales lines, not verified net revenue.", ""])
    lines.extend([_table(["Worksheet", "Quantity not positive", "Unit price not positive", "Cancellation-coded"], [[s["worksheet"], *[f'{s["exclusion_counts_overlap"][key]:,}' for key in ("quantity not positive", "unit price not positive", "cancellation-coded invoice")]] for s in stats]), ""])
    lines.extend(["## Duplicate treatment", "", "Exact duplicates are detected across all original source fields, within each worksheet. `exact_duplicate_flag` marks every member of a repeated group. The report distinguishes flagged rows from duplicate occurrences after the first. No rows are deduplicated, including in the sales subset.", ""])
    lines.extend([_table(["Worksheet", "Rows flagged as duplicate members", "Duplicate occurrences after first", "Removed"], [[s["worksheet"], f'{s["duplicate_flagged_rows"]:,}', f'{s["duplicate_later_rows"]:,}', "0"] for s in stats]), ""])
    lines.extend(["## Data validation", "", _table(["Worksheet", "Rows preserved", "Source IDs unique", "Source rows valid", "Required columns present", "Line totals valid", "Invalid numeric rows", "Invalid/missing dates", "Rows without line_total"], [[s["worksheet"], *["PASS" if s["validation"][key] else "FAIL" for key in ("row_count_preserved", "source_row_id_unique", "source_row_sequence_valid", "canonical_required_columns_present", "line_total_matches_operands")], f'{s["invalid_numeric_rows"]:,}', f'{s["invalid_dates"]:,}', f'{s["line_total_missing"]:,}'] for s in stats]), ""])
    lines.extend([
        "All validation checks must pass for each worksheet; a failed invariant raises an error before report completion. Invalid numeric or date values are retained and flagged; the numeric/date counts above are observed input findings.", "",
        "## Business rules and limitations", "",
        "- Original values are retained in `source_*` audit columns. Canonical fields are standardized, and `source_worksheet`, `source_excel_row`, and `source_row_id` preserve origin and row traceability.",
        "- Invoice numbers and stock codes are strings. Customer IDs remain missing when absent; no identifier is inferred. Descriptions remain missing and receive `description_missing_flag`.",
        "- Date strings are parsed to datetimes; unparseable or missing dates are retained as missing and flagged. Surrounding whitespace is stripped from canonical text fields only.",
        "- `transaction_status` precedence: required-field/invalid-value review, zero quantity, or negative unit price -> `requires_review`; otherwise cancellation-coded invoice -> `cancellation_coded_transaction`; otherwise negative quantity -> `other_negative_quantity_transaction`; otherwise zero unit price -> `zero_price_transaction`; otherwise `normal_sale`. Thus cancellation codes take precedence over negative quantity; review takes precedence over all operational labels.",
        "- Exact duplicate flags do not remove rows. Negative quantities can be returns, adjustments, or other events; cancellation codes do not prove a refund. Zero/negative prices are flagged and not imputed.",
        "- `line_total` is quantity * unit price when both parse to finite numbers, including zero/negative operands. Its sign is the arithmetic sign of those inputs; it is not verified net revenue.",
        "- Sales-subset exclusion reason counts overlap; they are not additive. Eligibility uses only positive quantity, positive price, and a non-cancellation invoice. Missing descriptions/customer IDs and duplicates are preserved if the three eligibility conditions pass.",
        "- Currency precision follows the workbook's floating-point price values; the pipeline does not apply currency rounding. Business meaning, invoice-level netting, and refund status require domain review.", "",
    ])
    return "\n".join(lines)


def run_pipeline(input_path: Path, cleaned_path: Path, sales_path: Path, report_path: Path) -> list[dict[str, Any]]:
    """Process each sheet in turn to limit peak memory and replace outputs deterministically."""
    if not input_path.is_file():
        raise FileNotFoundError(f"Input workbook not found: {input_path}")
    for path in (cleaned_path, sales_path, report_path):
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists():
            path.unlink()
    stats: list[dict[str, Any]] = []
    workbook = pd.ExcelFile(input_path, engine="openpyxl")
    for sheet in workbook.sheet_names:
        # Read identifier columns as strings to prevent loss of leading zeros.
        frame = pd.read_excel(workbook, sheet_name=sheet, dtype={"Invoice": "string", "InvoiceNo": "string", "StockCode": "string"})
        cleaned, sales, item = process_worksheet(frame, sheet)
        failed_checks = [name for name, passed in item["validation"].items() if not passed]
        if failed_checks:
            raise RuntimeError(f"Validation failed for worksheet {sheet!r}: {', '.join(failed_checks)}")
        cleaned.to_csv(cleaned_path, mode="a", header=not cleaned_path.exists(), index=False, date_format="%Y-%m-%d %H:%M:%S", encoding="utf-8")
        sales.to_csv(sales_path, mode="a", header=not sales_path.exists(), index=False, date_format="%Y-%m-%d %H:%M:%S", encoding="utf-8")
        stats.append(item)
    workbook.close()
    report_path.write_text(build_report(input_path, stats), encoding="utf-8")
    return stats


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--cleaned-output", type=Path, default=DEFAULT_CLEANED)
    parser.add_argument("--sales-output", type=Path, default=DEFAULT_SALES)
    parser.add_argument("--report", type=Path, default=DEFAULT_REPORT)
    args = parser.parse_args()
    stats = run_pipeline(args.input, args.cleaned_output, args.sales_output, args.report)
    print(f"Processed {sum(s['input_rows'] for s in stats):,} source rows across {len(stats)} worksheets.")
    print(f"Cleaned records: {args.cleaned_output}")
    print(f"Sales analysis subset: {args.sales_output}")
    print(f"Report: {args.report}")


if __name__ == "__main__":
    main()
