"""Read-only initial quality assessment for the Online Retail II workbook."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = PROJECT_ROOT / "data" / "online_retail_II.xlsx"
DEFAULT_OUTPUT = PROJECT_ROOT / "reports" / "initial_data_quality_report.md"
EXPECTED_START = pd.Timestamp("2009-12-01")
EXPECTED_END = pd.Timestamp("2011-12-09 23:59:59")
SAMPLE_SIZE = 5


def display(value: Any) -> str:
    """Format values for readable Markdown tables."""
    if pd.isna(value):
        return "(missing)"
    return str(value).replace("|", "\\|").replace("\n", " ")


def markdown_table(frame: pd.DataFrame) -> str:
    """Render a DataFrame as a Markdown table, including empty frames."""
    if frame.empty:
        return "_None found._"
    headers = [str(column) for column in frame.columns]
    lines = ["| " + " | ".join(headers) + " |", "| " + " | ".join("---" for _ in headers) + " |"]
    for row in frame.itertuples(index=False, name=None):
        lines.append("| " + " | ".join(display(value) for value in row) + " |")
    return "\n".join(lines)


def suspicious_sample(frame: pd.DataFrame, mask: pd.Series) -> pd.DataFrame:
    """Return a small sample with original Excel row numbers."""
    sample = frame.loc[mask].head(SAMPLE_SIZE).copy()
    if sample.empty:
        return sample
    sample.insert(0, "Excel row", sample.index + 2)
    return sample.reset_index(drop=True)


def inspect_sheet(name: str, frame: pd.DataFrame) -> tuple[list[str], dict[str, int]]:
    lines = [f"## Worksheet: {name}", "", f"- **Dimensions:** {len(frame):,} data rows x {len(frame.columns):,} columns"]
    lines.extend(["", "### Columns and inferred data types", ""])
    columns = pd.DataFrame({"Column": frame.columns, "Inferred dtype": [str(dtype) for dtype in frame.dtypes]})
    lines.extend([markdown_table(columns), "", "### Missing values", ""])
    missing = pd.DataFrame({"Column": frame.columns, "Missing count": frame.isna().sum().values})
    missing["Missing %"] = (missing["Missing count"] / len(frame) * 100).round(3) if len(frame) else 0.0
    lines.extend([markdown_table(missing), "", f"### Exact duplicate rows\n\n**{int(frame.duplicated(keep='first').sum()):,}** rows duplicate an earlier row across all columns.", ""])

    numeric_columns = frame.select_dtypes(include="number").columns
    numeric_rows: list[dict[str, Any]] = []
    masks: dict[str, pd.Series] = {}
    for column in numeric_columns:
        values = frame[column]
        neg_mask, zero_mask = values.lt(0), values.eq(0)
        numeric_rows.append({"Column": column, "Negative count": int(neg_mask.sum()), "Zero count": int(zero_mask.sum())})
        if neg_mask.any():
            masks[f"Negative {column}"] = neg_mask
        if zero_mask.any():
            masks[f"Zero {column}"] = zero_mask
    lines.extend(["### Negative and zero numeric values", "", markdown_table(pd.DataFrame(numeric_rows, columns=["Column", "Negative count", "Zero count"])), ""])

    stats: dict[str, int] = {"rows": len(frame), "duplicates": int(frame.duplicated(keep="first").sum())}
    if "InvoiceDate" in frame.columns:
        parsed = pd.to_datetime(frame["InvoiceDate"], errors="coerce")
        invalid = frame["InvoiceDate"].notna() & parsed.isna()
        outside = parsed.notna() & ((parsed < EXPECTED_START) | (parsed > EXPECTED_END))
        missing_date = frame["InvoiceDate"].isna()
        lines.extend(["### Date validity", "", f"Expected source coverage: {EXPECTED_START.date()} through {EXPECTED_END.date()} (inclusive).", "", markdown_table(pd.DataFrame([
            {"Check": "Missing InvoiceDate", "Count": int(missing_date.sum())},
            {"Check": "Non-missing values that fail datetime parsing", "Count": int(invalid.sum())},
            {"Check": "Parsed dates outside expected source coverage", "Count": int(outside.sum())},
        ])), ""])
        if invalid.any(): masks["Unparseable InvoiceDate"] = invalid
        if outside.any(): masks["InvoiceDate outside expected coverage"] = outside
        stats.update({"invalid_dates": int(invalid.sum()), "out_of_range_dates": int(outside.sum())})
        if parsed.notna().any():
            lines.append(f"Parsed date range observed: {parsed.min()} to {parsed.max()}.")
            lines.append("")

    invoice_column = next((column for column in ("InvoiceNo", "Invoice") if column in frame.columns), None)
    if invoice_column is not None:
        cancellation = frame[invoice_column].astype("string").str.strip().str.upper().str.startswith("C", na=False)
        lines.extend(["### Cancellation-coded invoices", "", f"Rows with {invoice_column} beginning with `C` (case-insensitive): **{int(cancellation.sum()):,}**. This is a potential business event, not an automatic error.", ""])
        if cancellation.any(): masks["Cancellation-coded invoice"] = cancellation
        stats["cancellations"] = int(cancellation.sum())

    if "Quantity" in frame.columns:
        for label, mask in (("Negative Quantity", frame["Quantity"].lt(0)), ("Zero Quantity", frame["Quantity"].eq(0))):
            if mask.any(): masks[label] = mask
    price_column = next((column for column in ("UnitPrice", "Price") if column in frame.columns), None)
    if price_column is not None:
        for label, mask in ((f"Negative {price_column}", frame[price_column].lt(0)), (f"Zero {price_column}", frame[price_column].eq(0))):
            if mask.any(): masks[label] = mask

    lines.extend(["### Samples of suspicious records", "", "Samples are capped at five rows per check. Row numbers refer to the worksheet including its header (header is row 1). Samples are for review only.", ""])
    if not masks:
        lines.extend(["_No records matched the implemented review checks._", ""])
    for label, mask in masks.items():
        lines.extend([f"#### {label} ({int(mask.sum()):,} rows)", "", markdown_table(suspicious_sample(frame, mask)), ""])
    return lines, stats


def build_report(input_path: Path) -> str:
    """Read every worksheet and assemble a Markdown assessment."""
    workbook = pd.ExcelFile(input_path, engine="openpyxl")
    lines = ["# Initial Data Quality Assessment", "", f"- **Input workbook:** `{input_path}`", f"- **Worksheets:** {len(workbook.sheet_names)}", "- **Assessment scope:** read-only inspection; no records were changed.", "- **Interpretation:** negative quantities and cancellation-coded invoices are flagged as potential legitimate business events.", ""]
    for sheet_name in workbook.sheet_names:
        frame = pd.read_excel(workbook, sheet_name=sheet_name)
        section, _ = inspect_sheet(sheet_name, frame)
        lines.extend(section)
    lines.extend(["## Limitations", "", "Date checks flag parsing failures and values outside the UCI-published coverage dates. They do not establish whether a transaction is commercially valid. Numeric checks cover columns inferred by pandas as numeric. Suspicious samples are not an exhaustive listing.", ""])
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="Source Excel workbook (read-only).")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Markdown report destination.")
    args = parser.parse_args()
    if not args.input.is_file():
        raise FileNotFoundError(f"Input workbook not found: {args.input}")
    report = build_report(args.input)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(report, encoding="utf-8")
    print(f"Report written to {args.output}")


if __name__ == "__main__":
    main()
