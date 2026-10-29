# Client delivery guide

This package documents a self-directed portfolio project using a public dataset. It does not describe a paid client engagement or a measured business outcome.

## Files and purpose

| File | Purpose |
| --- | --- |
| `reports/client_summary.pdf` | Three-page English handoff summary with verified technical results and limitations. |
| `reports/data_quality_comparison.png` | Worksheet comparison of detected/flagged records; categories overlap and do not represent errors fixed. |
| `outputs/cleaned_transactions.csv` | All 1,067,371 original records, canonical values, source audit columns, lineage and flags. |
| `outputs/sales_analysis_ready.csv` | 1,041,670 gross positive sales lines. The 25,701 other records are excluded only here, never deleted from the cleaned file. |
| `portfolio/sample_cleaned_data.xlsx` | 805 rows covering both worksheets and available transaction statuses; all columns retained. Read the sampling worksheet first. |
| `portfolio/data_dictionary.xlsx` | All 33 actual output columns, recommended types, observed examples and business rules. |
| `reports/final_verification_report.md` | Previously executed full-data and pytest checks: 7 passed, 0 failed, 0 skipped. |
| `portfolio/artifact_verification.json` | Generated artifact checks, SHA-256 hashes and protected-file checks. |

Full CSV files are deliberately excluded from the portfolio folder and from Git. The source workbook is not included in the delivery sample.

## Run and reproduce

From the project root, use the already configured environment:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe src\inspect_data.py
.\.venv\Scripts\python.exe src\clean_data.py
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe src\build_portfolio.py
```

The inspection and cleaning commands are for reproducing the underlying outputs when needed. To rebuild only this package from existing CSVs and reports, run only `build_portfolio.py`. It scans outputs in chunks and does not invoke cleaning or alter the source workbook. Default artifact paths are relative to the project root. Repeat builds with the same inputs and dependency environment produce the same sample and normalized artifact bytes; rendering may differ across library/font versions.

## Correctly read CSV identifiers and dates

```python
import pandas as pd

ids = {
    'invoice_no': 'string', 'stock_code': 'string', 'customer_id': 'string',
    'source_invoice_no': 'string', 'source_stock_code': 'string',
    'source_customer_id': 'string', 'source_row_id': 'string',
}
df = pd.read_csv('outputs/cleaned_transactions.csv', dtype=ids)
df['invoice_date'] = pd.to_datetime(df['invoice_date'], errors='coerce')
```

CSV stores no type metadata. Default readers may infer sales invoice IDs as integers and customer IDs as floats. Treat identifiers as labels, not arithmetic values. The sample XLSX explicitly stores identifier cells as Excel text, including source audit versions. Missing values remain blank. Source customer IDs may have a `.0` textual suffix because the original Pandas read represented numeric IDs as floats; canonical customer IDs are normalized by the existing pipeline. Never invent customer IDs.

## Interpret flags and statuses

Flags overlap. The dictionary explains each field. Missing customer IDs and descriptions are retained and flagged. Exact duplicates are tested within each worksheet across original fields; every group member is flagged, with no automatic deletion. The initial assessment counted 12,133 duplicate occurrences after the first; 23,430 rows are members of those groups.

Status precedence: missing required fields, zero quantity or negative unit price -> `requires_review`; otherwise C-coded invoice -> `cancellation_coded_transaction`; otherwise negative quantity -> `other_negative_quantity_transaction`; otherwise zero unit price -> `zero_price_transaction`; otherwise `normal_sale`. The status is a classification, not proof of a refund or commercial validity. Individual flags still expose conditions hidden by precedence. The required-field flag itself covers missing canonical invoice, stock, quantity, date or price.

Sales eligibility is exactly quantity > 0, unit_price > 0 and not cancellation-coded. It retains eligible duplicates and missing customer IDs/descriptions, and does not independently validate business meaning or require date/identifier validity. All required values happened to pass the existing full-data checks; do not silently assume broader filtering rules for future inputs.

## Sampling and limitations

The sample selects up to 100 rows per (worksheet, transaction_status) stratum by the smallest SHA-256 hash of seed plus source_row_id. Seed: `online-retail-ii-portfolio-v1`. It caps the combined sample at 1,000 by the same priority, then sorts by worksheet and numeric Excel row. It illustrates available cases rather than matching population frequencies. The Read me worksheet records actual stratum counts. Do not calculate population totals from it.

The sales subset represents gross positive sales lines, not verified net revenue. Returns, credits and cancellations require domain interpretation and invoice-level reconciliation. Negative quantities can also be corrections or stock adjustments. line_total is a signed finite quantity-times-price calculation, with no accounting policy or currency rounding; floating-point precision applies. Duplicate retention can affect aggregates. The source audit fields preserve read values, not Excel cell formatting or formulas, and source row positions are stable only for the unchanged workbook.

Before financial use, agree on return/refund interpretation, duplicate handling and monetary precision. Before public release, retain attribution, keep large source/output data out of Git, and review project paths and author details. This package has not been uploaded or published.

## Attribution

Chen, D. (2012). Online Retail II [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5CG6D. Source: https://archive.ics.uci.edu/dataset/502/online+retail+ii. License: [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). This project standardizes, annotates and samples the public source data; no endorsement by the original creator or UCI is implied.
