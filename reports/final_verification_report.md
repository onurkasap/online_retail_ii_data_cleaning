# Final Verification Report

**Verification date:** 2026-10-29  
**Scope:** Existing pipeline and outputs only. No cleaning logic or source data was changed.

## Test results

Command: `.venv/Scripts/python.exe -m pytest -q`

| Passed | Failed | Skipped | Result |
| ---: | ---: | ---: | --- |
| 7 | 0 | 0 | PASS |

Pytest completed successfully with exit code 0.

## Output existence, readability, and row counts

Both files exist and were read in 100,000-row Pandas chunks without parsing errors.

| File | Bytes | Rows found | Expected rows | Match |
| --- | ---: | ---: | ---: | --- |
| `outputs/cleaned_transactions.csv` | 332,953,903 | 1,067,371 | 1,067,371 | PASS |
| `outputs/sales_analysis_ready.csv` | 324,777,487 | 1,041,670 | 1,041,670 | PASS |

## Source preservation and row lineage

The cleaned output contains the complete workbook row count by worksheet:

| Worksheet | Expected/input rows | Cleaned output rows | Excel row positions present |
| --- | ---: | ---: | --- |
| Year 2009-2010 | 525,461 | 525,461 | 2–525,462 |
| Year 2010-2011 | 541,910 | 541,910 | 2–541,911 |
| **Total** | **1,067,371** | **1,067,371** | — |

`source_row_id` is populated and unique across all 1,067,371 cleaned rows, and every value matches its `source_worksheet` and `source_excel_row` fields. Per worksheet, the output contains the expected number of unique source Excel row positions and their full contiguous ranges. This verifies that no source row is silently dropped from `cleaned_transactions.csv`. The sales subset contains fewer rows by design.

## Original workbook

The workbook was not written by the verification commands. Its current size and modification time match the values recorded after the prior pipeline run: **45,622,278 bytes**, modified **2023-05-22 15:20:22**. Current SHA-256: `BCBE73B35F5B7BABF197FB0CB983A11F5D9FF929078D4AA53D171B1F2DF2E980`.

## CSV type preservation

CSV files do not store column type metadata. Identifier values are serialized as text, but a consumer that relies on Pandas type inference can read them back as numeric types:

| File | Pandas default inferred type (sampled) | Explicit string dtype |
| --- | --- | --- |
| `cleaned_transactions.csv` — `invoice_no` | string/object | Readable as string |
| `cleaned_transactions.csv` — `stock_code` | string/object | Readable as string |
| `cleaned_transactions.csv` — `customer_id` | `float64` | Readable as string |
| `sales_analysis_ready.csv` — `invoice_no` | `int64` | Readable as string |
| `sales_analysis_ready.csv` — `stock_code` | string/object | Readable as string |
| `sales_analysis_ready.csv` — `customer_id` | `float64` | Readable as string |

To preserve identifiers as strings on import, pass explicit dtypes, for example:

```python
pd.read_csv(
    "outputs/cleaned_transactions.csv",
    dtype={"invoice_no": "string", "stock_code": "string", "customer_id": "string"},
)
```

The same dtype mapping applies to `sales_analysis_ready.csv`. Customer ID missing values remain blank/missing in the CSV; string dtype does not invent identifiers. The defaults observed above are a consumer-side inference risk, not evidence that source rows were lost.

## Risks and limitations to disclose to a freelance customer

- **Use explicit CSV dtypes.** Pandas' default inference reads sales invoice identifiers as integers and customer IDs as floats. Downstream tools, especially spreadsheets, may similarly reinterpret identifier columns. Provide the import dtype mapping above and avoid treating identifiers as quantities.
- **The analysis-ready file is not net revenue.** It includes only positive-quantity, positive-price, non-cancellation lines. It does not net returns, cancellations, credits, or adjustments against sales.
- **Business events are not adjudicated.** Negative quantities and cancellation-coded invoices are retained in the cleaned output, but the dataset alone does not establish whether any line is a refund, return, correction, or another event.
- **Duplicates are retained.** The pipeline flags exact duplicates rather than removing them. Downstream totals can be overstated if a user sums rows without reviewing duplicate flags.
- **Currency precision is not normalized.** Prices and line totals use floating-point arithmetic and are not rounded to currency precision by this pipeline.
- **Large files require suitable tooling.** The outputs are approximately 333 MB and 325 MB. Chunked or database-backed downstream processing may be preferable to loading both outputs into memory at once.
- **Tests use synthetic fixtures.** The 7 passing pytest tests cover pipeline behavior on small constructed examples. Full-data checks in this report independently verify the generated files' readability, row counts, worksheet lineage ranges, and source row ID uniqueness.

## Overall result

**PASS with a documented CSV schema caveat.** The tests pass, both outputs are readable and have the expected row counts, the cleaned output preserves all expected source rows and unique row identifiers, and the original workbook's recorded size and modification time are unchanged. No project logic was modified during verification.
