# Data Cleaning Pipeline Report

- **Input:** `data/online_retail_II.xlsx` (read-only)
- **Cleaned output:** `outputs/cleaned_transactions.csv`
- **Sales subset:** `outputs/sales_analysis_ready.csv`
- **Processing:** worksheet-by-worksheet; outputs are recreated on each run.

## Row counts by worksheet

| Worksheet | Input | Cleaned | Sales analysis ready | Excluded from sales subset |
| --- | --- | --- | --- | --- |
| Year 2009-2010 | 525,461 | 525,461 | 511,565 | 13,896 |
| Year 2010-2011 | 541,910 | 541,910 | 530,105 | 11,805 |

## Missing values before and after standardization

### Year 2009-2010

| Source column | Missing before | Missing after |
| --- | --- | --- |
| Invoice | 0 | 0 |
| StockCode | 0 | 0 |
| Description | 2,928 | 2,928 |
| Quantity | 0 | 0 |
| InvoiceDate | 0 | 0 |
| Price | 0 | 0 |
| Customer ID | 107,927 | 107,927 |
| Country | 0 | 0 |

### Year 2010-2011

| Source column | Missing before | Missing after |
| --- | --- | --- |
| Invoice | 0 | 0 |
| StockCode | 0 | 0 |
| Description | 1,454 | 1,454 |
| Quantity | 0 | 0 |
| InvoiceDate | 0 | 0 |
| Price | 0 | 0 |
| Customer ID | 135,080 | 135,080 |
| Country | 0 | 0 |

## Quality flags

Counts are row counts; flags may overlap.

| Flag | Year 2009-2010 | Year 2010-2011 |
| --- | --- | --- |
| description_missing_flag | 2,928 | 1,454 |
| customer_id_missing_flag | 107,927 | 135,080 |
| exact_duplicate_flag | 13,283 | 10,147 |
| cancellation_coded_flag | 10,206 | 9,288 |
| negative_quantity_flag | 12,326 | 10,624 |
| zero_quantity_flag | 0 | 0 |
| zero_unit_price_flag | 3,687 | 2,515 |
| negative_unit_price_flag | 3 | 2 |
| invalid_quantity_flag | 0 | 0 |
| invalid_unit_price_flag | 0 | 0 |
| invalid_invoice_date_flag | 0 | 0 |
| required_field_review_flag | 0 | 0 |

## Sales subset exclusion reasons

These reason counts overlap because a row can fail more than one criterion. Eligibility requires positive quantity, positive unit price, and a non-cancellation invoice. The subset is gross positive sales lines, not verified net revenue.

| Worksheet | Quantity not positive | Unit price not positive | Cancellation-coded |
| --- | --- | --- | --- |
| Year 2009-2010 | 12,326 | 3,690 | 10,206 |
| Year 2010-2011 | 10,624 | 2,517 | 9,288 |

## Duplicate treatment

Exact duplicates are detected across all original source fields, within each worksheet. `exact_duplicate_flag` marks every member of a repeated group. The report distinguishes flagged rows from duplicate occurrences after the first. No rows are deduplicated, including in the sales subset.

| Worksheet | Rows flagged as duplicate members | Duplicate occurrences after first | Removed |
| --- | --- | --- | --- |
| Year 2009-2010 | 13,283 | 6,865 | 0 |
| Year 2010-2011 | 10,147 | 5,268 | 0 |

## Data validation

| Worksheet | Rows preserved | Source IDs unique | Source rows valid | Required columns present | Line totals valid | Invalid numeric rows | Invalid/missing dates | Rows without line_total |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Year 2009-2010 | PASS | PASS | PASS | PASS | PASS | 0 | 0 | 0 |
| Year 2010-2011 | PASS | PASS | PASS | PASS | PASS | 0 | 0 | 0 |

All validation checks must pass for each worksheet; a failed invariant raises an error before report completion. Invalid numeric or date values are retained and flagged; the numeric/date counts above are observed input findings.

## Business rules and limitations

- Original values are retained in `source_*` audit columns. Canonical fields are standardized, and `source_worksheet`, `source_excel_row`, and `source_row_id` preserve origin and row traceability.
- Invoice numbers and stock codes are strings. Customer IDs remain missing when absent; no identifier is inferred. Descriptions remain missing and receive `description_missing_flag`.
- Date strings are parsed to datetimes; unparseable or missing dates are retained as missing and flagged. Surrounding whitespace is stripped from canonical text fields only.
- `transaction_status` precedence: required-field/invalid-value review, zero quantity, or negative unit price -> `requires_review`; otherwise cancellation-coded invoice -> `cancellation_coded_transaction`; otherwise negative quantity -> `other_negative_quantity_transaction`; otherwise zero unit price -> `zero_price_transaction`; otherwise `normal_sale`. Thus cancellation codes take precedence over negative quantity; review takes precedence over all operational labels.
- Exact duplicate flags do not remove rows. Negative quantities can be returns, adjustments, or other events; cancellation codes do not prove a refund. Zero/negative prices are flagged and not imputed.
- `line_total` is quantity * unit price when both parse to finite numbers, including zero/negative operands. Its sign is the arithmetic sign of those inputs; it is not verified net revenue.
- Sales-subset exclusion reason counts overlap; they are not additive. Eligibility uses only positive quantity, positive price, and a non-cancellation invoice. Missing descriptions/customer IDs and duplicates are preserved if the three eligibility conditions pass.
- Currency precision follows the workbook's floating-point price values; the pipeline does not apply currency rounding. Business meaning, invoice-level netting, and refund status require domain review.
