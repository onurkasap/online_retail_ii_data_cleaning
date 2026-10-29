# Auditable Retail Data Cleaning with Python and Pandas

**Self-directed portfolio project using a public dataset. This was not a paid client engagement.**

## Business challenge

Prepare a two-worksheet retail transaction workbook for transparent downstream analysis without erasing legitimate business events. Missing customer details, repeated rows, cancellation codes and signed quantities/prices make an unqualified sales total unreliable.

## My role

I designed and implemented the data assessment, reusable cleaning workflow, explicit transaction classification, audit fields, validation tests and client delivery documentation for this portfolio exercise.

## Tools

Python, Pandas, openpyxl, pytest, Matplotlib, Pillow and pypdf. Excel/CSV deliverables and Markdown documentation support handoff and reproducibility.

## Approach

1. Inspected both source worksheets and documented quality observations.
2. Standardized headers, whitespace, identifiers, numerics and dates while retaining source read values and original row references.
3. Added independent flags and explicit status precedence for cancellations, other negative quantities, zero prices and review conditions.
4. Preserved missing customer IDs/descriptions and exact duplicates instead of silently filling or removing them.
5. Produced a complete cleaned view and a separately defined gross positive sales subset.
6. Verified output readability, row reconciliation and lineage, and added synthetic pytest coverage and reproducible delivery materials.

## Verified results

- 1,067,371 source rows preserved across two worksheets, with unique source row identifiers in the existing final verification.
- 1,041,670 rows meet the positive-quantity, positive-price, non-cancellation subset rule; 25,701 rows are excluded only from that view and retained in the full cleaned file.
- 243,007 missing customer IDs and 4,382 missing descriptions retained.
- 23,430 duplicate-group members flagged; 12,133 duplicate occurrences after the first, with no automatic deduplication.
- 7 pytest tests passed, 0 failed, 0 skipped in the recorded verification.
- No measured claim of revenue uplift, cost savings, correction rate or client business impact.

## Deliverables

- [Client summary PDF](../reports/client_summary.pdf) and [quality comparison chart](../reports/data_quality_comparison.png).
- [Sample workbook](sample_cleaned_data.xlsx) with 805 stratified rows, flags, lineage and explicit Excel text identifiers.
- [All-column data dictionary](data_dictionary.xlsx), [delivery guide](client_delivery_guide.md) and [artifact verification](artifact_verification.json).
- Reusable inspection, cleaning and portfolio generation scripts, plus existing validation reports and automated tests.

## Limitations

This is historical public data, with no claim of a commercial client relationship. The sample overrepresents rare cases and cannot estimate population metrics. Cancellations and negative quantities are potential business events, not verified refunds. Exact duplicates and missing customer IDs remain. The sales subset represents gross positive sales lines, not verified net revenue. Floating-point prices and line totals have no currency-rounding policy; financial reporting requires business review. CSV identifiers must be imported with explicit string dtypes.

## Dataset attribution

Chen, D. (2012). Online Retail II [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5CG6D. [UCI dataset](https://archive.ics.uci.edu/dataset/502/online+retail+ii); [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). Source data was standardized, annotated and sampled for this portfolio exercise.
