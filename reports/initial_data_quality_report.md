# Initial Data Quality Assessment

- **Input workbook:** `data/online_retail_II.xlsx`
- **Worksheets:** 2
- **Assessment scope:** read-only inspection; no records were changed.
- **Interpretation:** negative quantities and cancellation-coded invoices are flagged as potential legitimate business events.

## Worksheet: Year 2009-2010

- **Dimensions:** 525,461 data rows x 8 columns

### Columns and inferred data types

| Column | Inferred dtype |
| --- | --- |
| Invoice | object |
| StockCode | object |
| Description | object |
| Quantity | int64 |
| InvoiceDate | datetime64[us] |
| Price | float64 |
| Customer ID | float64 |
| Country | str |

### Missing values

| Column | Missing count | Missing % |
| --- | --- | --- |
| Invoice | 0 | 0.0 |
| StockCode | 0 | 0.0 |
| Description | 2928 | 0.557 |
| Quantity | 0 | 0.0 |
| InvoiceDate | 0 | 0.0 |
| Price | 0 | 0.0 |
| Customer ID | 107927 | 20.539 |
| Country | 0 | 0.0 |

### Exact duplicate rows

**6,865** rows duplicate an earlier row across all columns.

### Negative and zero numeric values

| Column | Negative count | Zero count |
| --- | --- | --- |
| Quantity | 12326 | 0 |
| Price | 3 | 3687 |
| Customer ID | 0 | 0 |

### Date validity

Expected source coverage: 2009-12-01 through 2011-12-09 (inclusive).

| Check | Count |
| --- | --- |
| Missing InvoiceDate | 0 |
| Non-missing values that fail datetime parsing | 0 |
| Parsed dates outside expected source coverage | 0 |

Parsed date range observed: 2009-12-01 07:45:00 to 2010-12-09 20:01:00.

### Cancellation-coded invoices

Rows with Invoice beginning with `C` (case-insensitive): **10,206**. This is a potential business event, not an automatic error.

### Samples of suspicious records

Samples are capped at five rows per check. Row numbers refer to the worksheet including its header (header is row 1). Samples are for review only.

#### Negative Quantity (12,326 rows)

| Excel row | Invoice | StockCode | Description | Quantity | InvoiceDate | Price | Customer ID | Country |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 180 | C489449 | 22087 | PAPER BUNTING WHITE LACE | -12 | 2009-12-01 10:33:00 | 2.95 | 16321.0 | Australia |
| 181 | C489449 | 85206A | CREAM FELT EASTER EGG BASKET | -6 | 2009-12-01 10:33:00 | 1.65 | 16321.0 | Australia |
| 182 | C489449 | 21895 | POTTING SHED SOW 'N' GROW SET | -4 | 2009-12-01 10:33:00 | 4.25 | 16321.0 | Australia |
| 183 | C489449 | 21896 | POTTING SHED TWINE | -6 | 2009-12-01 10:33:00 | 2.1 | 16321.0 | Australia |
| 184 | C489449 | 22083 | PAPER CHAIN KIT RETRO SPOT | -12 | 2009-12-01 10:33:00 | 2.95 | 16321.0 | Australia |

#### Negative Price (3 rows)

| Excel row | Invoice | StockCode | Description | Quantity | InvoiceDate | Price | Customer ID | Country |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 179405 | A506401 | B | Adjust bad debt | 1 | 2010-04-29 13:36:00 | -53594.36 | (missing) | United Kingdom |
| 276276 | A516228 | B | Adjust bad debt | 1 | 2010-07-19 11:24:00 | -44031.79 | (missing) | United Kingdom |
| 403474 | A528059 | B | Adjust bad debt | 1 | 2010-10-20 12:04:00 | -38925.87 | (missing) | United Kingdom |

#### Zero Price (3,687 rows)

| Excel row | Invoice | StockCode | Description | Quantity | InvoiceDate | Price | Customer ID | Country |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 265 | 489464 | 21733 | 85123a mixed | -96 | 2009-12-01 10:52:00 | 0.0 | (missing) | United Kingdom |
| 285 | 489463 | 71477 | short | -240 | 2009-12-01 10:52:00 | 0.0 | (missing) | United Kingdom |
| 286 | 489467 | 85123A | 21733 mixed | -192 | 2009-12-01 10:53:00 | 0.0 | (missing) | United Kingdom |
| 472 | 489521 | 21646 | (missing) | -50 | 2009-12-01 11:44:00 | 0.0 | (missing) | United Kingdom |
| 3116 | 489655 | 20683 | (missing) | -44 | 2009-12-01 17:26:00 | 0.0 | (missing) | United Kingdom |

#### Cancellation-coded invoice (10,206 rows)

| Excel row | Invoice | StockCode | Description | Quantity | InvoiceDate | Price | Customer ID | Country |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 180 | C489449 | 22087 | PAPER BUNTING WHITE LACE | -12 | 2009-12-01 10:33:00 | 2.95 | 16321.0 | Australia |
| 181 | C489449 | 85206A | CREAM FELT EASTER EGG BASKET | -6 | 2009-12-01 10:33:00 | 1.65 | 16321.0 | Australia |
| 182 | C489449 | 21895 | POTTING SHED SOW 'N' GROW SET | -4 | 2009-12-01 10:33:00 | 4.25 | 16321.0 | Australia |
| 183 | C489449 | 21896 | POTTING SHED TWINE | -6 | 2009-12-01 10:33:00 | 2.1 | 16321.0 | Australia |
| 184 | C489449 | 22083 | PAPER CHAIN KIT RETRO SPOT | -12 | 2009-12-01 10:33:00 | 2.95 | 16321.0 | Australia |

## Worksheet: Year 2010-2011

- **Dimensions:** 541,910 data rows x 8 columns

### Columns and inferred data types

| Column | Inferred dtype |
| --- | --- |
| Invoice | object |
| StockCode | object |
| Description | object |
| Quantity | int64 |
| InvoiceDate | datetime64[us] |
| Price | float64 |
| Customer ID | float64 |
| Country | str |

### Missing values

| Column | Missing count | Missing % |
| --- | --- | --- |
| Invoice | 0 | 0.0 |
| StockCode | 0 | 0.0 |
| Description | 1454 | 0.268 |
| Quantity | 0 | 0.0 |
| InvoiceDate | 0 | 0.0 |
| Price | 0 | 0.0 |
| Customer ID | 135080 | 24.927 |
| Country | 0 | 0.0 |

### Exact duplicate rows

**5,268** rows duplicate an earlier row across all columns.

### Negative and zero numeric values

| Column | Negative count | Zero count |
| --- | --- | --- |
| Quantity | 10624 | 0 |
| Price | 2 | 2515 |
| Customer ID | 0 | 0 |

### Date validity

Expected source coverage: 2009-12-01 through 2011-12-09 (inclusive).

| Check | Count |
| --- | --- |
| Missing InvoiceDate | 0 |
| Non-missing values that fail datetime parsing | 0 |
| Parsed dates outside expected source coverage | 0 |

Parsed date range observed: 2010-12-01 08:26:00 to 2011-12-09 12:50:00.

### Cancellation-coded invoices

Rows with Invoice beginning with `C` (case-insensitive): **9,288**. This is a potential business event, not an automatic error.

### Samples of suspicious records

Samples are capped at five rows per check. Row numbers refer to the worksheet including its header (header is row 1). Samples are for review only.

#### Negative Quantity (10,624 rows)

| Excel row | Invoice | StockCode | Description | Quantity | InvoiceDate | Price | Customer ID | Country |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 143 | C536379 | D | Discount | -1 | 2010-12-01 09:41:00 | 27.5 | 14527.0 | United Kingdom |
| 156 | C536383 | 35004C | SET OF 3 COLOURED  FLYING DUCKS | -1 | 2010-12-01 09:49:00 | 4.65 | 15311.0 | United Kingdom |
| 237 | C536391 | 22556 | PLASTERS IN TIN CIRCUS PARADE  | -12 | 2010-12-01 10:24:00 | 1.65 | 17548.0 | United Kingdom |
| 238 | C536391 | 21984 | PACK OF 12 PINK PAISLEY TISSUES  | -24 | 2010-12-01 10:24:00 | 0.29 | 17548.0 | United Kingdom |
| 239 | C536391 | 21983 | PACK OF 12 BLUE PAISLEY TISSUES  | -24 | 2010-12-01 10:24:00 | 0.29 | 17548.0 | United Kingdom |

#### Negative Price (2 rows)

| Excel row | Invoice | StockCode | Description | Quantity | InvoiceDate | Price | Customer ID | Country |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 299985 | A563186 | B | Adjust bad debt | 1 | 2011-08-12 14:51:00 | -11062.06 | (missing) | United Kingdom |
| 299986 | A563187 | B | Adjust bad debt | 1 | 2011-08-12 14:52:00 | -11062.06 | (missing) | United Kingdom |

#### Zero Price (2,515 rows)

| Excel row | Invoice | StockCode | Description | Quantity | InvoiceDate | Price | Customer ID | Country |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 624 | 536414 | 22139 | (missing) | 56 | 2010-12-01 11:52:00 | 0.0 | (missing) | United Kingdom |
| 1512 | 536545 | 21134 | (missing) | 1 | 2010-12-01 14:32:00 | 0.0 | (missing) | United Kingdom |
| 1987 | 536547 | 37509 | (missing) | 1 | 2010-12-01 14:33:00 | 0.0 | (missing) | United Kingdom |
| 1988 | 536546 | 22145 | (missing) | 1 | 2010-12-01 14:33:00 | 0.0 | (missing) | United Kingdom |
| 2024 | 536552 | 20950 | (missing) | 1 | 2010-12-01 14:34:00 | 0.0 | (missing) | United Kingdom |

#### Cancellation-coded invoice (9,288 rows)

| Excel row | Invoice | StockCode | Description | Quantity | InvoiceDate | Price | Customer ID | Country |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 143 | C536379 | D | Discount | -1 | 2010-12-01 09:41:00 | 27.5 | 14527.0 | United Kingdom |
| 156 | C536383 | 35004C | SET OF 3 COLOURED  FLYING DUCKS | -1 | 2010-12-01 09:49:00 | 4.65 | 15311.0 | United Kingdom |
| 237 | C536391 | 22556 | PLASTERS IN TIN CIRCUS PARADE  | -12 | 2010-12-01 10:24:00 | 1.65 | 17548.0 | United Kingdom |
| 238 | C536391 | 21984 | PACK OF 12 PINK PAISLEY TISSUES  | -24 | 2010-12-01 10:24:00 | 0.29 | 17548.0 | United Kingdom |
| 239 | C536391 | 21983 | PACK OF 12 BLUE PAISLEY TISSUES  | -24 | 2010-12-01 10:24:00 | 0.29 | 17548.0 | United Kingdom |

## Limitations

Date checks flag parsing failures and values outside the UCI-published coverage dates. They do not establish whether a transaction is commercially valid. Numeric checks cover columns inferred by pandas as numeric. Suspicious samples are not an exhaustive listing.
