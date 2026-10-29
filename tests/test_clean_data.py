import pandas as pd
import pytest

from src.clean_data import process_worksheet, run_pipeline


@pytest.fixture
def transactions():
    return pd.DataFrame(
        {
            "Invoice": ["00123", " C00124 ", "00125", "00126", "00126", "00127", "00128"],
            "StockCode": ["001A", "002B", "003C", "004D", "004D", "005E", "006F"],
            "Description": [" Item ", None, "Cancel", "Duplicate", "Duplicate", "Zero", "Bad price"],
            "Quantity": [2, -1, -2, 1, 1, 1, 1],
            "InvoiceDate": ["2011-01-01"] * 7,
            "Price": [3.5, 4.0, 5.0, 2.0, 2.0, 0.0, -1.0],
            "Customer ID": [None, None, "42", "42", "42", "43", "44"],
            "Country": [" UK "] * 7,
        }
    )


def test_row_preservation_traceability_and_source_values(transactions):
    cleaned, _, stats = process_worksheet(transactions, "Synthetic")
    assert len(cleaned) == len(transactions) == stats["cleaned_rows"]
    assert cleaned["source_row_id"].is_unique
    assert cleaned["source_excel_row"].tolist() == list(range(2, 9))
    assert cleaned["source_worksheet"].eq("Synthetic").all()
    assert cleaned.loc[0, "source_invoice_no"] == "00123"
    assert cleaned.loc[0, "invoice_no"] == "00123"
    assert cleaned.loc[0, "stock_code"] == "001A"
    assert cleaned.loc[0, "source_country"] == " UK "
    assert cleaned.loc[0, "country"] == "UK"


def test_missing_ids_and_descriptions_are_preserved(transactions):
    cleaned, _, _ = process_worksheet(transactions, "Synthetic")
    assert pd.isna(cleaned.loc[0, "customer_id"])
    assert pd.isna(cleaned.loc[1, "customer_id"])
    assert cleaned["customer_id_missing_flag"].sum() == 2
    assert pd.isna(cleaned.loc[1, "description"])
    assert cleaned.loc[1, "description_missing_flag"]


def test_status_classification_and_precedence(transactions):
    cleaned, _, _ = process_worksheet(transactions, "Synthetic")
    assert cleaned["transaction_status"].tolist() == [
        "normal_sale",
        "cancellation_coded_transaction",  # cancellation beats negative quantity
        "other_negative_quantity_transaction",
        "normal_sale",
        "normal_sale",
        "zero_price_transaction",
        "requires_review",  # negative price takes review precedence
    ]
    assert cleaned.loc[1, "negative_quantity_flag"]
    assert cleaned.loc[1, "cancellation_coded_flag"]


def test_duplicates_are_flagged_but_retained(transactions):
    cleaned, _, stats = process_worksheet(transactions, "Synthetic")
    assert len(cleaned) == len(transactions)
    assert cleaned["exact_duplicate_flag"].sum() == 2
    assert stats["duplicate_later_rows"] == 1


def test_sales_filter_and_signed_financial_calculation(transactions):
    cleaned, sales, stats = process_worksheet(transactions, "Synthetic")
    assert sales["invoice_no"].tolist() == ["00123", "00126", "00126"]
    assert sales["line_total"].tolist() == [7.0, 2.0, 2.0]
    assert len(cleaned) - len(sales) == 4
    assert stats["eligible_sales_rows"] == 3
    assert cleaned.loc[1, "line_total"] == -4.0
    assert cleaned.loc[2, "line_total"] == -10.0
    assert pd.isna(cleaned.loc[0, "customer_id"])


def test_missing_and_non_numeric_required_values_are_reviewed():
    frame = pd.DataFrame(
        {"Invoice": ["1", "2"], "StockCode": ["A", "B"], "Quantity": ["bad", 1],
         "InvoiceDate": ["not a date", "2011-01-01"], "Price": [2, None], "Description": [None, "x"]}
    )
    cleaned, sales, _ = process_worksheet(frame, "Errors")
    assert cleaned["transaction_status"].eq("requires_review").all()
    assert cleaned["line_total"].isna().all()
    assert sales.empty


def test_pipeline_rerun_replaces_outputs_deterministically(tmp_path, transactions):
    workbook = tmp_path / "fixture.xlsx"
    with pd.ExcelWriter(workbook) as writer:
        transactions.iloc[:3].to_excel(writer, sheet_name="Year A", index=False)
        transactions.iloc[3:].to_excel(writer, sheet_name="Year B", index=False)
    cleaned_path = tmp_path / "out" / "cleaned.csv"
    sales_path = tmp_path / "out" / "sales.csv"
    report_path = tmp_path / "reports" / "cleaning.md"

    run_pipeline(workbook, cleaned_path, sales_path, report_path)
    first_outputs = (cleaned_path.read_bytes(), sales_path.read_bytes(), report_path.read_bytes())
    run_pipeline(workbook, cleaned_path, sales_path, report_path)
    second_outputs = (cleaned_path.read_bytes(), sales_path.read_bytes(), report_path.read_bytes())

    assert first_outputs == second_outputs
    assert len(pd.read_csv(cleaned_path)) == len(transactions)

