"""Build and verify delivery artifacts from existing CSVs; never run cleaning.

Sampling uses SHA-256 priority within worksheet/status strata, at most 100 rows
per stratum and 1,000 overall. Fixed document metadata and normalized XLSX ZIP
timestamps make repeat builds byte reproducible within the same environment.
"""

from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import re
import textwrap
import zipfile
from importlib.metadata import version

# The application cache path can be read-only in managed workspaces.
os.environ.setdefault("MPLCONFIGDIR", str(Path(__file__).resolve().parents[1] / "outputs/.matplotlib"))
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import FancyBboxPatch
from matplotlib.ticker import StrMethodFormatter
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd
from PIL import Image
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]
FIXED_TIME = datetime(2026, 10, 29, tzinfo=timezone.utc)
SOURCE = "https://archive.ics.uci.edu/dataset/502/online+retail+ii"
CITATION = "Chen, D. (2012). Online Retail II [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5CG6D"
SEED = "online-retail-ii-portfolio-v1"
NAVY, TEAL, BLUE, GREY = "#142C42", "#087F8C", "#668CA5", "#526576"
IDENTIFIERS = {"invoice_no", "stock_code", "customer_id", "source_invoice_no", "source_stock_code", "source_customer_id", "source_row_id"}
FLAGS = {
    "exact_duplicate_flag": "Every member of an exact repeated group, detected within worksheet across original fields; retained.",
    "description_missing_flag": "Canonical description missing after trimming; missing value retained.",
    "customer_id_missing_flag": "Customer ID missing; no identifier is invented.",
    "cancellation_coded_flag": "Trimmed invoice starts with C, case-insensitive; does not prove a refund.",
    "negative_quantity_flag": "Quantity below zero; may indicate returns, adjustments or another business event.",
    "zero_quantity_flag": "Quantity equals zero; transaction status requires review.",
    "zero_unit_price_flag": "Unit price equals zero; no price imputation.",
    "negative_unit_price_flag": "Unit price below zero; transaction status requires review; no imputation.",
    "invalid_quantity_flag": "Quantity is missing, unparseable or non-finite.",
    "invalid_unit_price_flag": "Unit price is missing, unparseable or non-finite.",
    "invalid_invoice_date_flag": "Invoice date is missing or unparseable; not a test of commercial validity.",
    "required_field_review_flag": "At least one required canonical field is missing (invoice, stock, quantity, date, price).",
}
RULES = {
    "invoice_no": "Trim surrounding whitespace; retain textual representation and leading zeros; C prefix flags cancellations.",
    "stock_code": "Trim surrounding whitespace; retain identifiers as strings, including alphanumeric codes.",
    "description": "Trim surrounding whitespace; blanks become missing; no filling.",
    "quantity": "Parse as numeric; retain zero and negative quantities and flag them.",
    "invoice_date": "Parse datetime with errors coerced to missing; invalid/missing values flagged.",
    "unit_price": "Numeric unit price in GBP; retain zero/negative values; do not impute.",
    "customer_id": "Nullable string; integral numeric representations normalized (e.g. 13085.0 to 13085); source retained.",
    "country": "Trim surrounding whitespace; no geographic reconciliation.",
    "source_worksheet": "Original worksheet label; identifies year origin.",
    "source_excel_row": "Original worksheet row, counting header as row 1; data begins at row 2.",
    "source_row_id": "Worksheet label + colon + original Excel row; stable for this source workbook's row order.",
    "transaction_status": "Precedence: required-field review/zero quantity/negative price; then cancellation; then other negative quantity; then zero price; else normal sale.",
    "line_total": "Quantity * unit_price only for finite numeric operands; signed, no currency rounding, not verified net revenue.",
}
DESCRIPTIONS = {
    "invoice_no": "Standardized invoice identifier", "stock_code": "Standardized product or adjustment code",
    "description": "Standardized product or transaction description", "quantity": "Signed units on transaction line",
    "invoice_date": "Transaction datetime without timezone", "unit_price": "Signed unit price in pounds sterling",
    "customer_id": "Customer identifier (nullable)", "country": "Customer country label",
    "source_worksheet": "Workbook worksheet of origin", "source_excel_row": "Immutable reference to original row position",
    "source_row_id": "Composite source record reference", "transaction_status": "Mutually exclusive operational classification",
    "line_total": "Arithmetic signed extended line amount in GBP",
}


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def collect(root: Path) -> tuple[pd.DataFrame, dict]:
    """Scan CSVs in bounded chunks for metrics, examples and sample candidates."""
    cleaned = root / "outputs/cleaned_transactions.csv"
    counts, statuses, flags, missing, examples = Counter(), Counter(), Counter(), Counter(), {}
    candidates: dict[tuple, pd.DataFrame] = {}
    schema = None
    total = 0
    for chunk in pd.read_csv(cleaned, dtype="string", keep_default_na=False, chunksize=50_000):
        schema = list(chunk.columns)
        if not set(FLAGS).issubset(schema):
            raise ValueError("Existing output schema lacks expected flags")
        total += len(chunk)
        counts.update(chunk["source_worksheet"].value_counts().to_dict())
        statuses.update(chunk["transaction_status"].value_counts().to_dict())
        for col in schema:
            missing[col] += int(chunk[col].eq("").sum())
            if col not in examples:
                nonempty = chunk.loc[chunk[col].ne(""), col]
                if not nonempty.empty:
                    examples[col] = nonempty.iloc[0]
        for col in FLAGS:
            flags[col] += int(chunk[col].eq("True").sum())
        # SHA-256 ordering is stable across Python versions and CSV chunk sizes.
        chunk["_priority"] = chunk["source_row_id"].map(lambda value: hashlib.sha256(f"{SEED}|{value}".encode()).hexdigest())
        for key, group in chunk.groupby(["source_worksheet", "transaction_status"], sort=True):
            selected = group.sort_values(["_priority", "source_row_id"]).head(100)
            if key in candidates:
                selected = pd.concat([candidates[key], selected], ignore_index=True)
            candidates[key] = selected.sort_values(["_priority", "source_row_id"]).head(100)
    if schema is None:
        raise ValueError("Cleaned CSV has no rows")
    sample = pd.concat([candidates[key] for key in sorted(candidates)], ignore_index=True)
    sample = sample.sort_values(["_priority", "source_row_id"]).head(1000)
    sample["_excel_row"] = pd.to_numeric(sample["source_excel_row"])
    sample = sample.sort_values(["source_worksheet", "_excel_row"]).drop(columns=["_priority", "_excel_row"]).reset_index(drop=True)
    sales_counts = Counter()
    for chunk in pd.read_csv(root / "outputs/sales_analysis_ready.csv", usecols=["source_worksheet"], dtype="string", chunksize=100_000):
        sales_counts.update(chunk["source_worksheet"].value_counts().to_dict())
    verification = (root / "reports/final_verification_report.md").read_text(encoding="utf-8-sig")
    result = re.search(r"\|\s*(\d+)\s*\|\s*(\d+)\s*\|\s*(\d+)\s*\|\s*PASS\s*\|", verification)
    if not result:
        raise ValueError("Cannot locate recorded pytest counts in final verification report")
    tests = dict(zip(["passed", "failed", "skipped"], map(int, result.groups())))
    initial = (root / "reports/initial_data_quality_report.md").read_text(encoding="utf-8-sig")
    duplicates_after_first = sum(int(n.replace(",", "")) for n in re.findall(r"\*\*([\d,]+)\*\* rows duplicate an earlier row", initial))
    if total != 1_067_371 or sum(sales_counts.values()) != 1_041_670:
        raise ValueError("CSV counts differ from the final verified delivery baseline")
    metrics = {"cleaned_rows": total, "sales_rows": sum(sales_counts.values()), "excluded": total - sum(sales_counts.values()),
               "by_worksheet": dict(sorted(counts.items())), "sales_by_worksheet": dict(sorted(sales_counts.items())),
               "flags": dict(flags), "missing": dict(missing), "examples": examples, "columns": schema,
               "statuses": dict(sorted(statuses.items())), "recorded_pytest": tests,
               "duplicate_occurrences_after_first": duplicates_after_first, "sample_rows": len(sample),
               "sample_strata": {f"{sheet} | {status}": len(group) for (sheet, status), group in sample.groupby(["source_worksheet", "transaction_status"], sort=True)}}
    return sample, metrics


def column_info(column: str) -> tuple[str, str, str, str]:
    """Return logical type, role, description and business rule for actual columns."""
    if column in FLAGS:
        return "boolean", "quality flag", FLAGS[column], FLAGS[column]
    if column.startswith("source_") and column not in {"source_worksheet", "source_excel_row", "source_row_id"}:
        canonical = column.removeprefix("source_")
        if canonical not in RULES:
            raise ValueError(f"Undocumented source column: {column}")
        return "source value / lexical text", "source audit", f"Original {canonical} value as read by the cleaning pipeline", "Retain CSV representation as text in sample. Workbook cell formatting is not reproduced; original workbook is authoritative."
    if column not in RULES:
        raise ValueError(f"Undocumented output column: {column}")
    logical = {"quantity": "numeric", "unit_price": "floating-point numeric", "line_total": "floating-point numeric", "source_excel_row": "integer", "invoice_date": "datetime (no timezone)"}.get(column, "nullable string" if column == "customer_id" else "string")
    role = "lineage" if column.startswith("source_") else "derived" if column in {"line_total", "transaction_status"} else "canonical"
    return logical, role, DESCRIPTIONS[column], RULES[column]


def style_sheet(sheet, widths: dict[int, float] | None = None):
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    sheet.sheet_view.showGridLines = False
    for cell in sheet[1]:
        cell.fill = PatternFill("solid", fgColor="142C42")
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    sheet.row_dimensions[1].height = 32
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.font = Font(name="Calibri", size=10, color="142C42")
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if cell.row % 2 == 0:
                cell.fill = PatternFill("solid", fgColor="F0F5F7")
            if isinstance(cell.value, str):
                # Explicit literal strings also prevent Excel from treating =text as a formula.
                cell.data_type = "s"
    for index in range(1, sheet.max_column + 1):
        sheet.column_dimensions[get_column_letter(index)].width = (widths or {}).get(index, 23)


def save_workbook(workbook, path: Path):
    """Normalize document and ZIP timestamps for deterministic XLSX bytes."""
    workbook.properties.creator = "Online Retail II portfolio project"
    workbook.properties.created = FIXED_TIME.replace(tzinfo=None)
    workbook.properties.modified = FIXED_TIME.replace(tzinfo=None)
    workbook.save(path)
    with zipfile.ZipFile(path) as archive:
        files = {item.filename: archive.read(item.filename) for item in archive.infolist()}
    # openpyxl updates dcterms:modified during save.
    files["docProps/core.xml"] = re.sub(rb"(<dcterms:modified[^>]*>).*?(</dcterms:modified>)", rb"\g<1>2026-10-29T00:00:00Z\2", files["docProps/core.xml"])
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for name, content in sorted(files.items()):
            info = zipfile.ZipInfo(name, (2026, 10, 29, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(info, content)


def build_spreadsheets(root: Path, sample: pd.DataFrame, metrics: dict):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Sample transactions"
    sheet.append(list(sample.columns))
    for values in sample.itertuples(index=False, name=None):
        parsed = []
        for column, value in zip(sample.columns, values):
            if value == "":
                parsed.append(None)
            elif column in FLAGS:
                parsed.append(value == "True")
            elif column == "source_excel_row":
                parsed.append(int(value))
            elif column in {"quantity", "unit_price", "line_total"}:
                parsed.append(float(value))
            elif column == "invoice_date":
                parsed.append(datetime.fromisoformat(value))
            else:
                parsed.append(value)
        sheet.append(parsed)
    style_sheet(sheet)
    sheet.freeze_panes = "C2"
    sheet.row_dimensions[1].height = 45
    for row in range(2, len(sample) + 2):
        sheet.row_dimensions[row].height = 42
    for name in ("source_description", "description"):
        sheet.column_dimensions[get_column_letter(list(sample.columns).index(name) + 1)].width = 45
    for column_index, column in enumerate(sample.columns, 1):
        for row in range(2, len(sample) + 2):
            cell = sheet.cell(row, column_index)
            if column in IDENTIFIERS or column.startswith("source_") and column != "source_excel_row":
                cell.number_format = "@"
                if cell.value is not None:
                    cell.data_type = "s"
            elif column == "invoice_date":
                cell.number_format = "yyyy-mm-dd hh:mm:ss"
            elif column in {"unit_price", "line_total"}:
                cell.number_format = "0.00############"
    notes = workbook.create_sheet("Read me")
    notes.append(["Topic", "Explanation"])
    notes.append(["Purpose", "Illustrative coverage sample. Rare statuses are overrepresented; do not estimate population totals or net revenue from this sample."])
    notes.append(["Input", "outputs/cleaned_transactions.csv; all output columns, flags and source lineage retained."])
    notes.append(["Method", f"Within each (source_worksheet, transaction_status) stratum, select up to 100 smallest SHA-256(seed|source_row_id) values. Seed: {SEED}. If needed, retain the 1,000 smallest priorities overall."])
    notes.append(["Order", "Output sorted by worksheet then numeric original Excel row; same inputs/environment reproduce the selection and workbook."])
    notes.append(["Sample size", f"{len(sample)} rows; population {metrics['cleaned_rows']:,} cleaned rows."])
    notes.append(["Cell types", "Invoice, stock and customer identifiers (including source audit forms) are explicit Excel text. Flags are booleans, source row integer, canonical quantity/price/line total numeric, canonical invoice date datetime. Missing values stay blank."])
    notes.append(["Source audit fields", "Original CSV lexical values retained in source_* cells. Source workbook formatting and cell types are not copied."])
    notes.append(["Attribution", CITATION])
    notes.append(["Source and license", f"{SOURCE} | CC BY 4.0 | https://creativecommons.org/licenses/by/4.0/"])
    notes.append(["Changes", "Source values standardized and annotated in the existing pipeline; records sampled for this public portfolio. No claim of source publisher endorsement."])
    notes.append(["Project status", "Self-directed portfolio project using public data; not a paid client engagement."])
    for stratum, count in metrics["sample_strata"].items():
        notes.append(["Stratum: " + stratum, str(count)])
    style_sheet(notes, {1: 40, 2: 110})
    for row in range(2, notes.max_row + 1):
        notes.row_dimensions[row].height = 65 if row in {2, 4, 7} else 45
    save_workbook(workbook, root / "portfolio/sample_cleaned_data.xlsx")

    dictionary = openpyxl.Workbook()
    sheet = dictionary.active
    sheet.title = "Data dictionary"
    sheet.append(["Column", "Logical type", "Role", "Description", "Example from CSV", "Business rule", "Missing in full cleaned CSV", "CSV import guidance"])
    for column in metrics["columns"]:
        logical, role, description, rule = column_info(column)
        guidance = "dtype='string'" if column in IDENTIFIERS or role in {"source audit", "lineage"} and column != "source_excel_row" else "parse_dates after reading" if column == "invoice_date" else "boolean" if column in FLAGS else "numeric" if column in {"source_excel_row", "quantity", "unit_price", "line_total"} else "dtype='string'"
        sheet.append([column, logical, role, description, metrics["examples"].get(column, "(all missing)"), rule, metrics["missing"][column], guidance])
    style_sheet(sheet, {1: 31, 2: 27, 3: 18, 4: 52, 5: 45, 6: 90, 7: 24, 8: 27})
    for row in range(2, sheet.max_row + 1):
        sheet.row_dimensions[row].height = 65
    notes = dictionary.create_sheet("Schema notes")
    for row in [["Topic", "Explanation"], ["Coverage", f"All {len(metrics['columns'])} columns, in actual cleaned CSV order; sales CSV uses the same column schema."], ["CSV typing", "CSV has no type metadata. Logical types are recommended import types; examples are literal values from existing output. Missing counts count empty CSV fields."], ["Examples", "First nonempty value encountered per column, not fabricated. Examples may be repeated across source and canonical fields."], ["Flags", "Flags overlap and indicate review conditions/business events, not necessarily errors. Duplicate flag covers every member of repeated groups within a sheet."], ["Status precedence", RULES["transaction_status"]], ["Sales eligibility", "quantity > 0, unit_price > 0 and cancellation_coded_flag is False. This rule alone does not require valid dates/identifiers; do not silently imply additional eligibility rules."], ["Attribution", CITATION], ["Source/license", f"{SOURCE} | CC BY 4.0"], ["Privacy/context", "Public historical dataset; no client relationship or business outcome is implied."]]:
        notes.append(row)
    style_sheet(notes, {1: 30, 2: 115})
    for row in range(2, notes.max_row + 1):
        notes.row_dimensions[row].height = 55
    save_workbook(dictionary, root / "portfolio/data_dictionary.xlsx")


def build_chart(root: Path, metrics: dict):
    # Use separate, explicitly labelled linear scales to keep low counts legible.
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.spines.top": False, "axes.spines.right": False, "axes.spines.left": False, "axes.edgecolor": "#CFD9E0", "pdf.fonttype": 42})
    figure = plt.figure(figsize=(12, 8.2), facecolor="white")
    figure.text(.06, .955, "Data quality observations", fontsize=23, weight="bold", color=NAVY)
    figure.text(.06, .915, "Detected or flagged rows in the preserved cleaned output", fontsize=12, color=GREY)
    sheets = sorted(metrics["by_worksheet"])
    # Counts by sheet are measured directly during collection below.
    by_sheet = metrics["flags_by_worksheet"]
    panels = [(figure.add_axes([.27, .72, .61, .12]), [("Missing customer ID", "customer_id_missing_flag")], 160_000),
              (figure.add_axes([.27, .17, .61, .48]), [("Duplicate-group members", "exact_duplicate_flag"), ("Negative quantity", "negative_quantity_flag"), ("Cancellation coded", "cancellation_coded_flag"), ("Zero unit price", "zero_unit_price_flag"), ("Missing description", "description_missing_flag"), ("Negative unit price", "negative_unit_price_flag")], 16_500)]
    for axis, categories, limit in panels:
        for j, sheet in enumerate(sheets):
            values = [by_sheet[sheet][flag] for _, flag in categories]
            positions = [i + (j - .5) * .32 for i in range(len(categories))]
            bars = axis.barh(positions, values, height=.27, color=[TEAL, BLUE][j], label=sheet)
            for bar, count in zip(bars, values):
                axis.text(count + limit * .012, bar.get_y() + bar.get_height() / 2, f"{count:,}", va="center", fontsize=9, color=NAVY)
        axis.set_yticks(range(len(categories)), [label for label, _ in categories])
        axis.invert_yaxis()
        axis.set_xlim(0, limit)
        axis.xaxis.set_major_formatter(StrMethodFormatter("{x:,.0f}"))
        axis.grid(axis="x", alpha=.15)
        axis.set_axisbelow(True)
        axis.tick_params(axis="y", length=0, pad=12)
        axis.set_xlabel("Row count (linear scale)", color=GREY)
    handles, labels = panels[0][0].get_legend_handles_labels()
    figure.legend(handles, labels, loc="upper right", bbox_to_anchor=(.92, .905), frameon=False, ncol=2)
    figure.text(.06, .07, "Flags overlap. Panels use different linear scales. Counts do not represent errors fixed.", fontsize=10, color=GREY)
    figure.text(.06, .038, "Missing IDs, missing descriptions and duplicate rows remain in the cleaned data. Source: existing output CSV; UCI Online Retail II.", fontsize=9, color=GREY)
    figure.savefig(root / "reports/data_quality_comparison.png", dpi=180, metadata={"Software": "excel_data_clean portfolio generator"})
    plt.close(figure)


def draw_text(figure, x: float, y: float, text: str, width: int = 94, size: float = 10, color: str = NAVY) -> float:
    lines = textwrap.wrap(text, width=width, break_long_words=False, break_on_hyphens=False)
    figure.text(x, y, "\n".join(lines), fontsize=size, color=color, va="top", linespacing=1.45)
    return y - len(lines) * size * 1.45 / (72 * 11.69)


def pdf_page(number: int, title: str):
    figure = plt.figure(figsize=(8.27, 11.69), facecolor="white")
    figure.add_artist(plt.Rectangle((0, .875), 1, .125, transform=figure.transFigure, color=NAVY))
    figure.text(.075, .955, "ONLINE RETAIL II  /  DELIVERY BRIEF", fontsize=9, color="#B9D8DE", weight="bold")
    figure.text(.075, .91, title, fontsize=23, color="white", weight="bold")
    figure.text(.075, .04, "Public-data portfolio project | Verified baseline: 29 Oct 2026", fontsize=8, color=GREY)
    figure.text(.925, .04, f"{number:02d} / 03", fontsize=8, color=GREY, ha="right")
    return figure


def section(figure, y, title, text):
    figure.text(.075, y, title, fontsize=12, color=TEAL, weight="bold", va="top")
    return draw_text(figure, .075, y - .025, text) - .018


def finish_pdf_page(figure, pdf, number, preview_dir):
    """Reject overflowing body text before accepting a page."""
    figure.canvas.draw()
    renderer = figure.canvas.get_renderer()
    for text in figure.texts:
        if .05 < text.get_position()[1] < .875:
            box = text.get_window_extent(renderer).transformed(figure.transFigure.inverted())
            if box.y0 < .065 or box.x1 > .97:
                raise ValueError(f"PDF page {number} has overflowing body text: {text.get_text()[:70]}")
    if preview_dir:
        figure.savefig(preview_dir / f"client_summary_page_{number}.png", dpi=120)
    pdf.savefig(figure)
    plt.close(figure)


def pdf_table(figure, bottom, height, headers, rows, widths=None):
    axis = figure.add_axes([.075, bottom, .85, height])
    axis.axis("off")
    table = axis.table(cellText=rows, colLabels=headers, colWidths=widths, loc="center", cellLoc="left", colLoc="left", bbox=[0, 0, 1, 1])
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    for (row, column), cell in table.get_celld().items():
        cell.set_edgecolor("white")
        cell.PAD = .08
        if row == 0:
            cell.set_facecolor(NAVY)
            cell.set_text_props(color="white", weight="bold")
        else:
            cell.set_facecolor("#EEF4F6" if row % 2 else "#F8FAFB")
            cell.set_text_props(color=NAVY)


def build_pdf(root: Path, metrics: dict, preview_dir: Path | None):
    metadata = {"Title": "Online Retail II | Client delivery summary", "Author": "Self-directed public-data portfolio project", "Subject": "Data preparation, preservation and validation", "CreationDate": FIXED_TIME, "ModDate": FIXED_TIME}
    with PdfPages(root / "reports/client_summary.pdf", metadata=metadata) as pdf:
        figure = pdf_page(1, "Reliable data, with traceability")
        y = section(figure, .83, "Project overview", "Prepare a large historical retail workbook for transparent analysis while preserving its source records. Mixed identifiers, missing customer details, repeated rows and signed transaction values require explicit treatment before business users can interpret totals.")
        test_label = " / ".join(str(metrics["recorded_pytest"][key]) for key in ("passed", "failed", "skipped"))
        for x, value, caption in [(.075, f"{metrics['cleaned_rows']:,}", "source rows preserved"), (.375, f"{metrics['sales_rows']:,}", "gross positive sales lines"), (.675, test_label, "tests: pass / fail / skip")]:
            figure.add_artist(FancyBboxPatch((x, y - .085), .25, .085, boxstyle="round,pad=0.008", transform=figure.transFigure, facecolor="#EEF4F6", edgecolor="none"))
            figure.text(x + .012, y - .025, value, fontsize=17, weight="bold", color=TEAL, va="top")
            figure.text(x + .012, y - .06, caption, fontsize=8, color=GREY, va="top")
        y -= .14
        y = section(figure, y, "Dataset and scope", "UCI Online Retail II contains transaction lines from a UK-based online retailer, covering December 2009 to December 2011 across two worksheets. The eight original fields describe invoice and product identifiers, description, quantity, invoice date, price, customer ID and country. The historical workbook remains the authoritative source.")
        figure.text(.075, y, "Reconciled output counts", fontsize=12, color=TEAL, weight="bold")
        rows = [[sheet.replace("Year ", ""), f"{count:,}", f"{metrics['sales_by_worksheet'][sheet]:,}"] for sheet, count in metrics["by_worksheet"].items()]
        rows.append(["TOTAL", f"{metrics['cleaned_rows']:,}", f"{metrics['sales_rows']:,}"])
        pdf_table(figure, y - .185, .16, ["Worksheet", "Source = cleaned", "Sales subset"], rows, [.28, .36, .36])
        y -= .23
        y = section(figure, y, "Two delivery views", f"The cleaned CSV retains all {metrics['cleaned_rows']:,} lines, original source values, worksheet origin and source row references. The sales CSV contains only positive-quantity, positive-price, non-cancellation lines. The other {metrics['excluded']:,} records are excluded only from that sales view; they remain in the cleaned output.")
        figure.text(.075, y, "Source attribution", fontsize=12, color=TEAL, weight="bold", va="top")
        draw_text(figure, .075, y - .025, CITATION + ". CC BY 4.0; standardized and sampled. Self-directed portfolio project, not a paid client engagement.", width=105, size=9)
        finish_pdf_page(figure, pdf, 1, preview_dir)

        figure = pdf_page(2, "Observed quality and treatment")
        draw_text(figure, .075, .83, "Detected or flagged records are review observations, not necessarily errors. Categories overlap. The chart compares worksheets rather than implying that business events were corrected or removed.")
        axis = figure.add_axes([.065, .365, .87, .405])
        axis.imshow(plt.imread(root / "reports/data_quality_comparison.png"))
        axis.axis("off")
        figure.text(.075, .35, "Measured before / after preservation", fontsize=12, color=TEAL, weight="bold")
        pdf_table(figure, .185, .145, ["Measure", "Before", "After cleaning"], [
            ["Source row count", f"{metrics['cleaned_rows']:,}", f"{metrics['cleaned_rows']:,}"],
            ["Missing customer IDs", f"{metrics['missing']['source_customer_id']:,}", f"{metrics['missing']['customer_id']:,}"],
            ["Missing descriptions", f"{metrics['missing']['source_description']:,}", f"{metrics['missing']['description']:,}"],
            ["Duplicate occurrences*", f"{metrics['duplicate_occurrences_after_first']:,}", "Retained; no removal"],
        ], [.46, .24, .30])
        draw_text(figure, .075, .16, f"*Duplicate occurrences count rows after the first in each repeated group (initial assessment). The duplicate flag marks all {metrics['flags']['exact_duplicate_flag']:,} group members. No missing IDs or descriptions were filled; prices were not imputed.", size=9)
        finish_pdf_page(figure, pdf, 2, preview_dir)

        figure = pdf_page(3, "Validation and delivery guidance")
        y = section(figure, .83, "Cleaning and standardization", "The existing Pandas pipeline processes worksheets sequentially, standardizes headers and surrounding whitespace, parses numeric and date fields, and keeps invoice and stock identifiers textual. Original values remain in source audit columns; missing customer IDs and descriptions remain missing. Row-level quality flags support review without silent deletion, filling or deduplication.")
        tests = metrics["recorded_pytest"]
        y = section(figure, y, "Recorded validation evidence", f"The final verification report records {tests['passed']} pytest tests passed, {tests['failed']} failed and {tests['skipped']} skipped using the existing virtual environment. It independently verifies full CSV readability, exact row counts, unique populated source row IDs and complete worksheet row coverage. The workbook's recorded size and modification time were unchanged. These are technical checks; no commercial performance improvement has been measured.")
        y = section(figure, y, "Classification and signed amounts", "Status precedence is: review for missing required values, zero quantity or negative price; then cancellation-coded invoice; then other negative quantity; then zero price; otherwise normal sale. Independent flags preserve overlapping conditions. line_total is quantity multiplied by unit price when both are finite numbers; its sign follows the operands.")
        y = section(figure, y, "Business assumptions and limitations", "Negative quantities and C-coded invoices may represent returns, corrections or other business events; the data alone does not verify refunds. Duplicate records remain in both outputs when eligible. The sales subset is gross positive sales lines, not verified net revenue. No invoice-level netting or price imputation is performed. Currency calculations use floating-point values without a currency-rounding policy.")
        y = section(figure, y, "Handoff and correct import", f"Use the delivery guide and all-column data dictionary for interpretation. The sample workbook contains {metrics['sample_rows']} deterministic, stratified rows with explicit Excel text identifiers. It overrepresents rare statuses and must not be used for population totals. CSV has no type metadata: import invoice numbers, stock codes and customer IDs as strings; parse dates explicitly. Full CSVs stay outside the portfolio folder and are ignored by Git.")
        section(figure, y, "Reproduction and next decision", "Run src/build_portfolio.py to regenerate this package from the existing outputs and verification reports. The cleaning pipeline is not rerun by the generator. Business owners should agree on refund interpretation, duplicate handling and accounting precision before using these files for financial reporting. Nothing has been published or uploaded.")
        finish_pdf_page(figure, pdf, 3, preview_dir)


def write_markdown(root: Path, metrics: dict):
    guide = f'''# Client delivery guide

This package documents a self-directed portfolio project using a public dataset. It does not describe a paid client engagement or a measured business outcome.

## Files and purpose

| File | Purpose |
| --- | --- |
| `reports/client_summary.pdf` | Three-page English handoff summary with verified technical results and limitations. |
| `reports/data_quality_comparison.png` | Worksheet comparison of detected/flagged records; categories overlap and do not represent errors fixed. |
| `outputs/cleaned_transactions.csv` | All {metrics['cleaned_rows']:,} original records, canonical values, source audit columns, lineage and flags. |
| `outputs/sales_analysis_ready.csv` | {metrics['sales_rows']:,} gross positive sales lines. The {metrics['excluded']:,} other records are excluded only here, never deleted from the cleaned file. |
| `portfolio/sample_cleaned_data.xlsx` | {metrics['sample_rows']} rows covering both worksheets and available transaction statuses; all columns retained. Read the sampling worksheet first. |
| `portfolio/data_dictionary.xlsx` | All {len(metrics['columns'])} actual output columns, recommended types, observed examples and business rules. |
| `reports/final_verification_report.md` | Previously executed full-data and pytest checks: {metrics['recorded_pytest']['passed']} passed, {metrics['recorded_pytest']['failed']} failed, {metrics['recorded_pytest']['skipped']} skipped. |
| `portfolio/artifact_verification.json` | Generated artifact checks, SHA-256 hashes and protected-file checks. |

Full CSV files are deliberately excluded from the portfolio folder and from Git. The source workbook is not included in the delivery sample.

## Run and reproduce

From the project root, use the already configured environment:

```powershell
.\\.venv\\Scripts\\python.exe -m pip install -r requirements.txt
.\\.venv\\Scripts\\python.exe src\\inspect_data.py
.\\.venv\\Scripts\\python.exe src\\clean_data.py
.\\.venv\\Scripts\\python.exe -m pytest -q
.\\.venv\\Scripts\\python.exe src\\build_portfolio.py
```

The inspection and cleaning commands are for reproducing the underlying outputs when needed. To rebuild only this package from existing CSVs and reports, run only `build_portfolio.py`. It scans outputs in chunks and does not invoke cleaning or alter the source workbook. Default artifact paths are relative to the project root. Repeat builds with the same inputs and dependency environment produce the same sample and normalized artifact bytes; rendering may differ across library/font versions.

## Correctly read CSV identifiers and dates

```python
import pandas as pd

ids = {{
    'invoice_no': 'string', 'stock_code': 'string', 'customer_id': 'string',
    'source_invoice_no': 'string', 'source_stock_code': 'string',
    'source_customer_id': 'string', 'source_row_id': 'string',
}}
df = pd.read_csv('outputs/cleaned_transactions.csv', dtype=ids)
df['invoice_date'] = pd.to_datetime(df['invoice_date'], errors='coerce')
```

CSV stores no type metadata. Default readers may infer sales invoice IDs as integers and customer IDs as floats. Treat identifiers as labels, not arithmetic values. The sample XLSX explicitly stores identifier cells as Excel text, including source audit versions. Missing values remain blank. Source customer IDs may have a `.0` textual suffix because the original Pandas read represented numeric IDs as floats; canonical customer IDs are normalized by the existing pipeline. Never invent customer IDs.

## Interpret flags and statuses

Flags overlap. The dictionary explains each field. Missing customer IDs and descriptions are retained and flagged. Exact duplicates are tested within each worksheet across original fields; every group member is flagged, with no automatic deletion. The initial assessment counted {metrics['duplicate_occurrences_after_first']:,} duplicate occurrences after the first; {metrics['flags']['exact_duplicate_flag']:,} rows are members of those groups.

Status precedence: missing required fields, zero quantity or negative unit price -> `requires_review`; otherwise C-coded invoice -> `cancellation_coded_transaction`; otherwise negative quantity -> `other_negative_quantity_transaction`; otherwise zero unit price -> `zero_price_transaction`; otherwise `normal_sale`. The status is a classification, not proof of a refund or commercial validity. Individual flags still expose conditions hidden by precedence. The required-field flag itself covers missing canonical invoice, stock, quantity, date or price.

Sales eligibility is exactly quantity > 0, unit_price > 0 and not cancellation-coded. It retains eligible duplicates and missing customer IDs/descriptions, and does not independently validate business meaning or require date/identifier validity. All required values happened to pass the existing full-data checks; do not silently assume broader filtering rules for future inputs.

## Sampling and limitations

The sample selects up to 100 rows per (worksheet, transaction_status) stratum by the smallest SHA-256 hash of seed plus source_row_id. Seed: `{SEED}`. It caps the combined sample at 1,000 by the same priority, then sorts by worksheet and numeric Excel row. It illustrates available cases rather than matching population frequencies. The Read me worksheet records actual stratum counts. Do not calculate population totals from it.

The sales subset represents gross positive sales lines, not verified net revenue. Returns, credits and cancellations require domain interpretation and invoice-level reconciliation. Negative quantities can also be corrections or stock adjustments. line_total is a signed finite quantity-times-price calculation, with no accounting policy or currency rounding; floating-point precision applies. Duplicate retention can affect aggregates. The source audit fields preserve read values, not Excel cell formatting or formulas, and source row positions are stable only for the unchanged workbook.

Before financial use, agree on return/refund interpretation, duplicate handling and monetary precision. Before public release, retain attribution, keep large source/output data out of Git, and review project paths and author details. This package has not been uploaded or published.

## Attribution

{CITATION}. Source: {SOURCE}. License: [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). This project standardizes, annotates and samples the public source data; no endorsement by the original creator or UCI is implied.
'''
    (root / "portfolio/client_delivery_guide.md").write_text(guide, encoding="utf-8")
    case = f'''# Auditable Retail Data Cleaning with Python and Pandas

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

- {metrics['cleaned_rows']:,} source rows preserved across two worksheets, with unique source row identifiers in the existing final verification.
- {metrics['sales_rows']:,} rows meet the positive-quantity, positive-price, non-cancellation subset rule; {metrics['excluded']:,} rows are excluded only from that view and retained in the full cleaned file.
- {metrics['missing']['customer_id']:,} missing customer IDs and {metrics['missing']['description']:,} missing descriptions retained.
- {metrics['flags']['exact_duplicate_flag']:,} duplicate-group members flagged; {metrics['duplicate_occurrences_after_first']:,} duplicate occurrences after the first, with no automatic deduplication.
- {metrics['recorded_pytest']['passed']} pytest tests passed, {metrics['recorded_pytest']['failed']} failed, {metrics['recorded_pytest']['skipped']} skipped in the recorded verification.
- No measured claim of revenue uplift, cost savings, correction rate or client business impact.

## Deliverables

- [Client summary PDF](../reports/client_summary.pdf) and [quality comparison chart](../reports/data_quality_comparison.png).
- [Sample workbook](sample_cleaned_data.xlsx) with {metrics['sample_rows']} stratified rows, flags, lineage and explicit Excel text identifiers.
- [All-column data dictionary](data_dictionary.xlsx), [delivery guide](client_delivery_guide.md) and [artifact verification](artifact_verification.json).
- Reusable inspection, cleaning and portfolio generation scripts, plus existing validation reports and automated tests.

## Limitations

This is historical public data, with no claim of a commercial client relationship. The sample overrepresents rare cases and cannot estimate population metrics. Cancellations and negative quantities are potential business events, not verified refunds. Exact duplicates and missing customer IDs remain. The sales subset represents gross positive sales lines, not verified net revenue. Floating-point prices and line totals have no currency-rounding policy; financial reporting requires business review. CSV identifiers must be imported with explicit string dtypes.

## Dataset attribution

{CITATION}. [UCI dataset]({SOURCE}); [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/). Source data was standardized, annotated and sampled for this portfolio exercise.
'''
    (root / "portfolio/portfolio_case_study.md").write_text(case, encoding="utf-8")


def verify(root: Path, sample: pd.DataFrame, metrics: dict) -> dict:
    checks = {}
    pdf = PdfReader(root / "reports/client_summary.pdf", strict=True)
    text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    checks["pdf_three_readable_pages"] = len(pdf.pages) == 3 and all(len(page.extract_text() or "") > 500 for page in pdf.pages)
    checks["pdf_verified_counts_present"] = all(value in text for value in ["1,067,371", "1,041,670", "25,701"])
    with Image.open(root / "reports/data_quality_comparison.png") as image:
        checks["png_dimensions"] = list(image.size)
        image.verify()
    book = openpyxl.load_workbook(root / "portfolio/sample_cleaned_data.xlsx", data_only=False)
    sheet = book["Sample transactions"]
    headers = [cell.value for cell in sheet[1]]
    checks["sample_rows"] = sheet.max_row - 1
    checks["sample_at_most_1000"] = sheet.max_row - 1 == len(sample) <= 1000
    checks["sample_all_columns"] = headers == metrics["columns"]
    identifier_checks = []
    source_checks = []
    flag_checks = []
    lineage = []
    for sample_index, row in enumerate(sheet.iter_rows(min_row=2)):
        record = dict(zip(headers, row))
        lineage.append(record["source_row_id"].value)
        for column in headers:
            if column.startswith("source_"):
                original = sample.iloc[sample_index][column]
                actual = record[column].value
                source_checks.append(("" if actual is None else str(actual)) == original)
            if column in FLAGS:
                flag_checks.append(record[column].value is (sample.iloc[sample_index][column] == "True"))
        for name in IDENTIFIERS:
            cell = record[name]
            if cell.value is not None:
                identifier_checks.append(cell.data_type == "s" and cell.number_format == "@")
    checks["excel_identifier_cells_are_text"] = all(identifier_checks)
    checks["sample_source_values_match_csv"] = all(source_checks)
    checks["sample_flags_match_csv"] = all(flag_checks)
    checks["sample_covers_present_quality_flags"] = all(sample[name].eq("True").any() for name, count in metrics["flags"].items() if count > 0)
    checks["sample_lineage_matches_csv_selection"] = lineage == sample["source_row_id"].tolist() and len(lineage) == len(set(lineage))
    checks["sample_no_formula_cells"] = not any(cell.data_type == "f" for row in sheet for cell in row)
    book.close()
    dictionary = openpyxl.load_workbook(root / "portfolio/data_dictionary.xlsx", read_only=True)
    documented = [row[0] for row in dictionary["Data dictionary"].iter_rows(min_row=2, values_only=True)]
    checks["dictionary_covers_actual_schema"] = documented == metrics["columns"]
    checks["dictionary_column_count"] = len(documented)
    dictionary.close()
    for filename in ["client_delivery_guide.md", "portfolio_case_study.md"]:
        content = (root / "portfolio" / filename).read_text(encoding="utf-8")
        checks[filename + "_readable"] = len(content) > 1000 and "1,067,371" in content
    checks["no_full_csv_in_portfolio"] = not any((root / "portfolio").glob("*.csv"))
    if not all(value for value in checks.values()):
        raise RuntimeError(f"Artifact checks failed: {checks}")
    return checks


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--preview-dir", type=Path, help="Optional ignored-directory page renders for visual review")
    args = parser.parse_args()
    (ROOT / "portfolio").mkdir(exist_ok=True)
    if args.preview_dir:
        args.preview_dir.mkdir(parents=True, exist_ok=True)
    protected = ["data/online_retail_II.xlsx", "outputs/cleaned_transactions.csv", "outputs/sales_analysis_ready.csv", "src/clean_data.py", "src/inspect_data.py"]
    before = {name: sha256(ROOT / name) for name in protected}
    sample, metrics = collect(ROOT)
    # Worksheet flag breakdowns are measured in a lightweight separate chunk scan.
    by_sheet = {sheet: Counter() for sheet in metrics["by_worksheet"]}
    for chunk in pd.read_csv(ROOT / "outputs/cleaned_transactions.csv", usecols=["source_worksheet", *FLAGS], dtype="string", keep_default_na=False, chunksize=100_000):
        for sheet, group in chunk.groupby("source_worksheet", sort=True):
            by_sheet[sheet].update({name: int(group[name].eq("True").sum()) for name in FLAGS})
    metrics["flags_by_worksheet"] = {sheet: dict(counts) for sheet, counts in by_sheet.items()}
    build_spreadsheets(ROOT, sample, metrics)
    build_chart(ROOT, metrics)
    build_pdf(ROOT, metrics, args.preview_dir)
    write_markdown(ROOT, metrics)
    checks = verify(ROOT, sample, metrics)
    after = {name: sha256(ROOT / name) for name in protected}
    if before != after:
        raise RuntimeError("A protected input changed during package generation")
    files = ["reports/client_summary.pdf", "reports/data_quality_comparison.png", "portfolio/sample_cleaned_data.xlsx", "portfolio/data_dictionary.xlsx", "portfolio/client_delivery_guide.md", "portfolio/portfolio_case_study.md"]
    manifest = {"method": "Artifact checks executed by src/build_portfolio.py; recorded pytest counts are from existing final_verification_report.md, not a new pytest run.", "dependency_versions": {package: version(package) for package in ("pandas", "openpyxl", "matplotlib", "Pillow", "pypdf")}, "metrics": metrics, "checks": checks, "protected_input_sha256": before, "protected_inputs_unchanged": True, "artifacts": {name: {"bytes": (ROOT / name).stat().st_size, "sha256": sha256(ROOT / name)} for name in files}}
    (ROOT / "portfolio/artifact_verification.json").write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    for name in files + ["portfolio/artifact_verification.json"]:
        print(f"{name}: {(ROOT / name).stat().st_size:,} bytes")
    print(f"Verified: PDF {3} pages; sample {len(sample)} rows; dictionary {len(metrics['columns'])} columns; source/CSV/cleaning code hashes unchanged.")


if __name__ == "__main__":
    main()
