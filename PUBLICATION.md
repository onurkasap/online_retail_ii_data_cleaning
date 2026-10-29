# Publication checklist

The project is a self-directed portfolio exercise using the public UCI Online Retail II dataset. Publication and Git operations are left to the project owner.

## Recommended files to commit

- `README.md`, `PUBLICATION.md`, `.gitignore`, and `requirements.txt`.
- `src/inspect_data.py`, `src/clean_data.py`, and `src/build_portfolio.py`.
- `tests/test_clean_data.py`.
- `reports/initial_data_quality_report.md`, `reports/cleaning_report.md`, `reports/final_verification_report.md`, `reports/client_summary.pdf`, and `reports/data_quality_comparison.png`.
- `portfolio/client_delivery_guide.md`, `portfolio/portfolio_case_study.md`, `portfolio/sample_cleaned_data.xlsx`, `portfolio/data_dictionary.xlsx`, and `portfolio/artifact_verification.json`.

The sample workbook contains a selection from the public dataset and includes attribution and a CC BY 4.0 license link. The artifact manifest contains source and output checksums, schema examples and verification results; it does not contain the full CSVs.

## Exclude

- `.venv/` and local `.env` or `.env.*` secrets; `.env.example` may be committed if intentionally added without secrets.
- `data/*.xlsx`, `data/*.xls`, and `data/processed/`, including the original Excel workbook and generated datasets.
- `outputs/` and `portfolio/*.csv`, including both full transaction CSVs, plot caches and preview renders.
- Python, test and tool caches, coverage output, build artifacts, editor metadata, temporary files, logs and OS-specific files covered by `.gitignore`.

Ignore rules only prevent new files from being added automatically; they do not remove files that are already tracked. Review the staged file list and the repository's tracking status before committing.

## Review before public release

1. Confirm the staged files and Git ignore behavior after creating your own repository. Do not stage the original workbook, full CSVs or local environment.
2. Recheck Markdown links, binary metadata and text for personal paths or sensitive information. The current assessment and cleaning reports use project-relative paths.
3. If inspection or cleaning is rerun, its generated reports may include absolute paths from that machine. Replace them with project-relative paths before sharing those reports.
4. Keep the UCI attribution, dataset link, CC BY 4.0 link and modification notice alongside the sample and documentation. Select a separate code license if you want to grant reuse rights for this project's code.
5. Review your author details and case-study wording. The case study must remain clear that this is a self-directed public-data project and that results are technical checks, not measured client outcomes.

The source workbook, full generated outputs and cleaning rules are unchanged by publication preparation.
