# Automated Financial Data Warehouse

This repository documents the evolution of a financial data warehouse project used to standardize semi-structured monthly accounting reports into a clean, analysis-ready dataset.

The project has moved through three phases:

1. A Python prototype to validate parsing logic, warehouse schema, and idempotent monthly updates.
2. A Google Sheets and Apps Script implementation to support a shared workflow for non-technical users.
3. A current Python rebuild, developed with Codex support, to create a more maintainable local warehouse and a clearer reproducible ETL process.

## Current Status

The active direction of the project is the Python implementation. The current goal is to rebuild the warehouse as a local, version-controlled workflow that is easier to test, document, and extend while preserving lessons learned from the Google Suite phase.

## Repository Structure

```text
docs/                    Project notes, architecture, and process documentation
google_sheets_version/   Apps Script implementation used in the shared Google workflow
python_version/          Python implementation for local warehouse development
python_version/sample_data/  Synthetic templates and safe example files
```

## What The Pipeline Does

- Reads monthly income statement and balance sheet workbooks
- Extracts date metadata from filename conventions
- Parses semi-structured financial statement layouts
- Standardizes GL codes, descriptions, and categories
- Appends clean records into a long-format warehouse table
- Prevents duplicate monthly loads through idempotent update logic
- Flags missing GL mappings for quality review

## Confidentiality

This public repository does not include live company data or production workbooks. Any sample files included here are synthetic templates or sanitized examples used to illustrate the workflow.

## Why Both Python And Google Suite Appear Here

The Google Suite implementation reflects the stage where the project was optimized for accessibility and shared use. The Python implementation reflects the current stage, where the project is being restructured into a more maintainable local data engineering workflow.

Together, those versions show how the project evolved in response to changing requirements, user needs, and tooling.
