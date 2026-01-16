# Automated Financial Data Warehouse

This project implements a monthly financial data warehouse automation that transforms multi-sheet income statements into a clean, deduplicated, long-format dataset.

The system was built to support non-technical stakeholders while maintaining data integrity, reproducibility, and auditability.

## Features
- Automated ingestion of monthly Excel income statements
- Parsing of multi-department worksheets
- Revenue vs expense classification
- GL reference table normalization
- Idempotent updates (safe re-runs)
- QA reporting for missing GL mappings

## Implementations

### Python Version
Designed for local execution or shared-drive environments.
- Uses pandas for data wrangling
- Appends and deduplicates monthly data
- Preserves historical records

### Google Sheets Version
Designed for non-technical users.
- Runs entirely in Google Sheets
- Uses Apps Script and Drive API
- One-click execution via custom menu

## Repository Structure
```text
python_version/         # Python ETL implementation
google_sheets_version/  # Google Apps Script implementation
docs/                   # Architecture and design notes
