# Automated Financial Data Warehouse (Python & Google Suite)

This project automates monthly financial reporting by transforming multi-sheet Excel statements into a clean, deduplicated, long-format warehouse table.

It includes two implementations:

## ✅ Production Implementation (Google Suite)
The current production workflow runs entirely in Google Sheets using Google Apps Script:
- Reads monthly Excel inputs from Google Drive
- Converts XLSX → temporary Google Sheet (Drive API)
- Parses semi-structured statement layouts (income statements + balance sheets)
- Normalizes GL codes using a centralized reference table (GL Code → Description + Group)
- Appends and deduplicates records into a warehouse “Final” fact table
- Tracks missing GL mappings in an accumulating QA sheet until resolved

This approach was chosen to support non-technical users and provide a one-click update workflow in a shared environment.

## 🧪 Preliminary Prototype (Python)
A Python (pandas) version was created first as a prototype to validate parsing logic and schema design and to process initial historical data. After requirements expanded to non-technical users and shared-drive execution, the workflow was migrated to Google Suite.

## Repository Structure
```text
google_sheets_version/   # production Apps Script automation
python_version/          # prototype pandas implementation (legacy)
docs/                    # architecture + design notes
