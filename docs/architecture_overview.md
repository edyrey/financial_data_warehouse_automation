# Architecture Overview

## Goal
Automate monthly ingestion of semi-structured financial Excel reports into a clean warehouse fact table that supports reporting, trend analysis, and downstream BI.

## Inputs
Two monthly Excel report types:
1) **Income Statements**
   - Multi-sheet workbook (one worksheet per department)
   - Revenue/Expense sections represented as header rows
2) **Balance Sheets**
   - Statement-style layout with subtotals and section totals
   - GL codes appear in a fixed column; YTD amounts in a fixed column

Both file types use a filename convention containing `MM.YYYY` for date metadata.

## Reference Data
A Google Sheet tab (GL Reference) maintains:
- GL Code
- Description
- Group

This reference table is used to normalize descriptions/groups and resolve typos or inconsistencies from source reports.

## Output
A long-format fact table with columns:
- GL Code
- Description
- Category
- Group
- Year
- Month
- Department
- Amount

## Key Design Choices
- **Idempotent updates:** safe to rerun; deduplication key prevents duplicates
- **QA accumulation:** missing GL mappings persist across runs until resolved
- **Non-technical usability:** one-click execution via Google Sheets custom menu
- **Confidentiality:** repository uses synthetic sample content only

## Execution Flow (Google Suite)
1. Locate monthly XLSX files in Google Drive input folders
2. Convert XLSX → temporary Google Sheet via Drive API
3. Parse each report into normalized rows
4. Join to GL reference (Description + Group)
5. Append + dedupe into Final fact table
6. Update “Missing GL Mapping” QA table (accumulative + auto-resolve)
