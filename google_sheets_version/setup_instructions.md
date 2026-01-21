# Google Sheets Version — Setup Instructions (Production)

## 1) Create the Warehouse Google Sheet
Create a Google Sheet that will serve as the data warehouse.

Create these tabs (sheet names can be changed, but must match the constants in `Code.gs`):
- `GL` (GL reference table)
- `Final` (warehouse fact table output)
- `Missing_GL_Mapping` (QA log of missing GL reference mappings)

## 2) Prepare the GL reference table (GL tab)
Add headers in row 1:
- `GL Code` (or `GL`, `GL#`)
- `Description`
- `Group`

Populate with reference mappings (synthetic in public repo; real internally).

## 3) Create two input folders in Google Drive
Create:
- An **Income Statements** folder (monthly XLSX uploads)
- A **Balance Sheets** folder (monthly XLSX uploads)

Copy each folder’s ID from the URL:
`https://drive.google.com/drive/folders/<FOLDER_ID>`

## 4) Install the Apps Script
In the Warehouse Google Sheet:
- Extensions → Apps Script
- Paste the contents of `Code.gs`
- Update:
  - `INCOME_INPUTS_FOLDER_ID`
  - `BALANCE_INPUTS_FOLDER_ID`

## 5) Enable Drive API (Advanced Service)
Apps Script editor:
- Services (puzzle icon) → Add a service → **Drive API** (v3)

Ensure Drive API is enabled in the connected Google Cloud project if prompted.

## 6) Run
Reload the Google Sheet. You should see a custom menu:
**Financial Warehouse → Run Update (Income + Balance)**

Authorize the script on first run.

## Notes
- Safe re-runs (deduplication prevents duplicates)
- Missing GL mappings are tracked and accumulate until resolved in the GL tab
- Converted temporary files are trashed automatically
