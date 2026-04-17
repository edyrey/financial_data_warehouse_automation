# Google Sheets Version Setup

This folder preserves the Google Suite phase of the project.

## Purpose

The Google Sheets and Apps Script implementation was used when the workflow needed to be accessible to non-technical users in a shared environment.

## High-Level Setup

1. Create a warehouse Google Sheet with tabs for:
   - `GL`
   - `Final`
   - `Missing_GL_Mapping`
2. Upload monthly statement files into Google Drive folders.
3. Paste the script from `Code.gs` into Apps Script.
4. Update the folder IDs in the config section.
5. Enable the Drive API advanced service.
6. Run the update from the custom spreadsheet menu.

## Notes

- This version represents an important middle stage of the project.
- It is included here for documentation and historical context.
- The current development direction is the Python rebuild.
