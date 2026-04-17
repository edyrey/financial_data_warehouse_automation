# Architecture Overview

## Goal

Transform recurring financial statement workbooks into a standardized warehouse table that supports reporting, trend analysis, reconciliation, and downstream analytics.

## Source Inputs

The project works from monthly Excel reports such as:

- departmental income statements
- balance sheets
- supporting GL reference mappings

These files follow naming conventions that encode the reporting period and are structured enough to automate, but inconsistent enough to require custom parsing logic.

## Core Warehouse Pattern

The pipeline normalizes raw statement data into long-format records with fields such as:

- GL Code
- Description
- Category
- Group
- Year
- Month
- Department
- Amount

## Project Evolution

### Python prototype

The first version was built in Python to prove that the source reports could be parsed, normalized, and appended safely into a warehouse-style table.

### Google Suite implementation

The workflow then moved into Google Sheets and Apps Script so the process could run in a more familiar shared environment for non-technical users.

### Current Python rebuild

The project is now being rebuilt in Python to improve maintainability, support a local data warehouse workflow, and make the ETL process easier to document, test, and extend.

## Design Priorities

- idempotent monthly updates
- clear warehouse schema
- maintainable parsing rules
- reproducible ETL steps
- explicit handling of missing GL mappings
- safe separation of public examples from private operational files
