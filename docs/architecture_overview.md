# Architecture Overview

This project implements a monthly ETL pipeline for financial data.

## Inputs
- Multi-sheet Excel income statements
- One sheet per department
- File names encode month/year

## Processing
- Parse department sheets
- Extract GL-level records
- Normalize descriptions via GL reference table
- Detect revenue vs expense categories
- Deduplicate using natural keys

## Outputs
- Long-format financial warehouse
- QA table for missing GL mappings

Two implementations are provided:
- Python (local/shared drive)
- Google Apps Script (cloud-based, non-technical users)
