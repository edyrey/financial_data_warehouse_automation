# Python Version

This folder represents the current direction of the project.

The goal of the Python rebuild is to create a local, reproducible workflow for loading monthly financial statement files into a warehouse-style dataset while keeping the parsing and QA logic easier to test and maintain.

## Included Here

- `build_warehouse.py`: example ETL script for parsing statement files and updating a warehouse workbook
- `requirements.txt`: minimal Python dependencies
- `sample_data/`: safe templates and synthetic example files for documenting the workflow

## Intended Local Workflow

1. Place monthly source files in a local input directory.
2. Run the warehouse build script.
3. Append or refresh long-format warehouse records.
4. Review missing GL mappings in the QA output.

The live operational files are intentionally not included in this public-facing repository.
