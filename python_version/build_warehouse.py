import re
import os
import calendar
from pathlib import Path

import pandas as pd


SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SCRIPT_DIR.parent

INPUT_DIR = PROJECT_DIR / "input"
WAREHOUSE_PATH = PROJECT_DIR / "warehouse" / "EXAMPLE_COMPANY Data Warehouse.xlsx"

WAREHOUSE_GL_SHEET = "GL"
WAREHOUSE_FINAL_SHEET = "Final"
WAREHOUSE_QA_SHEET = "Missing_GL_Mapping"


def parse_month_year_from_filename(file_path: str) -> tuple[int, int]:
    name = Path(file_path).name
    m = re.search(r"(?P<month>\d{2})\.(?P<year>\d{4})", name)
    if not m:
        raise ValueError(f"Could not find mm.yyyy in filename: {name}")
    month = int(m.group("month"))
    year = int(m.group("year"))
    if not (1 <= month <= 12):
        raise ValueError(f"Month out of range in filename: {month}")
    return month, year


def extract_department_from_sheet(sheet_name: str) -> str | None:
    m = re.match(r"(?i)DEPARTMENT\s+(\d+)-F\s*$", sheet_name.strip())
    return m.group(1) if m else None


def clean_amount(x) -> float | None:
    if pd.isna(x):
        return None
    s = str(x).strip().replace("$", "").replace(",", "")
    neg = False
    if re.match(r"^\(.*\)$", s):
        neg = True
        s = s.strip("()").strip()
    if s == "":
        return None
    try:
        val = float(s)
        return -val if neg else val
    except ValueError:
        return None


def is_gl_code(val) -> bool:
    if pd.isna(val):
        return False
    return bool(re.fullmatch(r"\d{4}", str(val).strip()))


def load_gl_reference(warehouse_path: str, sheet_name: str) -> pd.DataFrame:
    gl = pd.read_excel(warehouse_path, sheet_name=sheet_name, engine="openpyxl")

    cols = {c: re.sub(r"\s+", " ", str(c)).strip().lower() for c in gl.columns}
    gl_code_col = None
    desc_col = None

    for c, lc in cols.items():
        if lc in {"gl", "gl code", "glcode", "number", "account", "account number", "account#", "account #"}:
            gl_code_col = c
        if lc in {"description", "account description", "gl description", "name"}:
            desc_col = c

    if gl_code_col is None:
        raise ValueError("Could not identify GL code column in GL reference sheet.")
    if desc_col is None:
        raise ValueError("Could not identify Description column in GL reference sheet.")

    out = gl[[gl_code_col, desc_col]].copy()
    out.columns = ["GL code", "description"]
    out["GL code"] = out["GL code"].astype(str).str.strip()
    out["description"] = out["description"].astype(str).str.strip()
    out = out[out["GL code"].str.fullmatch(r"\d{4}", na=False)].drop_duplicates(subset=["GL code"])
    return out


def parse_income_statement_sheet(df_raw: pd.DataFrame, dept: str, month: int, year: int) -> pd.DataFrame:
    df = df_raw.copy()
    df.columns = ["NUMBER", "DESCRIPTION", "ACTUAL"]

    num_str = df["NUMBER"].astype(str).str.strip()
    cat = pd.Series([None] * len(df), index=df.index, dtype="object")
    cat[num_str.str.upper() == "REVENUES"] = "Revenue"
    cat[num_str.str.upper() == "EXPENSES"] = "Expenses"
    df["category"] = cat.ffill()

    df = df[df["NUMBER"].apply(is_gl_code)].copy()
    df["amount"] = df["ACTUAL"].apply(clean_amount)
    df = df.dropna(subset=["amount"])

    df["GL code"] = df["NUMBER"].astype(str).str.strip()
    df["department number"] = dept
    df["month"] = month
    df["year"] = year

    return df[["GL code", "category", "year", "month", "department number", "amount"]]


def build_month_fact(monthly_file_path: str, warehouse_path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    month, year = parse_month_year_from_filename(monthly_file_path)
    gl_ref = load_gl_reference(warehouse_path, WAREHOUSE_GL_SHEET)

    xl = pd.ExcelFile(monthly_file_path, engine="openpyxl")
    all_rows = []

    for sheet in xl.sheet_names:
        dept = extract_department_from_sheet(sheet)
        if not dept:
            continue
        df_raw = pd.read_excel(
            monthly_file_path,
            sheet_name=sheet,
            usecols=[0, 1, 2],
            header=1,
            engine="openpyxl",
        )
        all_rows.append(parse_income_statement_sheet(df_raw, dept=dept, month=month, year=year))

    if not all_rows:
        raise ValueError("No department sheets found (expected 'DEPARTMENT XXX-F').")

    fact = pd.concat(all_rows, ignore_index=True)
    merged = fact.merge(gl_ref, on="GL code", how="left")
    merged["gl_missing_in_reference"] = merged["description"].isna()

    final = merged.rename(columns={
        "description": "Description",
        "category": "Category",
        "year": "Year",
        "month": "Month",
        "department number": "Department",
        "amount": "Amount",
    })[["GL code", "Description", "Category", "Year", "Month", "Department", "Amount"]].copy()

    return final, merged


def read_existing_final(warehouse_path: str) -> pd.DataFrame:
    cols = ["GL code", "Description", "Category", "Year", "Month", "Department", "Amount"]
    try:
        existing = pd.read_excel(warehouse_path, sheet_name=WAREHOUSE_FINAL_SHEET, engine="openpyxl")
        return existing[cols].copy()
    except ValueError as e:
        msg = str(e).lower()
        if "worksheet" in msg or "not found" in msg:
            return pd.DataFrame(columns=cols)
        raise


def append_and_dedupe(existing: pd.DataFrame, new_rows: pd.DataFrame) -> pd.DataFrame:
    combined = pd.concat([existing, new_rows], ignore_index=True)
    combined["GL code"] = combined["GL code"].astype(str).str.strip()
    combined["Department"] = combined["Department"].astype(str).str.strip()
    combined["Year"] = pd.to_numeric(combined["Year"], errors="raise").astype(int)
    combined["Month"] = pd.to_numeric(combined["Month"], errors="raise").astype(int)

    dedupe_key = ["GL code", "Year", "Month", "Department", "Category"]
    combined = combined.drop_duplicates(subset=dedupe_key, keep="last")
    combined = combined.sort_values(by=["Year", "Month", "Department", "Category", "GL code"]).reset_index(drop=True)
    return combined


def write_back_to_warehouse(warehouse_path: str, final_df: pd.DataFrame, missing_df: pd.DataFrame):
    book = pd.ExcelFile(warehouse_path, engine="openpyxl")
    sheets_to_preserve = [s for s in book.sheet_names if s not in {WAREHOUSE_FINAL_SHEET, WAREHOUSE_QA_SHEET}]
    preserved = {s: pd.read_excel(warehouse_path, sheet_name=s, engine="openpyxl") for s in sheets_to_preserve}

    with pd.ExcelWriter(warehouse_path, engine="openpyxl", mode="w") as writer:
        for s, df in preserved.items():
            df.to_excel(writer, sheet_name=s, index=False)
        final_df.to_excel(writer, sheet_name=WAREHOUSE_FINAL_SHEET, index=False)
        missing_df.to_excel(writer, sheet_name=WAREHOUSE_QA_SHEET, index=False)


def main():
    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"Input folder not found: {INPUT_DIR}")
    if not WAREHOUSE_PATH.exists():
        raise FileNotFoundError(f"Warehouse file not found: {WAREHOUSE_PATH}")

    files = [
        INPUT_DIR / f
        for f in os.listdir(INPUT_DIR)
        if f.lower().endswith(".xlsx") and "data warehouse" not in f.lower()
    ]
    if not files:
        raise FileNotFoundError("No monthly Excel files found in input folder.")

    monthly_path = max(files, key=lambda p: p.stat().st_mtime)
    new_final, merged = build_month_fact(str(monthly_path), str(WAREHOUSE_PATH))
    existing_final = read_existing_final(str(WAREHOUSE_PATH))
    updated_final = append_and_dedupe(existing_final, new_final)
    updated_final["Month"] = updated_final["Month"].apply(lambda m: calendar.month_name[int(m)])
    missing = merged[merged["gl_missing_in_reference"]].copy()
    write_back_to_warehouse(str(WAREHOUSE_PATH), updated_final, missing)

    print("Done.")
    print(f"Updated '{WAREHOUSE_FINAL_SHEET}' in: {WAREHOUSE_PATH}")


if __name__ == "__main__":
    main()
