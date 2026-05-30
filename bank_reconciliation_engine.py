"""
Bank Reconciliation Engine

Purpose:
    Match bank statement transactions against cash general ledger transactions.

Generates:
    - bank_reconciliation.xlsx
    - unmatched_bank_transactions.xlsx
    - unmatched_gl_cash_transactions.xlsx
"""

from pathlib import Path
import pandas as pd


def excel_date_to_datetime(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit="D", origin="1899-12-30")
    return pd.to_datetime(series)


def get_parameter(parameters: pd.DataFrame, name: str, default=None):
    if "Parameter" not in parameters.columns or "Value" not in parameters.columns:
        return default

    row = parameters.loc[parameters["Parameter"].astype(str).str.lower() == name.lower()]
    if row.empty:
        return default
    return row["Value"].iloc[0]


def get_amount_tolerance(parameters: pd.DataFrame) -> float:
    value = get_parameter(parameters, "Amount Tolerance", 0.01)
    try:
        return float(value)
    except Exception:
        return 0.01


def standardize_bank_statement(bank_statement: pd.DataFrame) -> pd.DataFrame:
    df = bank_statement.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Flexible column mapping for generated or manually adjusted files.
    rename = {
        "Bank_Transaction_ID": "bank_transaction_id",
        "Transaction_ID": "bank_transaction_id",
        "Transaction_Date": "transaction_date",
        "Date": "transaction_date",
        "Description": "description",
        "Amount": "amount",
        "Currency": "currency",
        "Bank_Account": "bank_account",
        "Reference": "reference",
        "Counterparty": "counterparty",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})

    required = ["bank_transaction_id", "transaction_date", "amount"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"bank_statement.xlsx missing required columns: {missing}")

    df["transaction_date"] = excel_date_to_datetime(df["transaction_date"])
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0.0)
    df["abs_amount"] = df["amount"].abs()
    df["matched"] = False
    df["match_id"] = ""

    return df


def standardize_gl_cash(general_ledger_cash: pd.DataFrame) -> pd.DataFrame:
    df = general_ledger_cash.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename = {
        "GL_Entry_ID": "gl_entry_id",
        "Entry_ID": "gl_entry_id",
        "Posting_Date": "posting_date",
        "Date": "posting_date",
        "Account_Name": "account_name",
        "Description": "description",
        "Debit": "debit",
        "Credit": "credit",
        "Amount": "amount",
        "Reference": "reference",
        "Source": "source",
    }
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})

    required = ["gl_entry_id", "posting_date"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"general_ledger_cash.xlsx missing required columns: {missing}")

    df["posting_date"] = excel_date_to_datetime(df["posting_date"])

    if "amount" not in df.columns:
        df["debit"] = pd.to_numeric(df.get("debit", 0.0), errors="coerce").fillna(0.0)
        df["credit"] = pd.to_numeric(df.get("credit", 0.0), errors="coerce").fillna(0.0)
        df["amount"] = df["debit"] - df["credit"]
    else:
        df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0.0)

    df["abs_amount"] = df["amount"].abs()
    df["matched"] = False
    df["match_id"] = ""

    return df


def match_bank_to_gl(bank: pd.DataFrame, gl: pd.DataFrame, tolerance: float) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Simple deterministic matching logic:
        1. Match exact opposite/same amount depending on sign conventions by absolute amount.
        2. Prefer same date.
        3. Use one-to-one matching.
    """
    bank = bank.copy()
    gl = gl.copy()
    matches = []

    match_counter = 1

    for bank_idx, bank_row in bank.iterrows():
        if bank.loc[bank_idx, "matched"]:
            continue

        candidates = gl[
            (~gl["matched"]) &
            ((gl["abs_amount"] - bank_row["abs_amount"]).abs() <= tolerance)
        ].copy()

        if candidates.empty:
            continue

        candidates["date_diff"] = (candidates["posting_date"] - bank_row["transaction_date"]).abs().dt.days
        candidates = candidates.sort_values(["date_diff", "gl_entry_id"])
        best_idx = candidates.index[0]

        match_id = f"BM-{match_counter:06d}"
        match_counter += 1

        bank.loc[bank_idx, "matched"] = True
        bank.loc[bank_idx, "match_id"] = match_id

        gl.loc[best_idx, "matched"] = True
        gl.loc[best_idx, "match_id"] = match_id

        matches.append({
            "Match_ID": match_id,
            "Bank_Transaction_ID": bank_row["bank_transaction_id"],
            "GL_Entry_ID": gl.loc[best_idx, "gl_entry_id"],
            "Bank_Date": bank_row["transaction_date"],
            "GL_Date": gl.loc[best_idx, "posting_date"],
            "Bank_Amount": bank_row["amount"],
            "GL_Amount": gl.loc[best_idx, "amount"],
            "Amount_Difference": round(bank_row["abs_amount"] - gl.loc[best_idx, "abs_amount"], 2),
            "Match_Type": "Amount match",
            "Status": "Matched",
        })

    matched_df = pd.DataFrame(matches)

    unmatched_bank = bank[~bank["matched"]].copy()
    unmatched_gl = gl[~gl["matched"]].copy()

    return matched_df, unmatched_bank, unmatched_gl


def generate_bank_reconciliation(
    bank_statement: pd.DataFrame,
    general_ledger_cash: pd.DataFrame,
    parameters: pd.DataFrame,
    output_dir: Path,
) -> dict:
    tolerance = get_amount_tolerance(parameters)

    bank = standardize_bank_statement(bank_statement)
    gl = standardize_gl_cash(general_ledger_cash)

    matched, unmatched_bank, unmatched_gl = match_bank_to_gl(bank, gl, tolerance)

    matched.to_excel(output_dir / "bank_reconciliation.xlsx", index=False)
    unmatched_bank.to_excel(output_dir / "unmatched_bank_transactions.xlsx", index=False)
    unmatched_gl.to_excel(output_dir / "unmatched_gl_cash_transactions.xlsx", index=False)

    return {
        "bank_reconciliation": matched,
        "unmatched_bank_transactions": unmatched_bank,
        "unmatched_gl_cash_transactions": unmatched_gl,
        "bank_total": bank["amount"].sum(),
        "gl_cash_total": gl["amount"].sum(),
        "matched_count": len(matched),
        "unmatched_bank_count": len(unmatched_bank),
        "unmatched_gl_count": len(unmatched_gl),
    }
