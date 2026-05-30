"""
Exceptions Engine

Purpose:
    Consolidate unmatched and open items into one exception report.

Generates:
    - exceptions_report.xlsx
"""

from pathlib import Path
import pandas as pd


def generate_exceptions_report(
    bank_outputs: dict,
    subledger_outputs: dict,
    output_dir: Path,
) -> pd.DataFrame:
    rows = []

    unmatched_bank = bank_outputs["unmatched_bank_transactions"]
    unmatched_gl = bank_outputs["unmatched_gl_cash_transactions"]

    for _, row in unmatched_bank.iterrows():
        rows.append({
            "Exception_Type": "Unmatched Bank Transaction",
            "Reference_ID": row.get("bank_transaction_id", ""),
            "Date": row.get("transaction_date", ""),
            "Amount": row.get("amount", 0.0),
            "Description": row.get("description", ""),
            "Recommended_Action": "Investigate bank transaction not found in GL cash account.",
            "Priority": "High",
        })

    for _, row in unmatched_gl.iterrows():
        rows.append({
            "Exception_Type": "Unmatched GL Cash Entry",
            "Reference_ID": row.get("gl_entry_id", ""),
            "Date": row.get("posting_date", ""),
            "Amount": row.get("amount", 0.0),
            "Description": row.get("description", ""),
            "Recommended_Action": "Investigate GL cash entry not found in bank statement.",
            "Priority": "High",
        })

    ar = subledger_outputs["ar_reconciliation"]
    ap = subledger_outputs["ap_reconciliation"]

    if "open_amount" in ar.columns:
        for _, row in ar[ar["open_amount"].abs() > 0.01].iterrows():
            rows.append({
                "Exception_Type": "Open AR Item",
                "Reference_ID": row.get("subledger_id", ""),
                "Date": row.get("document_date", ""),
                "Amount": row.get("open_amount", 0.0),
                "Description": row.get("party_name", ""),
                "Recommended_Action": "Review open receivable and follow collection process.",
                "Priority": "Medium",
            })

    if "open_amount" in ap.columns:
        for _, row in ap[ap["open_amount"].abs() > 0.01].iterrows():
            rows.append({
                "Exception_Type": "Open AP Item",
                "Reference_ID": row.get("subledger_id", ""),
                "Date": row.get("document_date", ""),
                "Amount": row.get("open_amount", 0.0),
                "Description": row.get("party_name", ""),
                "Recommended_Action": "Review payable status and payment schedule.",
                "Priority": "Medium",
            })

    exceptions = pd.DataFrame(rows)
    exceptions.to_excel(output_dir / "exceptions_report.xlsx", index=False)

    return exceptions
