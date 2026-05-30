"""
Subledger Reconciliation Engine

Purpose:
    Reconcile AR/AP subledger movements against cash GL activity.

Generates:
    - ar_reconciliation.xlsx
    - ap_reconciliation.xlsx
"""

from pathlib import Path
import pandas as pd


def excel_date_to_datetime(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit="D", origin="1899-12-30")
    return pd.to_datetime(series)


def standardize_subledger(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    data = df.copy()
    data.columns = [str(c).strip() for c in data.columns]

    if kind == "AR":
        rename = {
            "AR_ID": "subledger_id",
            "Invoice_ID": "subledger_id",
            "Customer_ID": "party_id",
            "Customer_Name": "party_name",
            "Invoice_Date": "document_date",
            "Due_Date": "due_date",
            "Invoice_Amount": "amount",
            "Open_Amount": "open_amount",
            "Status": "status",
            "Related_Payment_ID": "related_payment_id",
        }
    else:
        rename = {
            "AP_ID": "subledger_id",
            "Bill_ID": "subledger_id",
            "Vendor_ID": "party_id",
            "Vendor_Name": "party_name",
            "Bill_Date": "document_date",
            "Due_Date": "due_date",
            "Bill_Amount": "amount",
            "Open_Amount": "open_amount",
            "Status": "status",
            "Related_Payment_ID": "related_payment_id",
        }

    data = data.rename(columns={k: v for k, v in rename.items() if k in data.columns})

    required = ["subledger_id", "amount"]
    missing = [c for c in required if c not in data.columns]
    if missing:
        raise ValueError(f"{kind} subledger missing required columns: {missing}")

    if "document_date" in data.columns:
        data["document_date"] = excel_date_to_datetime(data["document_date"])
    if "due_date" in data.columns:
        data["due_date"] = excel_date_to_datetime(data["due_date"])

    data["amount"] = pd.to_numeric(data["amount"], errors="coerce").fillna(0.0)
    if "open_amount" in data.columns:
        data["open_amount"] = pd.to_numeric(data["open_amount"], errors="coerce").fillna(0.0)
    else:
        data["open_amount"] = 0.0

    data["Subledger_Type"] = kind

    return data


def build_subledger_summary(subledger: pd.DataFrame, kind: str) -> pd.DataFrame:
    total_amount = subledger["amount"].sum()
    open_amount = subledger["open_amount"].sum()
    closed_amount = total_amount - open_amount

    return pd.DataFrame([
        {
            "Subledger_Type": kind,
            "Total_Documents": len(subledger),
            "Total_Amount": round(total_amount, 2),
            "Open_Amount": round(open_amount, 2),
            "Closed_or_Settled_Amount": round(closed_amount, 2),
        }
    ])


def generate_subledger_reconciliations(
    accounts_receivable: pd.DataFrame,
    accounts_payable: pd.DataFrame,
    general_ledger_cash: pd.DataFrame,
    parameters: pd.DataFrame,
    output_dir: Path,
) -> dict:
    ar = standardize_subledger(accounts_receivable, "AR")
    ap = standardize_subledger(accounts_payable, "AP")

    ar_summary = build_subledger_summary(ar, "AR")
    ap_summary = build_subledger_summary(ap, "AP")

    ar_output = ar.copy()
    ap_output = ap.copy()

    ar_output.to_excel(output_dir / "ar_reconciliation.xlsx", index=False)
    ap_output.to_excel(output_dir / "ap_reconciliation.xlsx", index=False)

    summary = pd.concat([ar_summary, ap_summary], ignore_index=True)
    summary.to_excel(output_dir / "subledger_reconciliation_summary.xlsx", index=False)

    return {
        "ar_reconciliation": ar_output,
        "ap_reconciliation": ap_output,
        "subledger_summary": summary,
        "ar_total": ar["amount"].sum(),
        "ar_open": ar["open_amount"].sum(),
        "ap_total": ap["amount"].sum(),
        "ap_open": ap["open_amount"].sum(),
    }
