from pathlib import Path
import pandas as pd

OUTPUT_DIR = Path("outputs")

FILES_TO_CONSOLIDATE = {
    "Bank_Reconciliation": "bank_reconciliation.xlsx",
    "Unmatched_Bank": "unmatched_bank_transactions.xlsx",
    "Unmatched_GL": "unmatched_gl_cash_transactions.xlsx",
    "AR_Reconciliation": "ar_reconciliation.xlsx",
    "AP_Reconciliation": "ap_reconciliation.xlsx",
    "Subledger_Summary": "subledger_reconciliation_summary.xlsx",
    "Exceptions_Report": "exceptions_report.xlsx",
    "Control_Check": "reconciliation_control_check.xlsx",
}

CONSOLIDATED_FILE = OUTPUT_DIR / "Account_Reconciliation_Consolidated_Report.xlsx"


def main():
    if not OUTPUT_DIR.exists():
        raise FileNotFoundError("outputs folder not found. Run main.py first.")

    with pd.ExcelWriter(CONSOLIDATED_FILE, engine="openpyxl") as writer:
        for sheet_name, file_name in FILES_TO_CONSOLIDATE.items():
            file_path = OUTPUT_DIR / file_name

            if not file_path.exists():
                print(f"Skipping missing file: {file_path}")
                continue

            df = pd.read_excel(file_path)

            # Excel sheet names have a 31-character limit
            safe_sheet_name = sheet_name[:31]

            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

    print(f"Consolidated report created: {CONSOLIDATED_FILE}")


if __name__ == "__main__":
    main()