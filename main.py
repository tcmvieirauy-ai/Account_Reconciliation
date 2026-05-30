"""
Account Reconciliation Engine - Main Orchestrator

Run:
    python main.py

Expected structure:
    data/
        bank_statement.xlsx
        general_ledger_cash.xlsx
        accounts_receivable_subledger.xlsx
        accounts_payable_subledger.xlsx
        reconciliation_parameters.xlsx
        reconciliation_chart_of_accounts.xlsx

    outputs/
        generated automatically
"""

from pathlib import Path
import pandas as pd

from bank_reconciliation_engine import generate_bank_reconciliation
from subledger_reconciliation_engine import generate_subledger_reconciliations
from exceptions_engine import generate_exceptions_report
from controls_engine import generate_reconciliation_controls

DATA_DIR = Path("data")
OUTPUT_DIR = Path("outputs")


def read_inputs() -> dict:
    required_files = {
        "bank_statement": "bank_statement.xlsx",
        "general_ledger_cash": "general_ledger_cash.xlsx",
        "accounts_receivable": "accounts_receivable_subledger.xlsx",
        "accounts_payable": "accounts_payable_subledger.xlsx",
        "parameters": "reconciliation_parameters.xlsx",
        "chart_of_accounts": "reconciliation_chart_of_accounts.xlsx",
    }

    inputs = {}

    for key, filename in required_files.items():
        path = DATA_DIR / filename
        if not path.exists():
            raise FileNotFoundError(f"Missing required file: {path}")
        inputs[key] = pd.read_excel(path)

    return inputs


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)

    print("1. Reading input files...")
    inputs = read_inputs()

    print("2. Running bank reconciliation...")
    bank_outputs = generate_bank_reconciliation(
        bank_statement=inputs["bank_statement"],
        general_ledger_cash=inputs["general_ledger_cash"],
        parameters=inputs["parameters"],
        output_dir=OUTPUT_DIR,
    )

    print("3. Running AR/AP subledger reconciliations...")
    subledger_outputs = generate_subledger_reconciliations(
        accounts_receivable=inputs["accounts_receivable"],
        accounts_payable=inputs["accounts_payable"],
        general_ledger_cash=inputs["general_ledger_cash"],
        parameters=inputs["parameters"],
        output_dir=OUTPUT_DIR,
    )

    print("4. Generating exceptions report...")
    exceptions_output = generate_exceptions_report(
        bank_outputs=bank_outputs,
        subledger_outputs=subledger_outputs,
        output_dir=OUTPUT_DIR,
    )

    print("5. Running reconciliation controls...")
    generate_reconciliation_controls(
        bank_outputs=bank_outputs,
        subledger_outputs=subledger_outputs,
        exceptions_output=exceptions_output,
        output_dir=OUTPUT_DIR,
    )

    print("Done. Check the outputs folder.")


if __name__ == "__main__":
    main()
