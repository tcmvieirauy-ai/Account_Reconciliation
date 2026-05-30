# Account Reconciliation Engine

Python-based reconciliation automation project for accounting close activities.

## Scope

Version 1 includes:

- Bank reconciliation
- GL cash reconciliation
- AR subledger review
- AP subledger review
- Exceptions report
- Reconciliation control checks

## Inputs

Place these files inside `data/`:

- bank_statement.xlsx
- general_ledger_cash.xlsx
- accounts_receivable_subledger.xlsx
- accounts_payable_subledger.xlsx
- reconciliation_parameters.xlsx
- reconciliation_chart_of_accounts.xlsx

## Python files

- main.py
- bank_reconciliation_engine.py
- subledger_reconciliation_engine.py
- exceptions_engine.py
- controls_engine.py
- create_project_structure.py

## Outputs

Generated in `outputs/`:

- bank_reconciliation.xlsx
- unmatched_bank_transactions.xlsx
- unmatched_gl_cash_transactions.xlsx
- ar_reconciliation.xlsx
- ap_reconciliation.xlsx
- subledger_reconciliation_summary.xlsx
- exceptions_report.xlsx
- reconciliation_control_check.xlsx

## Run

```bash
pip install pandas openpyxl
python main.py
```
