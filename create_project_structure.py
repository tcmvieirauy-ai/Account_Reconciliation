from pathlib import Path

project_root = Path("Account-Reconciliation-Engine")

folders = [
    project_root,
    project_root / "data",
    project_root / "outputs",
]

for folder in folders:
    folder.mkdir(parents=True, exist_ok=True)

print("Project structure created successfully.")
print()
print("Place input Excel files inside data/:")
print("- bank_statement.xlsx")
print("- general_ledger_cash.xlsx")
print("- accounts_receivable_subledger.xlsx")
print("- accounts_payable_subledger.xlsx")
print("- reconciliation_parameters.xlsx")
print("- reconciliation_chart_of_accounts.xlsx")
print()
print("Place Python files in the project root.")
print("Run: python main.py")
