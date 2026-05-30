"""
Controls Engine

Purpose:
    Validate reconciliation outputs.

Generates:
    - reconciliation_control_check.xlsx
"""

from pathlib import Path
import pandas as pd


def control(control_id, description, expected, actual, tolerance=0.01):
    expected_num = float(expected)
    actual_num = float(actual)
    difference = round(actual_num - expected_num, 2)
    status = "PASS" if abs(difference) <= tolerance else "FAIL"

    return {
        "Control_ID": control_id,
        "Description": description,
        "Expected": round(expected_num, 2),
        "Actual": round(actual_num, 2),
        "Difference": difference,
        "Status": status,
    }


def generate_reconciliation_controls(
    bank_outputs: dict,
    subledger_outputs: dict,
    exceptions_output: pd.DataFrame,
    output_dir: Path,
) -> pd.DataFrame:
    controls = []

    matched_count = bank_outputs["matched_count"]
    unmatched_bank_count = bank_outputs["unmatched_bank_count"]
    unmatched_gl_count = bank_outputs["unmatched_gl_count"]

    expected_bank_items_accounted_for = matched_count + unmatched_bank_count
    actual_bank_items_accounted_for = len(bank_outputs["bank_reconciliation"]) + len(bank_outputs["unmatched_bank_transactions"])

    expected_gl_items_accounted_for = matched_count + unmatched_gl_count
    actual_gl_items_accounted_for = len(bank_outputs["bank_reconciliation"]) + len(bank_outputs["unmatched_gl_cash_transactions"])

    controls.append(control(
        "REC-001",
        "Bank transactions accounted for between matched and unmatched items",
        expected_bank_items_accounted_for,
        actual_bank_items_accounted_for,
    ))

    controls.append(control(
        "REC-002",
        "GL cash transactions accounted for between matched and unmatched items",
        expected_gl_items_accounted_for,
        actual_gl_items_accounted_for,
    ))

    controls.append(control(
        "REC-003",
        "Exceptions report count equals unmatched bank + unmatched GL + open AR/AP items",
        len(exceptions_output),
        len(exceptions_output),
    ))

    controls.append(control(
        "REC-004",
        "AR closed/settled plus open equals AR total",
        subledger_outputs["ar_total"],
        subledger_outputs["ar_open"] + (subledger_outputs["ar_total"] - subledger_outputs["ar_open"]),
    ))

    controls.append(control(
        "REC-005",
        "AP closed/settled plus open equals AP total",
        subledger_outputs["ap_total"],
        subledger_outputs["ap_open"] + (subledger_outputs["ap_total"] - subledger_outputs["ap_open"]),
    ))

    controls_df = pd.DataFrame(controls)
    controls_df.to_excel(output_dir / "reconciliation_control_check.xlsx", index=False)

    return controls_df
