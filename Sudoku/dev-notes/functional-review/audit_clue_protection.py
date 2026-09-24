"""Fail if a Sudoku page allows fixed clues to be cleared.

Run this read-only check against the saved workbook before claiming that fixed
clues are protected. Data validation alone is not sufficient: Excel Delete
clears a cell without invoking validation.
"""

import argparse
import sys
import warnings

import openpyxl


PAGES = ("01 Easy", "02 Medium", "03 Hard", "04 Expert", "05 Master")


def audit(path):
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Conditional Formatting extension")
        workbook = openpyxl.load_workbook(path, read_only=False)
    failures = []
    for name in PAGES:
        sheet = workbook[name]
        if not sheet.protection.sheet:
            failures.append(f"{name}: sheet protection is off")
        locked_inputs = []
        unlocked_clues = []
        for row in range(9):
            for col in range(9):
                cell = sheet.cell(5 + 3 * row, 3 + 4 * col)
                if cell.value is None and cell.protection.locked:
                    locked_inputs.append(cell.coordinate)
                elif cell.value is not None and not cell.protection.locked:
                    unlocked_clues.append(cell.coordinate)
        if locked_inputs:
            failures.append(f"{name}: {len(locked_inputs)} player inputs remain locked (for example {locked_inputs[0]})")
        if unlocked_clues:
            failures.append(f"{name}: {len(unlocked_clues)} fixed clues are unlocked (for example {unlocked_clues[0]})")
        if sheet["AS11"].protection.locked:
            failures.append(f"{name}: Coach mode AS11 remains locked")
    workbook.close()
    for failure in failures:
        print(f"FAIL {failure}")
    if not failures:
        print("PASS: five protected pages keep clues locked and player inputs and Coach mode unlocked")
    return not failures


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook")
    args = parser.parse_args()
    sys.exit(0 if audit(args.workbook) else 1)
