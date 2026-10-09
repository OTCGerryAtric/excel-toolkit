"""Load E1Test.xlam into the running Excel, run the harness, summarise results.

The harness works on a new scratch workbook that is closed without saving.
PERSONAL.XLSB is only called, never edited or saved.
"""

import csv
import hashlib
import statistics
from collections import defaultdict
from pathlib import Path

import win32com.client

HERE = Path(__file__).resolve().parent
XLAM = HERE / "E1Test.xlam"
RESULTS = HERE / "results"
PERSONAL = Path.home() / "AppData/Roaming/Microsoft/Excel/XLSTART/PERSONAL.XLSB"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def addin_state(excel) -> list[tuple[str, bool]]:
    return sorted((addin.Name, bool(addin.Installed)) for addin in excel.AddIns)


def p95(values: list[float]) -> float:
    ordered = sorted(values)
    return ordered[max(0, round(0.95 * len(ordered)) - 1)]


def main() -> None:
    RESULTS.mkdir(exist_ok=True)
    personal_before = sha256(PERSONAL)

    excel = win32com.client.GetActiveObject("Excel.Application")
    addins_before = addin_state(excel)
    open_before = [wb.Name for wb in excel.Workbooks]
    if "PERSONAL.XLSB" not in open_before:
        raise SystemExit("PERSONAL.XLSB is not open in Excel; cannot compare against the legacy macro.")

    if "E1Test.xlam" not in open_before:
        excel.Workbooks.Open(str(XLAM))
    excel.Run("E1Test.xlam!E1_Init")  # Auto_Open does not fire when opened from code
    excel.Run("E1Test.xlam!E1_RunAll", str(RESULTS))

    addins_after = addin_state(excel)
    personal_after = sha256(PERSONAL)

    with open(RESULTS / "e1_equivalence.csv", newline="") as f:
        equivalence = list(csv.DictReader(f))
    mismatches = [row for row in equivalence if row["match"] != "True"]

    timings: dict[tuple[str, str], list[float]] = defaultdict(list)
    with open(RESULTS / "e1_timings.csv", newline="") as f:
        for row in csv.DictReader(f):
            timings[(row["macro"], row["selection"])].append(float(row["ms"]))

    lines = ["# E1 results", "", f"Equivalence checks: {len(equivalence)}, mismatches: {len(mismatches)}", ""]
    lines += [f"- MISMATCH {r['test']} {r['selection']} from {r['start_state']}: legacy={r['legacy_result']} new={r['new_result']}" for r in mismatches]
    lines += ["", "| Macro | Selection | Runs | Median ms | p95 ms |", "|---|---|---|---|---|"]
    for (macro, selection), values in sorted(timings.items(), key=lambda item: (item[0][1], item[0][0])):
        lines.append(f"| {macro} | {selection} | {len(values)} | {statistics.median(values):.3f} | {p95(values):.3f} |")
    lines += [
        "",
        f"PERSONAL.XLSB unchanged on disk: {personal_before == personal_after}",
        f"Excel add-in states unchanged: {addins_before == addins_after}",
        f"Add-ins: {addins_after}",
    ]
    summary = "\n".join(lines)
    (RESULTS / "e1_summary.md").write_text(summary + "\n", encoding="utf-8")
    print(summary)


if __name__ == "__main__":
    main()
