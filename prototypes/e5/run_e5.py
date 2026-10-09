"""E5 orchestrator: Excel COM dependency tracing vs the static Python prototype.

Runs in a new Excel process on a disposable copy. Trace arrows and navigation change the
selection and active sheet (UI state), so both are restored, and workbook contents are
verified unchanged afterwards.
"""

import json
import re
import shutil
import sys
import time
from pathlib import Path

import pywintypes
import win32api
import win32com.client
import win32con
import win32process

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from deps_parser import Context, dependencies  # noqa: E402
from make_fixtures import SCENARIOS  # noqa: E402

FIXTURES = HERE / "fixtures"
WORK = HERE / "work" / time.strftime("run_%Y%m%d_%H%M%S")
RESULTS = HERE / "results"
EXTERNAL = "[E5_External.xlsx]Sheet1!A1"

# What each formula really depends on (canonical "Sheet!A1", no $). Run-time targets included.
TRUTH = {
    "S1": {"Calc!F1", "Calc!G1"},
    "S2": {"Inputs!A1"},
    "S3": {"Calc!F2:F101"},
    "S4": {"Calc!F1", "Calc!G2"},
    "S5": {"Data!B2:B6"},
    "S6": {"Inputs!B1"},
    "S7": {"Calc!H1", "Inputs!A2"},
    "S8": {"Calc!F1", "Calc!G1"},
    "S9": {"Calc!D1:D5"},
    "S10": {EXTERNAL},
    "S11": {"Calc!F1", "Calc!G1"},
    "S12": {"Inputs!A1", "Calc!A1", "Data!A1"},
    "S13": {"Calc!F:F"},
    "S14": {"My Sheet!A1"},
}
# Equivalent spellings: whole column F:F is F1:F101 in this workbook's used range;
# a spill anchor D1# is D1:D5 at run time.
EQUIVALENT = {"Calc!F1:F101": "Calc!F:F", "Calc!D1#": "Calc!D1:D5"}


def _column_number(letters: str) -> int:
    number = 0
    for char in letters:
        number = number * 26 + ord(char) - 64
    return number


def _column_letters(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def canonical(items: set[str]) -> set[str]:
    """Compare cells, not spellings: 'Calc!F1:G1' == {'Calc!F1', 'Calc!G1'}."""
    out = set()
    for item in items:
        item = EQUIVALENT.get(item, item)
        prefix, _, address = item.rpartition("!")
        bounded = re.match(r"^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$", address)
        if not bounded:
            out.add(item)  # whole column/row or symbolic: keep as is
            continue
        c1, r1 = _column_number(bounded.group(1)), int(bounded.group(2))
        c2 = _column_number(bounded.group(3)) if bounded.group(3) else c1
        r2 = int(bounded.group(4)) if bounded.group(4) else r1
        out |= {f"{prefix}!{_column_letters(c)}{r}" for c in range(c1, c2 + 1) for r in range(r1, r2 + 1)}
    return out


def split_areas(sheet: str, address: str, book: str | None = None) -> set[str]:
    prefix = f"[{book}]{sheet}" if book else sheet
    return {f"{prefix}!{area.replace('$', '')}" for area in address.split(",")}


def verdict(found: set[str], flag: str, truth: set[str]) -> str:
    found, truth = canonical(found), canonical(truth)
    if flag == "resolved" and found == truth:
        return "resolved"
    missing, extra = truth - found, found - truth
    if flag in ("partial", "indeterminate"):
        return f"{flag} (flagged){'; missing ' + ', '.join(sorted(missing)) if missing else ''}"
    detail = "; ".join(x for x in (f"missing {', '.join(sorted(missing))}" if missing else "",
                                   f"extra {', '.join(sorted(extra))}" if extra else "") if x)
    return f"WRONG, silently — {detail}"


# ---------- Excel COM ----------

def com_direct_precedents(cell) -> set[str]:
    try:
        precedents = cell.DirectPrecedents
    except pywintypes.com_error:
        return set()
    return split_areas(cell.Worksheet.Name, precedents.Address)


def com_navigate_arrows(app, cell, workbook_name: str) -> tuple[set[str], list[str]]:
    """Trace-precedents arrows, including the dashed arrow to other sheets and workbooks."""
    found, notes = set(), []
    sheet = cell.Worksheet
    sheet.Activate()
    cell.ShowPrecedents()
    for arrow in range(1, 30):
        link = 1
        while link < 60:
            try:
                target = cell.NavigateArrow(True, arrow, link)
            except pywintypes.com_error as error:
                if link == 1:
                    notes.append(f"arrow {arrow}: {error.excepinfo[2] if error.excepinfo else error}")
                break
            target_book = target.Worksheet.Parent.Name
            target_sheet = target.Worksheet.Name
            if target_book == workbook_name and target_sheet == sheet.Name and target.Address == cell.Address:
                break
            found |= split_areas(target_sheet, target.Address, None if target_book == workbook_name else target_book)
            link += 1
        if link == 1:
            break
    sheet.Activate()
    sheet.ClearArrows()
    return found, notes


def snapshot(workbook) -> dict:
    return {s.Name: (s.UsedRange.Address, s.UsedRange.Formula2, s.UsedRange.Value2) for s in workbook.Worksheets}


# ---------- Main ----------

def main() -> None:
    (WORK).mkdir(parents=True, exist_ok=True)
    RESULTS.mkdir(exist_ok=True)
    shutil.copy(FIXTURES / "E5_Deps.xlsx", WORK / "E5_Deps.xlsx")

    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    pid = win32process.GetWindowThreadProcessId(int(excel.Hwnd))[1]
    rows = []
    try:
        workbook = excel.Workbooks.Open(str(WORK / "E5_Deps.xlsx"), UpdateLinks=0)
        calc = workbook.Worksheets("Calc")
        context = Context(
            sheet="Calc",
            sheet_order=[s.Name for s in workbook.Worksheets],
            names={n.Name: n.RefersTo for n in workbook.Names},
            tables={lo.Name: {"sheet": ws.Name, "columns": {c.Name: c.DataBodyRange.Address.replace("$", "") for c in lo.ListColumns}}
                    for ws in workbook.Worksheets for lo in ws.ListObjects},
        )
        before = snapshot(workbook)
        ui_before = (excel.ActiveSheet.Name, excel.Selection.Address)

        for row, (scenario, label, _) in enumerate(SCENARIOS, start=1):
            cell = calc.Cells(row, 2)
            formula = cell.Formula2
            truth = TRUTH[scenario]

            started = time.perf_counter()
            direct = com_direct_precedents(cell)
            direct_ms = (time.perf_counter() - started) * 1000
            books_before = excel.Workbooks.Count
            started = time.perf_counter()
            arrows, arrow_notes = com_navigate_arrows(excel, cell, workbook.Name)
            arrows_ms = (time.perf_counter() - started) * 1000
            opened = excel.Workbooks.Count - books_before

            started = time.perf_counter()
            parsed = dependencies(formula, context)
            parser_ms = (time.perf_counter() - started) * 1000

            rows.append({
                "scenario": scenario, "label": label, "formula": formula, "truth": sorted(truth),
                "com_direct": {"found": sorted(direct), "verdict": verdict(direct, "resolved", truth), "ms": direct_ms},
                "com_arrows": {"found": sorted(arrows), "verdict": verdict(arrows, "resolved", truth), "ms": arrows_ms,
                               "notes": arrow_notes, "workbooks_opened": opened},
                "python": {"found": sorted(parsed.dependencies), "status": parsed.status, "notes": parsed.notes,
                           "anchoring": parsed.absolute, "verdict": verdict(parsed.dependencies, parsed.status, truth),
                           "ms": parser_ms},
            })
            print(f"{scenario:4} {label:28} COM direct: {rows[-1]['com_direct']['verdict'][:60]:60} | "
                  f"COM arrows: {rows[-1]['com_arrows']['verdict'][:60]:60} | Python: {rows[-1]['python']['verdict'][:70]}")

        # Absolute/relative semantics: compare with Excel's own R1C1 form.
        s4 = calc.Cells(4, 2)
        r1c1 = s4.FormulaR1C1
        workbook.Worksheets(ui_before[0]).Activate()
        workbook.Worksheets(ui_before[0]).Range(ui_before[1]).Select()
        after = snapshot(workbook)
        extra_books = [b.Name for b in excel.Workbooks if b.Name != workbook.Name]
        summary = {
            "rows": rows,
            "s4_r1c1": r1c1,
            "content_unchanged": after == before,
            "ui_before": ui_before,
            "ui_after_restore": (excel.ActiveSheet.Name, excel.Selection.Address),
            "other_workbooks_open_at_end": extra_books,
        }
    finally:
        try:
            for book in list(excel.Workbooks):
                book.Close(SaveChanges=False)
            excel.Quit()
        except pywintypes.com_error:
            pass
        excel = None
        deadline = time.monotonic() + 15
        while time.monotonic() < deadline:
            try:
                handle = win32api.OpenProcess(win32con.PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
                if win32process.GetExitCodeProcess(handle) != 259:
                    break
            except pywintypes.error:
                break
            time.sleep(0.2)
        else:
            win32api.TerminateProcess(win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid), 1)

    (RESULTS / "e5_results.json").write_text(json.dumps(summary, indent=1, default=str), encoding="utf-8")
    print(f"\nS4 R1C1: {summary['s4_r1c1']}; Python anchoring: {rows[3]['python']['anchoring']}")
    print(f"Content unchanged: {summary['content_unchanged']}; UI restored: {summary['ui_after_restore']} "
          f"(was {summary['ui_before']}); other workbooks open at end: {summary['other_workbooks_open_at_end']}")


if __name__ == "__main__":
    main()
