"""E3 prototype change journal for bulk and destructive operations.

Three distinct mechanisms, kept separate on purpose:
  - rollback: automatic, when an operation fails part-way
  - custom undo: an explicit command, refused if the cells changed since
  - Excel Ctrl+Z: only via Application.OnUndo from VBA (see E3Commands.bas)
"""

import json
import os
import time
import uuid
from dataclasses import asdict
from pathlib import Path

from excel_link import ExcelGone, Identity, attach, call

ERROR_NAMES = {
    2000: "#NULL!", 2007: "#DIV/0!", 2015: "#VALUE!", 2023: "#REF!", 2029: "#NAME?",
    2036: "#NUM!", 2042: "#N/A", 2043: "#GETTING_DATA", 2045: "#SPILL!", 2046: "#CONNECT!",
    2047: "#BLOCKED!", 2048: "#UNKNOWN!", 2049: "#FIELD!", 2050: "#CALC!",
}
COM_ERROR_BASE = -2146828288


class OperationFailed(Exception):
    pass


def as_2d(value) -> list[list]:
    if isinstance(value, (tuple, list)):
        return [list(row) for row in value]
    return [[value]]


def clean(value):
    if isinstance(value, int) and not isinstance(value, bool):
        code = value - COM_ERROR_BASE
        if 2000 <= code <= 2100:
            return {"error": ERROR_NAMES.get(code, f"#ERR{code}")}
    return value


def value_input(value):
    """What to write so Excel stores exactly this constant (no re-parsing of text)."""
    if isinstance(value, dict):
        return value["error"]
    if isinstance(value, str):
        return "'" + value
    if value is None:
        return ""
    return value


def restore_input(cell: list):
    formula, value = cell
    if isinstance(formula, str) and formula.startswith("="):
        return formula
    return value_input(value)


# ---------- Cell blocks ----------

def read_block(app, workbook, sheet: str, address: str) -> list[list[list]]:
    rng = call(app, lambda: workbook.Worksheets(sheet).Range(address))
    formulas = as_2d(call(app, lambda: rng.Formula2))
    values = as_2d(call(app, lambda: rng.Value2))
    return [[[formulas[r][c], clean(values[r][c])] for c in range(len(formulas[0]))] for r in range(len(formulas))]


def write_block(app, workbook, sheet: str, address: str, inputs: list[list]) -> None:
    rng = call(app, lambda: workbook.Worksheets(sheet).Range(address))
    payload = inputs[0][0] if len(inputs) == 1 and len(inputs[0]) == 1 else tuple(tuple(row) for row in inputs)
    call(app, lambda: setattr(rng, "Formula2", payload))


def write_block_cell_by_cell(app, workbook, sheet: str, address: str, inputs: list[list], progress) -> None:
    rng = call(app, lambda: workbook.Worksheets(sheet).Range(address))
    done = 0
    for r, row in enumerate(inputs, start=1):
        for c, item in enumerate(row, start=1):
            target = call(app, lambda r=r, c=c: rng.Cells(r, c))
            call(app, lambda t=target, i=item: setattr(t, "Formula2", i))
            done += 1
            if done % 100 == 0:
                progress(done)
    progress(done)


def differing_cells(app, workbook, block: dict, current: list[list[list]], expected: list[list[list]]) -> list[str]:
    rng = call(app, lambda: workbook.Worksheets(block["sheet"]).Range(block["address"]))
    out = []
    for r, row in enumerate(current):
        for c, cell in enumerate(row):
            if cell != expected[r][c]:
                out.append(f"{block['sheet']}!{call(app, lambda r=r, c=c: rng.Cells(r + 1, c + 1).Address)}")
    return out


# ---------- Journal ----------

class Journal:
    def __init__(self, directory: Path, identity: Identity, command: str, record: dict | None = None):
        self.directory = directory
        if record is None:
            record = {
                "id": uuid.uuid4().hex[:12],
                "created": time.strftime("%Y-%m-%d %H:%M:%S"),
                "identity": asdict(identity),
                "command": command,
                "status": "planned",
                "blocks": [],
                "names": None,
                "events": [],
            }
        self.record = record
        self.path = directory / f"{record['id']}_{command}.json"

    @classmethod
    def load(cls, path: Path) -> "Journal":
        record = json.loads(path.read_text(encoding="utf-8"))
        return cls(path.parent, Identity(**record["identity"]), record["command"], record)

    @property
    def identity(self) -> Identity:
        return Identity(**self.record["identity"])

    def save(self) -> None:
        self.directory.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.record, indent=1, default=str), encoding="utf-8")
        os.replace(tmp, self.path)

    def set_status(self, status: str, note: str = "") -> None:
        self.record["status"] = status
        self.record["events"].append(f"{time.strftime('%H:%M:%S')} {status}{': ' + note if note else ''}")
        self.save()


# ---------- Bulk operation with rollback ----------

def run_block_operation(identity: Identity, command: str, blocks: list[tuple[str, str]], transform,
                        journal_dir: Path, cell_by_cell: bool = False) -> Journal:
    """Capture every block, then apply transform(before) block by block.
    On failure: roll back what changed. If Excel disappears: mark the journal incomplete."""
    app, workbook = attach(identity)
    journal = Journal(journal_dir, identity, command)
    journal.record["blocks"] = [{"sheet": s, "address": a, "before": None, "after": None, "applied": False}
                                for s, a in blocks]
    for block in journal.record["blocks"]:
        block["before"] = read_block(app, workbook, block["sheet"], block["address"])
    journal.set_status("captured")

    try:
        for block in journal.record["blocks"]:
            inputs = transform(block["before"])
            if cell_by_cell:
                def progress(done, block=block):
                    block["cells_written"] = done
                    journal.save()
                write_block_cell_by_cell(app, workbook, block["sheet"], block["address"], inputs, progress)
            else:
                write_block(app, workbook, block["sheet"], block["address"], inputs)
            block["applied"] = True
            journal.save()
        for block in journal.record["blocks"]:
            block["after"] = read_block(app, workbook, block["sheet"], block["address"])
        journal.set_status("applied")
        return journal
    except ExcelGone as error:
        journal.set_status("incomplete", f"{error} Nothing could be restored; the workbook's unsaved state is lost with Excel.")
        raise
    except Exception as error:  # noqa: BLE001 - any failure triggers rollback
        journal.set_status("failed", f"{type(error).__name__}: {error}")
        rollback(app, workbook, journal)
        raise OperationFailed(f"{command} failed and was rolled back: {error}") from error


def rollback(app, workbook, journal: Journal) -> None:
    restored = []
    for block in journal.record["blocks"]:
        current = read_block(app, workbook, block["sheet"], block["address"])
        if current != block["before"]:
            write_block(app, workbook, block["sheet"], block["address"],
                        [[restore_input(cell) for cell in row] for row in block["before"]])
            restored.append(f"{block['sheet']}!{block['address']}")
    for block in journal.record["blocks"]:
        if read_block(app, workbook, block["sheet"], block["address"]) != block["before"]:
            journal.set_status("rollback_failed", f"{block['sheet']}!{block['address']} does not match its original state")
            raise OperationFailed("rollback could not restore every cell")
    journal.set_status("rolled_back", f"restored {restored or 'nothing (no cells had changed)'}")


def undo_blocks(journal_path: Path) -> dict:
    """Custom Undo: refuse if any affected cell changed since the operation."""
    journal = Journal.load(journal_path)
    if journal.record["status"] != "applied":
        return {"ok": False, "reason": f"journal status is {journal.record['status']!r}, not 'applied'"}
    app, workbook = attach(journal.identity)
    conflicts = []
    for block in journal.record["blocks"]:
        current = read_block(app, workbook, block["sheet"], block["address"])
        conflicts += differing_cells(app, workbook, block, current, block["after"])
    if conflicts:
        journal.set_status("applied", f"undo refused; cells changed since the operation: {conflicts}")
        return {"ok": False, "reason": "cells changed since the operation", "conflicts": conflicts}
    for block in journal.record["blocks"]:
        write_block(app, workbook, block["sheet"], block["address"],
                    [[restore_input(cell) for cell in row] for row in block["before"]])
    mismatched = [f"{b['sheet']}!{b['address']}" for b in journal.record["blocks"]
                  if read_block(app, workbook, b["sheet"], b["address"]) != b["before"]]
    if mismatched:
        journal.set_status("undo_incomplete", f"not restored exactly: {mismatched}")
        return {"ok": False, "reason": "restore verification failed", "blocks": mismatched}
    journal.set_status("undone")
    return {"ok": True}


# ---------- Destructive: names ----------

def capture_names(app, workbook) -> list[dict]:
    def read():
        names = []
        for name in workbook.Names:
            full = name.Name
            scope, local = ("workbook", full) if "!" not in full else full.split("!", 1)
            names.append({"full": full, "scope": scope.strip("'"), "local": local,
                          "refers_to": name.RefersTo, "visible": bool(name.Visible)})
        return sorted(names, key=lambda n: n["full"])
    return call(app, read)


def _add_name(app, workbook, name: dict) -> None:
    owner = workbook if name["scope"] == "workbook" else workbook.Worksheets(name["scope"])
    call(app, lambda: owner.Names.Add(name["local"], name["refers_to"], name["visible"]))


def preview_delete_names(identity: Identity) -> list[dict]:
    app, workbook = attach(identity)
    return capture_names(app, workbook)


def delete_names(identity: Identity, journal_dir: Path, confirmed: bool, fail_after: int | None = None) -> Journal:
    """Destructive: requires explicit confirmation; journals every name; rolls back on failure."""
    if not confirmed:
        raise OperationFailed("delete names needs explicit confirmation after the preview")
    app, workbook = attach(identity)
    journal = Journal(journal_dir, identity, "delete_names")
    journal.record["names"] = {"before": capture_names(app, workbook), "deleted": []}
    journal.set_status("captured")
    try:
        for index, name in enumerate(journal.record["names"]["before"]):
            if fail_after is not None and index == fail_after:
                raise RuntimeError(f"simulated failure after deleting {fail_after} names")
            call(app, lambda n=name: workbook.Names(n["full"]).Delete())
            journal.record["names"]["deleted"].append(name["full"])
            journal.save()
        journal.record["names"]["after"] = capture_names(app, workbook)
        journal.set_status("applied")
        return journal
    except ExcelGone:
        journal.set_status("incomplete", "Excel closed during the operation")
        raise
    except Exception as error:  # noqa: BLE001
        journal.set_status("failed", str(error))
        _restore_names(app, workbook, journal)
        raise OperationFailed(f"delete names failed and was rolled back: {error}") from error


def _restore_names(app, workbook, journal: Journal) -> None:
    before = journal.record["names"]["before"]
    present = {n["full"] for n in capture_names(app, workbook)}
    for name in before:
        if name["full"] not in present:
            _add_name(app, workbook, name)
    if capture_names(app, workbook) != before:
        journal.set_status("rollback_failed", "names differ from the original set")
        raise OperationFailed("names could not be restored exactly")
    journal.set_status("rolled_back")


def undo_delete_names(journal_path: Path) -> dict:
    journal = Journal.load(journal_path)
    if journal.record["status"] != "applied":
        return {"ok": False, "reason": f"journal status is {journal.record['status']!r}"}
    app, workbook = attach(journal.identity)
    current = {n["full"] for n in capture_names(app, workbook)}
    clashes = sorted(current & set(journal.record["names"]["deleted"]))
    if clashes:
        return {"ok": False, "reason": "names re-created since the operation", "conflicts": clashes}
    for name in journal.record["names"]["before"]:
        _add_name(app, workbook, name)
    if capture_names(app, workbook) != journal.record["names"]["before"]:
        journal.set_status("undo_incomplete", "names differ after restore")
        return {"ok": False, "reason": "restore verification failed"}
    journal.set_status("undone")
    return {"ok": True}
