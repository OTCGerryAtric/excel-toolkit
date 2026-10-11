"""Offline Phase B tests: audit adapter, JSON transport/wire-format, callback correlation,
delivery, error conditions and navigation logic — all with fake data / fake COM, no Excel.

The real detector regression suite (tests/formula_audit) is unchanged and still authoritative
for the detection rules; here we only test the Phase B worker/transport/front-end contract
around it.
"""

import json

import pytest

import audit_engine
from audit_engine import deliver, handle_audit
from phaseb_contract import PendingRequest, classify_callback

from excel_toolkit.formula_audit import (
    InMemoryWorksheet,
    display_json,
    navigation_from_record,
    scan_worksheet,
)
from excel_toolkit.formula_audit.worksheet import WorksheetFinding, WorksheetReport


# --- fake COM worksheet (mirrors pywin32: .Formula gives '=' text or a constant's string
#     form, .Value gives the typed value; single-cell ranges return scalars) -----------------

class _FakeUsedRange:
    def __init__(self, row, col, formula, value):
        self.Row, self.Column, self.Formula, self.Value = row, col, formula, value


class _FakeComWorksheet:
    def __init__(self, name, row, col, formula, value):
        self.Name, self.UsedRange = name, _FakeUsedRange(row, col, formula, value)


class _ExplodingWorksheet:
    """A worksheet whose used-range read raises, to exercise the failure path."""

    def __init__(self, name):
        self.Name = name

    @property
    def UsedRange(self):
        raise RuntimeError("COM blew up reading the used range")


class _FakeBook:
    def __init__(self, sheets):
        self._sheets = sheets  # name -> fake worksheet

    def Worksheets(self, name):
        return self._sheets[name]  # KeyError if missing -> handled as a failed audit


def _passthrough_call(app, op, timeout=5.0):
    return op()


def _attach_to(book):
    return lambda request: (object(), book)


def _request(sheet="Live", wb="Book.xlsx", full=r"C:\m\Book.xlsx", id="p-1-000001", token="tok-1"):
    return {"op": "audit", "id": id, "token": token, "hwnd": 123, "pid": 4242,
            "workbook": wb, "full_name": full, "sheet": sheet}


# A clean 2x4 block: row 1 constants, row 2 formulas with one odd one out at D2.
_FORMULA = (("1", "1", "1", "1"), ("=B1*2", "=C1*2", "=D1*3", "=E1*2"))
_VALUE = ((1, 1, 1, 1), (2, 2, 3, 2))


def _reference_report():
    snap = InMemoryWorksheet(
        {"B1": 1, "C1": 1, "D1": 1, "E1": 1,
         "B2": "=B1*2", "C2": "=C1*2", "D2": "=D1*3", "E2": "=E1*2"}, "Live")
    return scan_worksheet(snap)


# ======================= adapter + transport =======================

def test_handle_audit_matches_direct_detector_and_reports_ok():
    book = _FakeBook({"Live": _FakeComWorksheet("Live", 1, 2, _FORMULA, _VALUE)})
    resp = handle_audit(_request(), attach_fn=_attach_to(book), call_fn=_passthrough_call)

    assert resp["status"] == "ok" and resp["error"] == ""
    assert resp["count"] == 1
    assert resp["id"] == "p-1-000001" and resp["token"] == "tok-1"
    # findings JSON equals the direct detector's flat display JSON
    ref = _reference_report()
    assert json.loads(resp["findings_json"]) == json.loads(display_json(ref, "Book.xlsx", r"C:\m\Book.xlsx"))
    cells = {rec["cell"] for rec in json.loads(resp["findings_json"])}
    assert cells == {"D2"}


def test_findings_json_is_flat_strings_and_bools_only():
    book = _FakeBook({"Live": _FakeComWorksheet("Live", 1, 2, _FORMULA, _VALUE)})
    resp = handle_audit(_request(), attach_fn=_attach_to(book), call_fn=_passthrough_call)
    for rec in json.loads(resp["findings_json"]):
        assert set(rec) == {"workbook", "workbook_path", "sheet", "cell",
                            "confidence", "kind", "conflict", "message"}
        for key, val in rec.items():
            assert isinstance(val, (str, bool)), f"{key} is {type(val)} — not flat"


def test_wire_format_escapes_survive_round_trip():
    # A hostile message with tab/newline/quote/backslash must serialise and parse back intact,
    # pinning the exact escaping the VBA JSON reader must handle.
    nasty = 'tab\there\nnew "line" and a back\\slash'
    finding = WorksheetFinding(
        coord="B2", row=2, col=2, axes=("row",), observed="RC[-1]*3", expected="RC[-1]*2",
        expected_by_axis=(("row", "RC[-1]*2"),), conflict=False, kind="singleton",
        block_by_axis=(), confidence="high", message=nasty)
    report = WorksheetReport(sheet="Weird\tSheet")
    report.findings = [finding]

    text = display_json(report, 'Book "x".xlsx', r"C:\a\b.xlsx")
    assert '\\t' in text and '\\n' in text and '\\"' in text and '\\\\' in text  # escaped on the wire
    rec = json.loads(text)[0]
    assert rec["message"] == nasty                      # round-trips byte-for-byte
    assert rec["sheet"] == "Weird\tSheet"
    assert rec["workbook"] == 'Book "x".xlsx'


# ======================= error conditions (never an empty "success") =======================

def test_identity_failure_reports_failed_not_empty_success():
    def boom(_request):
        raise audit_engine.IdentityError("workbook is no longer open")
    resp = handle_audit(_request(), attach_fn=boom, call_fn=_passthrough_call)
    assert resp["status"] == "failed"
    assert "IdentityError" in resp["error"] and "no longer open" in resp["error"]
    assert resp["findings_json"] == "[]" and resp["count"] == 0


def test_worksheet_name_mismatch_is_a_failure():
    # Worksheets() returns a sheet whose Name differs from the requested one.
    book = _FakeBook({"Live": _FakeComWorksheet("NOT-Live", 1, 2, _FORMULA, _VALUE)})
    resp = handle_audit(_request(sheet="Live"), attach_fn=_attach_to(book), call_fn=_passthrough_call)
    assert resp["status"] == "failed" and "IdentityError" in resp["error"]
    assert resp["findings_json"] == "[]"


def test_missing_worksheet_is_a_failure():
    book = _FakeBook({})  # Worksheets("Live") -> KeyError
    resp = handle_audit(_request(sheet="Live"), attach_fn=_attach_to(book), call_fn=_passthrough_call)
    assert resp["status"] == "failed" and resp["findings_json"] == "[]" and resp["count"] == 0


def test_com_read_failure_is_a_failure():
    book = _FakeBook({"Live": _ExplodingWorksheet("Live")})
    resp = handle_audit(_request(sheet="Live"), attach_fn=_attach_to(book), call_fn=_passthrough_call)
    assert resp["status"] == "failed" and "COM blew up" in resp["error"]
    assert resp["findings_json"] == "[]"


# ======================= callback correlation (mirror of the VBA) =======================

def _registry(**over):
    rec = PendingRequest(id="p-1-1", token="tok", pid=4242, workbook="Book.xlsx",
                         full_name=r"C:\m\Book.xlsx")
    for k, v in over.items():
        setattr(rec, k, v)
    return {rec.id: rec}, rec


def _cb(**over):
    args = dict(id="p-1-1", token="tok", pid=4242, workbook="Book.xlsx", full_name=r"C:\m\Book.xlsx")
    args.update(over)
    return args


def test_callback_accepted_when_everything_matches():
    reg, _ = _registry()
    assert classify_callback(reg, "p-1-1", **_cb()) == "accepted"


def test_callback_rejected_unknown():
    reg, _ = _registry()
    assert classify_callback(reg, "p-1-1", **_cb(id="ghost")) == "rejected: unknown request"


def test_callback_rejected_not_pending():
    reg, rec = _registry(status="done")
    assert classify_callback(reg, "p-1-1", **_cb()) == "rejected: already done"


def test_callback_rejected_stale_when_superseded():
    reg, _ = _registry()
    # a newer request ("p-1-2") is now current; the callback for the old id is stale
    assert classify_callback(reg, "p-1-2", **_cb()) == "rejected: stale (superseded by a newer request)"


def test_callback_rejected_identity_mismatch():
    reg, _ = _registry()
    assert classify_callback(reg, "p-1-1", **_cb(pid=9999)) == "rejected: identity mismatch"
    assert classify_callback(reg, "p-1-1", **_cb(workbook="Other.xlsx")) == "rejected: identity mismatch"
    assert classify_callback(reg, "p-1-1", **_cb(full_name=r"C:\x\Other.xlsx")) == "rejected: identity mismatch"


def test_callback_rejected_token_mismatch():
    reg, _ = _registry()
    assert classify_callback(reg, "p-1-1", **_cb(token="stale-token")) == "rejected: correlation token mismatch"


# ======================= delivery wiring (fake Excel, mirrors the VBA front end) =======================

class _FakeExcel:
    """A fake Excel whose Run() routes to the same correlation logic the VBA implements."""

    def __init__(self, registry, current_id):
        self.registry, self.current_id = registry, current_id
        self.last_args = None

    def Run(self, callback, id, token, pid, workbook, full_name, status, error, count, findings_json):
        self.last_args = dict(callback=callback, id=id, token=token, pid=pid, workbook=workbook,
                              full_name=full_name, status=status, error=error, count=count,
                              findings_json=findings_json)
        verdict = classify_callback(self.registry, self.current_id, id=id, token=token, pid=pid,
                                    workbook=workbook, full_name=full_name)
        if verdict == "accepted":
            rec = self.registry[id]
            rec.status = "failed" if status != "ok" else "done"
        return verdict


def _response_from_audit():
    book = _FakeBook({"Live": _FakeComWorksheet("Live", 1, 2, _FORMULA, _VALUE)})
    return handle_audit(_request(), attach_fn=_attach_to(book), call_fn=_passthrough_call)


def test_deliver_passes_identity_token_and_json_and_is_accepted():
    resp = _response_from_audit()
    reg = {resp["id"]: PendingRequest(id=resp["id"], token=resp["token"], pid=resp["pid"],
                                      workbook=resp["workbook"], full_name=resp["full_name"])}
    excel = _FakeExcel(reg, current_id=resp["id"])
    verdict = deliver(resp, attach_fn=lambda r: (excel, None), call_fn=_passthrough_call)

    assert verdict == "accepted"
    assert excel.last_args["callback"] == audit_engine.CALLBACK
    assert excel.last_args["token"] == resp["token"]
    assert excel.last_args["status"] == "ok"
    assert json.loads(excel.last_args["findings_json"])[0]["cell"] == "D2"
    assert reg[resp["id"]].status == "done"


def test_callback_is_qualified_with_the_macro_workbook():
    # Live defect: an unqualified Application.Run name resolves only in the ACTIVE (audited)
    # workbook -> "Cannot run the macro". The callback must name the macro workbook.
    assert audit_engine.callback_name("Book2") == "'Book2'!PhaseBBridge.PhaseB_OnResult"
    assert audit_engine.callback_name("O'Brien.xlsm") == "'O''Brien.xlsm'!PhaseBBridge.PhaseB_OnResult"
    assert audit_engine.callback_name(None) == audit_engine.CALLBACK

    req = dict(_request(), macro_workbook="Book2")
    book = _FakeBook({"Live": _FakeComWorksheet("Live", 1, 2, _FORMULA, _VALUE)})
    resp = handle_audit(req, attach_fn=_attach_to(book), call_fn=_passthrough_call)
    reg = {resp["id"]: PendingRequest(id=resp["id"], token=resp["token"], pid=resp["pid"],
                                      workbook=resp["workbook"], full_name=resp["full_name"])}
    excel = _FakeExcel(reg, current_id=resp["id"])
    assert deliver(resp, attach_fn=lambda r: (excel, None), call_fn=_passthrough_call) == "accepted"
    assert excel.last_args["callback"] == "'Book2'!PhaseBBridge.PhaseB_OnResult"


def test_deliver_of_stale_result_is_rejected_by_front_end():
    resp = _response_from_audit()
    reg = {resp["id"]: PendingRequest(id=resp["id"], token=resp["token"], pid=resp["pid"],
                                      workbook=resp["workbook"], full_name=resp["full_name"])}
    excel = _FakeExcel(reg, current_id="a-newer-request")  # user already moved on
    verdict = deliver(resp, attach_fn=lambda r: (excel, None), call_fn=_passthrough_call)
    assert verdict.startswith("rejected: stale")
    assert reg[resp["id"]].status == "pending"             # stale result did not overwrite


def test_deliver_reports_failed_status_without_rendering_findings():
    resp = _response_from_audit()
    resp.update(status="failed", error="IdentityError: gone", findings_json="[]", count=0)
    reg = {resp["id"]: PendingRequest(id=resp["id"], token=resp["token"], pid=resp["pid"],
                                      workbook=resp["workbook"], full_name=resp["full_name"])}
    excel = _FakeExcel(reg, current_id=resp["id"])
    verdict = deliver(resp, attach_fn=lambda r: (excel, None), call_fn=_passthrough_call)
    assert verdict == "accepted"                           # delivered, but...
    assert excel.last_args["status"] == "failed"           # ...as a failure, not empty success
    assert excel.last_args["findings_json"] == "[]"
    assert reg[resp["id"]].status == "failed"


def test_deliver_attach_failure_is_reported_not_raised():
    resp = _response_from_audit()

    def boom(_r):
        raise audit_engine.ExcelGone("Excel closed")
    verdict = deliver(resp, attach_fn=boom, call_fn=_passthrough_call)
    assert verdict.startswith("not delivered:") and "ExcelGone" in verdict


# ======================= navigation logic =======================

def test_navigation_targets_original_workbook_from_a_results_row():
    # what the VBA double-click handler produces from a results row's hidden nav columns
    record = {"workbook": "Model.xlsx", "workbook_path": r"C:\m\Model.xlsx",
              "sheet": "P&L 'A'", "cell": "D3", "confidence": "high", "kind": "singleton",
              "conflict": False, "message": "..."}
    target = navigation_from_record(record)
    assert target["workbook"] == "Model.xlsx"              # the ORIGINAL audited workbook
    assert target["sheet"] == "P&L 'A'"
    assert target["cell"] == "D3"
    assert target["reference"] == "'P&L ''A'''!D3"         # apostrophes doubled, always quoted


def test_navigation_record_comes_straight_from_the_findings_json():
    resp = _response_from_audit()
    record = json.loads(resp["findings_json"])[0]
    target = navigation_from_record(record)
    assert (target["workbook"], target["sheet"], target["cell"]) == ("Book.xlsx", "Live", "D2")
