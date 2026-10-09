"""Unit checks of the bounded-retry policy in excel_link.call, with a simulated busy Excel.

Excel itself could not be made to reject calls automatically in E3 (dialogs did not
block COM), so the policy is exercised here against fakes that raise the real HRESULTs.
"""

import sys
import time
from pathlib import Path

import pywintypes

sys.path.insert(0, str(Path(__file__).resolve().parent))
from excel_link import RPC_E_CALL_REJECTED, ExcelBusy, ExcelGone, call  # noqa: E402

RPC_S_SERVER_UNAVAILABLE = -2147023174


def rejected() -> pywintypes.com_error:
    return pywintypes.com_error(RPC_E_CALL_REJECTED, "Call was rejected by callee.", None, None)


class FakeOle:
    def __init__(self, hresult: int | None):
        self.hresult = hresult

    def GetIDsOfNames(self, _name):
        if self.hresult is not None:
            raise pywintypes.com_error(self.hresult, "probe", None, None)
        return 1


class FakeApp:
    def __init__(self, probe_hresult: int | None = RPC_E_CALL_REJECTED):
        self._oleobj_ = FakeOle(probe_hresult)


def busy_for(seconds: float, error_factory=rejected):
    until = time.monotonic() + seconds

    def operation():
        if time.monotonic() < until:
            raise error_factory()
        return "ok"
    return operation


def main() -> None:
    results = []

    started = time.monotonic()
    try:
        call(FakeApp(), busy_for(60), timeout=2)
        results.append(("Permanently busy: gives up after the bounded wait", False, "returned"))
    except ExcelBusy as error:
        elapsed = time.monotonic() - started
        results.append(("Permanently busy: gives up after the bounded wait", 2.0 <= elapsed < 2.5, f"{elapsed:.2f} s: {error}"))

    started = time.monotonic()
    value = call(FakeApp(), busy_for(1), timeout=3)
    elapsed = time.monotonic() - started
    results.append(("Transient busy: succeeds once Excel is free", value == "ok" and 1.0 <= elapsed < 1.5, f"{elapsed:.2f} s"))

    def attribute_error_while_busy():
        raise AttributeError("Excel.Application.Workbooks")
    started = time.monotonic()
    try:
        call(FakeApp(RPC_E_CALL_REJECTED), attribute_error_while_busy, timeout=1)
        results.append(("Busy reported as AttributeError is recognised (probe) and bounded", False, "returned"))
    except ExcelBusy:
        elapsed = time.monotonic() - started
        results.append(("Busy reported as AttributeError is recognised (probe) and bounded", 1.0 <= elapsed < 1.5, f"{elapsed:.2f} s"))

    try:
        call(FakeApp(None), attribute_error_while_busy, timeout=1)
        results.append(("Genuine missing attribute is not retried", False, "returned"))
    except AttributeError:
        results.append(("Genuine missing attribute is not retried", True, "AttributeError raised at once"))

    started = time.monotonic()
    try:
        call(FakeApp(), busy_for(60, lambda: pywintypes.com_error(RPC_S_SERVER_UNAVAILABLE, "RPC server unavailable", None, None)), timeout=5)
        results.append(("Excel gone: fails at once, no retries", False, "returned"))
    except ExcelGone:
        results.append(("Excel gone: fails at once, no retries", time.monotonic() - started < 0.2, f"{time.monotonic() - started:.2f} s"))

    for name, ok, detail in results:
        print(f"[{'PASS' if ok else 'FAIL'}] {name}: {detail}")
    print(f"{sum(ok for _, ok, _ in results)}/{len(results)} passed")


if __name__ == "__main__":
    main()
