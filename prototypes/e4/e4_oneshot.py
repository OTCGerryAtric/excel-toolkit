"""E4 one-shot: one Python process per request (the E2 model), result delivered to VBA."""

import time

T0 = time.perf_counter()

import json  # noqa: E402
import sys  # noqa: E402
from pathlib import Path  # noqa: E402

import pythoncom  # noqa: E402

from e4_engine import deliver, handle  # noqa: E402


def main() -> None:
    request_path = Path(sys.argv[1])
    request = json.loads(request_path.read_text(encoding="mbcs"))
    pythoncom.CoInitialize()
    imports_ms = (time.perf_counter() - T0) * 1000
    response = handle(request)
    response["hwnd"] = request["hwnd"]
    response["verdict"] = deliver(response)
    response["timings"]["imports_ms"] = imports_ms
    response["timings"]["process_ms"] = (time.perf_counter() - T0) * 1000
    log = {k: v for k, v in response.items() if k != "findings"}
    (request_path.parent / f"oneshot_{request['id']}.json").write_text(json.dumps(log, default=str), encoding="utf-8")
    request_path.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
