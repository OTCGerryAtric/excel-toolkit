"""Phase B persistent worker (OFFLINE-PREPARED — start MANUALLY; NEVER launched by Excel).

Reuses E4's named-pipe channel and worker lifecycle VERBATIM
(``prototypes/e4/e4_worker.py`` -> ``pipe_name`` / ``create_pipe`` / ``accept_loop``): one
worker per Windows user, a local INBOUND-only pipe that rejects remote clients, no network,
no server, no admin rights. The ONLY differences from E4 are the op it accepts (``audit``)
and the engine it runs (the formula-inconsistency detector, via ``audit_engine``).

    <scratch-venv>\\Scripts\\python.exe audit_worker.py <state_dir> [--idle-seconds N]

Start it BY HAND from the development environment. The VBA front end only WRITES to the pipe;
it never starts this process (every Shell / one-shot / auto-start path has been removed from
the bridge). Exits on ``{"op": "shutdown"}`` or after the idle timeout.

STOP IMMEDIATELY (no retry) if endpoint protection intervenes, a process is unexpectedly
terminated, files disappear, or access-denied errors recur.
"""

from __future__ import annotations

import argparse
import json
import os
import queue
import sys
import threading
import time
from pathlib import Path

import pythoncom
import pywintypes

ROOT = Path(__file__).resolve().parents[2]
for _p in (ROOT / "prototypes" / "e4", ROOT / "prototypes" / "e6" / "engine"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from e4_worker import accept_loop, create_pipe, pipe_name  # reuse E4 pipe plumbing verbatim  # noqa: E402

from audit_engine import ALLOWED_OPS, deliver, handle_audit  # noqa: E402


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("state_dir")
    parser.add_argument("--idle-seconds", type=float, default=1800)
    args = parser.parse_args()
    state = Path(args.state_dir)
    state.mkdir(parents=True, exist_ok=True)

    try:
        first_pipe = create_pipe(first=True)
    except pywintypes.error:
        return  # another worker already owns the pipe (one per user)

    (state / "worker.pid").write_text(str(os.getpid()), encoding="ascii")
    log_path = state / "worker_log.jsonl"

    def log(entry: dict) -> None:
        entry["at"] = time.strftime("%H:%M:%S")
        with log_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry, default=str) + "\n")

    inbox: queue.Queue = queue.Queue()
    threading.Thread(target=accept_loop, args=(first_pipe, inbox, log), daemon=True).start()
    log({"event": "started", "pid": os.getpid(), "pipe": pipe_name()})

    pythoncom.CoInitialize()
    try:
        while True:
            try:
                request = inbox.get(timeout=args.idle_seconds)
            except queue.Empty:
                log({"event": "idle timeout"})
                break
            if request.get("op") == "shutdown":
                log({"event": "shutdown requested"})
                break
            if request.get("op") not in ALLOWED_OPS:
                log({"event": "refused", "id": request.get("id"), "op": request.get("op")})
                continue
            response = handle_audit(request)
            response["verdict"] = deliver(response)
            log({**{k: v for k, v in response.items() if k != "findings_json"},
                 "findings_bytes": len(response.get("findings_json", ""))})
    finally:
        pythoncom.CoUninitialize()
        (state / "worker.pid").unlink(missing_ok=True)


if __name__ == "__main__":  # pragma: no cover - started manually in the authorised test only
    main()
