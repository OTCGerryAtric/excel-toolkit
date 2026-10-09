"""E4 persistent worker: local named pipe in, VBA callback out. Read-only requests only.

  python e4_worker.py <state_dir> [--idle-seconds N]

- One worker per Windows user (the pipe name includes the user name; a second worker exits
  at once because the pipe already exists).
- Pipe rejects remote clients. No network, no server, no admin rights.
- Requests are handled one at a time in arrival order; each carries its own Excel identity
  and gets a fresh, short-lived COM connection.
- Exits after the idle timeout, or on {"op": "shutdown"}.
"""

import argparse
import json
import os
import queue
import threading
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32file
import win32pipe

from e4_engine import deliver, handle

PIPE_ACCESS_INBOUND = 0x00000001
FILE_FLAG_FIRST_PIPE_INSTANCE = 0x00080000
PIPE_REJECT_REMOTE_CLIENTS = 0x00000008
ALLOWED_OPS = {"classify", "read"}


def pipe_name() -> str:
    return rf"\\.\pipe\excel_toolkit_e4_{os.environ.get('USERNAME', 'user')}"


def create_pipe(first: bool):
    return win32pipe.CreateNamedPipe(
        pipe_name(),
        PIPE_ACCESS_INBOUND | (FILE_FLAG_FIRST_PIPE_INSTANCE if first else 0),
        win32pipe.PIPE_TYPE_BYTE | win32pipe.PIPE_READMODE_BYTE | win32pipe.PIPE_WAIT | PIPE_REJECT_REMOTE_CLIENTS,
        win32pipe.PIPE_UNLIMITED_INSTANCES, 0, 1 << 16, 0, None,
    )


def accept_loop(first_pipe, inbox: queue.Queue, log) -> None:
    pipe = first_pipe
    while True:
        win32pipe.ConnectNamedPipe(pipe, None)
        next_pipe = create_pipe(first=False)  # listen again before handling this one
        chunks = []
        try:
            while True:
                _, data = win32file.ReadFile(pipe, 1 << 16)
                chunks.append(data)
        except pywintypes.error:
            pass  # client closed its end
        win32pipe.DisconnectNamedPipe(pipe)
        win32file.CloseHandle(pipe)
        pipe = next_pipe
        text = b"".join(chunks).decode("mbcs").strip()
        if not text:
            continue
        try:
            inbox.put(json.loads(text))
        except json.JSONDecodeError as error:
            log({"event": "bad request", "error": str(error), "text": text[:200]})


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
        return  # another worker already owns the pipe

    (state / "worker.pid").write_text(str(os.getpid()), encoding="ascii")
    log_path = state / "worker_log.jsonl"

    def log(entry: dict) -> None:
        entry["at"] = time.strftime("%H:%M:%S")
        with log_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry, default=str) + "\n")

    started = time.perf_counter()
    inbox: queue.Queue = queue.Queue()
    threading.Thread(target=accept_loop, args=(first_pipe, inbox, log), daemon=True).start()
    log({"event": "started", "pid": os.getpid(), "startup_ms": (time.perf_counter() - started) * 1000})

    pythoncom.CoInitialize()
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
        handled = time.perf_counter()
        response = handle(request)
        response["hwnd"] = request["hwnd"]
        response["verdict"] = deliver(response)
        response["timings"]["worker_total_ms"] = (time.perf_counter() - handled) * 1000
        log({k: v for k, v in response.items() if k != "findings"})
    (state / "worker.pid").unlink(missing_ok=True)


if __name__ == "__main__":
    main()
