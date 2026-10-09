# Phase 0 — E4 results: one-shot vs persistent Python worker

Date: 2026-10-09. Machine: owner's PC (Excel M365 x64 16.0.20326.20142, Python 3.12.10, pywin32 312).

## Question

Is a persistent Python worker worth its extra complexity compared with starting Python for each request (~0.9 s in E2)?

## What was built (`prototypes/e4/`)

Both modes share one engine (`e4_engine.py`), so the comparison is like for like.

- **Request (VBA, `E4Bridge.bas`):** `E4_Request(mode, op, sheet, address, workbook)` returns a request id at once. The request carries the id plus the explicit identity (window, process id, workbook name, full name).
  - **One-shot:** VBA writes the request to a file and `Shell`s `e4_oneshot.py`.
  - **Worker:** VBA writes the request to a local named pipe. If no worker is listening, VBA starts `e4_worker.py` and retries for up to 10 s.
- **Work (Python):** attach with the E2/E3 explicit identity and bounded retries, read, classify, then release every COM reference (short-lived connections).
- **Result back to VBA:** Python attaches to the requesting Excel and calls `E4_OnResult` with request id, process id, workbook name, full name, status, error, a probe value, a count, and the findings as a **native 2D array** (sheet, address, kind, detail). VBA rejects unknown request ids, already-settled requests and any identity mismatch.
- **Watchdog (VBA `OnTime`, every 1 s while requests are pending):** marks a request failed if the worker process has died or after 60 s.
- **Worker:** one per Windows user (pipe name includes the user name; a second copy exits at once); pipe rejects remote clients; handles requests one at a time; read-only ops only (`classify`, `read`); idle timeout (30 min default); `shutdown` request.
- **Test runner (`run_e4.py`):** two new Excel processes — X (target + a second workbook) and Y (a decoy with the **same file name**) — each with the add-in. Timings are measured **inside Excel**, from request to callback.

## Results — 28/28 checks passed (final run)

### Latency, request → result in VBA (ms)

| | First request | Warm median | Warm p95 |
|---|---|---|---|
| One-shot (100 findings) | 820 | 817 | 834 |
| Worker (100 findings) | 822 (includes starting the worker) | **47** | **52** |

Worker breakdown when warm: attach ~13 ms, read and classify ~16 ms, deliver to VBA ~17 ms. Worker start-up itself (pipe ready) was under 1 ms after Python had loaded; the first request's ~820 ms is Python start-up and imports, the same cost as one-shot.

**Observed slowdown:** later in the same session, every COM call into that Excel (attach, read and deliver alike) became 3–4 times slower, and stayed slower: worker requests settled at ~150–200 ms. It began right after a batch of one-shot requests, including 10,000-finding results and 100,000-cell reads. The cause was not identified. It affects Excel's responsiveness, not the worker logic, but it means real-world warm latency may be nearer 150–200 ms than 50 ms. Worth re-measuring in E6 on a colleague's PC.

### Size scaling (median ms, 3 runs each; worker figures are from the slower period above)

| Operation | One-shot | Worker |
|---|---|---|
| Classify 100 cells → 100 findings to VBA | 821 | 173 |
| Classify 1,000 → 1,000 findings | 800 | 194 |
| Classify 10,000 → 10,000 findings | 849 | 389 |
| Read 1 cell | 821 | 176 |
| Read 1,000 cells | 801 | 180 |
| Read 10,000 cells | 818 | 216 |
| Read 100,000 cells | 1,001 | 690 |

10,000 findings went back to VBA as one 2D array in about 150 ms. Reading 100,000 cells (formulas + values) takes ~560 ms either way.

### Correctness and isolation

| Check | Result |
|---|---|
| Finding kinds: formula, formula with a typed-in number, number, text, formula returning an error, boolean | Pass |
| **69 interleaved requests** (60 worker, 9 one-shot) across 2 processes and 3 workbooks, two with the same name: every result from the right workbook and delivered to the right request | Pass |
| VBA rejects a result for an unknown request, and one carrying another process's identity | Pass |
| Audit requests left the target workbook unchanged (full snapshot) | Pass |

**Finding — don't rely on `ActiveWorkbook`.** In an earlier run, one request meant for the second workbook was issued for the target, because `ActiveWorkbook` had not yet caught up with a programmatic `Workbook.Activate`. Request, identity check and result were consistent with each other, so nothing was mixed between requests, but the request named the wrong workbook. Requests now name their workbook explicitly. In production, ribbon callbacks should take the workbook from the button's own context (the window it was clicked in), never `ActiveWorkbook`.

### Failure handling

| Check | Result |
|---|---|
| Worker killed mid-request (10,000-finding request) → VBA watchdog marks it failed: "the Python worker stopped before answering" | Pass (within the next 1 s tick) |
| Next request starts a new worker automatically and succeeds | Pass |
| Excel closed normally after the worker served it → the Excel process exits (worker holds no stale COM references) | Pass |
| Request for an Excel that has closed → clear `IdentityError`, worker keeps running | Pass |
| Excel killed while the worker was reading it → `ExcelGone`, result not delivered, worker survives and serves the next request | Pass |
| Worker still serves the other Excel afterwards | Pass |

### Idle cost of the worker

Working set 77 MB, private memory 60 MB, **0.000 s CPU over 10 s idle** (it blocks on the pipe and queue).

### Write timings (scratch workbook only, not via the worker)

| Cells | Number format: Python COM | Number format: VBA in-process | Write values: Python COM |
|---|---|---|---|
| 1 | 23 ms | 0.6 ms | 4 ms |
| 1,000 | 8 ms | 1.5 ms | 24 ms |
| 10,000 | 29 ms | 5.9 ms | 23 ms |
| 100,000 | 24 ms | 12.4 ms | 81 ms |

A whole-range format change is one call either way; both are fast. VBA stays the right place for instant commands (E1), and Python bulk writes are fine where safeguards (E3 journal) are needed.

## Comparison

| | One-shot | Persistent worker |
|---|---|---|
| Warm latency | ~0.8 s every time | ~50 ms (fresh) to ~200 ms (slow period); first request ~0.8 s |
| Large results | Same delivery path | Same delivery path |
| Isolation / identity | Per process, nothing shared | Shared process; identity checked per request; 69/69 correct |
| Failure modes | A failed run only affects its own request | Worker crash affects in-flight requests; watchdog reports it; auto-restart works |
| Lifecycle | None | Start on demand, single instance, idle timeout, shutdown, pid file |
| Stale COM risk | None | Avoided by short-lived connections (proven: Excel exits normally) |
| Idle cost | None | ~60–77 MB RAM, no CPU, up to the idle timeout |
| Extra code | — | ~120 lines Python (pipe server, queue) + ~60 lines VBA (pipe write, start, watchdog) |
| Deployment | Bundled Python | Same bundled Python; no service, no admin rights, no network, no firewall prompt |
| Security | Request file in a per-user folder | Local pipe, remote clients rejected; production should restrict the pipe to the current user (explicit security descriptor) |

## Recommendation: hybrid, worker first

- **Use the persistent worker for interactive audit commands.** Results arrive in about 50–200 ms instead of ~0.8 s, which makes an audit panel feel responsive. It passed every isolation and failure test, holds no Excel references between requests, and costs nothing while idle.
- **Start it on demand and let it exit after an idle period.** No Windows service, no start-up entry.
- **Keep one-shot as the fallback and for batch work** (full-workbook reports, command-line use, tests). It's the same engine, so it costs almost nothing to keep, and it covers the case where the worker can't start.
- **Hotkey commands stay in VBA** (E1), and bulk edits keep the E3 journal. The worker stays **read-only**.

## Gaps and limitations

- Excel busy at delivery time (e.g. the user is editing a cell when the result comes back) wasn't tested: the callback would retry for 5 s and then give up. Production needs a pickup path (result kept in a file, VBA collects it when idle).
- The session slowdown above is unexplained.
- Watchdog's process check uses the pid file only (no start-time check against pid reuse).
- Pipe uses the default security descriptor plus remote rejection; a user-only descriptor is a production task.
- Single machine, one session; idle measured over 10 s only.
- During the runs, an Excel the owner started (with PERSONAL.XLSB) had a disposable decoy copy from the first E4 run open; it was not touched. Each run now uses its own work folder so a leftover file cannot block a run.

Stopping before E5 as agreed.
