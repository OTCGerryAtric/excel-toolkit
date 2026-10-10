# Incident: the EDR removed project files during the E6 install test (2026-10-10)

Severity: high (development files and `claude.exe` deleted; no data loss — all recoverable).
Status: contained; blast radius provisionally complete; **source recovered 2026-10-11, full dependency
environment restoration deferred** (pywin32 cache lock). No new detection observed during recovery.
Full evidence (logs, snapshot, exports): `C:\Users\GerryPidgeon\IncidentEvidence\2026-10-10-ExcelAddIn\`.

## Summary

While the crashed session was running the E6 install lifecycle test (`run_e6_install.py`), the endpoint
protection on this company laptop — **the endpoint EDR (vendor/version redacted)** — raised a behavioural detection against
the process tree, killed it, and quarantined the executable and script files the tree had created or run. This
is what crashed the previous Claude session and what removed the files. No project script, Claude command, sync
client or other agent deleted them.

Microsoft Defender is passive/off on this machine; the EDR is the active product.

## 1. Affected and unaffected components (confirmed)

**Affected (removed/quarantined).** Deletion was by **file type within the flagged process tree's working area**,
not by location or age. From the surviving tree and `damaged_tree_listing.csv`:

- In the project: every `.py` (27 tracked + engine sources), `.ps1` (2), `.cmd` (2), and every
  `.exe`/`.dll`/`.pyd`/`.pyc` — including `.venv\Scripts\python.exe` and the PyInstaller-built engine in `build/`
  and `work/`.
- 10 `.py` files remain under `.venv` but are access-denied to every process (consistent with an unfinished quarantine).
- Outside the project: `claude.exe` in `%USERPROFILE%\.local\bin` (folder modified 16:24:23). **Resolved on its own:**
  Claude Code auto-updated to 2.1.296 at 16:48, which recreated the launcher; `claude` is on PATH and a new session
  starts normally. (The CLI *is* `claude.exe`; the claude.ai website is browser-based and was never affected.)

**Unaffected (confirmed intact).**

- Kept in the project: `.json`, `.md`, `.xlsx`, `.xlam`, `.bas`, `.zip` (packages still contain the engine),
  `.txt`, `.toml`, `.lock`, and **all git metadata** (`git fsck` clean; commit `cf67e37` has all 31 deleted tracked files).
- Shared toolchains (full scan, 2026-10-10): base Python 3.12 (`%LOCALAPPDATA%\Programs\Python`, 13,541),
  PowerShell 7 (840), `uv.exe`/`uvw`/`uvx`, sibling projects `BookScraping` (1,396) and `Scrapling` (89),
  `%TEMP%\claude` (6,253). The other (idle) Claude session was not killed.
- **uv cache — largely intact but NOT wholly untouched (corrected 2026-10-11).** The cache holds ~424,556
  exec/script files overall, but at least one extracted wheel archive is quarantine-locked: **all 43 `.py` files of
  the cached `pywin32` archive are access-denied** (same class the EDR locked in `.venv`). This surfaced during
  recovery when `uv sync` failed copying `pywin32`'s `adodbapi\apibase.py`. The earlier "cache untouched" wording
  (a file-count check) was wrong; specific packages in the cache are affected.

Blast radius is **confined to the project directory, `claude.exe`, and specific locked wheel archives in the uv
cache (notably `pywin32`)** — provisionally complete; no further filesystem investigation unless new evidence appears.

## 2. Evidence that the EDR was responsible (confirmed)

- **One the EDR threat** (TrueContextID `<redacted-threat-id>`) accounts for the whole episode. From
  `the EDR/Operational` (exported to the evidence folder):
  - 16:20:47 — event 31, detection, classification **`windows.penetration`**, triggering file logged only as
    "Unknown file" (the EDR does not record the specific file locally).
  - 16:22:19–16:23:24 — event 32, **178 × `Kill` Success**, same threat id.
  - 16:24:26–16:24:43 — event 25 ×5, could not open `C:\ProgramData\<EDR>\UnexcludedQuarantine\*.MAL`
    ("being used by another process").
  - 16:25:30 — event 32, **`Quarantine` Success**, same threat id.
- **The crash**: the interactive `pwsh.exe` hosting the crashed session died at 16:22:47 (Application log: .NET
  Runtime 1026 / Application Error 1000 / WER `AppCrash_pwsh.exe_…12aa2118`) — a lost console, consistent with its
  parent being killed.
- **Timing excludes our own code**: the crashed session's last command was 16:21:16; the removals ran
  16:23:37–16:25:24. All 206 tool calls of that session (and subagents) and every recovery-session call were
  searched for deletion commands — none target the repo or `.venv` broadly.
- **Ruled out by evidence**: Windows Defender (passive/off, no detections), Sysmon (not installed), scheduled
  cleanup, sync clients (project not in a synced folder; timing wrong), the other Claude session (idle).

## 3. Suspected triggering behaviour (hypothesis — not verified)

The detection fired during the first-run / worker-kill / upgrade phase. The install combines behaviours an EDR
associates with intrusion; the local logs do **not** say which rule fired. Ranked by likelihood:

1. **(High)** The install pattern: `powershell.exe -ExecutionPolicy Bypass`, stripping Mark of the Web
   (`Unblock-File`) from just-"downloaded" files, then writing an Excel auto-start entry
   (`HKCU\…\Excel\Options\OPEN1`) pointing at an **unsigned** PyInstaller `.exe` extracted seconds earlier — a
   classic Office-persistence shape.
2. **(Medium)** The test harness force-terminating Excel and the engine, and attaching to Excel cross-process via
   `WM_GETOBJECT` / `ObjectFromLresult`.
3. **(Low)** The unsigned PyInstaller one-dir bootloader on its own.

## 4. Exact rule and detection status (unverified — the EDR console required)

Not determinable from this machine:

- The **exact rule / signature** that fired.
- Whether the response was **automatic or analyst-initiated**.
- Whether the detection is **still open / actively responding** to this project — this governs whether a restore
  or re-run would re-trigger (possibly wider), and **is the gate on recovery** (see below).
- Whether any of the 16:23:22–24 kills touched the recovery session (it started 16:23:00; its commands kept working).

Console reference: threat `<redacted-threat-id>`, device WORKSTATION, 2026-10-10 16:20:47 BST.

## Current state (verified 2026-10-10, read-only)

- No `ExcelToolkitEngine`, worker, installer or E6 script running. `%LOCALAPPDATA%\ExcelToolkit` absent (uninstall completed).
- Excel registry back to baseline (Arixcel `OPEN` only; our `OPEN1` gone). No toolkit named pipes.
- One user-opened `EXCEL.EXE` (started 16:51) unrelated to the test.

## Recovery (authorised and executed 2026-10-11; partial)

Authorised after the overseer accepted that no new the EDR intervention was evident. Outcome:

1. Evidence preserved (`C:\Users\GerryPidgeon\IncidentEvidence\2026-10-10-ExcelAddIn\`) — untouched. **PASS**
2. Restored the 31 committed files from `cf67e37` via `git restore` (no `reset`/`clean`); content verified identical
   to the commit (hash-matched; "modified" flags were line-ending-only). **PASS**
3. Reapplied the 4 saved uncommitted fixes atomically; resulting diff matches the original pre-incident change
   exactly; `py_compile` passes. No installer/packaging/deployment script run. **PASS**
4. **Full `.venv` rebuild DEFERRED / INCOMPLETE.** `uv sync` failed: `pywin32` cannot install because its cached
   wheel archive is quarantine-locked (see uv-cache note above). This is residual incident damage, **not a new
   detection**. Instead, an **isolated minimal environment** (openpyxl + pytest only, `--no-cache`, created outside
   the project under the session scratchpad) was built for offline engine development. `pyproject.toml`/`uv.lock`
   left unchanged; the damaged `.venv` not reused or repaired; the quarantined `pywin32` cache not modified.
5. **No new the EDR detection observed during recovery** (only benign `id=143` performance-counter reads; no
   quarantine/kill events; no Defender detections; no project files disappeared — 89 tracked files intact).

`claude.exe` had already self-restored via the 16:48 auto-update. Outstanding: full dependency restoration
(blocked on the `pywin32` cache lock — likely cleared by a reboot or by IT releasing the quarantine); this is
needed only for Excel-integration work, not for the offline analysis engine. See
`docs/architecture/deployment-decision.md`.

## Governance

- **Gate G0 remains blocked** pending an acceptable deployment design and security validation.
- Recovery and deployment are **separate workstreams**.
- No file restore, installer execution, security-setting change, or move to Phase 1 until the respective gate is cleared.
