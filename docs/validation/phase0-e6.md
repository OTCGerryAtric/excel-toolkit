# Phase 0 — E6 results: packaging, install, and a security incident

Date: 2026-10-10. Machine: owner's PC (WORKSTATION), Excel M365 x64 16.0.20326.20142, Python 3.12.10,
the endpoint EDR (vendor/version redacted) active (Microsoft Defender passive).

## Question

Can the add-in and its Python engine be installed on a Windows PC with no manual Python/dependency setup, no
admin rights, and no global macro-security change — and does it coexist with the machine's security software?

## What was built and tested (part 1, committed cf67e37)

A packaged prototype: PyInstaller **one-dir** engine (`ExcelToolkitEngine.exe` + `_internal`), the VBA `.xlam`,
a per-user PowerShell installer, a signed-manifest integrity check, and a licence register, zipped as
`ExcelToolkit-<version>.zip`. Install writes only to `%LOCALAPPDATA%\ExcelToolkit` and the HKCU Excel `OPEN` key.

Two automated runs passed on this machine **before** the security response:

- **Runtime** (`run_e6_runtime.py`, `results/e6_runtime_0.1.0.json`): 11/11. Worker warm ~45 ms on a fresh Excel,
  one-shot fallback ~0.9 s, pipe restricted to the current user, malformed/oversized/injection requests refused,
  idle worker self-exits. **The E4 slowdown was explained**: Windows power-throttling of the hidden worker; the
  worker now opts out, and latency stayed ~40–45 ms across heavy workloads.
- **Install lifecycle** (`run_e6_install.py`, `results/e6_install.json`): 19/19. Standard-user install, tamper and
  Excel-running refusals, upgrade 0.1.0→0.2.0, rollback, uninstall to exact baseline, PERSONAL.XLSB unchanged,
  Mark of the Web removed from installed files, other add-ins untouched.

These results are retained as evidence. **They do not clear E6**, because the method by which they were achieved
is the method the endpoint protection rejected.

## Security incident (the decisive E6 finding)

During the install test, **the EDR** detected the process tree as `windows.penetration`, killed 178 processes
(crashing the session), and quarantined every executable and script file it had created — across the project and
`.venv`, plus `claude.exe`. Full record and evidence: `docs/incidents/2026-10-10-edr-file-deletion.md`.

**Confirmed**: the EDR did it; one threat id; deletion was by file type; git intact; our scripts and Claude
did not delete anything. **Not confirmed locally**: which specific behaviour triggered the rule (console only).
Leading hypothesis: PowerShell execution-policy bypass + Mark-of-the-Web stripping + Excel auto-start of an
unsigned, freshly-extracted PyInstaller exe — an Office-persistence shape.

The install route therefore **fails the E6 bar on this machine, and colleague machines are expected to run the
same policy.** The passing functional tests do not override this.

## Reference: why Arixcel Explorer coexists and ours did not

Arixcel Explorer is an add-in already installed on this machine that loads through the **same** HKCU `OPEN`
registry key ours used, with no the EDR issue. Inspected read-only (`C:\Users\…\AppData\Local\Arixcel Explorer`):

| Aspect | Arixcel (accepted) | Our E6 prototype (flagged) |
|---|---|---|
| Binaries | Every DLL **Authenticode-signed** — `Valid`, "Arixcel Limited" (own) + Microsoft/.NET (deps) | Unsigned PyInstaller exe + DLLs, written seconds before use |
| Technology | **VSTO** (`.vsto`, `.dll.manifest`, Office Tools runtime) — standard Office add-in | Python packaged by PyInstaller self-extracting bootloader |
| On disk | Files sit signed; `.xlam` carries **no** Mark of the Web | MOTW stripped with `Unblock-File` at install |
| Install behaviour | Conventional install; no bypass, no process-killing | `powershell -ExecutionPolicy Bypass`; harness force-kills Excel/engine |

**Caveat — do not over-read this.** The comparison shows a *plausible, consistent* contrast; it does **not**
prove that signing or VSTO alone guarantees the EDR acceptance. the EDR weighs reputation, prevalence,
publisher, behaviour and local policy together, none of which we can read from this machine. Arixcel is also a
widely-deployed vendor product with established reputation that our internal tool will never have. Treat signing
and conventional install as **necessary and strongly indicated**, not as a guarantee. Final acceptance can only be
confirmed by testing a redesigned build, ideally with security/IT visibility.

## Outcome

E6 is **not passed**. Functional packaging/runtime works; the deployment method is rejected. The redesign and the
open signing/deployment decision are in `docs/architecture/deployment-decision.md`. Approved as a **design and
investigation exercise only** (2026-10-10) — no Trusted Location, certificate, policy change, add-in registration
or installer execution until review. The flagged prototype must not be re-run.

## Limitations

- One machine, owner's PC. No colleague machine tested (the original E6 goal).
- Trigger inferred, not confirmed; no access here to the the EDR console.
- "Passing" runtime/install numbers predate and are independent of the security verdict.
