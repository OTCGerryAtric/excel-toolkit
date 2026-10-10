# Deployment & architecture decision (draft for review) — Excel Toolkit add-in

Date: 2026-10-10. Status: **approved as the working plan** (2026-10-10). Architecture and development direction
agreed; **deployment remains blocked** (Gate G0). No installation, certificate, Trusted Location, policy change or
deployment probe until security-cleared. No further architecture documents are needed — the next milestone is
working formula-analysis functionality with automated tests.
Context: E6's packaged installer was removed by the EDR (`docs/validation/phase0-e6.md`,
`docs/incidents/2026-10-10-edr-file-deletion.md`). This document compares front-end architectures and records
what IT/security must answer. The **zero paid-licence** constraint is mandatory — a pass/fail column, not an aspiration.

## Position (revised 2026-10-10 after review)

**VBA `.xlam` + Python is the preferred working baseline.** Excel-DNA + Python is a credible **alternative to
investigate**, not an equal implementation commitment at this stage. A full C# rewrite is a contingency only. The
**E6 failure blocks deployment, not all development**: the packaging approach was wrong; the application architecture
was not shown to be wrong. The most useful progress now is building the **framework-independent Python analysis
engine**, which is valuable under any front end.

## Guiding principle

**Preserve the independent Python analysis engine (E2/E4/E5)** regardless of front end. How Excel talks to it — VBA
or C# — is a secondary, reversible decision, and must not be allowed to consume the project. Excel-DNA additionally
allows *incremental* later migration of selected analysis features into C# without changing the front end again —
useful insurance if launching an external Python process proves problematic. We keep VBA because it is proven, not
merely because the project started there; we do not rewrite the engine without compelling evidence.

## What the evidence does and does not tell us

- **Arixcel Explorer coexists with the EDR on this machine/policy.** Registry evidence (read 2026-10-10): it
  registers **both** an `.xlam` under `HKCU\…\Excel\Options\OPEN` (→ `ArixcelExplorer.xlam`) **and** a VSTO
  add-in under `HKCU\Software\Microsoft\Office\Excel\Addins\ArixcelExplorer` (Manifest
  `…\ArixcelExplorer.vsto|vstolocal`, LoadBehavior 3). So our `.xlam`-via-`OPEN` matches only Arixcel's companion
  `.xlam`; Arixcel's main .NET component loads through the **VSTO Addins** mechanism, **not** `OPEN`. The earlier
  "same OPEN key" equivalence was imprecise and is corrected here. All Arixcel DLLs are Authenticode-signed ("Arixcel
  Limited"). This proves a professionally-deployed, signed add-in *can* be accepted here; it does **not** identify
  what triggered our E6 detection, nor that an unsigned `.xll` or a Python-backed add-in will be allowed.
- **Do not assume Excel-DNA avoids the EDR scrutiny.** `.xll` add-ins are a known malware vector; Microsoft began
  **blocking untrusted XLL add-ins from the internet by default in March 2023**. Signing and conventional install are
  necessary and indicated, **not proven sufficient**.

## Two independent trust problems (keep separate)

A solution must satisfy **both**.

### 1. Office add-in trust (will Excel load the front end?)

- **VBA `.xlam`** with Mark of the Web: a digital signature **does NOT** exempt it (Microsoft, since MS16-088/2016).
  Remedies: remove MOTW, or a **Trusted Location**.
- **`.xll`** (Excel-DNA): blocked from the internet by default since 2023. Microsoft's guidance is to **unblock the
  specific `.xll` first**, and a trusted-publisher certificate then applies to *future* add-ins — i.e. **signing
  alone does not exempt a MOTW-carrying `.xll`**; it must be unblocked (or delivered without MOTW / from a trusted
  location). This mirrors the `.xlam` limitation and is **not** a proven one-step deployment solution.
- **Trusted Locations** can load either reliably and are checked first, **but they bypass Office security checks for
  everything in that folder**, and on a managed device admin policy may enforce or forbid them. We will **not** adopt
  creating a new Trusted Location as the default install method without confirming it fits organisational policy.

### 2. Endpoint-protection (the EDR) acceptance (will the EDR allow it to run?)

Separate from #1, unresolved, the real E6 blocker. Applies to any option that keeps the local Python engine (A, B),
because launching a local Python process is the shared, unproven behaviour. Option C removes it.

## Tooling licences (verified; zero-cost column)

- **Excel-DNA** core — **Zlib licence** (confirmed, GitHub): permissive, free, commercial/internal use allowed.
- **.NET SDK** — free (MIT); builds C#/F# from the command line.
- **Excel-DNA build** — SDK-style project via `dotnet` CLI / NuGet (`ExcelDna.AddIn`); **no Visual Studio required**
  (confirm with a test build). Packs into a single `.xll`.
- **Visual Studio Community** — **not licensed** at this org size (>250 PCs/users or >$1M revenue) except
  OSS/education/etc. **VSTO** requires Visual Studio → needs **paid VS Professional**, **excluded on cost** (even
  though Arixcel itself is VSTO). The free .NET front-end route is **Excel-DNA**, not VSTO.
- **pywin32** (PSF), **openpyxl** (MIT), **python.org embeddable / uv standalone CPython** (PSF) — free.

## Correction: the Python runtime trust model

Earlier drafts described the embeddable Python as "Microsoft-built" with "signed `.py` scripts". Corrected:

- The **python.org embeddable distribution is CPython from the Python Software Foundation** (python.org build), **not
  a Microsoft build**. (A separately-packaged "Python from the Microsoft Store" exists, but that is not what we'd ship.)
- `python.exe`/`.pyd`/`.dll` are Authenticode-signed by their publisher, but **`.py` source files are not Windows
  Authenticode executables and cannot be code-signed like DLLs**. Integrity of our engine's `.py` files must be
  handled by **packaging + hashing / a signed manifest**, not by per-file Authenticode. The runtime trust model
  therefore needs explicit design and is an open question for the deployment workstream.

## The three architectures

| | **A. VBA `.xlam` + Python** | **B. Excel-DNA (`.xll`) + Python** | **C. Excel-DNA + C# engine** |
|---|---|---|---|
| Role | **Preferred working baseline** | Credible **alternative to investigate** (not committed) | **Contingency** |
| Free software | Yes | Yes (Zlib + MIT) | Yes |
| Existing VBA reuse (E1, E3) | **Excellent** | Partial — front end reimplemented in C# | Partial |
| Existing Python reuse (E2/E4/E5) | **Excellent** | **Excellent** | Requires porting to C# |
| External Python process | Yes | Yes | **No** (in-process .NET) |
| Office add-in trust (#1) | `.xlam` MOTW → Trusted Location (signing doesn't help) | `.xll` MOTW → unblock / Trusted Location (signing alone insufficient) | As B |
| the EDR acceptance (#2) | Python-launch + install — **unproven** | Python-launch + unsigned-`.xll` scrutiny — **unproven** | Lowest (Arixcel-like), still **unproven** for our tool |
| Main claimed benefit | Proven; preserves macros; fast (E1 ~1 ms) | Richer UI (custom task panes), .NET tooling — a **hypothesis**, not yet demonstrated for this toolkit | Single signed `.xll`, no Python runtime to trust |
| Assessment | **Keep / build on** | **Investigate where it offers measurable advantage** | **Last resort** |

### On Excel-DNA's advantage (qualified)

Excel-DNA *may* offer a richer, more maintainable front end — particularly custom task panes for navigating audit
findings. For the actual requirements (≈16 productivity macros, keyboard shortcuts, a ribbon/searchable command list,
formula auditing, navigable findings, 6–20 users), VBA already does the formatting tools in ~1 ms (E1) and Python
does the analysis (E2–E5). The concrete differentiator is UI sophistication. That is worth evaluating **when we know
we need it**, not a reason to rebuild a proven front end now. Switching front end also does **not** solve the E6
problem — the Python runtime and process-launch behaviour remain — so Excel-DNA is an **interface/maintainability**
decision, not a demonstrated security solution.

## Three workstreams (the central plan change)

E6 failing blocks **deployment**, not application development. Split the work so progress does not wait on the
deployment question.

1. **Python analysis engine (development priority).** Formula parsing, formula grouping, inconsistency detection,
   risk rules, audit findings, coverage statuses and test fixtures. Developed and tested **independently of Excel and
   without launching any add-in**, against offline workbook fixtures. Framework-independent: valid under VBA,
   Excel-DNA or any front end.
2. **Excel front end.** Keep the proven **VBA** implementation as the baseline. Evaluate **Excel-DNA** only where it
   offers a *measurable* advantage (especially navigating audit results). No commitment yet.
3. **Packaging & deployment.** **Gate G0 remains blocked.** Design a normal install/update strategy with **no
   execution-policy bypass, no EDR workaround, no unapproved trust change**. Resolve the Office-trust and the EDR
   questions here, separately, with IT/security.

Ordinary **offline C# compilation** (e.g. trying Excel-DNA locally) needs no certificate, Trusted Location or
security clearance. What needs clearance is **installing, auto-starting or launching the Python worker** on this
machine, and any Trusted Location / certificate / policy change.

**Scope of the offline-development authorisation.** Authorising workstream 1 means **ordinary offline Python
development and tests** (running the engine/parser against local fixtures from the command line). It is **not**
blanket approval to execute arbitrary code on the affected laptop: **Excel automation/COM, running packaged
executables, and any deployment experiment remain separately gated** and are not covered by it.

## Next development deliverable

A **standalone Python formula-inconsistency detector**:

- Pattern-based detection of a formula that breaks the pattern of its neighbours (e.g. in `=B2*(1+B3)`,
  `=C2*(1+C3)`, `=D2*(1+**D4**)`, `=E2*(1+E3)`, flag March's `D4` as *for review*, **not** asserted as an error).
- Built and tested **entirely in Python** against representative formula data and **offline workbook fixtures**.
- **Pattern analysis and dependency analysis are distinct, each with its own coverage/confidence.** *Pattern
  analysis* asks whether a formula follows the expected relative-reference pattern of its neighbours — this can be
  decided from the formula structure **without resolving every referenced cell**. *Dependency analysis* (the E5
  `resolved`/`partial`/`indeterminate` work) asks whether referenced cells/ranges can be determined. A formula can
  be structurally comparable to its neighbours *and* carry an unresolved (e.g. dynamic) dependency. **The full E5
  dependency parser is therefore NOT a prerequisite** for the inconsistency detector; the detector reports a
  pattern-coverage/confidence of its own and only consults dependency status where it genuinely needs it.
- **Automated regression tests** over the fixtures.
- **Precision is a first-class acceptance criterion, not just coverage.** An auditing tool that flags valid formulas
  becomes useless fast, so the fixtures must include **legitimate variations that must NOT be flagged** (e.g.
  deliberate first-column/edge differences, intentional constants, a row that legitimately breaks the pattern), and
  tests must measure **false-positive rate** alongside detection. Findings are **ranked by confidence** —
  high-confidence inconsistencies separated from weaker "for review" suggestions — never asserted as definite errors.
- The final Excel front end is **not** a prerequisite for this work.

## The minimal Excel-DNA test (feasibility only — not a deployment solution)

If/when Excel-DNA is investigated: a minimal `.xll` (ribbon button + read selection, **no Python, no installer**)
tests only whether an XLL front end **loads and functions** under this machine's policy. It would **not** prove the
complete application can be installed/updated/run alongside the EDR, nor that the Python worker is acceptable.
Treat it as an **Excel-DNA feasibility test**, gated on security clearance, never reusing the flagged E6 prototype.

## The precise question for IT / security

For threat **`<redacted-threat-id>`**, device **WORKSTATION**, 2026-10-10 16:20:47 BST:

1. **What triggered the the EDR detection** on the E6 install test?
2. **Is the incident still active** / responding to this project? (Gates repository recovery.)
3. **Approved process for distributing an internal Excel add-in with a local Python component** — delivery without
   MOTW or via an approved Trusted Location, and whether launching a bundled Python runtime is permitted.
4. **Are internal (company-CA) code-signing certificates available** at no cost (for a VBA `.xlam` and/or an `.xll`)?
5. On managed devices, can **Trusted Locations** be used (user- or IT-provisioned) or are they policy-locked, and is
   the **"block macros/XLL from the internet"** policy enforced?

## Decision requested at review

1. **Retain VBA `.xlam` + Python as the preferred working baseline.**
2. **Excel-DNA + Python** remains an alternative to investigate, **not** an equal implementation commitment.
3. **Preserve the Python analysis engine**; **no C# rewrite** (Option C contingency only).
4. Adopt the **three-workstream** split; **G0 blocked for deployment** does **not** block engine development after
   the dev environment is safely recovered.
5. **Next deliverable:** the standalone Python formula-inconsistency detector above.
6. The final front-end decision is **not** a prerequisite for that development.
7. Keep all recovery/installation restrictions: no file restore, installer run, Excel-trust change or deployment test
   without separate approval.

## Sources

- [Macros from the internet are blocked by default in Office — Microsoft Learn](https://learn.microsoft.com/en-us/deployoffice/security/internet-macros-blocked)
- [Excel is blocking untrusted XLL add-ins by default — Microsoft Support](https://support.microsoft.com/en-US/Excel/excel-is-blocking-untrusted-xll-add-ins-by-default) (unblock the specific `.xll`; then trusted publisher applies to future add-ins)
- [Excel-DNA — GitHub (Zlib licence)](https://github.com/Excel-DNA/ExcelDna) and [documentation](https://exceldna.readthedocs.io)
- [Visual Studio Community licence — Microsoft Q&A on the enterprise restriction](https://learn.microsoft.com/en-us/answers/questions/2101450/microsoft-visual-studio-community-2022-license-que)
- [Python Windows embeddable package — Python docs](https://docs.python.org/3/using/windows.html) (CPython/PSF build; `.py` are not Authenticode-signable)
