# Excel Toolkit add-in (development)

The VBA `.xlam` front end: ribbon, commands and one-level undo. Milestone M2 (see
`docs/product-roadmap.md` §D and `docs/validation/m2-convert-to-values.md`).

```
addin/
  customUI/customUI14.xml   ribbon (labels/tooltips/enabled state via callbacks to the registry)
  vba/TkMain.bas            command registry, ribbon callbacks, messages, timer
  vba/TkConvert.bas         Convert to Values: validate -> capture -> apply -> verify -> register undo
  vba/TkCellSnapshot.cls    the converted formula cells: exact before/after state, restore, conflict checks
  vba/TkUndo.bas            the single undo record; Ctrl+Z (OnUndo) and Undo Toolkit Action
  build_addin.py            builds build/ExcelToolkit.xlam (separate hidden Excel) and injects the ribbon
  live/m2_live_check.py     scripted live check in a disposable Excel process (authorised runs only)
```

Audit Formulas reuses Phase B (`prototypes/integration/PhaseBBridge.bas` and `PhaseBResultsHook.cls`,
imported into the same `.xlam`); the worker is still started by hand.

**Nothing is installed.** Open `build/ExcelToolkit.xlam` with File → Open for a session; close Excel to
unload it. VBA sources must stay ASCII with CRLF line endings (`.gitattributes`). Offline tests:
`tests/addin/`.
