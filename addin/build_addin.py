"""Build the development add-in ``addin/build/ExcelToolkit.xlam`` (M2).

  <scratch venv>/python addin/build_addin.py            # build (needs Excel - authorised use only)
  <scratch venv>/python addin/build_addin.py --check X  # only verify an existing .xlam package

Steps:
  1. CRLF copies of the VBA sources (an LF-only .cls imports as a standard module).
  2. A SEPARATE, hidden Excel process (DispatchEx: no XLSTART, no PERSONAL.XLSB, no add-ins)
     imports the modules into a new workbook and saves it as .xlam. Dev machine only: importing
     VBA needs "Trust access to the VBA project object model" (already on; not changed here).
     The process is terminated by pid if it has not exited 15 s after Quit.
  3. With Excel closed, customUI14.xml is added to the .xlam package (a zip edit) and the
     package is verified.

Nothing is installed: no Add-ins dialog, registry, Trusted Location or trust-setting change.
The add-in is opened by hand (File > Open) for each test session.
"""

from __future__ import annotations

import argparse
import shutil
import sys
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
HERE = Path(__file__).resolve().parent
BUILD = HERE / "build"
TARGET = BUILD / "ExcelToolkit.xlam"
CUSTOM_UI = HERE / "customUI" / "customUI14.xml"

VBA_SOURCES = [
    HERE / "vba" / "TkMain.bas",
    HERE / "vba" / "TkConvert.bas",
    HERE / "vba" / "TkUndo.bas",
    HERE / "vba" / "TkCellSnapshot.cls",
    HERE / "vba" / "TkFormat.bas",
    HERE / "vba" / "TkFormatSnapshot.cls",
    HERE / "vba" / "TkShortcuts.bas",
    HERE / "vba" / "TkStyle.bas",
    HERE / "vba" / "TkStyleSnapshot.cls",
    HERE / "vba" / "TkFormulas.bas",
    HERE / "vba" / "TkFormulaSnapshot.cls",
    ROOT / "prototypes" / "integration" / "PhaseBBridge.bas",
    ROOT / "prototypes" / "integration" / "PhaseBResultsHook.cls",
]
# VBComponent.Type: 1 = standard module, 2 = class module.
EXPECTED_TYPES = {"TkMain": 1, "TkConvert": 1, "TkUndo": 1, "TkCellSnapshot": 2, "TkFormat": 1,
                  "TkFormatSnapshot": 2, "TkShortcuts": 1, "TkStyle": 1,
                  "TkStyleSnapshot": 2, "TkFormulas": 1,
                  "TkFormulaSnapshot": 2, "PhaseBBridge": 1, "PhaseBResultsHook": 2}

XL_OPEN_XML_ADDIN = 55
UI_PART = "customUI/customUI14.xml"
UI_REL_TYPE = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
UI_REL_ID = "tkCustomUI"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


# ---------- Package editing (pure; tested offline) ----------

def inject_custom_ui(package: Path, ui_xml: bytes) -> None:
    """Add (or replace) customUI14.xml and its root relationship in an Office package.

    Every other part is copied byte for byte. Idempotent: a second run leaves one relationship.
    """
    ET.fromstring(ui_xml)  # refuse malformed XML before touching the package
    with zipfile.ZipFile(package) as source:
        entries = [(info, source.read(info.filename)) for info in source.infolist()]

    names = {info.filename for info, _ in entries}
    if "_rels/.rels" not in names or "[Content_Types].xml" not in names:
        raise ValueError(f"{package} is not an Office package")

    temp = package.with_suffix(package.suffix + ".tmp")
    with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as target:
        for info, data in entries:
            if info.filename == UI_PART:
                continue
            if info.filename == "_rels/.rels":
                data = _with_ui_relationship(data)
            elif info.filename == "[Content_Types].xml":
                data = _with_xml_default(data)
            target.writestr(info, data)
        target.writestr(UI_PART, ui_xml)
    temp.replace(package)


def _with_ui_relationship(rels: bytes) -> bytes:
    ET.register_namespace("", RELS_NS)
    root = ET.fromstring(rels)
    for rel in list(root):
        if rel.get("Type") == UI_REL_TYPE or rel.get("Id") == UI_REL_ID:
            root.remove(rel)
    ET.SubElement(root, f"{{{RELS_NS}}}Relationship", Id=UI_REL_ID, Type=UI_REL_TYPE, Target=UI_PART)
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")


def _with_xml_default(content_types: bytes) -> bytes:
    ET.register_namespace("", CT_NS)
    root = ET.fromstring(content_types)
    if not any(el.tag == f"{{{CT_NS}}}Default" and el.get("Extension", "").lower() == "xml" for el in root):
        ET.SubElement(root, f"{{{CT_NS}}}Default", Extension="xml", ContentType="application/xml")
    return ET.tostring(root, xml_declaration=True, encoding="UTF-8")


def verify_package(package: Path, ui_xml: bytes, expect_vba: bool = True) -> list[str]:
    """Return a list of problems (empty = the package is as expected)."""
    problems = []
    with zipfile.ZipFile(package) as z:
        names = set(z.namelist())
        if UI_PART not in names:
            problems.append(f"missing {UI_PART}")
        elif z.read(UI_PART) != ui_xml:
            problems.append(f"{UI_PART} differs from the source")
        rels = ET.fromstring(z.read("_rels/.rels"))
        ui_rels = [r for r in rels if r.get("Type") == UI_REL_TYPE]
        if len(ui_rels) != 1 or ui_rels[0].get("Target") != UI_PART:
            problems.append(f"expected one ui/extensibility relationship to {UI_PART}, found {len(ui_rels)}")
        if expect_vba and "xl/vbaProject.bin" not in names:
            problems.append("missing xl/vbaProject.bin (no VBA in the package)")
    return problems


# ---------- Build (needs Excel) ----------

def crlf_copies() -> list[Path]:
    out = BUILD / "vba"
    shutil.rmtree(out, ignore_errors=True)
    out.mkdir(parents=True)
    copies = []
    for source in VBA_SOURCES:
        text = source.read_bytes().replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
        target = out / source.name
        target.write_bytes(text)
        copies.append(target)
    return copies


def build_xlam(copies: list[Path]) -> None:
    import win32com.client  # noqa: PLC0415 - only needed for the Excel step
    import win32process  # noqa: PLC0415

    excel = win32com.client.DispatchEx("Excel.Application")  # new process; nothing else loaded
    _, pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Add()
        components = workbook.VBProject.VBComponents
        for copy in copies:
            components.Import(str(copy))
        found = {components.Item(i).Name: components.Item(i).Type for i in range(1, components.Count + 1)}
        wrong = {name: found.get(name) for name, kind in EXPECTED_TYPES.items() if found.get(name) != kind}
        if wrong:
            raise RuntimeError(f"VBA components missing or of the wrong type: {wrong}")
        TARGET.unlink(missing_ok=True)
        workbook.SaveAs(str(TARGET), FileFormat=XL_OPEN_XML_ADDIN)
        workbook.Close(SaveChanges=False)
    finally:
        excel.Quit()
        del excel
        _ensure_exited(pid)


def _ensure_exited(pid: int, seconds: float = 15.0) -> None:
    import win32api  # noqa: PLC0415
    import win32con  # noqa: PLC0415
    import win32event  # noqa: PLC0415

    try:
        handle = win32api.OpenProcess(win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid)
    except Exception:  # noqa: BLE001 - already gone
        return
    try:
        if win32event.WaitForSingleObject(handle, int(seconds * 1000)) != win32event.WAIT_OBJECT_0:
            win32api.TerminateProcess(handle, 1)
            print(f"build Excel {pid} did not exit within {seconds:.0f} s; terminated")
    finally:
        win32api.CloseHandle(handle)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--check", type=Path, help="only verify this existing .xlam")
    args = parser.parse_args()
    ui_xml = CUSTOM_UI.read_bytes()

    if args.check:
        problems = verify_package(args.check, ui_xml)
    else:
        BUILD.mkdir(exist_ok=True)
        started = time.perf_counter()
        build_xlam(crlf_copies())
        inject_custom_ui(TARGET, ui_xml)
        problems = verify_package(TARGET, ui_xml)
        print(f"built {TARGET} in {time.perf_counter() - started:.1f} s")
    for problem in problems:
        print("PROBLEM:", problem)
    print("package OK" if not problems else "package NOT OK")
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
