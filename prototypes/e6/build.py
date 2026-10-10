"""Build the E6 package: PyInstaller engine + .xlam + installer + manifest + licences -> zip.

  uv run python prototypes/e6/build.py 0.1.0

Dev machine only (the .xlam build imports VBA and needs "Trust access to the VBA project
object model"). Output: prototypes/e6/dist/ExcelToolkit-<version>.zip
"""

import hashlib
import json
import shutil
import subprocess
import sys
import sysconfig
import zipfile
from importlib import metadata
from pathlib import Path

import win32com.client

HERE = Path(__file__).resolve().parent
BUILD = HERE / "build"
DIST = HERE / "dist"
XL_OPEN_XML_ADDIN = 55


def build_engine(version: str, out: Path) -> Path:
    source = BUILD / "engine_src"
    shutil.rmtree(source, ignore_errors=True)
    shutil.copytree(HERE / "engine", source)
    engine = source / "toolkit_engine.py"
    engine.write_text(engine.read_text(encoding="utf-8").replace('VERSION = "0.0.0"', f'VERSION = "{version}"'), encoding="utf-8")
    subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--noconfirm", "--clean", "--onedir", "--name", "ExcelToolkitEngine",
         "--distpath", str(BUILD / "pyi_dist"), "--workpath", str(BUILD / "pyi_work"), "--specpath", str(BUILD),
         "--paths", str(source), "--exclude-module", "tkinter",
         "--hidden-import", "win32timezone",  # pywin32 imports it lazily for COM/Win32 dates
         str(engine)],
        check=True, capture_output=True, text=True,
    )
    target = out / "engine"
    shutil.copytree(BUILD / "pyi_dist" / "ExcelToolkitEngine", target)
    return target


def build_xlam(version: str, out: Path) -> Path:
    bas = BUILD / "ExcelToolkit.bas"
    text = (HERE / "ExcelToolkit.bas").read_text(encoding="ascii").replace("{{VERSION}}", version)
    bas.write_bytes(text.replace("\r\n", "\n").replace("\n", "\r\n").encode("ascii"))
    target = out / "ExcelToolkit.xlam"
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Add()
        workbook.VBProject.VBComponents.Import(str(bas))
        workbook.SaveAs(str(target), FileFormat=XL_OPEN_XML_ADDIN)
        workbook.Close(SaveChanges=False)
    finally:
        excel.Quit()
    return target


def licence_texts(out: Path) -> list[dict]:
    """Licences of everything shipped inside the package."""
    out.mkdir(parents=True, exist_ok=True)
    register = []
    python_licence = Path(sysconfig.get_config_var("prefix") or sys.base_prefix) / "LICENSE.txt"
    if not python_licence.exists():
        python_licence = Path(sys.base_prefix) / "LICENSE.txt"
    shutil.copy(python_licence, out / "Python-LICENSE.txt")
    register.append({"component": f"Python {sys.version.split()[0]} runtime (incl. its bundled DLLs)", "licence": "PSF License (see Python-LICENSE.txt, which also lists bundled third-party components)"})
    for dist_name, label in (("pywin32", "pywin32"), ("pyinstaller", "PyInstaller bootloader")):
        dist = metadata.distribution(dist_name)
        licence = dist.metadata.get("License-Expression") or dist.metadata.get("License") or ""
        for file in dist.files or []:
            if "licen" in file.name.lower() or "copying" in file.name.lower():
                shutil.copy(file.locate(), out / f"{dist_name}-{file.name}")
        register.append({"component": f"{label} {dist.version}", "licence": licence.splitlines()[0] if licence else "see files"})
    return register


def main() -> None:
    version = sys.argv[1] if len(sys.argv) > 1 else "0.1.0"
    BUILD.mkdir(exist_ok=True)
    DIST.mkdir(exist_ok=True)
    package = BUILD / f"ExcelToolkit-{version}"
    if package.exists():
        shutil.rmtree(package)  # fail loudly rather than build on top of stale files
    payload = package / "payload"
    payload.mkdir(parents=True)

    build_engine(version, payload)
    build_xlam(version, payload)
    for name in ("install.cmd", "install.ps1", "uninstall.cmd", "uninstall.ps1"):
        shutil.copy(HERE / "installer" / name, package / name)
    register = licence_texts(package / "licences")
    register.append({"component": "Excel Toolkit add-in, engine code and installer scripts", "licence": "internal"})
    (package / "THIRD_PARTY_LICENSES.json").write_text(json.dumps(register, indent=1), encoding="utf-8")

    files = sorted(p for p in payload.rglob("*") if p.is_file())
    manifest = {"version": version, "files": [
        {"path": str(p.relative_to(payload)), "sha256": hashlib.sha256(p.read_bytes()).hexdigest().upper()} for p in files]}
    (package / "manifest.json").write_text(json.dumps(manifest, indent=1), encoding="utf-8")

    archive = DIST / f"ExcelToolkit-{version}.zip"
    archive.unlink(missing_ok=True)
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(package.rglob("*")):
            if p.is_file():
                z.write(p, p.relative_to(package.parent))
    size = sum(p.stat().st_size for p in files)
    print(f"{archive} ({archive.stat().st_size / 1e6:.1f} MB zipped, payload {size / 1e6:.1f} MB, {len(files)} files)")


if __name__ == "__main__":
    main()
