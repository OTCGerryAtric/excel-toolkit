"""Build E1Test.xlam from E1Test.bas in a separate, hidden Excel instance.

Dev-machine only: importing VBA programmatically needs "Trust access to the VBA
project object model". Colleagues never run this.
"""

from pathlib import Path

import win32com.client

HERE = Path(__file__).resolve().parent
SOURCE = HERE / "E1Test.bas"
TARGET = HERE / "E1Test.xlam"
XL_OPEN_XML_ADDIN = 55


def main() -> None:
    # VBA import expects CRLF line endings.
    crlf_copy = HERE / "build" / SOURCE.name
    crlf_copy.parent.mkdir(exist_ok=True)
    crlf_copy.write_bytes(SOURCE.read_text(encoding="ascii").replace("\r\n", "\n").replace("\n", "\r\n").encode("ascii"))

    excel = win32com.client.DispatchEx("Excel.Application")  # new instance; no XLSTART/add-ins loaded
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        workbook = excel.Workbooks.Add()
        workbook.VBProject.VBComponents.Import(str(crlf_copy))
        TARGET.unlink(missing_ok=True)
        workbook.SaveAs(str(TARGET), FileFormat=XL_OPEN_XML_ADDIN)
        workbook.Close(SaveChanges=False)
    finally:
        excel.Quit()
    print(f"Built {TARGET}")


if __name__ == "__main__":
    main()
