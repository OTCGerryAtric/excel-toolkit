"""Make ``src/`` and the Phase B prototype modules importable for the offline Phase B tests.

Kept test-local so the project's pyproject.toml / uv.lock stay untouched. These tests import
``audit_engine`` (which imports ``excel_link`` -> pywin32), so they must run under the isolated
scratch environment that has pywin32 + openpyxl + pytest.
"""

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
for rel in ("src", "prototypes/integration", "prototypes/e6/engine", "prototypes/e4"):
    p = str(ROOT / rel)
    if p not in sys.path:
        sys.path.insert(0, p)
