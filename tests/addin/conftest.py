"""Make ``addin/`` importable for the offline M2 add-in tests (no Excel is started).

Kept test-local so the project's pyproject.toml / uv.lock stay untouched. Run under the isolated
scratch environment (openpyxl + pytest).
"""

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
p = str(ROOT / "addin")
if p not in sys.path:
    sys.path.insert(0, p)
