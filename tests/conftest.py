"""Make ``src/`` importable for the isolated detector dev environment.

Kept as test-local config so the project's pyproject.toml / uv.lock stay untouched
(the full project environment restoration is deferred; see the incident record).
"""

import pathlib
import sys

SRC = pathlib.Path(__file__).resolve().parent.parent / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))
