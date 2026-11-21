"""Shared fixtures for the p107 test suite.

Only the standard library is imported here on purpose: tests/test_data.py must
run before pandas and Streamlit are installed (Step 14), so nothing in this file
may depend on the runtime stack.

Every test opens data/app.db read-only. Nothing in this suite writes to the
warehouse (NFR-09, AS-39).
"""
from __future__ import annotations

import json
import sqlite3
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent
DB_PATH = ROOT / "data" / "app.db"
SUMMARY_PATH = ROOT / "data" / "validation_summary.json"
QUALITY_LOG = ROOT / "data" / "quality_log.csv"

# The suite imports src.kpi / src.stats directly, so the repository root must be
# importable even when pytest is invoked from another directory.
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


@pytest.fixture(scope="session")
def root() -> Path:
    return ROOT


@pytest.fixture(scope="session")
def summary() -> dict:
    return json.loads(SUMMARY_PATH.read_text(encoding="utf-8"))


@pytest.fixture(scope="session")
def db():
    """Read-only connection, opened in URI mode so a stray INSERT cannot succeed."""
    if not DB_PATH.exists():
        pytest.fail(f"data/app.db is missing — run scripts/generate_data.py first ({DB_PATH})")
    con = sqlite3.connect(f"file:{DB_PATH}?mode=ro", uri=True)
    con.row_factory = sqlite3.Row
    yield con
    con.close()


@pytest.fixture(scope="session")
def tables() -> dict[str, int]:
    """The five warehouse tables and their validated row counts (seed 42)."""
    return {
        "dim_department": 8,
        "dim_ward": 20,
        "dim_payer": 5,
        "fact_inpatient_episode": 53_915,
        "fact_opd_visit": 143_139,
    }
