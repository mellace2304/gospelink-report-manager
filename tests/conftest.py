import os
import sys

import openpyxl
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "tests", "data")
sys.path.insert(0, ROOT)

COVER_HEADERS = [
    "Type", "Status", "Note", "Sponsorship Amount", "Pledge", "Donor/Preacher Combo",
    "Account Number", "Preacher Number", "Preacher Name", "Envelope Number", "Envelope Name",
    "Primary Street", "Primary City", "Primary State", "Primary ZIP Code",
    "Primary Email Address", "Send Quarterly Report Via",
]
GIFT_HEADERS = ["ID", "Name", "Amount", "Donor #", "Donor Name"]


@pytest.fixture
def make_xlsx(tmp_path):
    """make_xlsx(name, rows, sheet="Sheet1") -> path. rows[0] is usually the header row."""
    def _make(name, rows, sheet="Sheet1"):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = sheet
        for row in rows:
            ws.append(row)
        path = tmp_path / f"{name}.xlsx"
        wb.save(path)
        return str(path)
    return _make


@pytest.fixture
def cover_rows():
    """Two donors: 100 (envelope 1001, preachers 900001 and 900002c) and 200 (no envelope #)."""
    def row(acct, pnum, pname, env, ename):
        return ["Sponsor", "Active", "", 30, 30, "", acct, pnum, pname, env, ename,
                "1 Main St\nApt 2", "Ames", "IA", 50010, f"{acct}@example.com", "E-Mail"]
    return [
        COVER_HEADERS,
        row(100, 900001, "Preacher, Alpha", 1001, "Test Donor One"),
        row(100, "900002c", "Preacher, Bravo", 1001, "Test Donor One"),
        row(200, 900003, "Preacher, Charlie", None, "Test Donor Two"),
    ]
