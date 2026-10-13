"""Spreadsheet loading: getDonors (cover sheet) and enforceTY (extra gift report)."""
import pytest

import Merge
from conftest import COVER_HEADERS, GIFT_HEADERS


def enforced(donors):
    return sorted((d.aNum, p.pNum) for d in donors.values() for p in d.preachers.values() if p.noteNecessary)


# ── Cover sheet ──────────────────────────────────────────────────────────────

def test_cover_sheet_loads(make_xlsx, cover_rows):
    donors = Merge.getDonors(make_xlsx("cover", cover_rows))
    assert set(donors) == {"100", "200"}
    d = donors["100"]
    assert d.eNum == "1001"
    assert set(d.preachers) == {"900001", "900002c", "900002"}  # 900002 is the child's guardian placeholder
    assert d.street == "1 Main St Apt 2"
    assert d.send_quarterly
    assert donors["200"].eNum == ""


def test_cover_sheet_missing_optional_columns(make_xlsx, cover_rows):
    drop = {COVER_HEADERS.index(c) for c in ["Pledge", "Note", "Primary Street"]}
    rows = [[v for i, v in enumerate(r) if i not in drop] for r in cover_rows]
    donors = Merge.getDonors(make_xlsx("cover", rows))
    assert donors["100"].street == ""


def test_cover_sheet_missing_required_column(make_xlsx, cover_rows):
    i = COVER_HEADERS.index("Envelope Number")
    rows = [r[:i] + r[i + 1:] for r in cover_rows]
    with pytest.raises(Merge.SpreadsheetFormatError, match="Cover Sheet Spreadsheet.*'Envelope Number'"):
        Merge.getDonors(make_xlsx("cover", rows))


# ── Extra gift report ────────────────────────────────────────────────────────

@pytest.fixture
def donors(make_xlsx, cover_rows):
    return Merge.getDonors(make_xlsx("cover", cover_rows))


GIFTS = [[900001, "Preacher, Alpha", 9, 1001, "Test Donor One"],
         [900002, "Preacher, Bravo", 80, 1001, "Test Donor One"],
         [900003, "Preacher, Charlie", 50, 9999, "Test Donor Unknown"]]
EXPECTED = [("100", "900001"), ("100", "900002")]


def test_extra_gifts_standard(make_xlsx, donors):
    unmatched = Merge.enforceTY(donors, make_xlsx("gifts", [GIFT_HEADERS] + GIFTS, "Spons Xtra"))
    assert enforced(donors) == EXPECTED
    assert unmatched == ["9999"]


@pytest.mark.parametrize("name,rows,sheet", [
    ("title_rows", [["Extra Gifts Title Row"], [], GIFT_HEADERS] + GIFTS, "Spons Xtra"),
    ("header_variants", [["ID ", "Name", "Amount", "Donor Number", "Donor Name"]] + GIFTS, "Spons Xtra"),
    ("tab_case_spacing", [GIFT_HEADERS] + GIFTS, " spons xtra"),
    ("tab_renamed", [GIFT_HEADERS] + GIFTS, "Sheet1"),
    ("float_ids", [GIFT_HEADERS] + [[float(r[0]), r[1], r[2], float(r[3]), r[4]] for r in GIFTS], "Spons Xtra"),
])
def test_extra_gifts_tolerated_layouts(make_xlsx, donors, name, rows, sheet):
    Merge.enforceTY(donors, make_xlsx(name, rows, sheet))
    assert enforced(donors) == EXPECTED


def test_extra_gifts_skips_blank_and_memo_rows(make_xlsx, donors):
    rows = [GIFT_HEADERS] + GIFTS + [[None, None, 10, 1001, ""], [None, None, None, "for church rent.", None]]
    unmatched = Merge.enforceTY(donors, make_xlsx("gifts", rows, "Spons Xtra"))
    assert enforced(donors) == EXPECTED
    assert unmatched == ["9999"]


def test_extra_gifts_unrecognized_headers(make_xlsx, donors):
    rows = [["Sponsor", "Name", "Amount", "Acct", "Donor Name"]] + GIFTS
    with pytest.raises(Merge.SpreadsheetFormatError) as e:
        Merge.enforceTY(donors, make_xlsx("bad", rows, "Spons Xtra"))
    msg = str(e.value)
    assert "Extra Gift Report" in msg and "'ID'" in msg and "'Donor #'" in msg
    assert "'Sponsor'" in msg  # shows what it did find


def test_extra_gifts_wrong_file(make_xlsx, donors, cover_rows):
    with pytest.raises(Merge.SpreadsheetFormatError, match="no tab named 'Spons Xtra'"):
        Merge.enforceTY(donors, make_xlsx("wrong", cover_rows, "Data"))
