"""Regression run against real client spreadsheets in tests/data/ (gitignored, local only).

Skipped when the files aren't present, e.g. in CI. See docs/testing.md.
"""
import os

import pytest

import Merge
from conftest import DATA_DIR

COVER = os.path.join(DATA_DIR, "cover_sheet.xlsx")
GIFTS = os.path.join(DATA_DIR, "extra_gifts.xlsx")

pytestmark = pytest.mark.skipif(
    not (os.path.exists(COVER) and os.path.exists(GIFTS)),
    reason="real data not in tests/data/",
)


def test_real_cover_sheet_and_gifts():
    donors = Merge.getDonors(COVER)
    unmatched = Merge.enforceTY(donors, GIFTS)
    enforced = sum(p.noteNecessary for d in donors.values() for p in d.preachers.values())
    # Baseline from the 2026.07.01 cover sheet + matching Spons Xtra export (v1.5.7 == v1.5.8)
    assert len(donors) == 2682
    assert enforced == 354
    assert all(n.isdigit() for n in unmatched)
