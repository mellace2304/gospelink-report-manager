# Testing

```bash
pip install pytest
python -m pytest tests -q
```

Run the tests from the repo root with Python 3.12+ and the packages in `requirements.txt` installed. `Merge.py` imports `pywin32`, so the tests need Windows.

## What's covered

- **`tests/test_spreadsheets.py`** builds small synthetic spreadsheets in a temp folder. It checks `getDonors` and `enforceTY` against:
  - the normal layout,
  - title rows above the headers,
  - variant header spellings,
  - renamed tabs,
  - float IDs,
  - blank and memo rows,
  - and the error messages for unrecognized headers and the wrong file.

  It needs no real data and runs in CI.
- **`tests/test_real_data.py`** is a regression run against real client exports. It only runs when they're present locally:

  | File | What it is |
  |---|---|
  | `tests/data/cover_sheet.xlsx` | a Cover Sheet Data export |
  | `tests/data/extra_gifts.xlsx` | the matching Extra Gift Report (Spons Xtra) export |

  The expected counts in the test (2682 donors, 354 TY flags) belong to the 2026.07.01 exports. Update them if you swap in newer files.

`tests/data/` is gitignored. Client spreadsheets hold donor names, addresses and emails, so they must never be committed.

## CI

`.github/workflows/main.yml` runs the suite before the PyInstaller build on every push to `main` and on every `v*` tag. A failing test stops the build, so a broken version never becomes a release. The real-data test is skipped there.

## Scratch work

`notebooks/scratch.ipynb` is for interactive poking at `Merge.py`. Its first cell moves the working directory to the repo root, so `import Merge` and the relative paths work.
