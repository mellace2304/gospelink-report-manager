# gospelink-report-manager

An automated quarterly report manager for Gospelink. It merges donor information, preacher reports, cover letters and thank-you notes into one PDF package per donor.

It runs as a single Windows executable: a Flask backend with a browser-based React UI. Download the latest `Gospelink.exe` from [Releases](../../releases). The app updates itself from there.

## Project structure

```
.
├── server.py              # Flask backend + REST API (entry point for the exe)
├── Merge.py               # Core logic: Donor/Preacher, spreadsheet loading, merging
├── static/index.html      # React UI
├── tests/                 # pytest suite (tests/data/ is local-only, gitignored)
├── notebooks/             # scratch notebook for interactive work
├── docs/                  # documentation
├── requirements.txt
└── .github/workflows/     # test + build + release
```

## Docs

- [Architecture](docs/architecture.md): how the pieces fit and the report workflow
- [Spreadsheet formats](docs/spreadsheet-formats.md): required columns, plus troubleshooting for spreadsheet errors
- [Development](docs/development.md): setup, running from source, building and releasing
- [Testing](docs/testing.md): running the tests and handling real client data

## Quick start (from source)

```bash
python -m venv .venv && .venv\Scripts\activate
pip install -r requirements.txt
python server.py
python -m pytest tests -q
```

You need Windows and Python 3.12+.

## License

See [LICENSE](LICENSE).
