# Architecture

The app is a Flask backend (`server.py`) serving a single-page React UI (`static/index.html`). The core logic lives in `Merge.py`. PyInstaller bundles all three into one `Gospelink.exe`.

## Merge.py (core logic)

**Classes**
- **`Donor`**: one account, with its envelope info, address, email preference and a dict of `Preacher`s.
- **`Preacher`**: one sponsored preacher under a donor. It tracks that preacher's report PDF, the thank-you note, and whether a note is required.

**Pipeline functions**
- **`getDonors`**: reads the Cover Sheet Spreadsheet.
- **`assignCoverLetters`**: matches cover letters to donors.
- **`enforceTY`**: reads the Extra Gift Report and flags preacher–donor pairs that need a TY note.
- **`addReports`**: attaches report and TY-note files.
- **`merge` / `createCoverLetter`**: build each donor's output PDF.

**Helpers**
- **`_read_required_columns`**: tolerant spreadsheet reading. See [spreadsheet-formats.md](spreadsheet-formats.md).
- **`_validate_pdf`**: checks files before PyMuPDF opens them.
- **`docx_to_pdf`**: converts through Word COM.

## server.py (REST API)

The pipeline steps map to `/api/load`, `/api/assign-coverletters`, `/api/enforce-ty`, `/api/add-reports` and `/api/run-all`. Merge, history, config and update-check endpoints sit around them.

Errors come back as `{"ok": false, "error": "..."}`, and the UI shows the message in a toast. A `SpreadsheetFormatError` returns a 400 with a plain-English message. Anything else returns a 500 and prints a traceback to the console.

## Workflow (UI)

1. **Configuration**: pick the input files, folders, template and output folder.
2. **Load donors**: cover sheet.
3. **Assign cover letters**.
4. **Enforce TY**: extra gift report.
5. **Add reports**: preacher reports and TY notes.
6. **Merge**: one PDF per donor in the output folder.

**Run All** runs steps 2–5 in a row.
