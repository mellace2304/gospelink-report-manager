# Development

## Setup

You need Windows and Python **3.12+**. `Merge.py` uses PEP 701 f-strings and Word COM automation through `pywin32`.

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python server.py          # opens http://localhost:5000
```

When run from source, the server keeps its state next to `server.py`. When run as the exe, it keeps it next to `Gospelink.exe`. The state files are `gospelink_config.json`, `gospelink_config_history.json` and `merged_donors.json`. All three are gitignored.

## Releasing

1. Bump `APP_VERSION` in `server.py`. It must match the tag, because the in-app updater compares the two.
2. Commit and push to `main`. CI runs the tests and builds a temporary artifact.
3. Tag and push the release:

   ```bash
   git tag vX.Y.Z
   git push origin vX.Y.Z
   ```

   The tag build attaches `Gospelink.exe` to a GitHub Release, and users' apps then offer the update.

The tag build uses the workflow file **as of the tagged commit**. If a release build fails, fix the problem and re-tag onto the fixed commit.

## Building locally

Use the same command as CI (`.github/workflows/main.yml`):

```bash
pyinstaller --onefile --name Gospelink --add-data "static;static" --paths . --hidden-import win32com --hidden-import win32com.client --collect-all pywin32 --collect-all fitz --icon=icon.ico server.py
```

## Gotcha: "No module named 'Merge'" in the built exe

In the past, the CI-built exe failed at runtime even though local builds worked. The cause was CI running Python 3.11. On 3.11, `Merge.py` is a `SyntaxError` (PEP 701 f-string), and PyInstaller **silently drops** modules it can't compile instead of failing the build.

Keep CI's `python-version` at 3.12 or newer. If "works locally, fails in CI" comes back, compare the two Python versions first, and try `python -c "import Merge"` on the CI interpreter.
