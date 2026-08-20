# pyinstaller --onefile --name Gospelink --add-data "static;static" --hidden-import win32com --hidden-import win32com.client --collect-all docx2pdf --collect-all fitz --icon=icon.ico server.py
"""
Gospelink Quarterly Report Manager — Flask Backend
Wraps Merge.py logic and exposes REST API endpoints for the frontend.
"""

from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
import os
import json
import traceback
from datetime import datetime
import webbrowser
import threading
import urllib.request, tempfile, subprocess, sys
    
from Merge import (
    Donor, Preacher, getDonors,
    enforceTY, addReports, assignCoverLetters,
    merge, createCoverLetter,
    docx_to_pdf
)

app = Flask(__name__, static_folder="static")
CORS(app)

# ── Paths ─────────────────────────────────────────────────────────────────────
if getattr(sys, 'frozen', False):
    # Running as a PyInstaller exe — use the exe's location
    BASE_DIR = os.path.dirname(sys.executable)
else:
    # Running as a normal .py script
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "gospelink_config.json")
CONFIG_HISTORY_FILE = os.path.join(BASE_DIR, "gospelink_config_history.json")
MERGE_HISTORY_FILE = os.path.join(BASE_DIR, "merged_donors.json")

# ── Global State ──────────────────────────────────────────────────────────────
state = {
    "donors": {},
    "loaded": False,
    "steps_completed": [],
    "config": {
        "coverSheetFile": "",
        "extraGiftFile": "",
        "coverLetterDir": "",
        "reportsDir": "",
        "outputDir": "",
        "templatePath": "",
    },
    "logs": [],
}


def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    entry = f"[{ts}] {msg}"
    state["logs"].append(entry)
    if len(state["logs"]) > 1000:
        del state["logs"][:-1000]
    print(entry)

# ── Config Persistence ────────────────────────────────────────────────────────

def load_saved_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                saved = json.load(f)
            for k in state["config"]:
                if k in saved:
                    state["config"][k] = saved[k]
            log("Loaded saved configuration")
        except Exception:
            pass


def save_config():
    with open(CONFIG_FILE, "w") as f:
        json.dump(state["config"], f, indent=2)


def save_config_to_history():
    history = load_config_history()
    entry = {
        "timestamp": datetime.now().isoformat(),
        **state["config"],
    }
    if history and all(history[-1].get(k) == entry.get(k) for k in state["config"]):
        return
    history.append(entry)
    history = history[-20:]
    with open(CONFIG_HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=2)


def load_config_history():
    if os.path.exists(CONFIG_HISTORY_FILE):
        try:
            with open(CONFIG_HISTORY_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return []
    return []


load_saved_config()


# ── Serialization ─────────────────────────────────────────────────────────────

def serialize_preacher(p: Preacher) -> dict:
    return {
        "pNum": p.pNum,
        "pName": p.pName,
        "report": p.report,
        "note": p.note,
        "noteNecessary": p.noteNecessary,
        "ready": p.ready(),
        "reasons": p.reasons(),
    }


def serialize_donor(d: Donor, merge_history: dict = None) -> dict:
    preachers = {k: serialize_preacher(v) for k, v in d.preachers.items()}
    entry = (merge_history or {}).get(d.aNum, {})
    return {
        "eNum": d.eNum,
        "eName": d.eName,
        "street": d.street,
        "city": d.city,
        "state": d.state,
        "zip": d.zip,
        "email": d.email,
        "aNum": d.aNum,
        "sendQuarterly": d.send_quarterly,
        "coverLetter": d.coverLetter,
        "extraNotes": d.extra_notes,
        "preachers": preachers,
        "ready": d.ready(),
        "reasons": d.reasons(),
        "fileCount": len(d.getFiles()),
        "merged": bool(entry.get("merged", False)),
        "mergedManual": bool(entry.get("manual", False)),
    }


def get_donor_filename(donor: Donor) -> str:
    return f"{'d' + donor.eNum if donor.eNum else 'a' + donor.aNum}{'e' if donor.send_quarterly else ''}.pdf"

# ── File / Folder Picker ─────────────────────────────────────────────────────

def _pick_file_dialog(title, filetypes):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.focus_force()
    ft = [(label, pat) for label, pat in filetypes]
    path = filedialog.askopenfilename(title=title, filetypes=ft)
    root.destroy()
    return path or ""


def _pick_folder_dialog(title):
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    root.focus_force()
    path = filedialog.askdirectory(title=title)
    root.destroy()
    return path or ""


@app.route("/api/pick-file", methods=["POST"])
def pick_file():
    data = request.json or {}
    title = data.get("title", "Select a file")
    filetypes = data.get("filetypes", [["All files", "*.*"]])
    path = _pick_file_dialog(title, filetypes)
    return jsonify({"path": path})


@app.route("/api/pick-folder", methods=["POST"])
def pick_folder():
    data = request.json or {}
    title = data.get("title", "Select a folder")
    path = _pick_folder_dialog(title)
    return jsonify({"path": path})

# ── Routes: Config ────────────────────────────────────────────────────────────

APP_VERSION = "1.5.6"
GITHUB_REPO = "mellace2304/gospelink-report-manager"

@app.route("/api/check-update", methods=["GET"])
def check_update():
    try:
        import urllib.request, json
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        req = urllib.request.Request(url, headers={"User-Agent": "Gospelink"})
        with urllib.request.urlopen(req, timeout=5) as r:
            data = json.loads(r.read())
        latest = data["tag_name"].lstrip("v")
        asset = next((a for a in data["assets"] if a["name"].endswith(".exe")), None)
        download_url = asset["browser_download_url"] if asset else None
        return jsonify({
            "current": APP_VERSION,
            "latest": latest,
            "updateAvailable": latest != APP_VERSION and download_url is not None,
            "downloadUrl": download_url,
        })
    except Exception as e:
        return jsonify({"error": str(e), "current": APP_VERSION})

@app.route("/api/apply-update", methods=["POST"])
def apply_update():
    """Download new exe and launch a helper script to swap it, then exit."""
    data = request.json or {}
    url = data.get("url")
    if not url:
        return jsonify({"ok": False, "error": "No URL provided"}), 400
    
    exe_path = sys.executable
    new_exe = exe_path + ".new"
    
    # Download the update
    urllib.request.urlretrieve(url, new_exe)
    
    # Write a tiny batch script that waits for this process to exit,
    # swaps the files, then relaunches
    bat = tempfile.NamedTemporaryFile(delete=False, suffix=".bat", mode="w")
    bat.write(f"""@echo off
    timeout /t 2 /nobreak >nul
    move /y "{new_exe}" "{exe_path}"
    start "" "{exe_path}"
    del "%~f0"
    """)
    bat.close()
    
    subprocess.Popen(["cmd", "/c", bat.name], 
                     creationflags=subprocess.CREATE_NO_WINDOW)
    
    # Shut down Flask — the bat script will relaunch the app
    os.kill(os.getpid(), 9)
    return jsonify({"ok": True})

@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/api/config", methods=["GET"])
def get_config():
    return jsonify(state["config"])


@app.route("/api/config", methods=["POST"])
def set_config():
    data = request.json
    for key in state["config"]:
        if key in data:
            state["config"][key] = data[key]
    save_config()
    save_config_to_history()
    log("Configuration saved")
    return jsonify({"ok": True, "config": state["config"]})


@app.route("/api/config/history", methods=["GET"])
def get_config_history():
    return jsonify(load_config_history())


@app.route("/api/config/restore", methods=["POST"])
def restore_config():
    data = request.json or {}
    idx = data.get("index")
    history = load_config_history()
    if idx is not None and 0 <= idx < len(history):
        entry = history[idx]
        for k in state["config"]:
            if k in entry:
                state["config"][k] = entry[k]
        save_config()
        log("Configuration restored from history")
        return jsonify({"ok": True, "config": state["config"]})
    return jsonify({"ok": False, "error": "Invalid history index"}), 400


# ── Routes: Status ────────────────────────────────────────────────────────────

@app.route("/api/status", methods=["GET"])
def get_status():
    donors = state["donors"]
    total = len(donors)
    ready = sum(1 for d in donors.values() if d.ready())

    missing_cover = sum(1 for d in donors.values() if d.coverLetter == "")
    missing_report = sum(1 for d in donors.values() for p in d.preachers.values() if p.report == "")
    missing_ty = sum(1 for d in donors.values() for p in d.preachers.values() if p.noteNecessary and p.note == "")

    # Authoritative merged count — mirrors the same merged_donors.json lookup
    # that /api/merge/ready and /api/merge/all use to decide what's "fresh",
    # so this number always matches what a fresh-only merge will actually do.
    # (Do not use the Merge_* folder/manifest scan here — individual merges
    # and manual "mark as merged" toggles never create folders/manifests, so
    # that scan undercounts and desyncs from the real fresh-only filter.)
    merged_status = get_merged_status(state["config"]["outputDir"])
    ready_merged = sum(
        1 for d in donors.values()
        if d.ready() and merged_status.get(d.aNum, {}).get("merged")
    )

    return jsonify({
        "version": APP_VERSION,
        "loaded": state["loaded"],
        "stepsCompleted": state["steps_completed"],
        "total": total,
        "ready": ready,
        "notReady": total - ready,
        "readyMerged": ready_merged,
        "readyFresh": ready - ready_merged,
        "missingCoverLetters": missing_cover,
        "missingReports": missing_report,
        "missingTYNotes": missing_ty,
        "config": state["config"],
    })


# ── Routes: Pipeline Steps ───────────────────────────────────────────────────

@app.route("/api/load", methods=["POST"])
def load_donors():
    try:
        fp = state["config"]["coverSheetFile"]
        if not fp or not os.path.exists(fp):
            return jsonify({"ok": False, "error": f"File not found: {fp}"}), 400
        state["donors"] = getDonors(fp)
        state["loaded"] = True
        if "load" not in state["steps_completed"]:
            state["steps_completed"].append("load")
        log(f"Loaded {len(state['donors'])} donors")
        return jsonify({"ok": True, "count": len(state["donors"])})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/assign-coverletters", methods=["POST"])
def assign_coverletters():
    try:
        d = state["config"]["coverLetterDir"]
        if not d or not os.path.isdir(d):
            return jsonify({"ok": False, "error": f"Directory not found: {d}"}), 400
        assignCoverLetters(state["donors"], d)
        assigned = sum(1 for dn in state["donors"].values() if dn.coverLetter != "")
        if "coverletters" not in state["steps_completed"]:
            state["steps_completed"].append("coverletters")
        log(f"Assigned {assigned} cover letters")
        return jsonify({"ok": True, "assigned": assigned})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/enforce-ty", methods=["POST"])
def enforce_ty():
    try:
        fp = state["config"]["extraGiftFile"]
        if not fp or not os.path.exists(fp):
            return jsonify({"ok": False, "error": f"File not found: {fp}"}), 400
        enforceTY(state["donors"], fp)
        enforced = sum(1 for d in state["donors"].values() for p in d.preachers.values() if p.noteNecessary)
        if "enforce_ty" not in state["steps_completed"]:
            state["steps_completed"].append("enforce_ty")
        log(f"Enforced TY on {enforced} preacher-donor combos")
        return jsonify({"ok": True, "enforced": enforced})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/add-reports", methods=["POST"])
def add_reports():
    try:
        d = state["config"]["reportsDir"]
        if not d or not os.path.isdir(d):
            return jsonify({"ok": False, "error": f"Directory not found: {d}"}), 400
        addReports(state["donors"], d)
        rpts = sum(1 for dn in state["donors"].values() for p in dn.preachers.values() if p.report != "")
        notes = sum(1 for dn in state["donors"].values() for p in dn.preachers.values() if p.note != "")
        if "reports" not in state["steps_completed"]:
            state["steps_completed"].append("reports")
        log(f"Found {rpts} reports, {notes} TY notes")
        return jsonify({"ok": True, "reports": rpts, "notes": notes})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/run-all", methods=["POST"])
def run_all():
    results = {}
    try:
        cfg = state["config"]
        state["donors"] = getDonors(cfg["coverSheetFile"])
        state["loaded"] = True
        state["steps_completed"] = ["load"]
        results["load"] = len(state["donors"])

        if cfg["coverLetterDir"] and os.path.isdir(cfg["coverLetterDir"]):
            assignCoverLetters(state["donors"], cfg["coverLetterDir"])
            state["steps_completed"].append("coverletters")
            results["coverletters"] = sum(1 for d in state["donors"].values() if d.coverLetter != "")

        if cfg["extraGiftFile"] and os.path.exists(cfg["extraGiftFile"]):
            enforceTY(state["donors"], cfg["extraGiftFile"])
            state["steps_completed"].append("enforce_ty")

        if cfg["reportsDir"] and os.path.isdir(cfg["reportsDir"]):
            addReports(state["donors"], cfg["reportsDir"])
            state["steps_completed"].append("reports")

        ready = sum(1 for d in state["donors"].values() if d.ready())
        results["ready"] = ready
        results["total"] = len(state["donors"])
        log(f"Pipeline complete: {ready}/{results['total']} ready")
        return jsonify({"ok": True, "results": results})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e), "results": results}), 500


# ── Routes: Cover Letter Creation ─────────────────────────────────────────────

@app.route("/api/create-coverletter/<aNum>", methods=["POST"])
def create_single_coverletter(aNum):
    try:
        donor = state["donors"].get(aNum)
        if not donor:
            return jsonify({"ok": False, "error": "Donor not found"}), 404

        cl_dir = state["config"]["coverLetterDir"]
        tmpl = state["config"]["templatePath"]
        if not cl_dir:
            return jsonify({"ok": False, "error": "Cover letter directory not configured"}), 400
        if not tmpl or not os.path.exists(tmpl):
            return jsonify({"ok": False, "error": f"Template not found: {tmpl}"}), 400

        os.makedirs(cl_dir, exist_ok=True)
        createCoverLetter(donor, output_path=cl_dir, template_path=tmpl)

        # Convert the single docx
        docx_path = os.path.join(cl_dir, f"{donor.aNum}.docx")
        pdf_path = os.path.join(cl_dir, f"{donor.aNum}.pdf")
        if os.path.exists(docx_path) and not os.path.exists(pdf_path):
            docx_to_pdf(cl_dir, cl_dir, files=[docx_path])

        if os.path.exists(pdf_path):
            donor.coverLetter = pdf_path

        log(f"Cover letter created for {donor.eName} (A#{aNum})")
        return jsonify({"ok": True, "aNum": aNum})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/create-coverletters", methods=["POST"])
def create_mass_coverletters():
    try:
        cl_dir = state["config"]["coverLetterDir"]
        tmpl = state["config"]["templatePath"]
        if not cl_dir:
            return jsonify({"ok": False, "error": "Cover letter directory not configured"}), 400
        if not tmpl or not os.path.exists(tmpl):
            return jsonify({"ok": False, "error": f"Template not found: {tmpl}"}), 400

        os.makedirs(cl_dir, exist_ok=True)
        data = request.json or {}
        target = data.get("target", "missing")  # "missing" | "all"

        created = 0
        for donor in state["donors"].values():
            if target == "missing" and donor.coverLetter != "":
                continue
            createCoverLetter(donor, output_path=cl_dir, template_path=tmpl)
            created += 1

        docx_to_pdf(cl_dir, cl_dir)
        assignCoverLetters(state["donors"], cl_dir)
        assigned = sum(1 for d in state["donors"].values() if d.coverLetter != "")

        log(f"Created {created} cover letters, {assigned} assigned as PDF")
        return jsonify({"ok": True, "created": created, "assigned": assigned})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


# ── Merge History Tracking ────────────────────────────────────────────────────

def get_merge_folders(output_dir: str) -> list[str]:
    if not os.path.isdir(output_dir):
        return []
    return sorted(
        n for n in os.listdir(output_dir)
        if n.startswith("Merge_") and os.path.isdir(os.path.join(output_dir, n))
    )


def get_previously_merged(output_dir: str) -> set[str]:
    """Legacy detection: filenames found by scanning Merge_* folders/manifests
    in the currently configured output dir. Kept as a fallback/migration
    source — see get_merged_status()."""
    merged = set()
    for folder_name in get_merge_folders(output_dir):
        folder_path = os.path.join(output_dir, folder_name)
        manifest_path = os.path.join(folder_path, "manifest.json")
        if os.path.exists(manifest_path):
            try:
                with open(manifest_path, "r") as f:
                    manifest = json.load(f)
                merged.update(manifest.get("files", []))
            except Exception:
                pass
        for f in os.listdir(folder_path):
            if f.endswith(".pdf"):
                merged.add(f)
    return merged


def load_merge_history() -> dict:
    if os.path.exists(MERGE_HISTORY_FILE):
        try:
            with open(MERGE_HISTORY_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_merge_history(history: dict):
    with open(MERGE_HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=2)


def get_merged_status(output_dir: str) -> dict:
    """Authoritative aNum -> {merged, filename, timestamp, label, manual} map.
    merged_donors.json is the source of truth. Only if the file has *never*
    been created (first run after upgrading) is it seeded once from the
    legacy Merge_* folder scan, so existing merge history isn't lost. Once
    the file exists — even if intentionally emptied by a quarterly reset —
    it is never re-derived from the folder scan again."""
    if not os.path.exists(MERGE_HISTORY_FILE):
        history = {}
        if output_dir:
            legacy_filenames = get_previously_merged(output_dir)
            if legacy_filenames:
                now_iso = datetime.now().isoformat()
                for donor in state["donors"].values():
                    if get_donor_filename(donor) in legacy_filenames:
                        history[donor.aNum] = {
                            "merged": True,
                            "filename": get_donor_filename(donor),
                            "timestamp": now_iso,
                            "label": "Migrated from existing output folders",
                            "manual": False,
                        }
        save_merge_history(history)
        return history
    return load_merge_history()


def create_merge_batch(output_dir: str, donors_to_merge: list[Donor], label: str = "") -> dict:
    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    folder_name = f"Merge_{ts}"
    folder_path = os.path.join(output_dir, folder_name)
    os.makedirs(folder_path, exist_ok=True)

    merged_files = []
    errors = []
    history = load_merge_history()
    now_iso = datetime.now().isoformat()
    for donor in donors_to_merge:
        try:
            files = list(dict.fromkeys(donor.getFiles()))
            if not files:
                continue
            filename = get_donor_filename(donor)
            path = os.path.join(folder_path, filename)
            merge(files, path)
            merged_files.append(filename)
            history[donor.aNum] = {
                "merged": True,
                "filename": filename,
                "timestamp": now_iso,
                "label": label,
                "manual": False,
            }
        except Exception as e:
            errors.append({"aNum": donor.aNum, "eNum": donor.eNum, "error": str(e)})

    manifest = {
        "timestamp": now_iso,
        "label": label,
        "count": len(merged_files),
        "files": merged_files,
        "errors": errors,
    }
    with open(os.path.join(folder_path, "manifest.json"), "w") as f:
        json.dump(manifest, f, indent=2)

    if merged_files:
        save_merge_history(history)

    return {"folder": folder_name, "path": folder_path, "merged": len(merged_files), "errors": errors}


# ── Routes: Merge ─────────────────────────────────────────────────────────────

@app.route("/api/merge/individual/<aNum>", methods=["POST"])
def merge_individual(aNum):
    try:
        donor = state["donors"].get(aNum)
        if not donor:
            return jsonify({"ok": False, "error": "Donor not found"}), 404

        output_dir = state["config"]["outputDir"]
        if not output_dir:
            return jsonify({"ok": False, "error": "Output directory not configured"}), 400
        os.makedirs(output_dir, exist_ok=True)

        files = list(dict.fromkeys(donor.getFiles()))
        if not files:
            return jsonify({"ok": False, "error": "Donor has no files to merge"}), 400

        filename = get_donor_filename(donor)
        path = os.path.join(output_dir, filename)
        merge(files, path)

        history = load_merge_history()
        history[donor.aNum] = {
            "merged": True,
            "filename": filename,
            "timestamp": datetime.now().isoformat(),
            "label": "Individual merge",
            "manual": False,
        }
        save_merge_history(history)

        log(f"Merged individual: {donor.eName} -> {filename}")
        return jsonify({"ok": True, "file": filename, "path": path, "fileCount": len(files)})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/merge/ready", methods=["POST"])
def merge_ready():
    try:
        output_dir = state["config"]["outputDir"]
        if not output_dir:
            return jsonify({"ok": False, "error": "Output directory not configured"}), 400
        os.makedirs(output_dir, exist_ok=True)

        ready = [d for d in state["donors"].values() if d.ready()]
        data = request.json or {}
        fresh_only = data.get("freshOnly", True)

        if fresh_only:
            merged_status = get_merged_status(output_dir)
            targets = [d for d in ready if not merged_status.get(d.aNum, {}).get("merged")]
        else:
            targets = ready

        if not targets:
            return jsonify({"ok": True, "merged": 0, "message": "No fresh donors to merge"})

        result = create_merge_batch(output_dir, targets, label="Ready donors")
        if "merge" not in state["steps_completed"]:
            state["steps_completed"].append("merge")
        log(f"Batch merge (ready, fresh={fresh_only}): {result['merged']} -> {result['folder']}")
        return jsonify({"ok": True, **result})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/merge/all", methods=["POST"])
def merge_all():
    try:
        output_dir = state["config"]["outputDir"]
        if not output_dir:
            return jsonify({"ok": False, "error": "Output directory not configured"}), 400
        os.makedirs(output_dir, exist_ok=True)

        all_donors = list(state["donors"].values())
        data = request.json or {}
        fresh_only = data.get("freshOnly", True)

        if fresh_only:
            merged_status = get_merged_status(output_dir)
            targets = [d for d in all_donors if not merged_status.get(d.aNum, {}).get("merged")]
        else:
            targets = all_donors

        targets = [d for d in targets if d.getFiles()]

        if not targets:
            return jsonify({"ok": True, "merged": 0, "message": "No donors to merge"})

        result = create_merge_batch(output_dir, targets, label="All donors")
        log(f"Batch merge (all, fresh={fresh_only}): {result['merged']} -> {result['folder']}")
        return jsonify({"ok": True, **result})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/merge/history", methods=["GET"])
def merge_history():
    output_dir = state["config"]["outputDir"]
    if not output_dir or not os.path.isdir(output_dir):
        return jsonify([])

    batches = []
    for folder_name in get_merge_folders(output_dir):
        folder_path = os.path.join(output_dir, folder_name)
        manifest_path = os.path.join(folder_path, "manifest.json")
        if os.path.exists(manifest_path):
            try:
                with open(manifest_path, "r") as f:
                    manifest = json.load(f)
                manifest["folder"] = folder_name
                batches.append(manifest)
            except Exception:
                pass
        else:
            pdfs = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]
            batches.append({
                "folder": folder_name, "count": len(pdfs),
                "files": pdfs, "label": "",
                "timestamp": folder_name.replace("Merge_", "").replace("_", " "),
            })
    return jsonify(list(reversed(batches)))


@app.route("/api/donors/<aNum>/merged", methods=["POST"])
def set_donor_merged(aNum):
    try:
        donor = state["donors"].get(aNum)
        if not donor:
            return jsonify({"ok": False, "error": "Donor not found"}), 404

        data = request.json or {}
        merged = bool(data.get("merged", True))

        history = load_merge_history()
        existing = history.get(aNum, {})
        history[aNum] = {
            "merged": merged,
            "filename": existing.get("filename", get_donor_filename(donor)),
            "timestamp": datetime.now().isoformat(),
            "label": existing.get("label", "Manual override"),
            "manual": True,
        }
        save_merge_history(history)

        log(f"{'Marked' if merged else 'Unmarked'} as merged: {donor.eName} (A#{aNum})")
        return jsonify({"ok": True, "donor": serialize_donor(donor, history)})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


@app.route("/api/merge-history/reset", methods=["POST"])
def reset_merge_history():
    try:
        save_merge_history({})
        if "merge" in state["steps_completed"]:
            state["steps_completed"].remove("merge")
        log("Merge history reset for new quarter")
        return jsonify({"ok": True})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"ok": False, "error": str(e)}), 500


# ── Routes: Donors / Issues ──────────────────────────────────────────────────

@app.route("/api/donors", methods=["GET"])
def list_donors():
    filter_status = request.args.get("status", "all")
    search = request.args.get("search", "").lower()
    sort_by = request.args.get("sort", "eName")
    page = int(request.args.get("page", 1))
    per_page = int(request.args.get("per_page", 50))

    merge_history = get_merged_status(state["config"]["outputDir"])

    candidates = []
    for d in state["donors"].values():
        if filter_status != "all":
            is_ready = d.ready()
            is_merged = merge_history.get(d.aNum, {}).get("merged", False)
            if filter_status == "ready" and not is_ready:
                continue
            if filter_status == "not_ready" and is_ready:
                continue
            if filter_status == "merged" and not is_merged:
                continue
            if filter_status == "not_merged" and is_merged:
                continue
        if search:
            haystack = f"{d.eName} {d.eNum} {d.aNum} {d.email} ".lower()
            haystack += " ".join(p.pName.lower() for p in d.preachers.values())
            if search not in haystack:
                continue
        candidates.append(d)

    if sort_by == "status":
        candidates.sort(key=lambda d: (d.ready(), d.eName))
    elif sort_by == "eNum":
        candidates.sort(key=lambda d: d.eNum)
    elif sort_by == "aNum":
        candidates.sort(key=lambda d: d.aNum)
    else:
        candidates.sort(key=lambda d: d.eName)

    total = len(candidates)
    start = (page - 1) * per_page
    page_slice = candidates[start:start + per_page]

    return jsonify({
        "donors": [serialize_donor(d, merge_history) for d in page_slice],
        "total": total,
        "page": page,
        "perPage": per_page,
        "totalPages": max(1, (total + per_page - 1) // per_page),
    })


@app.route("/api/donors/<aNum>", methods=["GET"])
def get_donor(aNum):
    if aNum in state["donors"]:
        merge_history = get_merged_status(state["config"]["outputDir"])
        return jsonify(serialize_donor(state["donors"][aNum], merge_history))
    return jsonify({"error": "Donor not found"}), 404


@app.route("/api/issues", methods=["GET"])
def get_issues():
    issues = {"missingCoverLetters": [], "missingReports": [], "missingTYNotes": [], "preachersWithIssues": []}
    preachersWithIssues = dict()
    for d in state["donors"].values():
        if d.coverLetter == "":
            issues["missingCoverLetters"].append({"eNum": d.eNum, "eName": d.eName, "aNum": d.aNum})
        for p in d.preachers.values():
            if (p.noteNecessary and p.note == "") or p.report == "":
                preachersWithIssues[p.pNum] = {
                    "pNum": p.pNum.replace(",",""), 
                    "pName": p.pName.replace(",",""), "noteMissing": (p.noteNecessary and p.note == ""), "reportMissing": (p.report == "")}
                
                if p.report == "":
                    issues["missingReports"].append({
                        "eNum": d.eNum, "eName": d.eName, "aNum": d.aNum,
                        "pNum": p.pNum, "pName": p.pName,
                    })
                if p.noteNecessary and p.note == "":
                    issues["missingTYNotes"].append({
                        "eNum": d.eNum, "eName": d.eName, "aNum": d.aNum,
                        "pNum": p.pNum, "pName": p.pName,
                    })
    issues["preachersWithIssues"] = sorted(preachersWithIssues.values(), key=lambda x: x["pNum"])
    return jsonify(issues)


@app.route("/api/logs", methods=["GET"])
def get_logs():
    return jsonify(state["logs"])


@app.route("/api/reset", methods=["POST"])
def reset():
    state["donors"] = {}
    state["loaded"] = False
    state["steps_completed"] = []
    state["logs"] = []
    return jsonify({"ok": True})
def open_browser():
    webbrowser.open_new("http://127.0.0.1:5000/")

if __name__ == "__main__":
    print("\n  Gospelink Report Manager")
    print("  http://localhost:5000\n")
    is_frozen = getattr(sys, 'frozen', False)
    debug = not is_frozen
    if not debug or not os.environ.get("WERKZEUG_RUN_MAIN"):
        threading.Timer(1.5, open_browser).start()
    app.run(debug=debug, use_reloader=debug, port=5000)
