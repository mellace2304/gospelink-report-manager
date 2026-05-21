# gospelink-report-manager

An automated quarterly report manager for generating consolidated PDF reports. This application streamlines the process of merging donor information, preacher reports, cover letters, and related documents into organized quarterly PDF packages.

## Overview

The Gospelink Report Manager is a desktop application that automates the workflow for creating quarterly reports. It features a modern web-based UI (built with React) paired with a Flask backend that handles PDF generation, document merging, and report compilation.

### Key Features

- **Multi-step workflow** — Guided process for report generation with progress tracking
- **Donor & Preacher Management** — Organize and manage donor information and preacher reports
- **Automated PDF Merging** — Combine multiple documents into cohesive quarterly reports
- **Cover Letter Generation** — Create and merge cover letters with reports
- **Document Processing** — Support for PDF, Word (.docx), and Excel (.xlsx) files
- **Configuration Management** — Save and manage project settings
- **Standalone Executable** — Bundle as a single .exe file with PyInstaller for easy distribution

## Project Structure

```
.
├── server.py              # Flask backend server with REST API endpoints
├── Merge.py               # Core report generation logic (Donor, Preacher classes)
├── static/
│   └── index.html         # React-based frontend UI
├── requirements.txt       # Python dependencies
├── test.ipynb             # Jupyter notebook for testing/development
├── gospelink_config.json  # Configuration settings (generated at runtime)
└── README.md              # This file
```

## Architecture

### Backend (Flask Server)

**server.py** exposes REST API endpoints that wrap the core logic:
- Configuration management endpoints
- File upload and processing
- Report generation and merging
- PDF creation

### Core Logic (Merge.py)

Contains the main business logic:
- `Donor` class — Represents donor information
- `Preacher` class — Represents preacher reports
- Report merging functions — Combine documents into consolidated reports
- PDF utilities — Validation and processing of PDF files
- Word document handling — Template filling and conversion

### Frontend (React UI)

Modern single-page application with:
- Sidebar navigation with step tracking
- Configuration forms
- Progress indicators
- File management interface

## Installation

### Requirements

- Python 3.7+
- Windows (for full `pywin32` support), or Linux/Mac with adaptations

### Setup

1. Clone the repository:
   ```bash
   git clone <repository-url>
   cd gospelink-report-manager
   ```

2. Create a virtual environment:
   ```bash
   python -m venv .venv
   .venv\Scripts\activate  # Windows
   source .venv/bin/activate  # Linux/Mac
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Development

Run the Flask server:
```bash
python server.py
```

The application will open in your default browser at `http://localhost:5000`

### Building Standalone Executable

Use PyInstaller to create a single executable file (Windows):
```bash
pyinstaller --onefile --name Gospelink --add-data "static;static" --hidden-import win32com --hidden-import win32com.client --collect-all docx2pdf --collect-all fitz --icon=icon.ico server.py
```

This creates a standalone `.exe` file that includes all dependencies and doesn't require Python to be installed.

## Dependencies

| Package | Purpose |
|---------|---------|
| **flask** | Web framework for REST API |
| **flask-cors** | Cross-Origin Resource Sharing support |
| **pandas** | Data processing and manipulation |
| **PyMuPDF** | PDF reading and processing (fitz) |
| **python-docx** | Word document (.docx) handling |
| **openpyxl** | Excel file (.xlsx) support |
| **pywin32** | Windows COM integration for Office apps |

## Workflow

1. **Configuration** — Set up paths for input files and output directories
2. **Load Data** — Import donor lists and preacher information
3. **Process Reports** — Validate and prepare individual reports
4. **Merge Documents** — Combine reports, cover letters, and supporting documents
5. **Generate PDFs** — Create final consolidated quarterly reports
6. **Output** — Save completed reports to designated directory

## Configuration

Settings are stored in `gospelink_config.json`:
- Input file paths (cover sheet, extra gift file, reports directory)
- Output directory path
- Template file path
- Cover letter directory

## Development Notes

- The application supports hot-reloading in development mode
- PDF validation is performed before processing to catch corrupted files early
- Configuration history is tracked in `gospelink_config_history.json`
- The UI communicates with the backend via REST API calls

## License

See [LICENSE](LICENSE) file for details.

---

For issues or questions, please refer to the project documentation or contact the maintainers.
