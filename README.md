# Excel Merger

Small Flask app to upload multiple Excel/CSV files and download a merged Excel file.

Requirements

-   Python 3.8+

Install

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Run

```bash
export FLASK_APP=app.py
export FLASK_ENV=development
flask run
```

Then open http://127.0.0.1:5000

Notes

-   Output modes:
    -   Single sheet (default): concatenates all files row-wise into one sheet named "Merged". You can optionally add a Group column with per-file labels; the first column cells will be merged per labeled group in the Excel output.
    -   Separate sheets: writes each uploaded file to its own sheet. You can optionally provide sheet names (one per line, in the same order as the files). If left blank, sheet names fall back to the uploaded filename (without extension). Names are sanitized to Excel rules (max 31 chars; cannot contain `: \/ ? * [ ]`). Duplicates are auto-suffixed with " (2)", " (3)", etc.
-   It reads the first sheet for Excel files.
-   For large files or different merge logic (join on keys), further development is needed.
