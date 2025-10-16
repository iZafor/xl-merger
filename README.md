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
    -   Single sheet (default): concatenates all files row-wise into one sheet named "Merged".
        -   Column alignment option (single-sheet only):
            -   Match by position (default): aligns to the first non-empty file's columns by index, truncating extras and padding missing columns with blanks.
            -   Match by column names (union): includes all columns seen across files; missing values are left blank.
        -   Average row: an "Average" summary row is appended after each file's data block. In position mode, averages align by column index of the first file.
        -   Optional Group column with per-file labels; the first column cells will be merged per labeled group in the Excel output.
    -   Separate sheets: writes each uploaded file to its own sheet. You can optionally provide sheet names (one per line, in the same order as the files). If left blank, sheet names fall back to the uploaded filename (without extension). Names are sanitized to Excel rules (max 31 chars; cannot contain `: \/ ? * [ ]`). Duplicates are auto-suffixed with " (2)", " (3)", etc. No Average rows are added in separate-sheets mode.
-   Freeze panes toggle: a checkbox lets you freeze the first row and first column in outputs (default ON). Uncheck it to disable freezing. Applies to both single-sheet and separate-sheets outputs.
-   Formatting and appearance:
    -   All cells are centered vertically and horizontally.
    -   Columns are auto-sized to fit content (with reasonable min/max width bounds).
    -   In single-sheet output, Average rows and the merged Group header cells are highlighted light yellow.
    -   In separate-sheets mode, for .xlsx uploads the app preserves the source sheet's merged cells and most styles (font, fill, border, alignment, number format). CSV or other formats are written from DataFrame values without source formatting.
-   It reads the first (active) sheet for Excel files.
-   For large files or different merge logic (join on keys), further development is needed.
