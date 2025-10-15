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

-   The app concatenates all files row-wise into a single sheet named "Merged".
-   It reads the first sheet for Excel files.
-   For large files or different merge logic (join on keys), further development is needed.
