from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import pandas as pd
import os
import uuid
import time
from pathlib import Path
from datetime import datetime, timezone
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET', 'devsecret')

ALLOWED_EXTENSIONS = {'xls', 'xlsx', 'csv'}
TMP_DIR = Path(app.root_path) / 'tmp'
TMP_DIR.mkdir(exist_ok=True)

# Cleanup configuration
TMP_RETENTION_SECONDS = int(os.environ.get('TMP_RETENTION_SECONDS', str(6 * 60 * 60)))  # default: 6 hours
TMP_CLEANUP_INTERVAL_SECONDS = int(os.environ.get('TMP_CLEANUP_INTERVAL_SECONDS', str(30 * 60)))  # default: 30 minutes
_LAST_TMP_CLEANUP = 0.0


def allowed_file(filename: str) -> bool:
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def cleanup_tmp_dir(now: float | None = None) -> int:
    """Delete files in TMP_DIR older than TMP_RETENTION_SECONDS.

    Returns number of files deleted.
    """
    ts_now = now if now is not None else time.time()
    deleted = 0
    try:
        for p in TMP_DIR.glob('merged-*.xlsx'):
            try:
                mtime = p.stat().st_mtime
            except FileNotFoundError:
                continue
            if ts_now - mtime > TMP_RETENTION_SECONDS:
                try:
                    p.unlink(missing_ok=True)
                    deleted += 1
                except Exception as e:
                    # Non-fatal; log and continue
                    app.logger.debug(f"Could not delete tmp file {p}: {e}")
    except Exception as e:
        app.logger.warning(f"tmp cleanup scan failed: {e}")
    if deleted:
        app.logger.info(f"tmp cleanup removed {deleted} old file(s)")
    return deleted


@app.before_request
def _maybe_cleanup_tmp():
    global _LAST_TMP_CLEANUP
    now = time.time()
    if now - _LAST_TMP_CLEANUP >= TMP_CLEANUP_INTERVAL_SECONDS:
        _LAST_TMP_CLEANUP = now
        cleanup_tmp_dir(now)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Validate files
        if 'files' not in request.files:
            flash('No files part')
            return redirect(request.url)
        files = request.files.getlist('files')
        if not files or all(f.filename == '' for f in files):
            flash('No selected files')
            return redirect(request.url)

        # Parse per-file labels and custom column title
        raw_labels = request.form.get('group_texts', '')
        labels = [l.strip() for l in raw_labels.splitlines()] if raw_labels else []
        group_col_title = (request.form.get('group_col_title') or '').strip() or 'Group'
        use_group_col = any(bool(l) for l in labels)

        dfs = []
        merge_ranges = []  # tuples of (start_idx, end_idx) for data rows per file in merged DF (0-based)
        running_total = 0

        for i, file in enumerate(files):
            if not (file and allowed_file(file.filename)):
                flash(f'File not allowed: {getattr(file, "filename", "")}')
                return redirect(request.url)

            filename = secure_filename(file.filename)
            try:
                if filename.lower().endswith('.csv'):
                    df = pd.read_csv(file)
                else:
                    df = pd.read_excel(file)
            except Exception as e:
                flash(f'Error reading {filename}: {e}')
                return redirect(request.url)

            n_data = len(df)
            original_cols = list(df.columns)

            if n_data > 0:
                # Build Average row using original columns
                avg_row = {original_cols[0]: 'Average'}
                for c in original_cols[1:]:
                    numeric = pd.to_numeric(df[c], errors='coerce')
                    mean_val = numeric.mean()
                    avg_row[c] = '' if pd.isna(mean_val) else mean_val
                avg_df = pd.DataFrame([avg_row], columns=original_cols)

                # Insert group column if needed (always insert to keep columns aligned across files)
                if use_group_col:
                    label = labels[i] if i < len(labels) else ''
                    df.insert(0, group_col_title, label)
                    avg_df.insert(0, group_col_title, '')
                    cols_after = [group_col_title] + original_cols
                else:
                    cols_after = original_cols

                # Record merge range for this file's data rows if it has a label
                if use_group_col and (labels[i] if i < len(labels) else '') and n_data > 0:
                    start_idx = running_total
                    end_idx = running_total + n_data - 1
                    merge_ranges.append((start_idx, end_idx))

                # Append Average and optional empty row (not after last file)
                if i < len(files) - 1:
                    empty_row = {c: '' for c in cols_after}
                    empty_df = pd.DataFrame([empty_row], columns=cols_after)
                    df = pd.concat([df, avg_df, empty_df], ignore_index=True)
                else:
                    df = pd.concat([df, avg_df], ignore_index=True)

                running_total += len(df)

            dfs.append(df)

        # Merge all
        try:
            merged = pd.concat(dfs, ignore_index=True)
        except Exception as e:
            flash(f'Error merging files: {e}')
            return redirect(request.url)

        # Write to Excel
        unique_name = f'merged-{uuid.uuid4().hex}.xlsx'
        out_path = TMP_DIR / unique_name
        try:
            with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
                merged.to_excel(writer, index=False, sheet_name='Merged')
        except Exception as e:
            flash(f'Error writing output file: {e}')
            return redirect(request.url)

        # Merge first column cells for labeled ranges
        if use_group_col and merge_ranges:
            try:
                wb = load_workbook(out_path)
                ws = wb['Merged']
                for (start_idx, end_idx) in merge_ranges:
                    excel_start = start_idx + 2  # header row + 1-based indexing
                    excel_end = end_idx + 2
                    if excel_start <= excel_end:
                        ws.merge_cells(start_row=excel_start, start_column=1, end_row=excel_end, end_column=1)
                wb.save(out_path)
            except Exception as e:
                flash(f'Warning: could not merge group cells in Excel: {e}')

        # Render preview
        merged_html = merged.to_html(classes='table table-striped table-sm', index=False, escape=False)
        return render_template('index.html', merged_table=merged_html, download_filename=unique_name, num_rows=len(merged), num_cols=len(merged.columns))

    # GET
    return render_template('index.html')


@app.route('/download/<path:filename>')
def download(filename: str):
    # sanitize and only allow files created by this app (merged-*.xlsx)
    if not filename.startswith('merged-') or not filename.endswith('.xlsx'):
        flash('Invalid download request')
        return redirect(url_for('index'))
    file_path = TMP_DIR / filename
    if not file_path.exists():
        flash('File not found or expired')
        return redirect(url_for('index'))
    try:
        return send_file(
            file_path,
            as_attachment=True,
            download_name='merged.xlsx'
        )
    except Exception as e:
        flash(f'Error sending file: {e}')
        return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)
