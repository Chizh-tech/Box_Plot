import os
import tempfile
import uuid

import numpy as np
import pandas as pd
from flask import Flask, jsonify, render_template, request

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = tempfile.mkdtemp()
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB

# In-memory store: file_id -> {path, filename, sheet_names}
uploaded_files: dict = {}

ALLOWED_EXTENSIONS = {".xlsx", ".xls"}


def _safe_float(value):
    """Convert a value to float, returning None on failure."""
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _cell_display(value):
    """Format a cell value for JSON serialisation."""
    if value is None:
        return ""
    if isinstance(value, float):
        if np.isnan(value):
            return ""
        return round(value, 6)
    return str(value)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    """Receive an Excel file, save it, and return sheet names."""
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "No file selected"}), 400

    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({"error": "Only .xlsx and .xls files are supported"}), 400

    fid = str(uuid.uuid4())
    save_path = os.path.join(app.config["UPLOAD_FOLDER"], fid + "_" + f.filename)
    f.save(save_path)

    try:
        xl = pd.ExcelFile(save_path)
        sheet_names = xl.sheet_names
    except Exception as exc:
        os.remove(save_path)
        return jsonify({"error": f"Cannot read Excel file: {exc}"}), 400

    uploaded_files[fid] = {"path": save_path, "filename": f.filename, "sheet_names": sheet_names}
    return jsonify({"file_id": fid, "filename": f.filename, "sheet_names": sheet_names})


@app.route("/api/preview", methods=["POST"])
def preview():
    """Return the first rows of a sheet for the user to inspect."""
    body = request.get_json(force=True)
    fid = body.get("file_id")
    sheet = body.get("sheet_name")

    if fid not in uploaded_files:
        return jsonify({"error": "File not found"}), 404

    try:
        df = pd.read_excel(uploaded_files[fid]["path"], sheet_name=sheet, header=None)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

    max_preview = 20
    rows = []
    for ri in range(min(max_preview, len(df))):
        rows.append([_cell_display(df.iloc[ri, ci]) for ci in range(len(df.columns))])

    return jsonify({"rows": rows, "total_rows": len(df), "total_cols": len(df.columns)})


@app.route("/api/dimensions", methods=["POST"])
def dimensions():
    """Extract dimension (feature) names from a configured sheet."""
    body = request.get_json(force=True)
    fid = body.get("file_id")
    sheet = body.get("sheet_name")
    cfg = body.get("config", {})
    fmt = cfg.get("format", "col_per_dim")

    if fid not in uploaded_files:
        return jsonify({"error": "File not found"}), 404

    try:
        df = pd.read_excel(uploaded_files[fid]["path"], sheet_name=sheet, header=None)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500

    dims = []
    if fmt == "col_per_dim":
        hrow = int(cfg.get("header_row", 0))
        skip_cols = int(cfg.get("skip_cols", 0))
        for ci in range(skip_cols, len(df.columns)):
            v = df.iloc[hrow, ci]
            name = str(v).strip() if pd.notna(v) else ""
            if name and name.lower() != "nan":
                dims.append(name)
    else:  # row_per_dim
        name_col = int(cfg.get("name_col", 0))
        first_row = int(cfg.get("first_data_row", 1))
        for ri in range(first_row, len(df)):
            v = df.iloc[ri, name_col]
            name = str(v).strip() if pd.notna(v) else ""
            if name and name.lower() != "nan":
                dims.append(name)

    return jsonify({"dimensions": dims})


@app.route("/api/plot_data", methods=["POST"])
def plot_data():
    """
    Build measurement + spec data for every requested file/stage.

    Request body::

        {
          "files": [
            {
              "file_id": "...",
              "stage_name": "Stage 1",
              "sheet_name": "Sheet1",
              "config": {
                "format": "col_per_dim",   // or "row_per_dim"

                // ---- col_per_dim options ----
                "header_row": 0,           // 0-based row index of dimension names
                "skip_cols": 1,            // leading columns to ignore
                "nominal_row": 1,          // row with nominal/target values (optional)
                "usl_row": 2,              // row with upper spec limit (optional)
                "lsl_row": 3,              // row with lower spec limit (optional)
                "upper_tol_row": null,     // row with upper tolerance (added to nominal)
                "lower_tol_row": null,     // row with lower tolerance (abs, subtracted)
                "data_start_row": 4,       // first row of actual measurements

                // ---- row_per_dim options ----
                "name_col": 0,
                "nominal_col": 1,
                "usl_col": null,
                "lsl_col": null,
                "upper_tol_col": 2,
                "lower_tol_col": 3,
                "data_start_col": 4,
                "first_data_row": 1
              }
            }
          ],
          "selected_dims": ["Dim1", "Dim2"]
        }
    """
    body = request.get_json(force=True)
    files_cfg = body.get("files", [])
    selected = set(body.get("selected_dims", []))

    result = []
    for fc in files_cfg:
        fid = fc.get("file_id")
        stage = fc.get("stage_name", "Stage")
        sheet = fc.get("sheet_name")
        cfg = fc.get("config", {})
        fmt = cfg.get("format", "col_per_dim")

        if fid not in uploaded_files:
            continue

        try:
            df = pd.read_excel(uploaded_files[fid]["path"], sheet_name=sheet, header=None)
        except Exception:
            continue

        dims = []

        if fmt == "col_per_dim":
            hrow = int(cfg.get("header_row", 0))
            skip_cols = int(cfg.get("skip_cols", 0))
            nominal_row = cfg.get("nominal_row")
            usl_row = cfg.get("usl_row")
            lsl_row = cfg.get("lsl_row")
            upper_tol_row = cfg.get("upper_tol_row")
            lower_tol_row = cfg.get("lower_tol_row")
            data_start = int(cfg.get("data_start_row", 1))

            def _row_val(row, col):
                if row is None:
                    return None
                r = int(row)
                if r < len(df):
                    return _safe_float(df.iloc[r, col])
                return None

            for ci in range(skip_cols, len(df.columns)):
                raw_name = df.iloc[hrow, ci]
                name = str(raw_name).strip() if pd.notna(raw_name) else ""
                if not name or name.lower() == "nan":
                    continue
                if selected and name not in selected:
                    continue

                nominal = _row_val(nominal_row, ci)
                usl = _row_val(usl_row, ci)
                lsl = _row_val(lsl_row, ci)
                upper_tol = _row_val(upper_tol_row, ci)
                lower_tol = _row_val(lower_tol_row, ci)

                if usl is None and upper_tol is not None and nominal is not None:
                    usl = nominal + upper_tol
                if lsl is None and lower_tol is not None and nominal is not None:
                    lsl = nominal - abs(lower_tol)

                meas = []
                for ri in range(data_start, len(df)):
                    v = _safe_float(df.iloc[ri, ci])
                    if v is not None:
                        meas.append(v)

                dims.append({"name": name, "nominal": nominal, "usl": usl, "lsl": lsl, "measurements": meas})

        else:  # row_per_dim
            name_col = int(cfg.get("name_col", 0))
            nominal_col = cfg.get("nominal_col")
            usl_col = cfg.get("usl_col")
            lsl_col = cfg.get("lsl_col")
            upper_tol_col = cfg.get("upper_tol_col")
            lower_tol_col = cfg.get("lower_tol_col")
            data_start_col = int(cfg.get("data_start_col", 4))
            first_row = int(cfg.get("first_data_row", 1))

            def _col_val(row_idx, col):
                if col is None:
                    return None
                c = int(col)
                if c < len(df.columns):
                    return _safe_float(df.iloc[row_idx, c])
                return None

            for ri in range(first_row, len(df)):
                raw_name = df.iloc[ri, name_col]
                name = str(raw_name).strip() if pd.notna(raw_name) else ""
                if not name or name.lower() == "nan":
                    continue
                if selected and name not in selected:
                    continue

                nominal = _col_val(ri, nominal_col)
                usl = _col_val(ri, usl_col)
                lsl = _col_val(ri, lsl_col)
                upper_tol = _col_val(ri, upper_tol_col)
                lower_tol = _col_val(ri, lower_tol_col)

                if usl is None and upper_tol is not None and nominal is not None:
                    usl = nominal + upper_tol
                if lsl is None and lower_tol is not None and nominal is not None:
                    lsl = nominal - abs(lower_tol)

                meas = []
                for ci in range(data_start_col, len(df.columns)):
                    v = _safe_float(df.iloc[ri, ci])
                    if v is not None:
                        meas.append(v)

                dims.append({"name": name, "nominal": nominal, "usl": usl, "lsl": lsl, "measurements": meas})

        result.append({"stage": stage, "dimensions": dims})

    return jsonify({"data": result})


if __name__ == "__main__":
    app.run(debug=True, port=5000)
