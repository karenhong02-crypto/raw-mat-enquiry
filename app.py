"""
app.py  –  Flask web front-end for generate_enquiry_raw_mat
Run:  python app.py
Open: http://localhost:5000
"""

import io
import re
import sys
import tempfile
import traceback
from copy import copy
from pathlib import Path

from flask import Flask, render_template, request, send_file, jsonify

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Missing dependency. Run:  pip install openpyxl")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024   # 32 MB upload limit

ALLOWED_EXT = {".xlsx", ".xlsm", ".xls"}


# ---------------------------------------------------------------------------
# Helpers  (same as generate_enquiry_raw_mat.py)
# ---------------------------------------------------------------------------

def allowed(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXT


def parse_size(s):
    """Parse 'T x W x L' string → (T, W, L) floats."""
    if not s:
        return None, None, None
    s = re.sub(r"[()]", "", str(s))
    parts = re.split(r"\s*x\s*", s.strip(), flags=re.IGNORECASE)
    if len(parts) == 3:
        try:
            return float(parts[0]), float(parts[1]), float(parts[2])
        except ValueError:
            pass
    return None, None, None


def primary_material(m):
    """Normalise material names.
    'AL5083/AL6061' → 'AL5083'
    'MS'            → 'MS (bandsaw)'
    """
    if m and "/" in str(m):
        m = str(m).split("/")[0].strip()
    else:
        m = str(m).strip() if m else ""
    if m.upper().startswith("MS"):
        return f"{m} (bandsaw)"
    return m


def capture_row(ws, row_num, max_col=16):
    saved = {}
    for c in range(1, max_col + 1):
        cell = ws.cell(row_num, c)
        saved[c] = {
            "value":         cell.value,
            "font":          copy(cell.font),
            "fill":          copy(cell.fill),
            "border":        copy(cell.border),
            "alignment":     copy(cell.alignment),
            "number_format": cell.number_format,
        }
    return saved


def restore_row(ws, row_num, saved):
    for c, props in saved.items():
        cell = ws.cell(row_num, c)
        cell.value         = props["value"]
        cell.font          = copy(props["font"])
        cell.fill          = copy(props["fill"])
        cell.border        = copy(props["border"])
        cell.alignment     = copy(props["alignment"])
        cell.number_format = props["number_format"]


def get_pmc_rows(ws_bom):
    rows = []
    for r in range(2, ws_bom.max_row + 1):
        col_o = ws_bom.cell(r, 15).value
        if col_o and str(col_o).strip().upper() == "PMC":
            rows.append({
                "afa":      str(ws_bom.cell(r, 1).value) if ws_bom.cell(r, 1).value else "",
                "qty":      ws_bom.cell(r, 5).value,
                "material": primary_material(ws_bom.cell(r, 6).value),
                "pmc_raw":  ws_bom.cell(r, 14).value,
            })
    return rows


def build_enquiry_bytes(bom_path: str, enq_path: str) -> tuple[bytes, dict]:
    """Run the full generation and return (xlsx_bytes, summary_dict)."""
    wb_bom = load_workbook(bom_path)
    wb_enq = load_workbook(enq_path)
    ws_bom = wb_bom.active
    ws_enq = wb_enq.active

    pmc_rows = get_pmc_rows(ws_bom)
    if not pmc_rows:
        raise ValueError("No PMC rows found in BOM file. Check that column O contains 'PMC'.")

    DATA_START  = 6
    FOOTER1_ROW = 62
    FOOTER2_ROW = 63

    footer1 = capture_row(ws_enq, FOOTER1_ROW)
    footer2 = capture_row(ws_enq, FOOTER2_ROW)
    ref_fmt  = capture_row(ws_enq, DATA_START)

    n_old = FOOTER1_ROW - DATA_START
    n_new = len(pmc_rows)

    if n_new > n_old:
        ws_enq.insert_rows(FOOTER1_ROW, n_new - n_old)
    elif n_new < n_old:
        surplus = n_old - n_new
        ws_enq.delete_rows(FOOTER1_ROW - surplus, surplus)

    for i, item in enumerate(pmc_rows):
        row = DATA_START + i
        t, w, l = parse_size(item["pmc_raw"])

        ws_enq.cell(row,  1).value = i + 1
        ws_enq.cell(row,  2).value = None
        ws_enq.cell(row,  3).value = item["afa"]
        ws_enq.cell(row,  4).value = item["material"]
        ws_enq.cell(row,  5).value = item["qty"]
        ws_enq.cell(row,  6).value = None
        ws_enq.cell(row,  7).value = item["pmc_raw"]
        ws_enq.cell(row,  8).value = t
        ws_enq.cell(row,  9).value = w
        ws_enq.cell(row, 10).value = l
        ws_enq.cell(row, 11).value = None
        ws_enq.cell(row, 12).value = t
        ws_enq.cell(row, 13).value = w
        ws_enq.cell(row, 14).value = l
        ws_enq.cell(row, 15).value = None
        ws_enq.cell(row, 16).value = f"=E{row}*O{row}"

        for c in range(1, 17):
            dst = ws_enq.cell(row, c)
            dst.font          = copy(ref_fmt[c]["font"])
            dst.fill          = copy(ref_fmt[c]["fill"])
            dst.border        = copy(ref_fmt[c]["border"])
            dst.alignment     = copy(ref_fmt[c]["alignment"])
            dst.number_format = ref_fmt[c]["number_format"]

    last_data_row = DATA_START + n_new - 1
    new_footer1   = last_data_row + 1
    new_footer2   = last_data_row + 2

    restore_row(ws_enq, new_footer1, footer1)
    restore_row(ws_enq, new_footer2, footer2)
    ws_enq.cell(new_footer1, 16).value = f"=SUM(P{DATA_START}:P{last_data_row})"

    buf = io.BytesIO()
    wb_enq.save(buf)
    buf.seek(0)

    summary = {
        "pmc_rows":   n_new,
        "data_start": DATA_START,
        "data_end":   last_data_row,
        "footer1":    new_footer1,
        "footer2":    new_footer2,
    }
    return buf.read(), summary


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    bom_file = request.files.get("bom")
    enq_file = request.files.get("enq")

    # ── Validate uploads ──────────────────────────────────────────────────
    errors = []
    if not bom_file or bom_file.filename == "":
        errors.append("BOM file is required.")
    elif not allowed(bom_file.filename):
        errors.append("BOM file must be an Excel file (.xlsx / .xlsm).")

    if not enq_file or enq_file.filename == "":
        errors.append("Enquiry template file is required.")
    elif not allowed(enq_file.filename):
        errors.append("Enquiry template must be an Excel file (.xlsx / .xlsm).")

    if errors:
        return jsonify({"error": " ".join(errors)}), 400

    # ── Save uploads to temp files ────────────────────────────────────────
    try:
        with (
            tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_bom,
            tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_enq,
        ):
            bom_file.save(tmp_bom.name)
            enq_file.save(tmp_enq.name)
            tmp_bom_path = tmp_bom.name
            tmp_enq_path = tmp_enq.name

        xlsx_bytes, summary = build_enquiry_bytes(tmp_bom_path, tmp_enq_path)

    except ValueError as e:
        return jsonify({"error": str(e)}), 422
    except Exception:
        return jsonify({"error": "Processing failed.\n" + traceback.format_exc()}), 500
    finally:
        Path(tmp_bom_path).unlink(missing_ok=True)
        Path(tmp_enq_path).unlink(missing_ok=True)

    # ── Stream the file back ──────────────────────────────────────────────
    buf = io.BytesIO(xlsx_bytes)
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name="Enquiry_Output.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5000)
