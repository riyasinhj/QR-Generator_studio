import os
import pandas as pd
import qrcode
from flask import Flask, render_template, request, send_file

from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# ---------------- FLASK APP ----------------
app = Flask(__name__)

# ---------------- PATHS ----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")
QR_FOLDER = os.path.join(BASE_DIR, "qr_codes")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(QR_FOLDER, exist_ok=True)

# ---------------- QR CONFIG ----------------
QR_MM = 9
PIXELS_PER_MM = 96 / 25.4
QR_PX = int(QR_MM * PIXELS_PER_MM)
ROW_HEIGHT = 30
QR_COL_WIDTH = 18
PATH_COL_WIDTH = 95
# ------------------------------------------


def normalize(text):
    return (
        str(text)
        .replace("\xa0", " ")
        .replace("\u200b", "")
        .strip()
        .lower()
    )


def generate_qr(data, path):
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=10,
        border=1,
    )
    qr.add_data(data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    img.save(path)


# ✅ READ EXACT DISPLAY VALUE FROM EXCEL (CRITICAL FIX)
def excel_display_value(ws, row_idx, col_idx):
    cell = ws.cell(row=row_idx, column=col_idx)
    val = cell.value

    if val is None:
        return ""

    # Keep numbers EXACTLY as Excel shows
    if isinstance(val, (int, float)):
        fmt = cell.number_format
        try:
            if "," in fmt:
                return format(val, ",")
            if "." in fmt:
                decimals = fmt.split(".")[-1].count("0")
                return f"{val:.{decimals}f}"
        except:
            pass
        return str(val)

    return str(val)


def build_qr_text(ws, row_idx, selected_col_indexes):
    parts = []
    for col_idx in selected_col_indexes:
        value = excel_display_value(ws, row_idx, col_idx)
        parts.append(value if value != "" else "|")
    return " ".join(parts)


# ---------------- ROUTES ----------------

@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        file = request.files["excel"]
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        df = pd.read_excel(file_path)
        columns = list(df.columns)

        return render_template(
            "columns.html",
            columns=columns,
            file_path=file_path
        )

    return render_template("upload.html")


@app.route("/generate", methods=["POST"])
def generate():
    file_path = request.form["file_path"]
    selected_cols = request.form.getlist("columns")

    # Load workbook for exact formatting
    wb_input = load_workbook(file_path, data_only=False)
    ws_input = wb_input.active

    headers = [normalize(c.value) for c in ws_input[1]]
    selected_indexes = [
        headers.index(normalize(c)) + 1 for c in selected_cols
    ]

    df = pd.read_excel(file_path)

    df["QR_Image_Path"] = ""
    df["QR"] = ""

    for r in range(2, ws_input.max_row + 1):
        qr_text = build_qr_text(ws_input, r, selected_indexes)
        img_name = f"qr_{r-1}.png"
        img_path = os.path.join(QR_FOLDER, img_name)
        generate_qr(qr_text, img_path)
        df.at[r-2, "QR_Image_Path"] = img_path

    output_path = os.path.join(OUTPUT_FOLDER, "output_with_qr.xlsx")
    df.to_excel(output_path, index=False)

    # Insert QR images
    wb = load_workbook(output_path)
    ws = wb.active

    headers_out = {normalize(c.value): i + 1 for i, c in enumerate(ws[1])}
    qr_col = headers_out["qr"]
    path_col = headers_out["qr_image_path"]

    for r in range(2, ws.max_row + 1):
        img_path = ws.cell(r, path_col).value
        if not img_path or not os.path.exists(img_path):
            continue

        img = Image(img_path)
        img.width = QR_PX
        img.height = QR_PX
        ws.add_image(img, ws.cell(r, qr_col).coordinate)
        ws.row_dimensions[r].height = ROW_HEIGHT
        ws.cell(r, qr_col).alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions[get_column_letter(qr_col)].width = QR_COL_WIDTH
    ws.column_dimensions[get_column_letter(path_col)].width = PATH_COL_WIDTH

    wb.save(output_path)

    return render_template("download.html")


@app.route("/download")
def download_file():
    output_path = os.path.join(OUTPUT_FOLDER, "output_with_qr.xlsx")
    return send_file(output_path, as_attachment=True)


# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(debug=True)
