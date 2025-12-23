import os
import pandas as pd
import qrcode
from flask import Flask, render_template, request, send_file, send_from_directory

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

# 🔴 IMPORTANT: Your Render public URL
PUBLIC_BASE_URL = "https://qr-generator-studio-vf7g.onrender.com"

# ---------------- QR CONFIG ----------------
QR_MM = 9
PIXELS_PER_MM = 96 / 25.4
QR_PX = int(QR_MM * PIXELS_PER_MM)
ROW_HEIGHT = 30
QR_COL_WIDTH = 18
PATH_COL_WIDTH = 95
# ------------------------------------------


def normalize(text):
    return str(text).strip().lower()


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


def preserve_exact_value(val):
    """
    VERY IMPORTANT:
    - Keeps 1.00 as '1.00'
    - Keeps 95,000 as '95,000'
    - Keeps text EXACTLY as in Excel
    """
    if pd.isna(val):
        return "|"

    return str(val).strip()


def build_qr_text(row, selected_cols):
    parts = []
    for col in selected_cols:
        parts.append(preserve_exact_value(row[col]))
    return " ".join(parts)


# ---------------- ROUTES ----------------

@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        file = request.files["excel"]
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        # IMPORTANT: read as STRING to preserve format
        df = pd.read_excel(file_path, dtype=str)
        columns = list(df.columns)

        return render_template(
            "columns.html",
            columns=columns,
            file_path=file.filename
        )

    return render_template("upload.html")


@app.route("/generate", methods=["POST"])
def generate():
    file_name = request.form["file_path"]
    selected_cols = request.form.getlist("columns")

    file_path = os.path.join(UPLOAD_FOLDER, file_name)

    # READ AS STRING (CRITICAL FIX)
    df = pd.read_excel(file_path, dtype=str)

    df["QR_Image_URL"] = ""
    df["QR"] = ""

    for i, row in df.iterrows():
        qr_text = build_qr_text(row, selected_cols)

        img_name = f"qr_{i+1}.png"
        img_path = os.path.join(QR_FOLDER, img_name)

        generate_qr(qr_text, img_path)

        # ✅ PUBLIC URL (MAIL MERGE SAFE)
        qr_url = f"{PUBLIC_BASE_URL}/qr/{img_name}"
        df.at[i, "QR_Image_URL"] = qr_url
        df.at[i, "QR"] = qr_text

    output_path = os.path.join(OUTPUT_FOLDER, "output_with_qr.xlsx")
    df.to_excel(output_path, index=False)

    # ---------- INSERT QR IMAGE INTO EXCEL ----------
    wb = load_workbook(output_path)
    ws = wb.active

    headers = {normalize(c.value): i + 1 for i, c in enumerate(ws[1])}
    qr_col = headers["qr"]

    for r in range(2, ws.max_row + 1):
        img_name = f"qr_{r-1}.png"
        img_path = os.path.join(QR_FOLDER, img_name)

        if not os.path.exists(img_path):
            continue

        img = Image(img_path)
        img.width = QR_PX
        img.height = QR_PX

        ws.add_image(img, ws.cell(r, qr_col).coordinate)
        ws.row_dimensions[r].height = ROW_HEIGHT
        ws.cell(r, qr_col).alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions[get_column_letter(qr_col)].width = QR_COL_WIDTH
    wb.save(output_path)

    return render_template("download.html")


# ✅ PUBLIC QR IMAGE ROUTE (THIS FIXES YOUR PATH ISSUE)
@app.route("/qr/<filename>")
def serve_qr(filename):
    return send_from_directory(QR_FOLDER, filename)


@app.route("/download")
def download_file():
    output_path = os.path.join(OUTPUT_FOLDER, "output_with_qr.xlsx")
    return send_file(output_path, as_attachment=True)


# ---------------- LOCAL RUN ----------------
if __name__ == "__main__":
    app.run(debug=True)
