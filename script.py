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


# ✅ ONLY CHANGE IS HERE (decimal-safe)
def build_qr_text(row, selected_cols):
    """
    RULE:
    - Value present → keep exact formatting (1.00 → 1.00)
    - Value missing/empty → |
    - Separator → single space
    """
    parts = []

    for col in selected_cols:
        val = row.get(col, "")

        if pd.isna(val) or str(val).strip() == "":
            parts.append("|")
        else:
            # Preserve decimal values exactly
            if isinstance(val, float):
                text = format(val, ".2f")
            else:
                text = str(val).strip()

            parts.append(text)

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

    df = pd.read_excel(file_path)
    df_norm = df.copy()
    df_norm.columns = [normalize(c) for c in df_norm.columns]
    selected_norm = [normalize(c) for c in selected_cols]

    df["QR_Image_Path"] = ""
    df["QR"] = ""

    for i, row in df_norm.iterrows():
        qr_text = build_qr_text(row, selected_norm)
        img_name = f"qr_{i+1}.png"
        img_path = os.path.join(QR_FOLDER, img_name)
        generate_qr(qr_text, img_path)
        df.at[i, "QR_Image_Path"] = img_path

    output_path = os.path.join(OUTPUT_FOLDER, "output_with_qr.xlsx")
    df.to_excel(output_path, index=False)

    # Insert QR images
    wb = load_workbook(output_path)
    ws = wb.active

    headers = {normalize(c.value): i + 1 for i, c in enumerate(ws[1])}
    qr_col = headers["qr"]
    path_col = headers["qr_image_path"]

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
    output_path = os.path.join(BASE_DIR, "outputs", "output_with_qr.xlsx")
    return send_file(output_path, as_attachment=True)


# ---------------- RUN LOCAL ----------------
if __name__ == "__main__":
    app.run(debug=True)
