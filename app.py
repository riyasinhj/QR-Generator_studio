import os
import qrcode
from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import format_cell

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


def get_visible_cell_value(cell):
    """
    🔒 CRITICAL FUNCTION
    Returns EXACT text visible in Excel
    Keeps 1.00, 56,000, 00123 exactly as-is
    """
    if cell.value is None:
        return ""
    return format_cell(cell)


# ---------------- ROUTES ----------------

@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        file = request.files["excel"]
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        wb = load_workbook(file_path, data_only=False)
        ws = wb.active

        columns = [cell.value for cell in ws[1]]

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
    selected_norm = [normalize(c) for c in selected_cols]

    wb = load_workbook(file_path, data_only=False)
    ws = wb.active

    headers = {normalize(c.value): i + 1 for i, c in enumerate(ws[1])}

    # Add QR headers
    qr_col = ws.max_column + 1
    path_col = ws.max_column + 2
    ws.cell(1, qr_col).value = "QR"
    ws.cell(1, path_col).value = "QR_Image_Path"

    for r in range(2, ws.max_row + 1):
        parts = []

        for col in selected_norm:
            if col not in headers:
                parts.append("|")
            else:
                cell = ws.cell(r, headers[col])
                text = get_visible_cell_value(cell)
                parts.append(text if text else "|")

        qr_text = " ".join(parts)

        img_path = os.path.join(QR_FOLDER, f"qr_{r}.png")
        generate_qr(qr_text, img_path)

        ws.cell(r, path_col).value = img_path

        img = Image(img_path)
        img.width = QR_PX
        img.height = QR_PX

        ws.add_image(img, ws.cell(r, qr_col).coordinate)
        ws.row_dimensions[r].height = ROW_HEIGHT
        ws.cell(r, qr_col).alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions[get_column_letter(qr_col)].width = QR_COL_WIDTH
    ws.column_dimensions[get_column_letter(path_col)].width = PATH_COL_WIDTH

    output_path = os.path.join(OUTPUT_FOLDER, "output_with_qr.xlsx")
    wb.save(output_path)

    return render_template("download.html")


@app.route("/download")
def download_file():
    output_path = os.path.join(OUTPUT_FOLDER, "output_with_qr.xlsx")
    return send_file(output_path, as_attachment=True)


# ---------------- RUN LOCAL ----------------
if __name__ == "__main__":
    app.run(debug=True)
