import os
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


# ---------------- HELPERS ----------------
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


def get_excel_data_exact(path):
    """
    Reads Excel EXACTLY as displayed (NO formatting change)
    """
    wb = load_workbook(path, data_only=False)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        row_data = {}
        for h, cell in zip(headers, row):
            row_data[h] = cell.value if cell.value is not None else ""
        rows.append(row_data)

    return headers, rows


def build_qr_text(row, selected_cols):
    parts = []
    for col in selected_cols:
        val = row.get(col, "")
        if val == "":
            parts.append("|")
        else:
            parts.append(str(val))
    return " ".join(parts)


# ---------------- ROUTES ----------------
@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        file = request.files["excel"]
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        headers, _ = get_excel_data_exact(file_path)

        return render_template(
            "columns.html",
            columns=headers,
            file_path=file_path
        )

    return render_template("upload.html")


@app.route("/generate", methods=["POST"])
def generate():
    file_path = request.form["file_path"]
    selected_cols = request.form.getlist("columns")

    headers, rows = get_excel_data_exact(file_path)

    # Create output workbook
    wb = load_workbook(file_path)
    ws = wb.active

    qr_path_col = ws.max_column + 1
    qr_img_col = ws.max_column + 2

    ws.cell(1, qr_path_col, "QR_Image_Path")
    ws.cell(1, qr_img_col, "QR")

    for i, row in enumerate(rows, start=2):
        qr_text = build_qr_text(row, selected_cols)

        img_name = f"qr_{i}.png"
        img_path = os.path.join(QR_FOLDER, img_name)
        generate_qr(qr_text, img_path)

        ws.cell(i, qr_path_col, img_path)

        img = Image(img_path)
        img.width = QR_PX
        img.height = QR_PX
        ws.add_image(img, ws.cell(i, qr_img_col).coordinate)

        ws.row_dimensions[i].height = ROW_HEIGHT
        ws.cell(i, qr_img_col).alignment = Alignment(horizontal="center", vertical="center")

    ws.column_dimensions[get_column_letter(qr_img_col)].width = QR_COL_WIDTH
    ws.column_dimensions[get_column_letter(qr_path_col)].width = PATH_COL_WIDTH

    output_path = os.path.join(OUTPUT_FOLDER, "output_with_qr.xlsx")
    wb.save(output_path)

    return render_template("download.html")


@app.route("/download")
def download_file():
    return send_file(
        os.path.join(OUTPUT_FOLDER, "output_with_qr.xlsx"),
        as_attachment=True
    )


# ---------------- RUN ----------------
if __name__ == "__main__":
    app.run(debug=True)
