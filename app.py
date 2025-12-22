import os
import qrcode
from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")
QR_FOLDER = os.path.join(BASE_DIR, "qr_codes")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(QR_FOLDER, exist_ok=True)

QR_MM = 9
PIXELS_PER_MM = 96 / 25.4
QR_PX = int(QR_MM * PIXELS_PER_MM)
ROW_HEIGHT = 30
QR_COL_WIDTH = 18
PATH_COL_WIDTH = 95


def excel_display_value(cell):
    """
    Returns EXACT value as shown in Excel (CRITICAL FIX)
    """
    if cell.value is None:
        return ""

    val = cell.value
    fmt = cell.number_format

    # Text stays text
    if isinstance(val, str):
        return val.strip()

    # Indian / comma formats
    if isinstance(val, (int, float)):
        if "," in fmt and ".00" in fmt:
            return f"{val:,.2f}"
        elif "," in fmt:
            return f"{val:,.0f}"
        elif ".00" in fmt:
            return f"{val:.2f}"
        else:
            return str(val)

    return str(val)


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


@app.route("/", methods=["GET", "POST"])
def upload():
    if request.method == "POST":
        file = request.files["excel"]
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        wb = load_workbook(file_path)
        ws = wb.active
        headers = [c.value for c in ws[1]]

        return render_template("columns.html", columns=headers, file_path=file_path)

    return render_template("upload.html")


@app.route("/generate", methods=["POST"])
def generate():
    file_path = request.form["file_path"]
    selected_cols = request.form.getlist("columns")

    wb = load_workbook(file_path)
    ws = wb.active

    header_map = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}

    qr_path_col = ws.max_column + 1
    qr_img_col = ws.max_column + 2

    ws.cell(1, qr_path_col, "QR_Image_Path")
    ws.cell(1, qr_img_col, "QR")

    for r in range(2, ws.max_row + 1):
        qr_parts = []

        for col in selected_cols:
            cell = ws.cell(r, header_map[col])
            display_val = excel_display_value(cell)
            qr_parts.append(display_val if display_val else "|")

        qr_text = " ".join(qr_parts)

        img_path = os.path.join(QR_FOLDER, f"qr_{r}.png")
        generate_qr(qr_text, img_path)

        ws.cell(r, qr_path_col, img_path)

        img = Image(img_path)
        img.width = QR_PX
        img.height = QR_PX
        ws.add_image(img, ws.cell(r, qr_img_col).coordinate)
        ws.row_dimensions[r].height = ROW_HEIGHT
        ws.cell(r, qr_img_col).alignment = Alignment(horizontal="center", vertical="center")

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


if __name__ == "__main__":
    app.run(debug=True)
