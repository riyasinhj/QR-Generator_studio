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


def normalize(text):
    return str(text).strip().lower()


def cell_display_value(cell):
    """
    Returns EXACT Excel displayed value
    """
    if cell.value is None:
        return ""
    if cell.number_format and isinstance(cell.value, (int, float)):
        return cell._value if isinstance(cell._value, str) else format(cell.value, ",")
    return str(cell.value)


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

        wb = load_workbook(file_path, data_only=False)
        ws = wb.active
        columns = [cell.value for cell in ws[1]]

        return render_template("columns.html", columns=columns, file_path=file_path)

    return render_template("upload.html")


@app.route("/generate", methods=["POST"])
def generate():
    file_path = request.form["file_path"]
    selected_cols = request.form.getlist("columns")
    selected_norm = [normalize(c) for c in selected_cols]

    wb = load_workbook(file_path, data_only=False)
    ws = wb.active

    headers = [cell.value for cell in ws[1]]
    header_map = {normalize(h): idx for idx, h in enumerate(headers)}

    output_wb = load_workbook(file_path)
    output_ws = output_wb.active

    output_ws.cell(1, output_ws.max_column + 1, "QR_Image_Path")
    output_ws.cell(1, output_ws.max_column + 1, "QR")

    qr_col = output_ws.max_column

    for r in range(2, ws.max_row + 1):
        parts = []
        for col in selected_norm:
            if col not in header_map:
                parts.append("|")
            else:
                cell = ws.cell(r, header_map[col] + 1)
                val = cell_display_value(cell)
                parts.append(val if val else "|")

        qr_text = " ".join(parts)
        img_path = os.path.join(QR_FOLDER, f"qr_{r}.png")
        generate_qr(qr_text, img_path)

        output_ws.cell(r, qr_col - 1, img_path)

        img = Image(img_path)
        img.width = QR_PX
        img.height = QR_PX
        output_ws.add_image(img, output_ws.cell(r, qr_col).coordinate)

        output_ws.row_dimensions[r].height = ROW_HEIGHT
        output_ws.cell(r, qr_col).alignment = Alignment(horizontal="center", vertical="center")

    output_ws.column_dimensions[get_column_letter(qr_col)].width = QR_COL_WIDTH
    output_ws.column_dimensions[get_column_letter(qr_col - 1)].width = PATH_COL_WIDTH

    output_path = os.path.join(OUTPUT_FOLDER, "output_with_qr.xlsx")
    output_wb.save(output_path)

    return render_template("download.html")


@app.route("/download", methods=["GET"])
def download_file():
    return send_file(
        os.path.join(OUTPUT_FOLDER, "output_with_qr.xlsx"),
        as_attachment=True
    )


if __name__ == "__main__":
    app.run(debug=True)
