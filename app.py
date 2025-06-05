from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime, timedelta
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

app = Flask(__name__)

# Register Thai font
pdfmetrics.registerFont(TTFont("THSarabun", "static/THSarabunNew.ttf"))

def safe_str(val):
    return "" if pd.isna(val) else str(val)

def generate_pdf(data):
    PAGE_WIDTH = 100 * mm
    PAGE_HEIGHT = 30 * mm
    BOX_WIDTH = 25 * mm
    BOX_HEIGHT = 30 * mm
    filename_pdf = f"StickerCulture_{datetime.today().strftime('%Y%m%d%H%M%S')}.pdf"
    pdf = canvas.Canvas(filename_pdf, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))
    pdf.setFont("THSarabun", 8)

    def draw_culture_name(pdf_canvas, text, x, y, fontname, fontsize, max_width):
        prefix = "Culture name : "
        full_text = text[len(prefix):].strip() if text.startswith(prefix) else text.strip()
        words = full_text.split()
        display_name = " ".join(words[:2]) if len(words) > 2 else full_text
        pdf_canvas.setFont(fontname, fontsize)
        width = pdf_canvas.stringWidth(display_name, fontname, fontsize)
        while width > max_width and fontsize > 5:
            fontsize -= 0.5
            width = pdf_canvas.stringWidth(display_name, fontname, fontsize)
        pdf_canvas.drawString(x, y, display_name)

    for _, row in data.iterrows():
        for i in range(4):
            x_left = i * BOX_WIDTH
            y_bottom = 0
            pdf.rect(x_left, y_bottom, BOX_WIDTH, BOX_HEIGHT)
            text_y = y_bottom + BOX_HEIGHT - 4 * mm

            lsn_value = safe_str(row.get('LSN :', ''))
            fontsize_lsn = 12
            pdf.setFont("THSarabun", fontsize_lsn)
            text_width = pdf.stringWidth(lsn_value, "THSarabun", fontsize_lsn)
            x_centered = x_left + (BOX_WIDTH - text_width) / 2
            pdf.rect(x_centered - 1*mm, text_y - 1*mm, text_width + 2*mm, fontsize_lsn + 1)
            pdf.drawString(x_centered, text_y, lsn_value)
            text_y -= 4 * mm

            pdf.setFont("THSarabun", 8)
            fullname = safe_str(row.get('Fullname :', ''))
            lines = fullname.split()
            lines = [' '.join(lines[:2])] if len(lines) > 2 else [fullname]
            for line in lines:
                pdf.drawString(x_left + 1.5 * mm, text_y, line)
                text_y -= 3 * mm

            fields = [
                safe_str(row.get('Hospital :', '')),
                safe_str(row.get('Criteria :', '')),
                f"Direct AFB : {safe_str(row.get('Direct AFB', ''))}",
                f"Culture date :  {safe_str(row.get('Culture date : ', ''))}",
                f"Culture name : {safe_str(row.get('Culture name :', ''))}",
                f"Left date : {safe_str(row.get('Left date', ''))}",
            ]

            for text in fields:
                if text.startswith("Culture name :"):
                    draw_culture_name(pdf, text, x_left + 1.5 * mm, text_y, "THSarabun", 8, BOX_WIDTH - 3 * mm)
                else:
                    pdf.drawString(x_left + 1.5 * mm, text_y, text[:30])
                text_y -= 3 * mm

        pdf.showPage()

    pdf.save()
    return filename_pdf

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        if file.filename.endswith(".csv"):
            df = pd.read_csv(file)
        elif file.filename.endswith((".xls", ".xlsx")):
            df = pd.read_excel(file)
        else:
            return "Unsupported file format", 400

        # เติมข้อมูลวันที่ Left date จาก Culture date
        if 'Culture date : ' in df.columns:
            df['Culture date : '] = pd.to_datetime(df['Culture date : '], errors='coerce').dt.strftime('%d/%m/%Y')
            df['Left date'] = pd.to_datetime(df['Culture date : '], format='%d/%m/%Y', errors='coerce') \
                                .apply(lambda d: (d + timedelta(days=56)).strftime('%d/%m/%Y') if pd.notna(d) else '')

        pdf_filename = generate_pdf(df)
        return send_file(pdf_filename, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
