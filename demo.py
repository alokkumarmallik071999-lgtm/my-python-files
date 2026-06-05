import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.graphics.barcode import code128
from textwrap import wrap
from datetime import datetime
import json
import os
from pystrich.datamatrix import DataMatrixEncoder
from reportlab.platypus import (
    SimpleDocTemplate,
    Image,
    Paragraph,
    PageBreak,
    Spacer,
    Table,
    TableStyle
)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT

def generate_labels(
    company_choice,
    header_choice,
    generate_barcode,
    generate_transparency
):  
    # ================= LABEL SIZE =================
    LABEL_WIDTH = 20.026 * cm
    LABEL_HEIGHT = 11.7 * cm

    LEFT_MARGIN = 1.3 * cm
    RIGHT_MARGIN = 2 * cm
    TOP_MARGIN = 4.5 * cm
    BOTTOM_MARGIN = 1 * cm

    USABLE_HEIGHT = LABEL_HEIGHT - TOP_MARGIN - BOTTOM_MARGIN

    # ================= COMPANY =================
    if company_choice == "2":
        company_name = "OSAKATEK INDIA PVT LTD"
    else:
        company_name = "MAHAJAN INDUSTRIES"

    # ================= HEADER =================
    if header_choice == "2":
        header_text = "Imported & Marketed by:"
    else:
        header_text = "Manufactured & Marketed by:"

    MANUFACTURER = [
        header_text,
        company_name,
        "Khasra no. 57/3/2/2/2, wazidpur,",
        "saboli, Sonipat Haryana -131029"
    ]

    df = pd.read_excel("labels.xlsx")

    for _, row in df.iterrows():
        asin = str(row["ASIN"]).strip().upper()
        product = str(row["ProductCode"]).strip()
        title = str(row["TITLE"]).strip()
        mrp = f"{float(row['MRP']):.0f}"
        qty = int(row["QTY"])

        mfg_date = datetime.now().strftime("%b-%Y").upper()

        c = canvas.Canvas(f"{asin}_labels.pdf", pagesize=(LABEL_WIDTH, LABEL_HEIGHT))

        # ================= FIRST: ALL MRP LABELS =================
        for _ in range(qty):

            line_height = 1 * cm

            total_lines = len(MANUFACTURER) + len(wrap(product, 40)) + 4
            total_height = total_lines * line_height

            y = BOTTOM_MARGIN + (USABLE_HEIGHT + total_height) / 2

            c.setFont("Helvetica-Bold", 28)

            for line in MANUFACTURER:
                c.drawString(LEFT_MARGIN, y, line)
                y -= line_height

            y -= 0.2 * cm

            wrapped_product = wrap(f"Product - {product}", 40)
            for line in wrapped_product:
                c.drawString(LEFT_MARGIN, y, line)
                y -= line_height

            c.drawString(LEFT_MARGIN, y, f"MFG Date : {mfg_date}")
            y -= line_height
            c.drawString(LEFT_MARGIN, y, "Qty: 1PC")
            y -= line_height
            c.drawString(LEFT_MARGIN, y, f"MRP : Rs.{mrp}/- (Inclusive of all taxes)")

            c.showPage()

        # ================= SECOND: ALL BARCODE LABELS =================
        if generate_barcode:
            for _ in range(qty):

                line_height = 1 * cm
                wrapped_title = wrap(title, 40)

                total_lines = len(wrapped_title) + 2 + 3
                total_height = total_lines * line_height

                y = BOTTOM_MARGIN + (USABLE_HEIGHT + total_height) / 2

                # TITLE (left aligned)
                c.setFont("Helvetica-Bold", 26)
                for line in wrapped_title:
                    c.drawString(LEFT_MARGIN /2, y, line)
                    y -= line_height

                y -= 0.5 * cm

                # ASIN (center)
                c.setFont("Helvetica-Bold", 22)
                c.drawCentredString(LABEL_WIDTH / 2, y, asin)
                y -= 2 * cm

                # BARCODE (center)
                barcode_obj = code128.Code128(
                    asin,
                    barHeight=2.8 * cm,
                    barWidth=0.12 * cm
                )

                barcode_x = (LABEL_WIDTH - barcode_obj.width) / 2
                barcode_obj.drawOn(c, barcode_x, y - 2 * cm)

                c.showPage()

                # ================= THIRD: TRANSPARENCY =================

        if generate_transparency:

            try:

                with open(
                    "artifact.json",
                    "r",
                    encoding="utf-8"
                ) as f:
                    artifact = json.load(f)

                transparency_codes = []

                for item in artifact.get("codesList", []):
                    transparency_codes.extend(
                        item.get("codes", [])
                    )

                transparency_sku = ""

                if artifact.get("codesList"):
                    transparency_sku = artifact["codesList"][0].get(
                        "sku",
                        ""
                    )

                transparency_qty = min(
                    qty,
                    len(transparency_codes)
                )

                for page_no in range(transparency_qty):

                    code = transparency_codes[page_no]

                    dm_file = f"dm_{page_no}.png"

                    encoder = DataMatrixEncoder(code)
                    encoder.save(dm_file)

                    if page_no > 0:
                        c.showPage()

                    c.drawImage(
                    "logo.png",
                     1.3 * cm,
                     7.0 * cm,
                     width=3 * cm,
                    height=2.2 * cm,
                    preserveAspectRatio=True,
                     mask='auto'
                        )

                    c.setFont(
                        "Helvetica-Bold",
                        50
                    )

                    c.drawString(
                        5 * cm,
                        8.0 * cm,
                        "Scan with the"
                    )

                    c.drawString(
                        5 * cm,
                        7.2 * cm,
                        "Transparency App"
                    )

                    c.drawImage(
                        dm_file,
                        14 * cm,
                        5.8 * cm,
                        width=3.8 * cm,
                        height=3.8 * cm
                    )

                    c.setFont(
                        "Helvetica-Bold",
                        50
                    )

                    c.drawString(
                        14 * cm,
                        5.0 * cm,
                        transparency_sku
                    )

                    c.drawRightString(
                        LABEL_WIDTH - 1 * cm,
                        0.8 * cm,
                        str(page_no + 1)
                    )

                    if os.path.exists(dm_file):
                        os.remove(dm_file)

            except Exception as e:

                print(
                    "Transparency Error:",
                    e
                )

        c.save()

    print("✅ Done")
