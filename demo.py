import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.graphics.barcode import code128
from textwrap import wrap
from datetime import datetime
import json
import os
from pystrich.datamatrix import DataMatrixEncoder
from datetime import datetime
def create_datamatrix(code, filename):

    encoder = DataMatrixEncoder(code)

    encoder.save(filename)

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
    with open(
    "artifact.json",
    "r",
    encoding="utf-8"
) as f:
    artifact = json.load(f)

codes = []

for item in artifact.get("codesList", []):
    codes.extend(item.get("codes", []))

transparency_sku = ""

if artifact.get("codesList"):
    transparency_sku = artifact["codesList"][0].get(
        "sku",
        ""
    )

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
 # ==========================================
 # THIRD: TRANSPARENCY LABELS
# ==========================================

                if generate_transparency:
                    transparency_qty = min(
                    qty,
                    len(codes)
                        )

                    print(
                "Transparency Labels:",
                    transparency_qty
                        )

    
LABEL_WIDTH = 20.026 * cm
LABEL_HEIGHT = 11.7 * cm

LEFT_MARGIN = 1.3 * cm
RIGHT_MARGIN = 1.0 * cm
TOP_MARGIN = 1.5 * cm
BOTTOM_MARGIN = 0.1 * cm

# ==========================================
# LOAD JSON
# ==========================================

print("Loading artifact.json...")

with open(
    "artifact.json",
    "r",
    encoding="utf-8"
) as f:
    artifact = json.load(f)

print("JSON Loaded")

# ==========================================
# EXTRACT CODES
# ==========================================

codes = []

for item in artifact.get("codesList", []):
    codes.extend(item.get("codes", []))

sku = ""

if artifact.get("codesList"):
    sku = artifact["codesList"][0].get(
        "sku",
        ""
    )

print("Codes Found:", len(codes))

# ==========================================
# DATAMATRIX
# ==========================================

def create_datamatrix(code, filename):
    encoder = DataMatrixEncoder(code)
    encoder.save(filename)

# ==========================================
# PDF
# ==========================================

pdf = SimpleDocTemplate(
    "amazon_style_labels.pdf",
    pagesize=(LABEL_WIDTH, LABEL_HEIGHT),
    leftMargin=LEFT_MARGIN,
    rightMargin=RIGHT_MARGIN,
    topMargin=TOP_MARGIN,
    bottomMargin=BOTTOM_MARGIN
)

# ==========================================
# STYLES
# ==========================================

scan_style = ParagraphStyle(
    "SCAN",
    fontSize=20,
    leading=28,
    alignment=TA_LEFT
)

sku_style = ParagraphStyle(
    "SKU",
    fontSize=20,
    leading=25,
    alignment=TA_LEFT
)

page_style = ParagraphStyle(
    "PAGE",
    fontSize=18,
    alignment=TA_LEFT
)

# ==========================================
# LOGO
# ==========================================

LOGO_FILE = "logo.png"

elements = []

# ==========================================
# LABEL LOOP
# ==========================================

for page_no, code in enumerate(codes, start=1):

    dm_file = f"dm_{page_no}.png"

    create_datamatrix(
        code,
        dm_file
    )

    logo = Image(
         LOGO_FILE,
    width=90,
    height=70
    )

    scan_text = Paragraph(
        "Scan with the<br/>Transparency App",
        scan_style
    )

    dm_img = Image(
        dm_file,
        width=150,
        height=150
    )

    # ------------------------
    # TOP SECTION
    # ------------------------
    elements.append(Spacer(1,10))
    top_table = Table(
        [[logo, scan_text, dm_img]],
        colWidths=[120, 230, 170]
    )

    top_table.setStyle(
        TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
            ("TOPPADDING", (0,0), (-1,-1), 0),
            ("BOTTOMPADDING", (0,0), (-1,-1), 0),
        ])
    )

    elements.append(top_table)

    elements.append(
        Spacer(1, 10)
    )

    # ------------------------
    # SKU BELOW CODE
    # ------------------------

    sku_text = sku.replace(
        "_",
        "<br/>_"
    )

    sku_table = Table(
        [[
            "",
            "",
            Paragraph(
                sku_text,
                sku_style
            )
        ]],
        colWidths=[120, 230, 230]
    )

    sku_table.setStyle(
        TableStyle([
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ])
    )

    elements.append(
        sku_table
    )

    elements.append(
        Spacer(1, 15)
    )

    # ------------------------
    # PAGE NUMBER
    # ------------------------

    page_table = Table(
        [[
            "",
            "",
            Paragraph(
                str(page_no),
                page_style
            )
        ]],
        colWidths=[400, 230, 170]
    )

    page_table.setStyle(
        TableStyle([
            ("ALIGN", (2,0), (2,0), "RIGHT"),
            ("LEFTPADDING", (0,0), (-1,-1), 0),
            ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ])
    )

    elements.append(page_table)

    if page_no < len(codes):
        elements.append(PageBreak())

# ==========================================
# BUILD PDF
# ==========================================

pdf.build(elements)

# ==========================================
# DELETE TEMP DATAMATRIX FILES
# ==========================================

for i in range(1, len(codes) + 1):

    file_name = f"dm_{i}.png"

    if os.path.exists(file_name):
        os.remove(file_name)

print()
print("=" * 50)
print("PDF CREATED SUCCESSFULLY")
print("amazon_style_labels.pdf")
print("=" * 50)


        c.save()

    print("✅ Done")
