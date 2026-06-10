import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from reportlab.graphics.barcode import code128
from textwrap import wrap
from datetime import datetime
import json
import os
from pystrich.datamatrix import DataMatrixEncoder
import requests
import uuid
import time


def get_token():
    CLIENT_ID = "6peh3gn3crdveh15aajar7gabf"
    CLIENT_SECRET = "7993b2622g2vnjcpoamj44bm5v592t14s86erniigi8k0igaap"

    AUTH_URL = "https://tpncy-web-services.auth.us-east-1.amazoncognito.com/oauth2/token"
   

    token_response = requests.post(
        AUTH_URL,
        data={
            "grant_type": "client_credentials",
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET
        }
    )

    token_response.raise_for_status()

    return token_response.json()["access_token"]

def get_transparency_codes(gtin, count, token):
    API_URL = "https://api.transparency.com/v1.2"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    payload = {
        "gtin": gtin,
        "count": count,
        "requestId": str(uuid.uuid4())
    }

    response = requests.post(
        f"{API_URL}/serial/sgtin",
        headers=headers,
        json=payload
    )
    print("Payload:", payload)
    print("Response:", response.text)


    response.raise_for_status()

    job_id = response.headers["Location"].split("/")[-1]

    while True:

        status_response = requests.get(
            f"{API_URL}/serial/job/{job_id}",
            headers={
                "Authorization": f"Bearer {token}"
            }
        )

        status_response.raise_for_status()

        data = status_response.json()

        status = str(
            data.get("status", "")
        ).upper()

        if status in ["DONE", "COMPLETED"]:

            artifact_url = data["url"]

            artifact = requests.get(
                artifact_url
            ).json()

            codes = []
            sku = ""

            for item in artifact.get(
                "codesList",
                []
            ):
                codes.extend(
                    item.get(
                        "codes",
                        []
                    )
                )

                if not sku:
                    sku = item.get(
                        "sku",
                        ""
                    )

            return codes, sku

        if status in ["FAILED", "ERROR"]:
            raise Exception(
                f"Transparency failed for GTIN {gtin}"
            )

        time.sleep(10)

def generate_labels(
    company_choice="2",
    header_choice="1",
    generate_barcode=True,
    generate_transparency=True,
    excel_file="labels.xlsx"
):
    token = get_token()

    # ================= LABEL SIZE =================

    LABEL_WIDTH = 20.026 * cm
    LABEL_HEIGHT = 11.7 * cm

    LEFT_MARGIN = 1.3 * cm
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

    # ================= EXCEL =================
    if not os.path.exists(excel_file):
        raise FileNotFoundError(
        f"Excel file not found: {excel_file}"
    )

    df = pd.read_excel(
    excel_file,
    dtype={"GTIN": str}
)

    # ================= TRANSPARENCY DATA =================

    
    # ================= LOOP PRODUCTS =================

    for _, row in df.iterrows():

        asin = str(row["ASIN"]).strip().upper()
        product = str(row["ProductCode"]).strip()
        title = str(row["TITLE"]).strip()

        mrp = f"{float(row['MRP']):.0f}"
        qty = int(row["QTY"])
        gtin = str(row.get("GTIN", "")).strip()

    

        transparency_codes = []
        transparency_sku = ""

        valid_gtin = (
        gtin
        and gtin.lower() != "nan"
        and gtin.upper() != "NA"
)



        if generate_transparency and valid_gtin:

            try:

                transparency_codes, transparency_sku = (
                    get_transparency_codes(
                        gtin,   
                        qty,
                        token
                    )
                )

                print(
                    f"{asin} -> {len(transparency_codes)} codes generated"
                )

            except Exception as e:

                print(
                    f"Transparency Error ({asin}): {e}"
                )
        elif generate_transparency:
            print(
        f"{asin} -> GTIN missing, transparency skipped"
    )

        mfg_date = datetime.now().strftime("%b-%Y").upper()

        pdf_name = f"{asin}_labels.pdf"

        c = canvas.Canvas(
            pdf_name,
            pagesize=(LABEL_WIDTH, LABEL_HEIGHT)
        )

        # ==================================================
        # 1. MRP LABELS
        # ==================================================

        for _ in range(qty):

            line_height = 1 * cm

            total_lines = (
                len(MANUFACTURER)
                + len(wrap(product, 40))
                + 4
            )

            total_height = total_lines * line_height

            y = (
                BOTTOM_MARGIN
                + (USABLE_HEIGHT + total_height) / 2
            )

            c.setFont("Helvetica-Bold", 28)

            for line in MANUFACTURER:
                c.drawString(
                    LEFT_MARGIN,
                    y,
                    line
                )
                y -= line_height

            y -= 0.2 * cm

            wrapped_product = wrap(
                f"Product - {product}",
                40
            )

            for line in wrapped_product:
                c.drawString(
                    LEFT_MARGIN,
                    y,
                    line
                )
                y -= line_height

            c.drawString(
                LEFT_MARGIN,
                y,
                f"MFG Date : {mfg_date}"
            )
            y -= line_height

            c.drawString(
                LEFT_MARGIN,
                y,
                "Qty : 1PC"
            )
            y -= line_height

            c.drawString(
                LEFT_MARGIN,
                y,
                f"MRP : Rs.{mrp}/- (Inclusive of all taxes)"
            )

            c.showPage()

        # ==================================================
        # 2. BARCODE LABELS
        # ==================================================

        if generate_barcode:

            for _ in range(qty):

                line_height = 1 * cm

                wrapped_title = wrap(
                    title,
                    40
                )

                total_lines = len(wrapped_title) + 5
                total_height = total_lines * line_height

                y = (
                    BOTTOM_MARGIN
                    + (USABLE_HEIGHT + total_height) / 2
                )

                c.setFont(
                    "Helvetica-Bold",
                    26
                )

                for line in wrapped_title:
                    c.drawString(
                        LEFT_MARGIN / 3,
                        y,
                        line
                    )
                    y -= line_height

                y -= 0.5 * cm

                c.setFont(
                    "Helvetica-Bold",
                    22
                )

                c.drawCentredString(
                    LABEL_WIDTH / 2,
                    y,
                    asin
                )

                y -= 2 * cm

                barcode_obj = code128.Code128(
                    asin,
                    barHeight=2.8 * cm,
                    barWidth=0.1 * cm
                )

                barcode_x = (
                    LABEL_WIDTH
                    - barcode_obj.width
                ) / 2

                barcode_obj.drawOn(
                    c,
                    barcode_x,
                    y - 2 * cm
                )

                c.showPage()

        # ==================================================
        # 3. TRANSPARENCY LABELS
        # ==================================================

        if (
            generate_transparency
            and transparency_codes
        ):

            transparency_qty = min(
                qty,
                len(transparency_codes)
            )

            for page_no in range(
                transparency_qty
            ):

                code = transparency_codes[
                    page_no
                ]

                dm_file = (
                    f"dm_{asin}_{page_no}.png"
                )

                try:

                    encoder = DataMatrixEncoder(
                        code
                    )

                    encoder.save(
                        dm_file
                    )

                    c.drawImage(
                        "logo.png",
                        -1 * cm,
                        6 * cm,
                        width=6.5 * cm,
                        height=3.0 * cm,
                        preserveAspectRatio=True,
                        mask="auto"
                    )

                    c.setFont(
                        "Helvetica-Bold",
                        26
                    )

                    c.drawString(
                        4 * cm,
                        8 * cm,
                        "Scan with the"
                    )

                    c.drawString(
                        4 * cm,
                        7 * cm,
                        "Transparency App"
                    )

                    c.drawImage(
                        dm_file,
                        12.5 * cm,
                        2 * cm,
                        width=7.4 * cm,
                        height=7.4 * cm
                    )

                                        # ================= SKU =================

                    font_size = 22

                    if len(transparency_sku) > 25:
                        font_size = 18

                    if len(transparency_sku) > 40:
                        font_size = 14

                    c.setFont(
                        "Helvetica-Bold",
                        font_size
                    )

                    sku_lines = wrap(
                        transparency_sku,
                        width=20
                    )

                    sku_y = 1.3 * cm

                    for line in sku_lines:
                        c.drawCentredString(
                            15.5 * cm,
                            sku_y,
                            line
                        )
                        sku_y -= 0.6 * cm
                    

                    c.setFont(
                        "Helvetica",
                        18
                    )

                    c.drawRightString(
                        LABEL_WIDTH - 1 * cm,
                        0.2 * cm,
                        str(page_no + 1)
                    )

                    c.showPage()

                except Exception as e:
                    print(
                        f"Transparency page error: {e}"
                    )

                finally:

                    if os.path.exists(
                        dm_file
                    ):
                        os.remove(
                            dm_file
                        )

        # ==================================================
        # SAVE PDF
        # ==================================================

        c.save()

        print(
            f"✅ Generated: {pdf_name}"
        )

    print(
        "\n🎉 All PDFs generated successfully."
    )


