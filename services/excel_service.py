import os
import time
import pandas as pd
import requests

from datetime import datetime

from utils.constants import PAYABLE_MAP

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


# ==============================
# 📋 COPY TEMPLATE ON ONEDRIVE
# ==============================

def copy_template_on_onedrive(
    token,
    user_email,
    template_path,
    ticket_id,
    output_file_name
):
    """
    Server-side copy of the template into the ticket folder.
    Preserves 100% of the original file — dropdowns, images,
    merged cells, styles — because the file is never downloaded
    or parsed locally.
    Deletes the destination file first if it already exists,
    because the Graph /copy endpoint ignores conflictBehavior.
    """

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    destination_folder = (
        f"New Sales RPA/DEV/"
        f"BotShareDrive/InProgress/"
        f"{ticket_id}"
    )

    destination_path = (
        f"{destination_folder}/{output_file_name}"
    )

    # ── Delete existing file if present ──────────────────────
    delete_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{destination_path}"
    )

    delete_resp = requests.delete(
        delete_url,
        headers=headers
    )

    # 204 = deleted, 404 = didn't exist — both are fine
    if delete_resp.status_code not in [204, 404]:
        print(delete_resp.text)
        raise Exception("Failed to delete existing file before copy")

    if delete_resp.status_code == 204:
        print("\n🗑️ Existing file deleted")

    # ── Copy template ─────────────────────────────────────────
    copy_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{template_path}:/copy"
    )

    body = {
        "parentReference": {
            "path": f"/drive/root:/{destination_folder}"
        },
        "name": output_file_name
    }

    response = requests.post(
        copy_url,
        headers=headers,
        json=body
    )

    # Graph returns 202 Accepted for async copy operations
    if response.status_code not in [200, 201, 202]:
        print(response.text)
        raise Exception("Template Copy Failed")

    monitor_url = response.headers.get("Location")

    if monitor_url:
        print("\n⏳ Waiting for copy to complete...")
        _poll_copy_operation(monitor_url)

    print("\n✅ Template Copied Successfully")

    return destination_path


def _poll_copy_operation(
    monitor_url,
    max_retries=20
):
    """Poll the async copy operation until it completes."""

    for _ in range(max_retries):

        resp = requests.get(monitor_url)
        data = resp.json()
        status = data.get("status", "")

        if status == "completed":
            return

        elif status == "failed":
            raise Exception(
                f"Copy operation failed: {data}"
            )

        time.sleep(2)

    raise Exception("Copy operation timed out")


# ==============================
# 📊 CREATE WORKBOOK SESSION
# ==============================

def create_workbook_session(
    token,
    user_email,
    file_path
):
    """
    Opens a persistent Excel session on the OneDrive file.
    persistChanges=true means edits are saved to the actual file.
    """

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}:/workbook/createSession"
    )

    body = {"persistChanges": True}

    response = requests.post(
        url,
        headers=headers,
        json=body
    )

    if response.status_code != 201:
        print(response.text)
        raise Exception(
            "Failed to create workbook session"
        )

    session_id = response.json()["id"]

    print("\n✅ Workbook Session Created")

    return session_id


# ==============================
# ✏️ UPDATE CELL RANGE
# ==============================

def update_cell_range(
    token,
    user_email,
    file_path,
    session_id,
    sheet_name,
    cell_range,
    values
):
    """
    PATCH a range of cells with values.
    values must be a 2D list: [[row1col1, row1col2], [row2col1, ...]]
    Styles, dropdowns and images on the sheet are completely untouched.
    """

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "workbook-session-id": session_id
    }

    encoded_sheet = requests.utils.quote(sheet_name)

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}:/workbook"
        f"/worksheets/{encoded_sheet}"
        f"/range(address='{cell_range}')"
    )

    body = {"values": values}

    response = requests.patch(
        url,
        headers=headers,
        json=body
    )

    if response.status_code != 200:
        print(response.text)
        raise Exception(
            f"Failed to update range {cell_range}"
        )


# ==============================
# 🔒 CLOSE WORKBOOK SESSION
# ==============================

def close_workbook_session(
    token,
    user_email,
    file_path,
    session_id
):

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "workbook-session-id": session_id
    }

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}:/workbook/closeSession"
    )

    requests.post(url, headers=headers)

    print("\n✅ Workbook Session Closed")


# ==============================
# 📝 POPULATE EXCEL TEMPLATE
# ==============================

def populate_excel_template(
    token,
    user_email,
    file_path,
    session_id,
    closing_ticket_df,
    invoice_df
):
    """
    Writes all data into the copied template via Graph API.
    The file stays on OneDrive throughout — no download, no openpyxl.
    Dropdowns, images, and all formatting are fully preserved.
    """

    sheet = "Closing Check Transmittal Form"

    row = closing_ticket_df.iloc[0]

    purchase_price = row.get(
        "cr109_saleprice", ""
    )

    closing_date = row.get(
        "cr7de_closingdate", ""
    )

    seller_tcode = row.get(
        "cr7de_sellertcode", ""
    )

    property_address = row.get(
        "cr7de_buildingaddress", ""
    )

    unit = row.get(
        "cr7de_unitnumber", ""
    )

    seller1_name = row.get(
        "cr7de_sellername", ""
    )

    deal = row.get(
        "cr7de_deal", ""
    )

    buyer1_name = row.get(
        "cr7de_buyername", ""
    )

    shares = row.get(
        "cr109_shares", ""
    )

    closing_agent = row.get(
        "cr7de_closingagentname", ""
    )

    closing_agent_phone = row.get(
        "cr7de_closingagentphone", ""
    )

    closing_agent_email = row.get(
        "cr7de_closingagentemail", ""
    )

    closing_agent_title = row.get(
        "cr7de_titlerole", ""
    )

    notes = row.get(
        "cr7de_notes", ""
    )

    current_date = datetime.now().strftime(
        "%m/%d/%Y"
    )

    try:

        if closing_date:

            closing_date = (
                pd.to_datetime(closing_date)
                .strftime("%m/%d/%Y")
            )

    except Exception:
        pass

    # Shorthand so every patch call stays readable
    def patch(cell_range, values):
        update_cell_range(
            token,
            user_email,
            file_path,
            session_id,
            sheet,
            cell_range,
            values
        )

    # ==============================
    # 📝 HEADER VALUES
    # Each patch call sends only the value — formatting,
    # dropdowns and images on the sheet are never touched.
    # ==============================

    patch("D1", [[current_date]])
    patch("D2", [[purchase_price]])
    patch("D3", [[deal]])
    patch("D4", [[shares]])
    patch("D5", [[closing_date]])
    patch("D6", [[seller_tcode]])
    patch("D7", [[property_address]])
    patch("D8", [[unit]])

    patch("C13", [[seller1_name]])
    patch("C43", [[buyer1_name]])

    patch("B86", [[closing_agent]])
    patch("B87", [[closing_agent_email]])
    patch("B88", [[closing_agent_phone]])
    patch("B89", [[closing_agent_title]])
    patch("B91", [[notes]])

    # ==============================
    # 📄 SELLER TABLE
    # ==============================

    seller_df = invoice_df[
        invoice_df["cr7de_paidby"] == 716070000
    ]

    seller_start_row = 15

    for idx, (_, inv_row) in enumerate(
        seller_df.iterrows()
    ):

        r = seller_start_row + idx

        patch(
            f"A{r}:D{r}",
            [[
                inv_row.get("cr7de_chequenumber", ""),
                inv_row.get("cr7de_dueatclosing", ""),
                inv_row.get("cr7de_amount", ""),
                PAYABLE_MAP.get(
                    inv_row.get("cr7de_payableto", ""), ""
                )
            ]]
        )

    # ==============================
    # 📄 BUYER TABLE
    # ==============================

    buyer_df = invoice_df[
        invoice_df["cr7de_paidby"] == 716070001
    ]

    buyer_start_row = 45

    for idx, (_, inv_row) in enumerate(
        buyer_df.iterrows()
    ):

        r = buyer_start_row + idx

        patch(
            f"A{r}:D{r}",
            [[
                inv_row.get("cr7de_chequenumber", ""),
                inv_row.get("cr7de_dueatclosing", ""),
                inv_row.get("cr7de_amount", ""),
                PAYABLE_MAP.get(
                    inv_row.get("cr7de_payableto", ""), ""
                )
            ]]
        )

    print("\n✅ Excel Populated Successfully")


# ==============================
# 📄 CONVERT EXCEL TO PDF
# ==============================

def convert_onedrive_file_to_pdf(
    token,
    user_email,
    file_path,
    pdf_output_path
):

    headers = {
        "Authorization": f"Bearer {token}"
    }

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}:/content"
        f"?format=pdf"
    )

    response = requests.get(
        url,
        headers=headers
    )

    print(
        "\n📡 PDF Conversion Response:",
        response.status_code
    )

    if response.status_code != 200:
        print(response.text)
        raise Exception("PDF Conversion Failed")

    with open(pdf_output_path, "wb") as f:
        f.write(response.content)

    print("\n✅ PDF Generated Successfully")