
import time
import pandas as pd
import requests

from datetime import datetime

from utils.constants import (
    PAYABLE_MAP,
    TRANSACTION_TYPE_DEAL_MAP,
    DUE_AT_CLOSING_MAP
)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

BASE_FOLDER = "New Sales RPA/DEV/BotShareDrive/InProgress"


# ======================================================
# 🛠️  LOW-LEVEL GRAPH HELPERS
# ======================================================

def _headers(token, extra=None):
    h = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    if extra:
        h.update(extra)

    return h


def _ensure_folder(token, user_email, folder_path):
    """
    Create a folder (and any missing parents) if it does not exist.
    """

    check_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{folder_path}"
    )

    resp = requests.get(
        check_url,
        headers=_headers(token)
    )

    if resp.status_code == 200:
        return

    parts = folder_path.rsplit("/", 1)

    parent_path = parts[0] if len(parts) == 2 else ""
    folder_name = parts[-1]

    if parent_path:
        create_url = (
            f"{GRAPH_BASE}/users/{user_email}"
            f"/drive/root:/{parent_path}:/children"
        )
    else:
        create_url = (
            f"{GRAPH_BASE}/users/{user_email}"
            f"/drive/root/children"
        )

    body = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "replace"
    }

    resp = requests.post(
        create_url,
        headers=_headers(token),
        json=body
    )

    if resp.status_code not in [200, 201]:
        raise Exception(
            f"Failed to create folder '{folder_path}': {resp.text}"
        )

    print(f"\n📁 Folder created: {folder_path}")


def _get_file_metadata(token, user_email, file_path):

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
    )

    resp = requests.get(
        url,
        headers=_headers(token)
    )

    if resp.status_code == 200:
        return resp.json()

    if resp.status_code == 404:
        return None

    raise Exception(
        f"Error checking file '{file_path}': {resp.text}"
    )


def _move_and_rename_file(
    token,
    user_email,
    source_path,
    destination_folder_path,
    new_name
):

    folder_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{destination_folder_path}"
    )

    folder_resp = requests.get(
        folder_url,
        headers=_headers(token)
    )

    if folder_resp.status_code != 200:
        raise Exception(
            f"Cannot find destination folder "
            f"'{destination_folder_path}': {folder_resp.text}"
        )

    folder_id = folder_resp.json()["id"]

    move_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{source_path}"
    )

    body = {
        "parentReference": {"id": folder_id},
        "name": new_name
    }

    resp = requests.patch(
        move_url,
        headers=_headers(token),
        json=body
    )

    if resp.status_code != 200:
        raise Exception(
            f"Failed to move/rename file: {resp.text}"
        )

    print(f"\n📦 Archived: {new_name}")


def _delete_file(token, user_email, file_path):

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
    )

    resp = requests.delete(
        url,
        headers=_headers(token)
    )

    if resp.status_code not in [204, 404]:
        raise Exception(
            f"Failed to delete '{file_path}': {resp.text}"
        )


# ======================================================
# 📁  FOLDER MANAGEMENT
# ======================================================

def setup_ticket_folders(token, user_email, ticket_id):

    ticket_folder = f"{BASE_FOLDER}/{ticket_id}"

    active_folder = f"{ticket_folder}/Active"

    inactive_folder = f"{ticket_folder}/Inactive"

    _ensure_folder(token, user_email, ticket_folder)

    _ensure_folder(token, user_email, active_folder)

    _ensure_folder(token, user_email, inactive_folder)

    return active_folder, inactive_folder


# ======================================================
# 📋  ARCHIVE + COPY LOGIC
# ======================================================

def archive_active_file_if_exists(
    token,
    user_email,
    active_folder,
    inactive_folder,
    output_file_name
):

    active_file_path = f"{active_folder}/{output_file_name}"

    existing = _get_file_metadata(
        token,
        user_email,
        active_file_path
    )

    if existing is None:
        print("\n📂 No existing file in Active — skipping archive")
        return

    now = datetime.now()

    date_str = now.strftime("%Y-%m-%d")

    datetime_str = now.strftime("%Y-%m-%d_%H-%M-%S")

    name_parts = output_file_name.rsplit(".", 1)

    if len(name_parts) == 2:
        archived_name = (
            f"{name_parts[0]}_{datetime_str}.{name_parts[1]}"
        )
    else:
        archived_name = f"{output_file_name}_{datetime_str}"

    date_folder = f"{inactive_folder}/{date_str}"

    _ensure_folder(token, user_email, date_folder)

    _move_and_rename_file(
        token,
        user_email,
        active_file_path,
        date_folder,
        archived_name
    )

    print(
        f"\n✅ Archived to: "
        f"Inactive/{date_str}/{archived_name}"
    )


# ======================================================
# 📋  COPY TEMPLATE → ACTIVE FOLDER
# ======================================================

def copy_template_to_active(
    token,
    user_email,
    template_path,
    active_folder,
    output_file_name
):

    headers = _headers(token)

    copy_url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{template_path}:/copy"
    )

    body = {
        "parentReference": {
            "path": f"/drive/root:/{active_folder}"
        },
        "name": output_file_name
    }

    response = requests.post(
        copy_url,
        headers=headers,
        json=body
    )

    if response.status_code not in [200, 201, 202]:
        print(response.text)
        raise Exception("Template Copy Failed")

    monitor_url = response.headers.get("Location")

    if monitor_url:
        print("\n⏳ Waiting for copy to complete...")
        _poll_copy_operation(monitor_url)

    print("\n✅ Template Copied to Active/")

    return f"{active_folder}/{output_file_name}"


def _poll_copy_operation(
    monitor_url,
    max_retries=20
):

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


# ======================================================
# 📊  WORKBOOK SESSION
# ======================================================

def create_workbook_session(
    token,
    user_email,
    file_path,
    max_retries=10
):
    """
    Opens a persistent Excel session on the OneDrive file.
    persistChanges=true means edits are saved to the actual file.

    Newly copied files sometimes take a few seconds
    before Excel APIs can access them.
    """

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
        f":/workbook/createSession"
    )

    for attempt in range(max_retries):

        response = requests.post(
            url,
            headers=_headers(token),
            json={"persistChanges": True}
        )

        if response.status_code == 201:

            session_id = response.json()["id"]

            print("\n✅ Workbook Session Created")

            return session_id

        print(
            f"\n⏳ Waiting for workbook availability "
            f"(Attempt {attempt + 1}/{max_retries})"
        )

        time.sleep(3)

    print(response.text)

    raise Exception(
        "Failed to create workbook session"
    )

def close_workbook_session(
    token,
    user_email,
    file_path,
    session_id
):

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
        f":/workbook/closeSession"
    )

    requests.post(
        url,
        headers=_headers(
            token,
            {"workbook-session-id": session_id}
        )
    )

    print("\n✅ Workbook Session Closed")


# ======================================================
# ✏️  CELL UPDATE
# ======================================================

def update_cell_range(
    token,
    user_email,
    file_path,
    session_id,
    sheet_name,
    cell_range,
    values
):

    headers = _headers(
        token,
        {"workbook-session-id": session_id}
    )

    encoded_sheet = requests.utils.quote(sheet_name)

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}"
        f":/workbook/worksheets/{encoded_sheet}"
        f"/range(address='{cell_range}')"
    )

    response = requests.patch(
        url,
        headers=headers,
        json={"values": values}
    )

    if response.status_code != 200:
        print(response.text)

        raise Exception(
            f"Failed to update range {cell_range}"
        )


# ======================================================
# 📝  POPULATE EXCEL TEMPLATE
# ======================================================

def populate_excel_template(
    token,
    user_email,
    file_path,
    session_id,
    closing_ticket_df,
    invoice_df
):

    sheet = "Closing Check Transmittal Form"

    row = closing_ticket_df.iloc[0]

    purchase_price = row.get("cr109_saleprice", "")

    closing_date = row.get("cr7de_closingdate", "")

    seller_tcode = row.get("cr7de_sellertcode", "")

    property_address = row.get(
        "cr7de_buildingaddress",
        ""
    )

    unit = row.get("cr7de_unitnumber", "")

    seller1_name = row.get("cr7de_sellername", "")

    deal = TRANSACTION_TYPE_DEAL_MAP.get(
        row.get("cr109_transactiontypedeal", ""),
        ""
    )

    buyer1_name = row.get("cr7de_buyername", "")

    shares = row.get("cr109_shares", "")

    closing_agent = row.get(
        "cr7de_closingagentname",
        ""
    )

    closing_agent_phone = row.get(
        "cr7de_closingagentphone",
        ""
    )

    closing_agent_email = row.get(
        "cr7de_closingagentemail",
        ""
    )

    closing_agent_title = row.get(
        "cr7de_titlerole",
        ""
    )

    notes = row.get("cr7de_notes", "")

    current_date = datetime.now().strftime("%m/%d/%Y")

    try:
        if closing_date:
            closing_date = (
                pd.to_datetime(closing_date)
                .strftime("%m/%d/%Y")
            )

    except Exception:
        pass

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

    # ==================================================
    # HEADER
    # ==================================================

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

    # ==================================================
    # SELLER TABLE
    # ==================================================

    seller_df = invoice_df[
        invoice_df["cr7de_paidby"] == 716070000
    ]

    for idx, (_, inv_row) in enumerate(
        seller_df.iterrows()
    ):

        r = 15 + idx

        due_at_closing = DUE_AT_CLOSING_MAP.get(
            inv_row.get("cr109_dueatclosing", ""),
            ""
        )

        payable_to = PAYABLE_MAP.get(
            inv_row.get("cr7de_payableto", ""),
            ""
        )

        patch(f"A{r}:D{r}", [[
            inv_row.get("cr7de_chequenumber", ""),
            due_at_closing,
            inv_row.get("cr7de_amount", ""),
            payable_to
        ]])

    # ==================================================
    # BUYER TABLE
    # ==================================================

    buyer_df = invoice_df[
        invoice_df["cr7de_paidby"] == 716070001
    ]

    for idx, (_, inv_row) in enumerate(
        buyer_df.iterrows()
    ):

        r = 45 + idx

        due_at_closing = DUE_AT_CLOSING_MAP.get(
            inv_row.get("cr109_dueatclosing", ""),
            ""
        )

        payable_to = PAYABLE_MAP.get(
            inv_row.get("cr7de_payableto", ""),
            ""
        )

        patch(f"A{r}:D{r}", [[
            inv_row.get("cr7de_chequenumber", ""),
            due_at_closing,
            inv_row.get("cr7de_amount", ""),
            payable_to
        ]])

    print("\n✅ Excel Populated Successfully")


# ======================================================
# 📄  CONVERT EXCEL TO PDF
# ======================================================

def convert_onedrive_file_to_pdf(
    token,
    user_email,
    file_path,
    pdf_output_path
):

    url = (
        f"{GRAPH_BASE}/users/{user_email}"
        f"/drive/root:/{file_path}:/content"
        f"?format=pdf"
    )

    response = requests.get(
        url,
        headers=_headers(token)
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