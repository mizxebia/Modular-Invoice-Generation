import os
import pandas as pd

from core.config import config
from core.auth import get_access_token

from services.dataverse_service import fetch_table

from services.graph_service import (
    create_onedrive_folder,
    upload_file_to_onedrive
)

from services.excel_service import (
    setup_ticket_folders,
    archive_active_file_if_exists,
    copy_template_to_active,
    create_workbook_session,
    populate_excel_template,
    close_workbook_session,
    convert_onedrive_file_to_pdf
)


def main():

    # ==============================
    # 🔐 TOKENS
    # ==============================

    dv_token = get_access_token(
        config["auth"]["dataverse_scope"]
    )

    graph_token = get_access_token(
        config["auth"]["graph_scope"]
    )

    # ==============================
    # 📄 CONFIG
    # ==============================

    dv = config["dataverse"]
    storage = config["storage"]

    ticket_column = dv["columns"]["ticket_id"]
    ticket_value  = dv["filter"]["ticket_id"]

    template_path = (
        storage["paths"]["invoice_template_path"]
    )

    user_email = storage["user_email_dev"]

    output_excel_name = (
        f"{ticket_value}_Closing_Form.xlsx"
    )

    pdf_output_path = (
        f"{ticket_value}_Closing_Form.pdf"
    )

    # ==============================
    # 📥 FETCH DATAVERSE
    # ==============================

    closing_data = fetch_table(
        dv["tables"]["closing_ticket_details"],
        dv_token,
        ticket_column,
        ticket_value
    )

    invoice_data = fetch_table(
        dv["tables"]["invoice_details"],
        dv_token,
        ticket_column,
        ticket_value
    )

    df_closing = pd.DataFrame(closing_data)
    df_invoice = pd.DataFrame(invoice_data)

    # ==============================
    # 📁 ENSURE FOLDER STRUCTURE
    #
    # Creates any missing folders:
    #   InProgress/TICKET_ID/
    #   InProgress/TICKET_ID/Active/
    #   InProgress/TICKET_ID/Inactive/
    # ==============================

    active_folder, inactive_folder = setup_ticket_folders(
        graph_token,
        user_email,
        ticket_value
    )

    # ==============================
    # 📦 ARCHIVE EXISTING FILE
    #
    # If Active/ already has a file:
    #   → move it to Inactive/YYYY-MM-DD/
    #   → rename with datetime suffix
    # If Active/ is empty → skip
    # ==============================

    archive_active_file_if_exists(
        graph_token,
        user_email,
        active_folder,
        inactive_folder,
        output_excel_name
    )

    # ==============================
    # 📋 COPY TEMPLATE → ACTIVE/
    #
    # Server-side copy — dropdowns,
    # images, styles preserved 100%
    # ==============================

    output_file_path = copy_template_to_active(
        graph_token,
        user_email,
        template_path,
        active_folder,
        output_excel_name
    )

    # ==============================
    # 📊 OPEN WORKBOOK SESSION
    # ==============================

    session_id = create_workbook_session(
        graph_token,
        user_email,
        output_file_path
    )

    # ==============================
    # 📝 POPULATE EXCEL
    # Session always closed even
    # if population fails
    # ==============================

    try:

        populate_excel_template(
            graph_token,
            user_email,
            output_file_path,
            session_id,
            df_closing,
            df_invoice
        )

    finally:

        close_workbook_session(
            graph_token,
            user_email,
            output_file_path,
            session_id
        )

    # ==============================
    # 📄 CONVERT TO PDF
    # ==============================

    convert_onedrive_file_to_pdf(
        graph_token,
        user_email,
        output_file_path,
        pdf_output_path
    )

    # ==============================
    # ☁️ UPLOAD PDF → ACTIVE/
    # ==============================

    upload_file_to_onedrive(
        graph_token,
        user_email,
        f"{ticket_value}/Active",
        pdf_output_path,
        pdf_output_path
    )

    # ==============================
    # 🧹 CLEANUP LOCAL PDF
    # ==============================

    if os.path.exists(pdf_output_path):
        os.remove(pdf_output_path)

    print("\n✅ PROCESS COMPLETED")


if __name__ == "__main__":

    main()
