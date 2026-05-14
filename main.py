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
    copy_template_on_onedrive,
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

    ticket_column = (
        dv["columns"]["ticket_id"]
    )

    ticket_value = (
        dv["filter"]["ticket_id"]
    )

    template_path = (
        storage["paths"]
        ["invoice_template_path"]
    )

    user_email = (
        storage["user_email_dev"]
    )

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
        dv["tables"]
        ["closing_ticket_details"],
        dv_token,
        ticket_column,
        ticket_value
    )

    invoice_data = fetch_table(
        dv["tables"]
        ["invoice_details"],
        dv_token,
        ticket_column,
        ticket_value
    )

    # ==============================
    # 📄 DATAFRAMES
    # ==============================

    df_closing = pd.DataFrame(
        closing_data
    )

    df_invoice = pd.DataFrame(
        invoice_data
    )

    # ==============================
    # 📁 CREATE FOLDER
    # ==============================

    create_onedrive_folder(
        graph_token,
        user_email,
        ticket_value
    )

    # ==============================
    # 📋 COPY TEMPLATE (SERVER-SIDE)
    # No download — preserves dropdowns,
    # images, and styles 100%
    # ==============================

    output_file_path = copy_template_on_onedrive(
        graph_token,
        user_email,
        template_path,
        ticket_value,
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
    # Session is always closed even
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
    # ☁️ UPLOAD PDF
    # ==============================

    upload_file_to_onedrive(
        graph_token,
        user_email,
        ticket_value,
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
