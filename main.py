import pandas as pd

from core.config import config
from core.auth import get_access_token

from services.dataverse_service import fetch_table
from services.graph_service import (
    get_onedrive_file,
    create_onedrive_folder,
    upload_file_to_onedrive
)
from services.excel_service import (
    populate_excel_template,
    download_template
)


def main():

    dv_token = get_access_token(
        config["auth"]["dataverse_scope"]
    )

    graph_token = get_access_token(
        config["auth"]["graph_scope"]
    )

    dv = config["dataverse"]
    storage = config["storage"]

    ticket_column = dv["columns"]["ticket_id"]
    ticket_value = dv["filter"]["ticket_id"]

    template_path = (
        storage["paths"]["invoice_template_path"]
    )

    user_email = storage["user_email_dev"]

    template_file = get_onedrive_file(
        graph_token,
        user_email,
        template_path
    )

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

    local_template = download_template(
        template_file
    )

    output_file = (
        f"{ticket_value}_Closing_Form.xlsx"
    )

    populate_excel_template(
        local_template,
        output_file,
        df_closing,
        df_invoice
    )

    create_onedrive_folder(
        graph_token,
        user_email,
        ticket_value
    )

    upload_file_to_onedrive(
        graph_token,
        user_email,
        ticket_value,
        output_file,
        output_file
    )

    print("✅ PROCESS COMPLETED")


if __name__ == "__main__":
    main()
