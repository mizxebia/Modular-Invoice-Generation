import requests

from core.config import config

auth = config["auth"]


def fetch_table(
    table_name,
    token,
    ticket_column,
    ticket_value
):

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    base_url = auth["dataverse_url"]

    url = (
        f"{base_url}/api/data/v9.2/{table_name}"
        f"?$filter={ticket_column} eq '{ticket_value}'"
    )

    all_records = []

    while url:

        response = requests.get(
            url,
            headers=headers
        )

        if response.status_code != 200:
            raise Exception(response.text)

        data = response.json()

        all_records.extend(
            data.get("value", [])
        )

        url = data.get("@odata.nextLink")

    return all_records
