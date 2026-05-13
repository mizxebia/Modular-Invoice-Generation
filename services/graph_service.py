import requests


def get_onedrive_file(
    token,
    user_email,
    file_path
):

    headers = {
        "Authorization": f"Bearer {token}"
    }

    url = (
        f"https://graph.microsoft.com/v1.0/"
        f"users/{user_email}"
        f"/drive/root:/{file_path}:/"
    )

    response = requests.get(
        url,
        headers=headers
    )

    if response.status_code != 200:
        raise Exception(response.text)

    return response.json()


def create_onedrive_folder(
    token,
    user_email,
    ticket_id
):

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    parent_path = (
        "New Sales RPA/DEV/"
        "BotShareDrive/InProgress"
    )

    url = (
        f"https://graph.microsoft.com/v1.0/"
        f"users/{user_email}"
        f"/drive/root:/{parent_path}:/children"
    )

    body = {
        "name": ticket_id,
        "folder": {},
        "@microsoft.graph.conflictBehavior": "replace"
    }

    response = requests.post(
        url,
        headers=headers,
        json=body
    )

    if response.status_code not in [200, 201]:
        raise Exception(response.text)

    print("✅ Folder Created Successfully")


def upload_file_to_onedrive(
    token,
    user_email,
    ticket_id,
    local_file_path,
    onedrive_file_name
):

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }

    onedrive_folder = (
        f"New Sales RPA/DEV/"
        f"BotShareDrive/InProgress/"
        f"{ticket_id}"
    )

    url = (
        f"https://graph.microsoft.com/v1.0/"
        f"users/{user_email}"
        f"/drive/root:/{onedrive_folder}/"
        f"{onedrive_file_name}:/content"
    )

    with open(local_file_path, "rb") as f:
        file_content = f.read()

    response = requests.put(
        url,
        headers=headers,
        data=file_content
    )

    if response.status_code not in [200, 201]:
        raise Exception(response.text)

    print("✅ File Uploaded Successfully")
