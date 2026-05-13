import traceback

from msal import ConfidentialClientApplication

from core.config import config
from core.logger import logger

auth = config["auth"]


def get_access_token(scope):

    try:

        authority_url = (
            f"https://login.microsoftonline.com/"
            f"{auth['tenant_id']}"
        )

        app = ConfidentialClientApplication(
            auth["client_id"],
            authority=authority_url,
            client_credential=auth["client_secret"]
        )

        token_response = app.acquire_token_for_client(
            scopes=[scope]
        )

        if "access_token" not in token_response:
            raise Exception(token_response)

        return token_response["access_token"]

    except Exception as e:

        logger.error(traceback.format_exc())
        print(e)

        return None
