import json
from datetime import datetime, timedelta, timezone
import base64
import requests
from azure.keyvault.secrets import SecretClient
import fabric.functions as fn

udf = fn.UserDataFunctions()

IN_PROGRESS_STATUSES = {"unknown", "inprogress", "notstarted"}


class FabricApiClient:
    def __init__(self, client_id, client_secret, tenant_id):
        self.__client_id = client_id
        self.__client_secret = client_secret
        self.__tenant_id = tenant_id

        self.__access_token = None
        self.__access_token_expiration = None

    def __decode_token_expiration(self, token):
        payload_segment = token.split(".")[1]
        padded = payload_segment + "=" * (-len(payload_segment) % 4)
        payload = json.loads(base64.urlsafe_b64decode(padded))
        return datetime.fromtimestamp(payload["exp"], tz=timezone.utc)

    def __get_access_token(self):
        if self.__access_token and self.__access_token_expiration > datetime.now(timezone.utc):
            return self.__access_token

        body = {
            "client_id": self.__client_id,
            "client_secret": self.__client_secret,
            "scope": "https://analysis.windows.net/powerbi/api/.default",
            "grant_type": "client_credentials",
        }
        response = requests.post(
            f"https://login.microsoftonline.com/{self.__tenant_id}/oauth2/v2.0/token",
            data=body,
        )

        self.__handle_response(response)
        self.__access_token = response.json()["access_token"]

        self.__access_token_expiration = self.__decode_token_expiration(self.__access_token) - timedelta(seconds=30)
        return self.__access_token

    def __handle_response(self, response):
        if response.status_code >= 400:
            raise Exception(f"HTTP {response.status_code} on {response.request.method} {response.request.url}: {response.text}")
        content_type = response.headers.get("Content-Type", "")
        response_body = response.json() if "json" in content_type else response.content
        return {"status_code": response.status_code, "body": response_body, "headers": response.headers}

    def request(self, method, uri, body=None, headers=None):
        token = self.__get_access_token()
        final_headers = {"Authorization": f"Bearer {token}"}
        if headers:
            final_headers.update(headers)
        response = requests.request(method, uri, json=body, headers=final_headers)
        return self.__handle_response(response)





def parse_items(serialized_items):
    # eg. <workspace_id>,<item_id>,<type>;<workspace_id>,<item_id>,<type>
    
    max_items = 25
    raw_items = serialized_items.split(";")
    if len(raw_items) > max_items:
        raise ValueError(f"Too many items: {len(raw_items)} (max {max_items})")

    for item in raw_items:
        workspace_id, item_id, item_type = item.split(",")
        yield {"workspace_id": workspace_id, "item_id": item_id, "item_type": item_type}


def _refresh_in_progress(fabric_api_client, item):
    workspace_id, item_id, item_type = item["workspace_id"], item["item_id"], item["item_type"]

    if item_type == "dataflow":
        values = fabric_api_client.request(
            "GET",
            f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/dataflows/{item_id}/transactions?$top=1",
        )["body"].get("value", [])
        return bool(values) and str(values[0].get("status", "")).lower() in IN_PROGRESS_STATUSES

    values = fabric_api_client.request(
        "GET",
        f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{item_id}/refreshes?$top=1",
    )["body"].get("value", [])

    if not values:
        return False
    last = values[0]
    return last.get("endTime") is None or str(last.get("status", "")).lower() in IN_PROGRESS_STATUSES


def _start_refresh(fabric_api_client, item):
    workspace_id, item_id, item_type = item["workspace_id"], item["item_id"], item["item_type"]

    if item_type == "dataflow":
        fabric_api_client.request(
            "POST",
            f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/dataflows/{item_id}/refreshes?processType=default",
            body={"notifyOption": "NoNotification"},
        )
    else:
        fabric_api_client.request(
            "POST",
            f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{item_id}/refreshes",
            body={"type": "Full", "commitMode": "transactional"},
        )


def refresh_items(fabric_api_client, serialized_items):
    started, already_running, errors = 0, 0, []
    for item in parse_items(serialized_items):
        try:
            if _refresh_in_progress(fabric_api_client, item):
                already_running += 1
                continue
            _start_refresh(fabric_api_client, item)
            started += 1
        except Exception as exc:
            errors.append({"item_id": item["item_id"], "error": str(exc)})
    return {
        "started": started,
        "alreadyInProgress": already_running,
        "failed": len(errors),
        "errors": errors,
    }


@udf.generic_connection(argName="kv", audienceType="KeyVault")
@udf.connection(argName="var", alias="Variables")
@udf.function()
def trigger_refreshes(items: str, kv: fn.FabricItem, var: fn.FabricVariablesClient) -> str:
    variables = var.getVariables()

    tenant_id = variables.get("tenant_id")
    key_vault_url = variables.get("key_vault_url")
    client_id = variables.get("refresh_client_id")
    client_secret_key = variables.get("refresh_client_secret_key")

    client = SecretClient(vault_url=key_vault_url, credential=kv.get_access_token())
    client_secret = client.get_secret(client_secret_key).value

    fabric_api_client = FabricApiClient(client_id, client_secret, tenant_id)
    result = refresh_items(fabric_api_client, items)
    return json.dumps(result)