# Fabric notebook source

# METADATA ********************

# META {
# META   "kernel_info": {
# META     "name": "jupyter",
# META     "jupyter_kernel_name": "python3.11"
# META   },
# META   "dependencies": {}
# META }

# CELL ********************

from datetime import datetime, timedelta
import base64
import json
import requests
import notebookutils as nu

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# Fabric API client
class FabricApiClient:
    def __init__(self, client_id=None, client_secret=None, tenant_id=None):
        self.__client_id = client_id
        self.__client_secret = client_secret
        self.__tenant_id = tenant_id
        self.__use_client_credentials = bool(client_id and client_secret and tenant_id)

        self.__token_url = (
            f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token" if tenant_id else None
        )
        self.__resource = "https://api.fabric.microsoft.com"
        # self.__resource = "https://analysis.windows.net/powerbi/api"

        self.__access_token = None
        self.__access_token_expiration = None

    def get_client_id(self):
        return self.__client_id

    def __decode_token_expiration(self, token):
        payload_segment = token.split(".")[1]
        padded = payload_segment + "=" * (-len(payload_segment) % 4)
        payload = json.loads(base64.urlsafe_b64decode(padded))

        return datetime.utcfromtimestamp(payload["exp"])

    def __get_access_token(self):
        if self.__access_token and self.__access_token_expiration > datetime.utcnow():
            return self.__access_token

        if self.__use_client_credentials:
            body = {
                "client_id": self.__client_id,
                "client_secret": self.__client_secret,
                "scope": f"{self.__resource}/.default",
                "grant_type": "client_credentials",
            }
            response = requests.post(self.__token_url, data=body)
            self.__handle_response(response)
            self.__access_token = response.json()["access_token"]
        else:
            self.__access_token = notebookutils.credentials.getToken("pbi")
        self.__access_token_expiration = self.__decode_token_expiration(self.__access_token) - timedelta(seconds=30)
        return self.__access_token

    def __handle_response(self, response):
        if response.status_code >= 400:
            print(response.headers.get("WWW-Authenticate"))
            print(self.__access_token)

            raise Exception(f"HTTP {response.status_code} on {response.request.method} {response.request.url}: {response.text}")

        content_type = response.headers.get("Content-Type", "")
        response_body = response.json() if "json" in content_type else response.content

        return {
            "status_code": response.status_code,
            "body": response_body,
            "headers": response.headers,
        }

    def request(self, method, uri, body=None, headers=None):
        token = self.__get_access_token()

        final_headers = {"Authorization": f"Bearer {token}"}
        if headers:
            final_headers.update(headers)

        response = requests.request(method, uri, json=body, headers=final_headers)
        return self.__handle_response(response)

    def get_paginated(self, uri, headers=None):
        results = []
        next_uri = uri

        while next_uri:
            response = self.request("GET", next_uri, headers=headers)
            data = response["body"]

            results.extend(data.get("value", []))
            next_uri = data.get("@odata.nextLink", None) or data.get("continuationUri", None)

        return results

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

obo_fabric_api_client = FabricApiClient()

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

TARGET_CAPACITY_ID = "37982985-f05e-453e-9ea7-fd8fd1b8a7ee"

# AssignWorkspaces doesn't document a max array size for workspacesToAssign, so
# batch conservatively. Adjust if the service rejects/accepts a different size.
ASSIGN_BATCH_SIZE = 500

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

def get_personal_workspaces(client):
    # The admin/groups endpoint paginates via $skip, not @odata.nextLink/continuationUri,
    # so FabricApiClient.get_paginated doesn't apply here - page manually instead.
    workspaces = []
    page_size = 5000  # max allowed value for $top
    skip = 0

    while True:
        print(f"🔍 Trace | Listing personal workspaces, skip={skip}")
        uri = (
            "https://api.powerbi.com/v1.0/myorg/admin/groups"
            f"?$filter=type eq 'PersonalGroup'&$top={page_size}&$skip={skip}"
        )
        response = client.request("GET", uri)
        page = response["body"].get("value", [])
        workspaces.extend(page)

        if len(page) < page_size:
            break
        skip += page_size

    return workspaces

personal_workspaces = get_personal_workspaces(obo_fabric_api_client)
print(f"🔍 Trace | Found {len(personal_workspaces)} personal workspaces")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# capacityId is already included on each workspace from the listing above, so
# skipping workspaces already on the target capacity requires no extra API call.
workspace_ids_to_assign = [
    workspace["id"]
    for workspace in personal_workspaces
    if workspace.get("capacityId") != TARGET_CAPACITY_ID
]

print(f"🔍 Trace | {len(workspace_ids_to_assign)} of {len(personal_workspaces)} personal workspaces need reassignment")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

def chunked(items, size):
    for i in range(0, len(items), size):
        yield items[i:i + size]

batches = list(chunked(workspace_ids_to_assign, ASSIGN_BATCH_SIZE))
print(f"🔍 Trace | Assigning {len(workspace_ids_to_assign)} workspaces to capacity {TARGET_CAPACITY_ID} in {len(batches)} batch(es)")

for batch_number, batch in enumerate(batches, start=1):
    print(f"🔍 Trace | Assigning batch {batch_number}/{len(batches)} ({len(batch)} workspaces)")
    request_body = {
        "capacityMigrationAssignments": [
            {
                "targetCapacityObjectId": TARGET_CAPACITY_ID,
                "workspacesToAssign": batch
            }
        ]
    }
    obo_fabric_api_client.request(
        "POST",
        "https://api.powerbi.com/v1.0/myorg/admin/capacities/AssignWorkspaces",
        request_body
    )
    print(f"🔍 Trace | Assigned batch {batch_number}/{len(batches)}")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }
