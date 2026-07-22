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

request_body = {
  "capacityMigrationAssignments": [
    {
      "targetCapacityObjectId": "37982985-f05e-453e-9ea7-fd8fd1b8a7ee",
      "workspacesToAssign": ["69964bac-492a-4bf5-84cf-37a28765b960"]
    }
  ]
}

response = obo_fabric_api_client.request(
    "POST",
    "https://api.powerbi.com/v1.0/myorg/admin/capacities/AssignWorkspaces",
    request_body
)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }
