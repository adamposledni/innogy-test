# Fabric notebook source

# METADATA ********************

# META {
# META   "kernel_info": {
# META     "name": "jupyter",
# META     "jupyter_kernel_name": "python3.11"
# META   },
# META   "dependencies": {}
# META }

# MARKDOWN ********************

# ### Deploy solution
# Downloads a packaged copy of this solution (Lakehouse, notebooks, pipelines, variable library) from
# blob storage and creates the requested items in **the workspace this notebook runs in**. Already existing
# items (matched by type + display name) are left untouched. Cross-item references (pipeline -> notebook,
# notebook -> default lakehouse) are rebound to the newly created items' real IDs.
#
# #### One-time prerequisites (customer tenant)
# 1. Package this repository into a zip (everything except this notebook's own item folder) and upload it
#    to a blob storage container, then generate a read-only, time-limited **SAS URL** for that blob.
# 2. Import this notebook manually into the target Fabric workspace.
# 3. Optional, only if you want these wired automatically instead of by hand afterwards:
#     - Create a **OneDrive/SharePoint connection** in the Fabric portal (for the Lakehouse's SharePoint shortcut).
#     - Create an **Office365 Outlook connection** in the Fabric portal (for the pipelines' failure e-mails).
#
# #### Usage
# 1. Fill in the **Parameters** cell below.
# 2. Run all cells.
# 3. Check the trace output and the **Manual follow-up steps** printed at the end.

# MARKDOWN ********************

# ### Setup

# CELL ********************

import base64
import io
import json
import os
import re
import tempfile
import time
import zipfile
from datetime import datetime, timedelta

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

fabric_api_client = FabricApiClient()

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Parameters

# CELL ********************

# ── Target ────────────────────────────────────────────────────────────────
# Items are always created in the workspace this notebook runs in.
target_workspace_id = nu.runtime.context["currentWorkspaceId"]

# ── Artifact source ──────────────────────────────────────────────────────
# Full blob URL including the SAS query string (e.g. from Azure Portal "Generate SAS" on the zip blob).
artifact_zip_url = ""

# ── Components to deploy (toggle True/False) ──────────────────────────────
deploy_variables = True
deploy_lakehouse = True
deploy_semantic_model_overtake = True
deploy_semantic_model_refresh_monitor = True
deploy_switch_personal_workspace_capacity = True

# ── Variable Library overrides ─────────────────────────────────────────────
# Values baked into the "Variables" library at creation time. Leave a value blank to keep whatever is
# already in the packaged variables.json (fine for non-tenant-specific defaults, not for the blank ones below).
variable_overrides = {
    "tenant_id": "",
    "key_vault_url": "",
    "semantic_models_management_xlsx_relative_file_path": "",
    "semantic_models_ownerships_sheet_name": "",
    "semantic_models_refreshes_sheet_name": "",
    "semantic_models_ownerships_notification_email": "",
    "semantic_models_refreshes_notification_email": "",
}

# ── Switch personal workspace capacity ─────────────────────────────────────
# Required if deploy_switch_personal_workspace_capacity is True - the Fabric capacity to assign personal
# workspaces to in this tenant.
target_capacity_id = ""

# ── Optional Fabric connections (create these once via the Fabric portal, "Manage connections and gateways") ─
# Leave blank to skip the corresponding wiring - it can always be configured by hand afterwards.
sharepoint_connection_id = ""
sharepoint_site_url = ""
sharepoint_folder_subpath = ""
office365_connection_id = ""

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Download artifact

# CELL ********************

if not artifact_zip_url:
    raise Exception("artifact_zip_url is required")

print("🔍 Trace | Downloading artifact from blob storage")
response = requests.get(artifact_zip_url)
response.raise_for_status()
print(f"🔍 Trace | Downloaded {len(response.content):,} bytes")

extract_dir = os.path.join(tempfile.mkdtemp(), "artifact")
with zipfile.ZipFile(io.BytesIO(response.content)) as zip_file:
    zip_file.extractall(extract_dir)

print(f"🔍 Trace | Extracted artifact to {extract_dir}")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# discover every item folder in the artifact (any directory containing a .platform file)

items_by_key = {}

for root, dirs, files in os.walk(extract_dir):
    if ".platform" not in files:
        continue

    with open(os.path.join(root, ".platform"), "r", encoding="utf-8") as f:
        platform = json.load(f)

    item_type = platform["metadata"]["type"]
    display_name = platform["metadata"]["displayName"]
    logical_id = platform["config"]["logicalId"]

    items_by_key[(item_type, display_name)] = {
        "folder": root,
        "logicalId": logical_id,
    }
    print(f"🔍 Trace | Discovered {item_type}: '{display_name}' (logicalId={logical_id})")

    dirs[:] = []  # don't recurse into item internals

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# map the 5 solution components to the (type, displayName) pairs they own, and apply the toggles

components = {
    "Variables": {
        "enabled": deploy_variables,
        "items": [("VariableLibrary", "Variables")],
    },
    "Lakehouse": {
        "enabled": deploy_lakehouse,
        "items": [("Lakehouse", "Lakehouse")],
    },
    "Semantic model overtake": {
        "enabled": deploy_semantic_model_overtake,
        "items": [("Notebook", "Semantic model overtake"), ("DataPipeline", "Semantic model overtake")],
    },
    "Semantic model refresh monitor": {
        "enabled": deploy_semantic_model_refresh_monitor,
        "items": [("Notebook", "Semantic model refresh monitor"), ("DataPipeline", "Semantic model refresh monitor")],
    },
    "Switch personal workspace capacity": {
        "enabled": deploy_switch_personal_workspace_capacity,
        "items": [("Notebook", "Switch personal workspace capacity")],
    },
}

enabled_keys = set()
for component_name, component in components.items():
    print(f"🔍 Trace | Component '{component_name}': {'enabled' if component['enabled'] else 'disabled (skipped)'}")
    if component["enabled"]:
        enabled_keys.update(component["items"])

existing_items = fabric_api_client.get_paginated(f"https://api.fabric.microsoft.com/v1/workspaces/{target_workspace_id}/items")
print(f"🔍 Trace | {len(existing_items)} item(s) already exist in the target workspace")

def find_existing_item(item_type, display_name):
    for item in existing_items:
        if item["type"] == item_type and item["displayName"] == display_name:
            return item
    return None

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Item creation helpers

# CELL ********************

def _resolve_lro(response):
    if response["status_code"] in (200, 201):
        return response["body"]

    if response["status_code"] != 202:
        raise Exception(f"Unexpected response status {response['status_code']}: {response['body']}")

    operation_url = response["headers"]["Location"]
    while True:
        time.sleep(3)
        operation = fabric_api_client.request("GET", operation_url)
        status = operation["body"]["status"]
        print(f"🔍 Trace | Operation status: {status}")

        if status == "Succeeded":
            try:
                result = fabric_api_client.request("GET", f"{operation_url}/result")
                return result["body"]
            except Exception:
                return None
        if status == "Failed":
            raise Exception(f"Operation failed: {operation['body']}")

def create_item(body):
    response = fabric_api_client.request("POST", f"https://api.fabric.microsoft.com/v1/workspaces/{target_workspace_id}/items", body)
    return _resolve_lro(response)

def build_parts_from_texts(files_by_relative_path):
    return [
        {
            "path": relative_path,
            "payload": base64.b64encode(text.encode("utf-8")).decode("utf-8"),
            "payloadType": "InlineBase64",
        }
        for relative_path, text in files_by_relative_path.items()
    ]

# id_map: logicalId (from the artifact) -> real item id in the target workspace
# created_items: (type, displayName) -> real item id
id_map = {}
created_items = {}

def deploy_item(item_type, display_name, build_parts_fn=None, creation_payload=None):
    key = (item_type, display_name)

    if key not in enabled_keys:
        print(f"⏭️ Skip | {item_type} '{display_name}' - component disabled")
        return None

    if key not in items_by_key:
        print(f"⚠️ Warn | {item_type} '{display_name}' - not found in artifact, skipping")
        return None

    existing = find_existing_item(item_type, display_name)
    if existing:
        print(f"⏭️ Skip | {item_type} '{display_name}' - already exists ({existing['id']})")
        item_id = existing["id"]
    else:
        print(f"🚀 Create | {item_type} '{display_name}'")
        body = {"displayName": display_name, "type": item_type}
        if build_parts_fn:
            body["definition"] = {"parts": build_parts_fn(items_by_key[key]["folder"])}
        if creation_payload:
            body["creationPayload"] = creation_payload
        item = create_item(body)
        item_id = item["id"]
        print(f"✅ Created | {item_type} '{display_name}' -> {item_id}")

    created_items[key] = item_id
    id_map[items_by_key[key]["logicalId"]] = item_id
    return item_id

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Rebinding helpers
# Notebooks store their default lakehouse binding in a JSON block wrapped in `# META` comment lines at the
# top of the file - `patch_notebook_default_lakehouse` rewrites that block in place. Pipelines reference
# notebooks by ID directly in their JSON - `rebind_pipeline_definition` walks the whole structure and swaps
# any `notebookId` it finds for the newly created notebook's real ID via `id_map`.

# CELL ********************

def _extract_meta_block(lines, header_idx):
    i = header_idx + 1
    while lines[i].strip() == "":
        i += 1

    block_start = i
    depth = 0
    raw_lines = []
    while True:
        content = lines[i][len("# META"):]
        if content.startswith(" "):
            content = content[1:]
        raw_lines.append(content.rstrip("\n"))
        depth += content.count("{") - content.count("}")
        if depth == 0 and content.strip():
            break
        i += 1

    return json.loads("\n".join(raw_lines)), block_start, i

def _replace_meta_block(lines, block_start, block_end, obj):
    new_lines = [f"# META {line}\n" if line else "# META\n" for line in json.dumps(obj, indent=2).split("\n")]
    return lines[:block_start] + new_lines + lines[block_end + 1:]

def patch_notebook_default_lakehouse(source_text, lakehouse_id, lakehouse_name, workspace_id):
    lines = source_text.splitlines(keepends=True)
    header_idx = next(i for i, line in enumerate(lines) if line.strip() == "# METADATA ********************")

    meta, block_start, block_end = _extract_meta_block(lines, header_idx)
    meta.setdefault("dependencies", {})["lakehouse"] = {
        "default_lakehouse": lakehouse_id,
        "default_lakehouse_name": lakehouse_name,
        "default_lakehouse_workspace_id": workspace_id,
        "known_lakehouses": [{"id": lakehouse_id}],
    }

    return "".join(_replace_meta_block(lines, block_start, block_end, meta))

def rebind_pipeline_definition(node, workspace_id, office365_connection_id):
    if isinstance(node, dict):
        if "notebookId" in node:
            old_id = node["notebookId"]
            new_id = id_map.get(old_id)
            if new_id:
                node["notebookId"] = new_id
                node["workspaceId"] = workspace_id
            else:
                print(f"⚠️ Warn | Could not rebind notebookId {old_id} - target notebook was not deployed")

        if "externalReferences" in node and "connection" in node["externalReferences"]:
            if office365_connection_id:
                node["externalReferences"]["connection"] = office365_connection_id
            else:
                print("⚠️ Warn | office365_connection_id not provided - pipeline e-mail step needs a manual connection binding")

        for value in node.values():
            rebind_pipeline_definition(value, workspace_id, office365_connection_id)
    elif isinstance(node, list):
        for value in node:
            rebind_pipeline_definition(value, workspace_id, office365_connection_id)

    return node

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Variables

# CELL ********************

def build_variables_parts(folder):
    with open(os.path.join(folder, ".platform"), "r", encoding="utf-8") as f:
        platform_text = f.read()
    with open(os.path.join(folder, "settings.json"), "r", encoding="utf-8") as f:
        settings_text = f.read()
    with open(os.path.join(folder, "variables.json"), "r", encoding="utf-8") as f:
        variables_obj = json.load(f)

    for variable in variables_obj["variables"]:
        override = variable_overrides.get(variable["name"])
        if override:
            variable["value"] = override

    missing = [v["name"] for v in variables_obj["variables"] if not v["value"]]
    if missing:
        print(f"⚠️ Warn | Variables left blank: {', '.join(missing)} - fill these in via the Fabric portal after deployment")

    return build_parts_from_texts({
        ".platform": platform_text,
        "settings.json": settings_text,
        "variables.json": json.dumps(variables_obj, indent=2),
    })

deploy_item("VariableLibrary", "Variables", build_variables_parts)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Lakehouse

# CELL ********************

def deploy_lakehouse():
    key = ("Lakehouse", "Lakehouse")

    if key not in enabled_keys:
        print("⏭️ Skip | Lakehouse 'Lakehouse' - component disabled")
        return None

    if key not in items_by_key:
        print("⚠️ Warn | Lakehouse 'Lakehouse' - not found in artifact, skipping")
        return None

    existing = find_existing_item("Lakehouse", "Lakehouse")
    if existing:
        print(f"⏭️ Skip | Lakehouse 'Lakehouse' - already exists ({existing['id']})")
        lakehouse_id = existing["id"]
    else:
        folder = items_by_key[key]["folder"]
        with open(os.path.join(folder, "lakehouse.metadata.json"), "r", encoding="utf-8") as f:
            lakehouse_metadata = json.load(f)

        creation_payload = {"enableSchemas": True} if lakehouse_metadata.get("defaultSchema") else None

        print("🚀 Create | Lakehouse 'Lakehouse'")
        body = {"displayName": "Lakehouse", "type": "Lakehouse"}
        if creation_payload:
            body["creationPayload"] = creation_payload

        try:
            item = create_item(body)
        except Exception as e:
            if not creation_payload:
                raise
            print(f"⚠️ Warn | Schema-enabled creation failed ({e}), retrying as a standard lakehouse")
            item = create_item({"displayName": "Lakehouse", "type": "Lakehouse"})

        lakehouse_id = item["id"]
        print(f"✅ Created | Lakehouse 'Lakehouse' -> {lakehouse_id}")

    created_items[key] = lakehouse_id
    id_map[items_by_key[key]["logicalId"]] = lakehouse_id

    if not sharepoint_connection_id:
        print("⚠️ Warn | sharepoint_connection_id not provided - Lakehouse shortcut(s) not created, add manually in the Fabric portal")
        return lakehouse_id

    folder = items_by_key[key]["folder"]
    shortcuts_path = os.path.join(folder, "shortcuts.metadata.json")
    if not os.path.exists(shortcuts_path):
        return lakehouse_id

    with open(shortcuts_path, "r", encoding="utf-8") as f:
        shortcuts = json.load(f)

    existing_shortcuts = fabric_api_client.request(
        "GET", f"https://api.fabric.microsoft.com/v1/workspaces/{target_workspace_id}/items/{lakehouse_id}/shortcuts"
    )
    existing_shortcut_names = {s["name"] for s in existing_shortcuts["body"].get("value", [])}

    for shortcut in shortcuts:
        if shortcut["name"] in existing_shortcut_names:
            print(f"⏭️ Skip | Shortcut '{shortcut['name']}' - already exists")
            continue

        print(f"🚀 Create | Shortcut '{shortcut['name']}'")
        shortcut_body = {
            "name": shortcut["name"],
            "path": shortcut["path"],
            "target": {
                "oneDriveSharePoint": {
                    "connectionId": sharepoint_connection_id,
                    "location": sharepoint_site_url,
                    "subpath": sharepoint_folder_subpath,
                }
            },
        }
        fabric_api_client.request(
            "POST",
            f"https://api.fabric.microsoft.com/v1/workspaces/{target_workspace_id}/items/{lakehouse_id}/shortcuts",
            shortcut_body,
        )
        print(f"✅ Created | Shortcut '{shortcut['name']}'")

    return lakehouse_id

lakehouse_item_id = deploy_lakehouse()

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Semantic model overtake & Semantic model refresh monitor
# Both components share the same notebook/pipeline shape: a notebook with a default lakehouse binding, and a
# pipeline that runs it and e-mails on failure.

# CELL ********************

def build_notebook_parts(folder, lakehouse_id):
    with open(os.path.join(folder, ".platform"), "r", encoding="utf-8") as f:
        platform_text = f.read()
    with open(os.path.join(folder, "notebook-content.py"), "r", encoding="utf-8") as f:
        notebook_text = f.read()

    if lakehouse_id:
        notebook_text = patch_notebook_default_lakehouse(notebook_text, lakehouse_id, "Lakehouse", target_workspace_id)
    else:
        print("⚠️ Warn | Lakehouse not deployed - notebook created without a default lakehouse binding, set it manually")

    return build_parts_from_texts({".platform": platform_text, "notebook-content.py": notebook_text})

def build_pipeline_parts(folder):
    with open(os.path.join(folder, ".platform"), "r", encoding="utf-8") as f:
        platform_text = f.read()
    with open(os.path.join(folder, "pipeline-content.json"), "r", encoding="utf-8") as f:
        pipeline_json = json.load(f)

    pipeline_json = rebind_pipeline_definition(pipeline_json, target_workspace_id, office365_connection_id)

    return build_parts_from_texts({
        ".platform": platform_text,
        "pipeline-content.json": json.dumps(pipeline_json, indent=2),
    })

deploy_item("Notebook", "Semantic model overtake", lambda folder: build_notebook_parts(folder, lakehouse_item_id))
deploy_item("DataPipeline", "Semantic model overtake", build_pipeline_parts)

deploy_item("Notebook", "Semantic model refresh monitor", lambda folder: build_notebook_parts(folder, lakehouse_item_id))
deploy_item("DataPipeline", "Semantic model refresh monitor", build_pipeline_parts)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Switch personal workspace capacity

# CELL ********************

def build_switch_capacity_notebook_parts(folder):
    with open(os.path.join(folder, ".platform"), "r", encoding="utf-8") as f:
        platform_text = f.read()
    with open(os.path.join(folder, "notebook-content.py"), "r", encoding="utf-8") as f:
        notebook_text = f.read()

    if not target_capacity_id:
        print("⚠️ Warn | target_capacity_id not set - deploying with the source tenant's capacity ID, update TARGET_CAPACITY_ID manually before running")
    else:
        match = re.search(r'TARGET_CAPACITY_ID = "[0-9a-fA-F-]+"', notebook_text)
        if not match:
            raise Exception("Could not locate TARGET_CAPACITY_ID literal in notebook source")
        notebook_text = notebook_text.replace(match.group(0), f'TARGET_CAPACITY_ID = "{target_capacity_id}"')

    return build_parts_from_texts({".platform": platform_text, "notebook-content.py": notebook_text})

deploy_item("Notebook", "Switch personal workspace capacity", build_switch_capacity_notebook_parts)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Summary

# CELL ********************

print("")
print("=" * 70)
print("Deployment summary")
print("=" * 70)

for component_name, component in components.items():
    if not component["enabled"]:
        print(f"⏭️ {component_name}: skipped (disabled)")
        continue

    item_ids = [created_items.get(item_key) for item_key in component["items"]]
    if all(item_ids):
        print(f"✅ {component_name}: deployed")
    else:
        print(f"⚠️ {component_name}: incomplete - {dict(zip(component['items'], item_ids))}")

print("")
print("Manual follow-up steps:")
if not sharepoint_connection_id:
    print("  - Create the SharePoint shortcut on the Lakehouse manually (no sharepoint_connection_id was provided)")
if not office365_connection_id:
    print("  - Bind the Office365 Outlook connection on both pipelines' 'Notify' e-mail activity manually")
print("  - Upload the semantic model ownership/refresh XLSX definition file to the lakehouse's SharePoint shortcut")
print("  - Create schedules (with failure notifications) for the deployed notebooks/pipelines")
print("  - Double-check the 'Variables' library values in the Fabric portal, especially any left blank above")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }
