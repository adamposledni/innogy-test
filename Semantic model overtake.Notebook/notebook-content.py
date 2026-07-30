# Fabric notebook source

# METADATA ********************

# META {
# META   "kernel_info": {
# META     "name": "jupyter",
# META     "jupyter_kernel_name": "python3.11"
# META   },
# META   "dependencies": {
# META     "lakehouse": {
# META       "default_lakehouse": "ba07d60c-b568-4186-b230-75203072eec3",
# META       "default_lakehouse_name": "Lakehouse",
# META       "default_lakehouse_workspace_id": "18b1a08f-f304-4d84-9085-be0c8e4d692a",
# META       "known_lakehouses": [
# META         {
# META           "id": "ba07d60c-b568-4186-b230-75203072eec3"
# META         }
# META       ]
# META     }
# META   }
# META }

# MARKDOWN ********************

# ### Overtake semantic models
# This notebooks overtakes all semantic models in all workspaces for specified domains.
# 
# Each domain has dedicated service principal account.
# 
# The mapping is controlled by XLSX definition file.
# 
# #### Deployment steps
# 1. Create **XLSX definition file** (with specified structure) in any SharePoint document library folder.
# 
# 1. Create **lakehouse** in the same workspace as this notebook.
# 
# 1. Create **variable library** called "Variables" in the same workspace as this notebook.
# 
# 1. Create variables:
#     - tenant_id
#     - key_vault_url
#     - semantic_models_management_xlsx_relative_file_path
#     - semantic_models_ownerships_sheet_name
# 
# 1. Create **shortcut** to the SharePoint document library folder in the lakehouse.
# 
# 1. Set the lakehouse as **default lakehouse** for this notebook.
# 
# 1. Update values in the **parameters cell** in this notebook.
# 
# 1. Create **schedule with failure notification** for this notebook.

# MARKDOWN ********************

# ### Setup

# CELL ********************

import pandas as pd
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

%run FabricApiClient

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Parameters

# CELL ********************

variables = nu.variableLibrary.getLibrary("Variables")

tenant_id = variables.tenant_id
key_vault_url = variables.key_vault_url 
semantic_models_management_xlsx_relative_file_path = variables.semantic_models_management_xlsx_relative_file_path 
semantic_models_ownerships_sheet_name = variables.semantic_models_ownerships_sheet_name 
semantic_models_ownerships_notification_email = variables.semantic_models_ownerships_notification_email

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Run

# CELL ********************

# initialize OBO Fabric API client
obo_fabric_api_client = FabricApiClient()

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# process definition XLSX file

domain_fabric_api_clients = {}
domains_to_process = set()

df_definition = pd.read_excel(
    pd.ExcelFile(f"/lakehouse/default/{semantic_models_management_xlsx_relative_file_path}"), 
    sheet_name = semantic_models_ownerships_sheet_name, 
    dtype = str
)

for row in df_definition.itertuples():
    domain_id = row[1]
    if not domain_id:
        continue

    client_id = row[2]
    client_secret_key_name = row[3]
    client_secret = nu.credentials.getSecret(key_vault_url, client_secret_key_name)

    cc_fabric_api_client = FabricApiClient(client_id, client_secret, tenant_id)
    
    domains_to_process.add(domain_id)
    domain_fabric_api_clients[domain_id] = cc_fabric_api_client

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# get relevant workspaces

workspaces = obo_fabric_api_client.get_paginated("https://api.fabric.microsoft.com/v1/workspaces")
print(f"🔍 Trace | {len(workspaces)} workspaces")

workspaces = [w for w in workspaces if w.get("domainId") in domains_to_process]
print(f"🔍 Trace | {len(workspaces)} workspaces from relevant domains")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# get relavant semantic models

semantic_models = []
for workspace in workspaces:
    workspace_id = workspace["id"]
    domain_id = workspace["domainId"]
    workspace_name = workspace["displayName"]

    semantic_models_tmp = obo_fabric_api_client.get_paginated(f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets")
    semantic_models_tmp = [{**s, "domainId": domain_id, "workspaceName": workspace_name, "workspaceId": workspace_id} for s in semantic_models_tmp]
    
    semantic_models.extend(semantic_models_tmp)

print(f"🔍 Trace | {len(semantic_models)} semantic models")

semantic_models_to_overtake = semantic_models

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

nu.fs.put(
    "/lakehouse/default/Files/semantic_models_to_overtake.json", 
    json.dumps(semantic_models_to_overtake),
    overwrite = True
)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

semantic_models_overtake_results = []

for semantic_model in semantic_models_to_overtake:
    domain_id = semantic_model["domainId"]
    semantic_model_id = semantic_model["id"]
    workspace_id = semantic_model["workspaceId"]
    configured_by = semantic_model["configuredBy"]

    cc_fabric_api_client = domain_fabric_api_clients[domain_id]
    if not cc_fabric_api_client:
        raise Exception("No Fabric API client for specified domain")

    cc_fabric_api_client_id = cc_fabric_api_client.get_client_id()

    status = "Succeeded"
    message = None

    try:
        print(f"🔍 Trace | Taking over semantic model {semantic_model_id} in workspace {workspace_id}")
        cc_fabric_api_client.request("POST", f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{semantic_model_id}/Default.TakeOver")
    except Exception as e:
        status = "Failed"
        message = str(e)
        print(f"🔍 Trace | Error: {message}")

    semantic_models_overtake_results.append({
        "semantic_model": semantic_model,
        "status": status,
        "message": message 
    })

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

semantic_models_overtake_results_fails = [
    d for d in semantic_models_overtake_results
    if d["status"] == "Failed"
]

notifications = []

if semantic_models_overtake_results_fails:
    notification_messages = [
        f"🔴 | {d['semantic_model']['workspaceName']} {d['semantic_model']['name']}" 
        for d in semantic_models_overtake_results_fails
    ]

    notification_body = f"""
        <p>
            {'<br/>'.join(notification_messages)}
        </p>
    """
    notifications.append({
        "subject": "Sematic model overtake failed",
        "body": notification_body,
        "email": semantic_models_ownerships_notification_email
    })

nu.notebook.exit(json.dumps(notifications))

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }
