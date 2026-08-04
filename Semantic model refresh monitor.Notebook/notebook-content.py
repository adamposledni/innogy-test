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

# ### Monitor semantic model refreshes
# This notebook monitors the refresh history of semantic models that were overtaken by service principals (see the *Overtake semantic models* notebook).
# 
# Under service principal ownership, the built-in scheduled refresh failure email notifications do not work, so this notebook reads the refresh history through the API instead and produces its own notifications for any failed or disabled refresh (scheduled, manual, or via API).
# For each semantic model listed in `semantic_models_to_overtake.json`, it keeps track of the last refresh it has already checked, so every run only looks at refreshes that happened since the previous run.
# Notification emails are looked up in an XLSX mapping file (workspace id, semantic model id, notification emails). A semantic model id of `*` configures the emails for every semantic model in that workspace. If a semantic model has no matching emails configured there, a fallback email from the variable library is used.
# 
# If checking a semantic model's refresh history itself fails (for example, a permissions issue), that is not sent to the model's delegated recipients - all such check failures are aggregated into a single notification sent to a dedicated monitoring email from the variable library.


# MARKDOWN ********************

# ### Setup

# CELL ********************

import pandas as pd
from datetime import datetime, timedelta
import base64
import json
import re
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

semantic_models_management_xlsx_relative_file_path = variables.semantic_models_management_xlsx_relative_file_path
semantic_models_refreshes_sheet_name = variables.semantic_models_refreshes_sheet_name
semantic_models_refreshes_notification_email = variables.semantic_models_refreshes_notification_email

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

def get_semantic_model_key(workspace_id, semantic_model_id):
    return f"{workspace_id}|{semantic_model_id}"

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# load semantic models to monitor

semantic_models_to_overtake_file_path = "/lakehouse/default/Files/semantic_models_to_overtake.json"

try:
    with open(semantic_models_to_overtake_file_path, "r") as f:
        semantic_models_to_monitor = json.load(f)
        semantic_models_to_monitor = [
            s for s in semantic_models_to_monitor if s.get("isRefreshable", False)
        ]
except FileNotFoundError:
    print(f"🔍 Trace | {semantic_models_to_overtake_file_path} not found, nothing to monitor yet")
    semantic_models_to_monitor = []

print(f"🔍 Trace | {len(semantic_models_to_monitor)} semantic models to monitor")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# load notification email mapping up front (workspace id, semantic model id -> emails)
# a semantic model id of "*" configures the emails for every semantic model in that workspace

notification_emails_by_semantic_model = {}
notification_emails_by_workspace = {}

df_notification_emails = pd.read_excel(
    pd.ExcelFile(f"/lakehouse/default{semantic_models_management_xlsx_relative_file_path}"),
    sheet_name = semantic_models_refreshes_sheet_name,
    dtype = str
)

for row in df_notification_emails.itertuples():
    workspace_id = row[1]
    semantic_model_id = row[2]
    notification_emails = row[3]

    if not workspace_id or not semantic_model_id:
        continue

    if semantic_model_id == "*":
        notification_emails_by_workspace[workspace_id] = notification_emails
    else:
        notification_emails_by_semantic_model[get_semantic_model_key(workspace_id, semantic_model_id)] = notification_emails

print(f"🔍 Trace | {len(notification_emails_by_semantic_model)} semantic models with configured notification emails")
print(f"🔍 Trace | {len(notification_emails_by_workspace)} workspaces with wildcard configured notification emails")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# load refresh checkpoints (last checked refresh per semantic model)

refresh_checkpoints_relative_file_path = "Files/refresh_checkpoints.json"

try:
    with open(f"/lakehouse/default/{refresh_checkpoints_relative_file_path}", "r") as f:
        refresh_checkpoints = json.load(f)
except FileNotFoundError:
    refresh_checkpoints = {}

print(f"🔍 Trace | {len(refresh_checkpoints)} semantic models with a prior checkpoint")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# check refresh history per semantic model, keeping only refreshes since the last checkpoint

refresh_check_results = []

for semantic_model in semantic_models_to_monitor:
    workspace_id = semantic_model["workspaceId"]
    semantic_model_id = semantic_model["id"]
    checkpoint_key = get_semantic_model_key(workspace_id, semantic_model_id)

    last_checkpoint_start_time = refresh_checkpoints.get(checkpoint_key, {}).get("checkpointStartTime")
    is_first_check = last_checkpoint_start_time is None

    new_refreshes = []
    check_error = None

    try:
        print(f"🔍 Trace | Checking refresh history for semantic model {semantic_model_id} in workspace {workspace_id}")
        response = obo_fabric_api_client.request(
            "GET",
            f"https://api.powerbi.com/v1.0/myorg/groups/{workspace_id}/datasets/{semantic_model_id}/refreshes?$top=50"
        )
        refreshes = sorted(response["body"].get("value", []), key=lambda r: r["startTime"])

        if is_first_check:
            # first ever check for this semantic model: establish a checkpoint only
            new_refreshes = []
        else:
            new_refreshes = [r for r in refreshes if r["startTime"] > last_checkpoint_start_time]

        if refreshes:
            refresh_checkpoints[checkpoint_key] = {"checkpointStartTime": refreshes[-1]["startTime"]}
    except Exception as e:
        check_error = str(e)
        print(f"🔍 Trace | Error: {check_error}")

    new_relevant_refreshes = [r for r in new_refreshes if r.get("status") in ("Failed", "Disabled")]

    refresh_check_results.append({
        "semantic_model": semantic_model,
        "new_relevant_refreshes": new_relevant_refreshes,
        "check_error": check_error,
    })

print(f"🔍 Trace | {sum(1 for r in refresh_check_results if r['new_relevant_refreshes'] or r['check_error'])} semantic models with new failures, disabled refreshes, or check errors")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

# persist refresh checkpoints for the next run

nu.fs.put(
    f"/lakehouse/default/{refresh_checkpoints_relative_file_path}",
    json.dumps(refresh_checkpoints),
    overwrite = True
)

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Notifications

# CELL ********************

# build one notification per semantic model with new failed or disabled refreshes

notifications = []

for result in refresh_check_results:
    relevant_refreshes = result["new_relevant_refreshes"]
    if not relevant_refreshes:
        continue

    semantic_model = result["semantic_model"]
    workspace_id = semantic_model["workspaceId"]
    semantic_model_id = semantic_model["id"]
    semantic_model_key = get_semantic_model_key(workspace_id, semantic_model_id)

    email = (
        notification_emails_by_semantic_model.get(semantic_model_key)
        or notification_emails_by_workspace.get(workspace_id)
        or semantic_models_refreshes_notification_email
    )

    message_lines = [
        f"🔴 | {r.get('refreshType')} refresh started {r.get('startTime')} {'disabled' if r.get('status') == 'Disabled' else 'failed'}"
        for r in relevant_refreshes
    ]

    notification_body = f"""
        <p>
            {'<br/>'.join(message_lines)}
        </p>
    """

    notifications.append({
        "subject": f"Semantic model refresh issue: {semantic_model['name']}",
        "body": notification_body,
        "emails": email,
    })

# aggregate refresh-history check errors into a single notification to the dedicated monitoring email

check_error_results = [r for r in refresh_check_results if r["check_error"]]

if check_error_results:
    message_lines = [
        f"🔴 | {r['semantic_model']['workspaceName']} {r['semantic_model']['name']}: {r['check_error']}"
        for r in check_error_results
    ]

    notification_body = f"""
        <p>
            {'<br/>'.join(message_lines)}
        </p>
    """

    notifications.append({
        "subject": "Semantic model refresh monitor failed",
        "body": notification_body,
        "emails": semantic_models_refreshes_notification_email,
    })

nu.notebook.exit(json.dumps(notifications))

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }
