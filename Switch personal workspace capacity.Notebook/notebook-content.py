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

# ### Setup

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

%run FabricApiClient

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

# MARKDOWN ********************

# ### Parameters

# CELL ********************

variables = nu.variableLibrary.getLibrary("Variables")

personal_workspace_capacity_id = variables.personal_workspace_capacity_id

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# MARKDOWN ********************

# ### Run

# CELL ********************

personal_workspaces = obo_fabric_api_client.get_paginated(
    "https://api.powerbi.com/v1.0/myorg/admin/groups?$filter=type eq 'PersonalGroup'&$top=1000"
)
print(f"🔍 Trace | Found {len(personal_workspaces)} personal workspaces")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

workspace_ids_to_assign = [
    workspace["id"]
    for workspace in personal_workspaces
    if (workspace.get("capacityId") or "").casefold() != (personal_workspace_capacity_id or "").casefold()
]

print(f"🔍 Trace | {len(workspace_ids_to_assign)} / {len(personal_workspaces)} personal workspaces need reassignment")

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }

# CELL ********************

def chunked(items, size):
    for i in range(0, len(items), size):
        yield items[i:i + size]

batches = list(chunked(workspace_ids_to_assign, 500))

for batch_number, batch in enumerate(batches, start=1):
    print(f"🔍 Trace | Assigning batch {batch_number}/{len(batches)} ({len(batch)} workspaces)")
    request_body = {
        "capacityMigrationAssignments": [
            {
                "targetCapacityObjectId": personal_workspace_capacity_id,
                "workspacesToAssign": batch
            }
        ]
    }
    obo_fabric_api_client.request(
        "POST",
        "https://api.powerbi.com/v1.0/myorg/admin/capacities/AssignWorkspaces",
        request_body
    )

# METADATA ********************

# META {
# META   "language": "python",
# META   "language_group": "jupyter_python"
# META }
