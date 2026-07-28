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

# CELL ********************

TARGET_CAPACITY_ID = "37982985-f05e-453e-9ea7-fd8fd1b8a7ee"

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

batches = list(chunked(workspace_ids_to_assign, 500))
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
