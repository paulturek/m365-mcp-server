"""
Power BI service layer.

Thin wrappers around PowerBIClient that shape raw API responses into
clean, consistent dicts for the MCP tool handlers.
"""

from __future__ import annotations

from m365_mcp.clients.powerbi_client import PowerBIClient


# ---------------------------------------------------------------------------
# Workspaces
# ---------------------------------------------------------------------------

async def list_workspaces(user_email: str) -> dict:
    client = PowerBIClient(user_email)
    result = await client.get("/groups")
    workspaces = result.get("value", [])
    return {
        "count": len(workspaces),
        "workspaces": [
            {
                "id": w["id"],
                "name": w["name"],
                "type": w.get("type", "Workspace"),
                "is_on_dedicated_capacity": w.get("isOnDedicatedCapacity", False),
            }
            for w in workspaces
        ],
    }


# ---------------------------------------------------------------------------
# Datasets / Models
# ---------------------------------------------------------------------------

async def list_datasets(user_email: str, workspace_id: str) -> dict:
    client = PowerBIClient(user_email)
    path = "/datasets" if workspace_id == "me" else f"/groups/{workspace_id}/datasets"
    result = await client.get(path)
    datasets = result.get("value", [])
    return {
        "count": len(datasets),
        "datasets": [
            {
                "id": d["id"],
                "name": d["name"],
                "configured_by": d.get("configuredBy"),
                "is_refreshable": d.get("isRefreshable", False),
                "is_effective_identity_required": d.get("isEffectiveIdentityRequired", False),
                "created_date": d.get("createdDate"),
                "web_url": d.get("webUrl"),
            }
            for d in datasets
        ],
    }


async def get_dataset(user_email: str, workspace_id: str, dataset_id: str) -> dict:
    client = PowerBIClient(user_email)
    path = (
        f"/datasets/{dataset_id}"
        if workspace_id == "me"
        else f"/groups/{workspace_id}/datasets/{dataset_id}"
    )
    d = await client.get(path)
    return {
        "id": d["id"],
        "name": d["name"],
        "configured_by": d.get("configuredBy"),
        "is_refreshable": d.get("isRefreshable", False),
        "created_date": d.get("createdDate"),
        "web_url": d.get("webUrl"),
        "upstream_datasets": d.get("upstreamDatasets", []),
    }


# ---------------------------------------------------------------------------
# Refresh operations
# ---------------------------------------------------------------------------

async def refresh_dataset(
    user_email: str,
    workspace_id: str,
    dataset_id: str,
    notify_option: str = "MailOnFailure",
) -> dict:
    client = PowerBIClient(user_email)
    path = (
        f"/datasets/{dataset_id}/refreshes"
        if workspace_id == "me"
        else f"/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
    )
    result = await client.post(path, body={"notifyOption": notify_option})
    return result or {"status": "accepted", "dataset_id": dataset_id}


async def get_refresh_history(
    user_email: str,
    workspace_id: str,
    dataset_id: str,
    top: int = 10,
) -> dict:
    client = PowerBIClient(user_email)
    path = (
        f"/datasets/{dataset_id}/refreshes?$top={top}"
        if workspace_id == "me"
        else f"/groups/{workspace_id}/datasets/{dataset_id}/refreshes?$top={top}"
    )
    result = await client.get(path)
    refreshes = result.get("value", [])
    return {
        "dataset_id": dataset_id,
        "count": len(refreshes),
        "refreshes": [
            {
                "request_id": r.get("requestId"),
                "status": r.get("status"),
                "refresh_type": r.get("refreshType"),
                "start_time": r.get("startTime"),
                "end_time": r.get("endTime"),
                "service_exception_json": r.get("serviceExceptionJson"),
            }
            for r in refreshes
        ],
    }


async def get_refresh_schedule(
    user_email: str, workspace_id: str, dataset_id: str
) -> dict:
    client = PowerBIClient(user_email)
    path = (
        f"/datasets/{dataset_id}/refreshSchedule"
        if workspace_id == "me"
        else f"/groups/{workspace_id}/datasets/{dataset_id}/refreshSchedule"
    )
    return await client.get(path)


# ---------------------------------------------------------------------------
# Reports
# ---------------------------------------------------------------------------

async def list_reports(user_email: str, workspace_id: str) -> dict:
    client = PowerBIClient(user_email)
    path = "/reports" if workspace_id == "me" else f"/groups/{workspace_id}/reports"
    result = await client.get(path)
    reports = result.get("value", [])
    return {
        "count": len(reports),
        "reports": [
            {
                "id": r["id"],
                "name": r["name"],
                "dataset_id": r.get("datasetId"),
                "web_url": r.get("webUrl"),
                "embed_url": r.get("embedUrl"),
            }
            for r in reports
        ],
    }
