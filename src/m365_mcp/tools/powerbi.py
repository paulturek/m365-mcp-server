"""Power BI MCP tools.

Covers: list reports, list datasets, trigger dataset refresh.
Uses the Power BI REST API (via PowerBIClient).
"""
import logging

from ..auth.oauth_web import get_access_token
from ..clients.powerbi_client import PowerBIClient

logger = logging.getLogger(__name__)

_USER_ID_PROP = {
    "user_id": {
        "type": "string",
        "description": "Your user identifier (email recommended)",
    }
}

TOOLS = [
    {
        "name": "powerbi_list_reports",
        "description": "List Power BI reports the user has access to.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "group_id": {
                    "type": "string",
                    "description": "Workspace/group ID (omit for 'My Workspace')",
                },
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "powerbi_list_datasets",
        "description": "List Power BI datasets.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "group_id": {
                    "type": "string",
                    "description": "Workspace/group ID (omit for 'My Workspace')",
                },
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "powerbi_refresh_dataset",
        "description": "Trigger a refresh for a Power BI dataset.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "dataset_id": {"type": "string", "description": "Dataset ID"},
                "group_id": {"type": "string", "description": "Workspace/group ID (optional)"},
            },
            "required": ["user_id", "dataset_id"],
        },
    },
]


async def _list_reports(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = PowerBIClient(token)
    group_id = params.get("group_id")
    reports = await client.list_reports(group_id=group_id)
    return {
        "count": len(reports),
        "reports": [
            {
                "id": r.get("id"),
                "name": r.get("name"),
                "webUrl": r.get("webUrl"),
                "datasetId": r.get("datasetId"),
            }
            for r in reports
        ],
    }


async def _list_datasets(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = PowerBIClient(token)
    group_id = params.get("group_id")
    datasets = await client.list_datasets(group_id=group_id)
    return {
        "count": len(datasets),
        "datasets": [
            {
                "id": d.get("id"),
                "name": d.get("name"),
                "configuredBy": d.get("configuredBy"),
                "isRefreshable": d.get("isRefreshable"),
            }
            for d in datasets
        ],
    }


async def _refresh_dataset(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = PowerBIClient(token)
    dataset_id = params["dataset_id"]
    group_id = params.get("group_id")
    await client.refresh_dataset(dataset_id, group_id=group_id)
    return {"triggered": True, "dataset_id": dataset_id}


HANDLERS = {
    "powerbi_list_reports": _list_reports,
    "powerbi_list_datasets": _list_datasets,
    "powerbi_refresh_dataset": _refresh_dataset,
}
