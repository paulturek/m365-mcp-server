"""
Power BI MCP tools.

Follows the same pattern as all other tool modules:
  - TOOLS  : list of MCP tool definition dicts
  - HANDLERS: dict mapping tool name -> async handler coroutine
"""

from __future__ import annotations

from m365_mcp.services import powerbi_service as svc

# ---------------------------------------------------------------------------
# Tool definitions
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "name": "powerbi_list_workspaces",
        "description": (
            "List all Power BI workspaces (groups) the authenticated user has access to. "
            "Returns workspace IDs and names needed for other Power BI tools."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {
                    "type": "string",
                    "description": "User email address (e.g. paul.turek@bolthousefresh.com)",
                }
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "powerbi_list_datasets",
        "description": (
            "List all datasets (semantic models) in a Power BI workspace. "
            "Use workspace_id='me' for the user's personal 'My Workspace'."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string", "description": "User email address"},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID from powerbi_list_workspaces, or 'me' for My Workspace",
                },
            },
            "required": ["user_id", "workspace_id"],
        },
    },
    {
        "name": "powerbi_get_dataset",
        "description": "Get details of a specific Power BI dataset/semantic model.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string", "description": "User email address"},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID, or 'me' for My Workspace",
                },
                "dataset_id": {"type": "string", "description": "Dataset ID"},
            },
            "required": ["user_id", "workspace_id", "dataset_id"],
        },
    },
    {
        "name": "powerbi_refresh_dataset",
        "description": (
            "Trigger an on-demand refresh of a Power BI dataset/semantic model. "
            "Returns immediately (202 Accepted) — the refresh runs asynchronously. "
            "Use powerbi_get_refresh_history to check completion status. "
            "Refresh limits: Power BI Pro = 8/day, Premium = 48/day per dataset."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string", "description": "User email address"},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID, or 'me' for My Workspace",
                },
                "dataset_id": {"type": "string", "description": "Dataset ID to refresh"},
                "notify_option": {
                    "type": "string",
                    "enum": ["MailOnFailure", "MailOnCompletion", "NoNotification"],
                    "description": "Email notification preference. Defaults to MailOnFailure.",
                    "default": "MailOnFailure",
                },
            },
            "required": ["user_id", "workspace_id", "dataset_id"],
        },
    },
    {
        "name": "powerbi_get_refresh_history",
        "description": (
            "Get the refresh history for a Power BI dataset. "
            "Use this to check the status of a triggered refresh (Completed, Failed, Unknown)."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string", "description": "User email address"},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID, or 'me' for My Workspace",
                },
                "dataset_id": {"type": "string", "description": "Dataset ID"},
                "top": {
                    "type": "integer",
                    "description": "Number of most recent refresh entries to return. Defaults to 10.",
                    "default": 10,
                },
            },
            "required": ["user_id", "workspace_id", "dataset_id"],
        },
    },
    {
        "name": "powerbi_get_refresh_schedule",
        "description": "Get the configured automatic refresh schedule for a Power BI dataset.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string", "description": "User email address"},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID, or 'me' for My Workspace",
                },
                "dataset_id": {"type": "string", "description": "Dataset ID"},
            },
            "required": ["user_id", "workspace_id", "dataset_id"],
        },
    },
    {
        "name": "powerbi_list_reports",
        "description": "List all reports in a Power BI workspace.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string", "description": "User email address"},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID, or 'me' for My Workspace",
                },
            },
            "required": ["user_id", "workspace_id"],
        },
    },
]

# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------

async def _list_workspaces(args: dict) -> dict:
    return await svc.list_workspaces(args["user_id"])


async def _list_datasets(args: dict) -> dict:
    return await svc.list_datasets(args["user_id"], args["workspace_id"])


async def _get_dataset(args: dict) -> dict:
    return await svc.get_dataset(args["user_id"], args["workspace_id"], args["dataset_id"])


async def _refresh_dataset(args: dict) -> dict:
    return await svc.refresh_dataset(
        args["user_id"],
        args["workspace_id"],
        args["dataset_id"],
        args.get("notify_option", "MailOnFailure"),
    )


async def _get_refresh_history(args: dict) -> dict:
    return await svc.get_refresh_history(
        args["user_id"],
        args["workspace_id"],
        args["dataset_id"],
        args.get("top", 10),
    )


async def _get_refresh_schedule(args: dict) -> dict:
    return await svc.get_refresh_schedule(
        args["user_id"], args["workspace_id"], args["dataset_id"]
    )


async def _list_reports(args: dict) -> dict:
    return await svc.list_reports(args["user_id"], args["workspace_id"])


HANDLERS: dict[str, object] = {
    "powerbi_list_workspaces": _list_workspaces,
    "powerbi_list_datasets": _list_datasets,
    "powerbi_get_dataset": _get_dataset,
    "powerbi_refresh_dataset": _refresh_dataset,
    "powerbi_get_refresh_history": _get_refresh_history,
    "powerbi_get_refresh_schedule": _get_refresh_schedule,
    "powerbi_list_reports": _list_reports,
}
