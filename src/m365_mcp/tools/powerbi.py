"""
Power BI MCP tool definitions and handlers.

Tools exposed:
  - powerbi_list_workspaces
  - powerbi_list_datasets
  - powerbi_get_dataset
  - powerbi_refresh_dataset
  - powerbi_get_refresh_history
  - powerbi_get_refresh_schedule
  - powerbi_list_reports
"""

from __future__ import annotations

from m365_mcp.services import powerbi_service

# ---------------------------------------------------------------------------
# Tool schemas
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "name": "powerbi_list_workspaces",
        "description": (
            "List all Power BI workspaces (groups) the authenticated user has access to. "
            "Returns workspace IDs, names, and capacity information."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {
                    "type": "string",
                    "description": "Email address of the authenticated user.",
                },
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "powerbi_list_datasets",
        "description": (
            "List all datasets (semantic models) in a Power BI workspace. "
            "Use workspace_id='me' for the user's personal My Workspace."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string", "description": "Authenticated user email."},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID, or 'me' for My Workspace.",
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
                "user_id": {"type": "string"},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID, or 'me' for My Workspace.",
                },
                "dataset_id": {"type": "string", "description": "Dataset GUID."},
            },
            "required": ["user_id", "workspace_id", "dataset_id"],
        },
    },
    {
        "name": "powerbi_refresh_dataset",
        "description": (
            "Trigger an on-demand refresh of a Power BI dataset/semantic model. "
            "Returns immediately with status 'accepted' — the refresh runs asynchronously. "
            "Use powerbi_get_refresh_history to track completion. "
            "Limits: Power BI Pro = 8 refreshes/day; Premium = 48/day."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string"},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID, or 'me' for My Workspace.",
                },
                "dataset_id": {"type": "string", "description": "Dataset GUID."},
                "notify_option": {
                    "type": "string",
                    "enum": ["MailOnFailure", "MailOnCompletion", "NoNotification"],
                    "description": "Email notification preference. Default: MailOnFailure.",
                },
            },
            "required": ["user_id", "workspace_id", "dataset_id"],
        },
    },
    {
        "name": "powerbi_get_refresh_history",
        "description": (
            "Get the refresh history for a Power BI dataset. "
            "Use this to check whether a triggered refresh has completed or failed."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string"},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID, or 'me' for My Workspace.",
                },
                "dataset_id": {"type": "string", "description": "Dataset GUID."},
                "top": {
                    "type": "integer",
                    "description": "Number of most recent refreshes to return. Default: 10.",
                },
            },
            "required": ["user_id", "workspace_id", "dataset_id"],
        },
    },
    {
        "name": "powerbi_get_refresh_schedule",
        "description": "Get the configured scheduled refresh settings for a Power BI dataset.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string"},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID, or 'me' for My Workspace.",
                },
                "dataset_id": {"type": "string", "description": "Dataset GUID."},
            },
            "required": ["user_id", "workspace_id", "dataset_id"],
        },
    },
    {
        "name": "powerbi_list_reports",
        "description": "List all Power BI reports in a workspace.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {"type": "string"},
                "workspace_id": {
                    "type": "string",
                    "description": "Workspace/group ID, or 'me' for My Workspace.",
                },
            },
            "required": ["user_id", "workspace_id"],
        },
    },
]

# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------

HANDLERS: dict[str, any] = {}


async def _handle_list_workspaces(args: dict) -> dict:
    return await powerbi_service.list_workspaces(args["user_id"])


async def _handle_list_datasets(args: dict) -> dict:
    return await powerbi_service.list_datasets(args["user_id"], args["workspace_id"])


async def _handle_get_dataset(args: dict) -> dict:
    return await powerbi_service.get_dataset(
        args["user_id"], args["workspace_id"], args["dataset_id"]
    )


async def _handle_refresh_dataset(args: dict) -> dict:
    return await powerbi_service.refresh_dataset(
        args["user_id"],
        args["workspace_id"],
        args["dataset_id"],
        args.get("notify_option", "MailOnFailure"),
    )


async def _handle_get_refresh_history(args: dict) -> dict:
    return await powerbi_service.get_refresh_history(
        args["user_id"],
        args["workspace_id"],
        args["dataset_id"],
        args.get("top", 10),
    )


async def _handle_get_refresh_schedule(args: dict) -> dict:
    return await powerbi_service.get_refresh_schedule(
        args["user_id"], args["workspace_id"], args["dataset_id"]
    )


async def _handle_list_reports(args: dict) -> dict:
    return await powerbi_service.list_reports(args["user_id"], args["workspace_id"])


HANDLERS = {
    "powerbi_list_workspaces": _handle_list_workspaces,
    "powerbi_list_datasets": _handle_list_datasets,
    "powerbi_get_dataset": _handle_get_dataset,
    "powerbi_refresh_dataset": _handle_refresh_dataset,
    "powerbi_get_refresh_history": _handle_get_refresh_history,
    "powerbi_get_refresh_schedule": _handle_get_refresh_schedule,
    "powerbi_list_reports": _handle_list_reports,
}
