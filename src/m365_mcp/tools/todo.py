"""Microsoft To Do MCP tools.

Covers: list task lists, list tasks, create task, update task.
"""
import logging

from ..auth.oauth_web import get_access_token
from ..clients.graph_client import GraphClient

logger = logging.getLogger(__name__)

_USER_ID_PROP = {
    "user_id": {
        "type": "string",
        "description": "Your user identifier (email recommended)",
    }
}

TOOLS = [
    {
        "name": "todo_list_task_lists",
        "description": "List all Microsoft To Do task lists.",
        "inputSchema": {
            "type": "object",
            "properties": {**_USER_ID_PROP},
            "required": ["user_id"],
        },
    },
    {
        "name": "todo_list_tasks",
        "description": "List tasks in a To Do task list.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "list_id": {"type": "string", "description": "Task list ID"},
                "top": {"type": "integer", "default": 25},
                "filter": {"type": "string", "description": "OData $filter"},
            },
            "required": ["user_id", "list_id"],
        },
    },
    {
        "name": "todo_create_task",
        "description": "Create a new task in a To Do list.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "list_id": {"type": "string"},
                "title": {"type": "string", "description": "Task title"},
                "body": {"type": "string", "description": "Task notes/body"},
                "due_date": {"type": "string", "description": "Due date (YYYY-MM-DD)"},
                "importance": {
                    "type": "string",
                    "enum": ["low", "normal", "high"],
                    "default": "normal",
                },
            },
            "required": ["user_id", "list_id", "title"],
        },
    },
    {
        "name": "todo_update_task",
        "description": "Update an existing task (mark complete, change title, etc.).",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "list_id": {"type": "string"},
                "task_id": {"type": "string"},
                "title": {"type": "string"},
                "status": {
                    "type": "string",
                    "enum": ["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"],
                },
                "importance": {"type": "string", "enum": ["low", "normal", "high"]},
                "due_date": {"type": "string"},
                "body": {"type": "string"},
            },
            "required": ["user_id", "list_id", "task_id"],
        },
    },
]


async def _list_task_lists(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    data = await client.get("/me/todo/lists")
    lists = data.get("value", [])
    return {
        "count": len(lists),
        "lists": [
            {
                "id": l.get("id"),
                "displayName": l.get("displayName"),
                "isOwner": l.get("isOwner"),
                "wellknownListName": l.get("wellknownListName"),
            }
            for l in lists
        ],
    }


async def _list_tasks(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    list_id = params["list_id"]
    top = params.get("top", 25)
    qp = f"?$top={top}"
    if params.get("filter"):
        qp += f"&$filter={params['filter']}"
    data = await client.get(f"/me/todo/lists/{list_id}/tasks{qp}")
    tasks = data.get("value", [])
    return {
        "count": len(tasks),
        "tasks": [
            {
                "id": t.get("id"),
                "title": t.get("title"),
                "status": t.get("status"),
                "importance": t.get("importance"),
                "createdDateTime": t.get("createdDateTime"),
                "dueDateTime": t.get("dueDateTime"),
                "completedDateTime": t.get("completedDateTime"),
            }
            for t in tasks
        ],
    }


async def _create_task(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    list_id = params["list_id"]
    body: dict = {"title": params["title"]}
    if params.get("body"):
        body["body"] = {"content": params["body"], "contentType": "text"}
    if params.get("importance"):
        body["importance"] = params["importance"]
    if params.get("due_date"):
        body["dueDateTime"] = {
            "dateTime": f"{params['due_date']}T00:00:00",
            "timeZone": "UTC",
        }
    result = await client.post(f"/me/todo/lists/{list_id}/tasks", json=body)
    return {
        "id": result.get("id"),
        "title": result.get("title"),
        "status": result.get("status"),
    }


async def _update_task(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    list_id = params["list_id"]
    task_id = params["task_id"]
    body: dict = {}
    if params.get("title"):
        body["title"] = params["title"]
    if params.get("status"):
        body["status"] = params["status"]
    if params.get("importance"):
        body["importance"] = params["importance"]
    if params.get("body"):
        body["body"] = {"content": params["body"], "contentType": "text"}
    if params.get("due_date"):
        body["dueDateTime"] = {
            "dateTime": f"{params['due_date']}T00:00:00",
            "timeZone": "UTC",
        }
    if not body:
        return {"error": "No fields to update"}
    result = await client.patch(
        f"/me/todo/lists/{list_id}/tasks/{task_id}", json=body
    )
    return {
        "id": result.get("id"),
        "title": result.get("title"),
        "status": result.get("status"),
    }


HANDLERS = {
    "todo_list_task_lists": _list_task_lists,
    "todo_list_tasks": _list_tasks,
    "todo_create_task": _create_task,
    "todo_update_task": _update_task,
}
