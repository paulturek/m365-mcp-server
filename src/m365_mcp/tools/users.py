"""Users / Directory MCP tools.

Covers: get current user profile, look up user, list directory, search.
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
        "name": "users_get_me",
        "description": "Get the authenticated user's profile (display name, email, job title, etc.).",
        "inputSchema": {
            "type": "object",
            "properties": {**_USER_ID_PROP},
            "required": ["user_id"],
        },
    },
    {
        "name": "users_get_user",
        "description": "Look up another user in the directory by their UPN or object ID.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "target": {
                    "type": "string",
                    "description": "User principal name (email) or Azure AD object ID",
                },
            },
            "required": ["user_id", "target"],
        },
    },
    {
        "name": "users_list_users",
        "description": "List users in the Azure AD directory.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "top": {"type": "integer", "default": 25},
                "filter": {"type": "string", "description": "OData $filter"},
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "users_search",
        "description": "Search for users by name or email.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "query": {"type": "string", "description": "Search query (name or email fragment)"},
                "top": {"type": "integer", "default": 10},
            },
            "required": ["user_id", "query"],
        },
    },
]

_USER_FIELDS = "id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones"


async def _get_me(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    data = await client.get(f"/me?$select={_USER_FIELDS}")
    return _format_user(data)


async def _get_user(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    target = params["target"]
    data = await client.get(f"/users/{target}?$select={_USER_FIELDS}")
    return _format_user(data)


async def _list_users(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    top = params.get("top", 25)
    qp = f"?$top={top}&$select={_USER_FIELDS}"
    if params.get("filter"):
        qp += f"&$filter={params['filter']}"
    data = await client.get(f"/users{qp}")
    users = data.get("value", [])
    return {"count": len(users), "users": [_format_user(u) for u in users]}


async def _search_users(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    q = params["query"]
    top = params.get("top", 10)
    # Use startswith on displayName; Graph also supports $search with ConsistencyLevel
    filter_expr = f"startswith(displayName,'{q}') or startswith(mail,'{q}')"
    data = await client.get(f"/users?$filter={filter_expr}&$top={top}&$select={_USER_FIELDS}")
    users = data.get("value", [])
    return {"count": len(users), "users": [_format_user(u) for u in users]}


def _format_user(u: dict) -> dict:
    return {
        "id": u.get("id"),
        "displayName": u.get("displayName"),
        "email": u.get("mail") or u.get("userPrincipalName"),
        "jobTitle": u.get("jobTitle"),
        "department": u.get("department"),
        "officeLocation": u.get("officeLocation"),
        "phone": u.get("mobilePhone") or (u.get("businessPhones") or [None])[0],
    }


HANDLERS = {
    "users_get_me": _get_me,
    "users_get_user": _get_user,
    "users_list_users": _list_users,
    "users_search": _search_users,
}
