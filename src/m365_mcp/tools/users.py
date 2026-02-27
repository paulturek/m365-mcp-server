"""Users / Directory MCP tools.

Covers: get current user profile, look up user, list directory, search,
        get manager, get direct reports, get user photo.
"""
import logging

from ..auth.oauth_web import get_access_token
from ..clients.graph_client import GraphClient, GraphAPIError, AuthenticationRequiredError

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
        "description": "Search for users by name or email (uses startsWith on displayName and mail).",
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
    {
        "name": "users_get_manager",
        "description": "Get the manager of the authenticated user or a specified user.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "target": {
                    "type": "string",
                    "description": "User UPN or object ID (omit for current user)",
                },
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "users_get_direct_reports",
        "description": "Get the direct reports of the authenticated user or a specified user.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "target": {
                    "type": "string",
                    "description": "User UPN or object ID (omit for current user)",
                },
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "users_get_photo",
        "description": "Get the profile photo metadata and download URL for a user.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "target": {
                    "type": "string",
                    "description": "User UPN or object ID (omit for current user)",
                },
                "size": {
                    "type": "string",
                    "enum": ["48x48", "64x64", "96x96", "120x120", "240x240", "360x360", "432x432", "504x504", "648x648"],
                    "description": "Photo size (omit for largest available)",
                },
            },
            "required": ["user_id"],
        },
    },
]

_USER_FIELDS = "id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation,mobilePhone,businessPhones"


async def _get_me(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    data = await client.get("/me", params={"$select": _USER_FIELDS})
    return _format_user(data)


async def _get_user(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    target = params["target"]
    data = await client.get(f"/users/{target}", params={"$select": _USER_FIELDS})
    return _format_user(data)


async def _list_users(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    qparams: dict = {"$top": params.get("top", 25), "$select": _USER_FIELDS}
    if params.get("filter"):
        qparams["$filter"] = params["filter"]
    data = await client.get("/users", params=qparams)
    users = data.get("value", [])
    return {"count": len(users), "users": [_format_user(u) for u in users]}


async def _search_users(params: dict) -> dict:
    """Search users via $filter with startsWith — passed as params dict for proper URL encoding."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    q = params["query"]
    top = params.get("top", 10)
    # Pass filter via params= so httpx handles URL encoding of spaces, quotes, etc.
    filter_expr = f"startswith(displayName,'{q}') or startswith(mail,'{q}')"
    data = await client.get(
        "/users",
        params={"$filter": filter_expr, "$top": top, "$select": _USER_FIELDS},
    )
    users = data.get("value", [])
    return {"count": len(users), "users": [_format_user(u) for u in users]}


async def _get_manager(params: dict) -> dict:
    """GET /me/manager or /users/{id}/manager — get user's manager."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    target = params.get("target")
    endpoint = f"/users/{target}/manager" if target else "/me/manager"
    data = await client.get(endpoint, params={"$select": _USER_FIELDS})
    return _format_user(data)


async def _get_direct_reports(params: dict) -> dict:
    """GET /me/directReports or /users/{id}/directReports."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    target = params.get("target")
    endpoint = f"/users/{target}/directReports" if target else "/me/directReports"
    data = await client.get(endpoint, params={"$select": _USER_FIELDS})
    reports = data.get("value", [])
    return {
        "count": len(reports),
        "directReports": [_format_user(r) for r in reports],
    }


async def _get_photo(params: dict) -> dict:
    """GET user photo metadata. Returns photo info and content endpoint."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    target = params.get("target")
    size = params.get("size")

    base = f"/users/{target}" if target else "/me"
    photo_endpoint = f"{base}/photos/{size}" if size else f"{base}/photo"

    try:
        meta = await client.get(photo_endpoint)
        return {
            "width": meta.get("width"),
            "height": meta.get("height"),
            "id": meta.get("id"),
            "contentType": meta.get("@odata.mediaContentType"),
            "downloadEndpoint": f"{photo_endpoint}/$value",
        }
    except AuthenticationRequiredError:
        raise
    except GraphAPIError as e:
        if e.status_code == 404:
            return {"error": "No photo set for this user"}
        if e.status_code == 403:
            return {"error": f"Permission denied — requires User.Read.All: {e.message}"}
        return {"error": f"Graph API error [{e.status_code}]: {e.message}"}
    except Exception as e:
        return {"error": f"Unexpected error retrieving photo: {e}"}


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
    "users_get_manager": _get_manager,
    "users_get_direct_reports": _get_direct_reports,
    "users_get_photo": _get_photo,
}
