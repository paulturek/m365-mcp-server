"""SharePoint MCP tools.

Covers: list sites, get site by URL, list library items, download, upload, search,
        create/update/delete list items.
"""
import logging
from urllib.parse import quote

from ..auth.oauth_web import get_access_token
from ..clients.graph_client import GraphClient

logger = logging.getLogger(__name__)


def _encode_id(entity_id: str) -> str:
    """URL-encode a Graph entity ID for safe use in URL path segments.

    Only encodes characters invalid in RFC 3986 path segments (like /).
    Preserves =, +, -, _ which are path-safe.
    """
    return quote(entity_id, safe=":@!$&'()*+,;=")


_USER_ID_PROP = {
    "user_id": {
        "type": "string",
        "description": "Your user identifier (email recommended)",
    }
}

TOOLS = [
    {
        "name": "sharepoint_list_sites",
        "description": "List SharePoint sites the user has access to, or search by keyword.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "search": {"type": "string", "description": "Site name search query"},
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "sharepoint_get_site",
        "description": "Get a SharePoint site by its hostname and path (e.g. contoso.sharepoint.com and /sites/TeamSite).",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "hostname": {
                    "type": "string",
                    "description": "SharePoint hostname (e.g. contoso.sharepoint.com)",
                },
                "site_path": {
                    "type": "string",
                    "description": "Site path (e.g. /sites/TeamSite)",
                },
            },
            "required": ["user_id", "hostname", "site_path"],
        },
    },
    {
        "name": "sharepoint_list_items",
        "description": "List files/folders in a SharePoint document library.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "site_id": {"type": "string", "description": "SharePoint site ID"},
                "drive_id": {"type": "string", "description": "Document library drive ID (optional, uses default)"},
                "path": {"type": "string", "default": "/", "description": "Folder path within the library"},
            },
            "required": ["user_id", "site_id"],
        },
    },
    {
        "name": "sharepoint_download_file",
        "description": "Download a file from a SharePoint document library.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "site_id": {"type": "string"},
                "item_id": {"type": "string", "description": "Item ID in the library"},
                "drive_id": {"type": "string", "description": "Drive ID (optional)"},
            },
            "required": ["user_id", "site_id", "item_id"],
        },
    },
    {
        "name": "sharepoint_upload_file",
        "description": "Upload a file to a SharePoint document library.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "site_id": {"type": "string"},
                "path": {"type": "string", "description": "Destination path including filename"},
                "content": {"type": "string", "description": "Base64-encoded file content"},
                "drive_id": {"type": "string", "description": "Drive ID (optional)"},
            },
            "required": ["user_id", "site_id", "path", "content"],
        },
    },
    {
        "name": "sharepoint_search",
        "description": "Search for files and content across SharePoint.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "query": {"type": "string", "description": "Search query"},
                "top": {"type": "integer", "default": 10},
            },
            "required": ["user_id", "query"],
        },
    },
    {
        "name": "sharepoint_list_lists",
        "description": "List all lists in a SharePoint site (includes document libraries and custom lists).",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "site_id": {"type": "string", "description": "SharePoint site ID"},
            },
            "required": ["user_id", "site_id"],
        },
    },
    {
        "name": "sharepoint_list_list_items",
        "description": "List items in a SharePoint list (not a document library — use sharepoint_list_items for that).",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "site_id": {"type": "string", "description": "SharePoint site ID"},
                "list_id": {"type": "string", "description": "SharePoint list ID"},
                "top": {"type": "integer", "default": 25},
                "expand_fields": {
                    "type": "boolean",
                    "default": True,
                    "description": "Include column values (fields) in results",
                },
            },
            "required": ["user_id", "site_id", "list_id"],
        },
    },
    {
        "name": "sharepoint_create_list_item",
        "description": "Create a new item in a SharePoint list.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "site_id": {"type": "string", "description": "SharePoint site ID"},
                "list_id": {"type": "string", "description": "SharePoint list ID"},
                "fields": {
                    "type": "object",
                    "description": "Column name/value pairs for the new item",
                },
            },
            "required": ["user_id", "site_id", "list_id", "fields"],
        },
    },
    {
        "name": "sharepoint_update_list_item",
        "description": "Update an existing item in a SharePoint list.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "site_id": {"type": "string", "description": "SharePoint site ID"},
                "list_id": {"type": "string", "description": "SharePoint list ID"},
                "item_id": {"type": "string", "description": "List item ID"},
                "fields": {
                    "type": "object",
                    "description": "Column name/value pairs to update",
                },
            },
            "required": ["user_id", "site_id", "list_id", "item_id", "fields"],
        },
    },
    {
        "name": "sharepoint_delete_list_item",
        "description": "Delete an item from a SharePoint list.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "site_id": {"type": "string", "description": "SharePoint site ID"},
                "list_id": {"type": "string", "description": "SharePoint list ID"},
                "item_id": {"type": "string", "description": "List item ID"},
            },
            "required": ["user_id", "site_id", "list_id", "item_id"],
        },
    },
]


# ---- Handlers -----------------------------------------------------------


async def _list_sites(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    search = params.get("search")
    endpoint = f"/sites?search={search}" if search else "/sites?$top=50"
    data = await client.get(endpoint)
    sites = data.get("value", [])
    return {
        "count": len(sites),
        "sites": [
            {
                "id": s.get("id"),
                "displayName": s.get("displayName"),
                "webUrl": s.get("webUrl"),
                "description": s.get("description", "")[:200],
            }
            for s in sites
        ],
    }


async def _get_site(params: dict) -> dict:
    """GET /sites/{hostname}:/{path} — get site by URL."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    hostname = params["hostname"]
    site_path = params["site_path"].strip("/")
    data = await client.get(f"/sites/{hostname}:/{site_path}")
    return {
        "id": data.get("id"),
        "displayName": data.get("displayName"),
        "webUrl": data.get("webUrl"),
        "description": data.get("description"),
    }


async def _list_items(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    site_id = _encode_id(params["site_id"])
    drive_id = _encode_id(params["drive_id"]) if params.get("drive_id") else None
    path = params.get("path", "/").strip("/")
    if drive_id:
        base = f"/sites/{site_id}/drives/{drive_id}"
    else:
        base = f"/sites/{site_id}/drive"
    endpoint = (
        f"{base}/root/children"
        if path in ("", "/")
        else f"{base}/root:/{path}:/children"
    )
    data = await client.get(endpoint)
    items = data.get("value", [])
    return {
        "count": len(items),
        "items": [
            {
                "name": i.get("name"),
                "id": i.get("id"),
                "type": "folder" if "folder" in i else "file",
                "size": i.get("size"),
                "webUrl": i.get("webUrl"),
                "lastModified": i.get("lastModifiedDateTime"),
            }
            for i in items
        ],
    }


async def _download_file(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    site_id = _encode_id(params["site_id"])
    item_id = _encode_id(params["item_id"])
    drive_id = _encode_id(params["drive_id"]) if params.get("drive_id") else None
    base = f"/sites/{site_id}/drives/{drive_id}" if drive_id else f"/sites/{site_id}/drive"
    meta = await client.get(f"{base}/items/{item_id}")
    return {
        "name": meta.get("name"),
        "size": meta.get("size"),
        "downloadUrl": meta.get("@microsoft.graph.downloadUrl", ""),
        "webUrl": meta.get("webUrl"),
    }


async def _upload_file(params: dict) -> dict:
    import base64

    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    site_id = _encode_id(params["site_id"])
    drive_id = _encode_id(params["drive_id"]) if params.get("drive_id") else None
    path = params["path"].strip("/")
    content_bytes = base64.b64decode(params["content"])
    base = f"/sites/{site_id}/drives/{drive_id}" if drive_id else f"/sites/{site_id}/drive"
    result = await client.put(f"{base}/root:/{path}:/content", data=content_bytes)
    return {
        "name": result.get("name"),
        "id": result.get("id"),
        "size": result.get("size"),
        "webUrl": result.get("webUrl"),
    }


async def _search(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    body = {
        "requests": [
            {
                "entityTypes": ["driveItem"],
                "query": {"queryString": params["query"]},
                "from": 0,
                "size": params.get("top", 10),
            }
        ]
    }
    data = await client.post("/search/query", json=body)
    hits_containers = data.get("value", [{}])[0].get("hitsContainers", [{}])
    hits = hits_containers[0].get("hits", []) if hits_containers else []
    return {
        "count": len(hits),
        "results": [
            {
                "name": h.get("resource", {}).get("name"),
                "webUrl": h.get("resource", {}).get("webUrl"),
                "size": h.get("resource", {}).get("size"),
                "lastModified": h.get("resource", {}).get("lastModifiedDateTime"),
                "summary": h.get("summary", "")[:200],
            }
            for h in hits
        ],
    }


async def _list_lists(params: dict) -> dict:
    """GET /sites/{id}/lists — list all SP lists (custom lists + libraries)."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    site_id = _encode_id(params["site_id"])
    data = await client.get(f"/sites/{site_id}/lists")
    lists = data.get("value", [])
    return {
        "count": len(lists),
        "lists": [
            {
                "id": l.get("id"),
                "displayName": l.get("displayName"),
                "description": l.get("description", "")[:200],
                "webUrl": l.get("webUrl"),
                "template": l.get("list", {}).get("template"),
            }
            for l in lists
        ],
    }


async def _list_list_items(params: dict) -> dict:
    """GET /sites/{id}/lists/{id}/items — list items in a SP list."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    site_id = _encode_id(params["site_id"])
    list_id = _encode_id(params["list_id"])
    top = params.get("top", 25)
    expand = "&$expand=fields" if params.get("expand_fields", True) else ""
    data = await client.get(
        f"/sites/{site_id}/lists/{list_id}/items?$top={top}{expand}"
    )
    items = data.get("value", [])
    return {
        "count": len(items),
        "items": [
            {
                "id": i.get("id"),
                "createdDateTime": i.get("createdDateTime"),
                "lastModifiedDateTime": i.get("lastModifiedDateTime"),
                "webUrl": i.get("webUrl"),
                "fields": i.get("fields", {}),
            }
            for i in items
        ],
    }


async def _create_list_item(params: dict) -> dict:
    """POST /sites/{id}/lists/{id}/items — create a new list item."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    site_id = _encode_id(params["site_id"])
    list_id = _encode_id(params["list_id"])
    body = {"fields": params["fields"]}
    result = await client.post(
        f"/sites/{site_id}/lists/{list_id}/items", json=body
    )
    return {
        "id": result.get("id"),
        "fields": result.get("fields", {}),
        "webUrl": result.get("webUrl"),
    }


async def _update_list_item(params: dict) -> dict:
    """PATCH /sites/{id}/lists/{id}/items/{id}/fields — update list item fields."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    site_id = _encode_id(params["site_id"])
    list_id = _encode_id(params["list_id"])
    item_id = _encode_id(params["item_id"])
    result = await client.patch(
        f"/sites/{site_id}/lists/{list_id}/items/{item_id}/fields",
        json=params["fields"],
    )
    return {
        "id": params["item_id"],
        "fields": result,
    }


async def _delete_list_item(params: dict) -> dict:
    """DELETE /sites/{id}/lists/{id}/items/{id} — delete a list item."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    site_id = _encode_id(params["site_id"])
    list_id = _encode_id(params["list_id"])
    item_id = _encode_id(params["item_id"])
    await client.delete(
        f"/sites/{site_id}/lists/{list_id}/items/{item_id}"
    )
    return {"deleted": True, "item_id": params["item_id"]}


HANDLERS = {
    "sharepoint_list_sites": _list_sites,
    "sharepoint_get_site": _get_site,
    "sharepoint_list_items": _list_items,
    "sharepoint_download_file": _download_file,
    "sharepoint_upload_file": _upload_file,
    "sharepoint_search": _search,
    "sharepoint_list_lists": _list_lists,
    "sharepoint_list_list_items": _list_list_items,
    "sharepoint_create_list_item": _create_list_item,
    "sharepoint_update_list_item": _update_list_item,
    "sharepoint_delete_list_item": _delete_list_item,
}
