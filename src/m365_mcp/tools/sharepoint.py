"""SharePoint MCP tools.

Covers: list sites, list library items, download, upload, search.
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
]


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


async def _list_items(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    site_id = params["site_id"]
    drive_id = params.get("drive_id")
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
    site_id = params["site_id"]
    item_id = params["item_id"]
    drive_id = params.get("drive_id")
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
    site_id = params["site_id"]
    drive_id = params.get("drive_id")
    path = params["path"].strip("/")
    content_bytes = base64.b64decode(params["content"])
    base = f"/sites/{site_id}/drives/{drive_id}" if drive_id else f"/sites/{site_id}/drive"
    result = await client.put(f"{base}/root:/{path}:/content", content=content_bytes)
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


HANDLERS = {
    "sharepoint_list_sites": _list_sites,
    "sharepoint_list_items": _list_items,
    "sharepoint_download_file": _download_file,
    "sharepoint_upload_file": _upload_file,
    "sharepoint_search": _search,
}
