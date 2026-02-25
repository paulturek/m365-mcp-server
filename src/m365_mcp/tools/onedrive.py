"""OneDrive MCP tools.

Covers: list files/folders, download, upload, delete, create folder, share.
"""
import logging

from ..auth.oauth_web import get_access_token
from ..clients.graph_client import GraphClient

logger = logging.getLogger(__name__)

# ---- Shared schema fragment for user_id --------------------------------
_USER_ID_PROP = {
    "user_id": {
        "type": "string",
        "description": "Your user identifier (email recommended)",
    }
}

# ---- Tool definitions ---------------------------------------------------

TOOLS = [
    {
        "name": "onedrive_list_files",
        "description": "List files and folders in a OneDrive directory.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "path": {
                    "type": "string",
                    "description": "Folder path (e.g. '/' or '/Documents')",
                    "default": "/",
                },
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "onedrive_download_file",
        "description": "Download a file from OneDrive. Returns file content or a download URL.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {
                    "type": "string",
                    "description": "Full path to the file (e.g. '/Documents/report.xlsx')",
                },
                "item_id": {
                    "type": "string",
                    "description": "OneDrive item ID (alternative to item_path)",
                },
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "onedrive_upload_file",
        "description": "Upload a file to OneDrive.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "path": {
                    "type": "string",
                    "description": "Destination path including filename (e.g. '/Documents/report.pdf')",
                },
                "content": {
                    "type": "string",
                    "description": "Base64-encoded file content",
                },
            },
            "required": ["user_id", "path", "content"],
        },
    },
    {
        "name": "onedrive_delete_item",
        "description": "Delete a file or folder from OneDrive.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {
                    "type": "string",
                    "description": "Path to the item to delete",
                },
                "item_id": {
                    "type": "string",
                    "description": "OneDrive item ID (alternative to item_path)",
                },
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "onedrive_create_folder",
        "description": "Create a new folder in OneDrive.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "parent_path": {
                    "type": "string",
                    "description": "Parent folder path (e.g. '/Documents')",
                    "default": "/",
                },
                "folder_name": {
                    "type": "string",
                    "description": "Name of the new folder",
                },
            },
            "required": ["user_id", "folder_name"],
        },
    },
    {
        "name": "onedrive_share_item",
        "description": "Create a sharing link for a OneDrive file or folder.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {
                    "type": "string",
                    "description": "Path to the item to share",
                },
                "item_id": {
                    "type": "string",
                    "description": "OneDrive item ID (alternative)",
                },
                "type": {
                    "type": "string",
                    "enum": ["view", "edit"],
                    "default": "view",
                    "description": "Permission level",
                },
                "scope": {
                    "type": "string",
                    "enum": ["anonymous", "organization"],
                    "default": "organization",
                    "description": "Sharing scope",
                },
            },
            "required": ["user_id"],
        },
    },
]

# ---- Handlers -----------------------------------------------------------


async def _list_files(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    path = params.get("path", "/")
    endpoint = (
        "/me/drive/root/children"
        if path in ("/", "")
        else f"/me/drive/root:/{path.strip('/')}:/children"
    )
    data = await client.get(endpoint)
    items = data.get("value", [])
    return {
        "count": len(items),
        "items": [
            {
                "name": i.get("name"),
                "type": "folder" if "folder" in i else "file",
                "size": i.get("size"),
                "id": i.get("id"),
                "lastModified": i.get("lastModifiedDateTime"),
                "webUrl": i.get("webUrl"),
            }
            for i in items
        ],
    }


async def _download_file(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    if params.get("item_id"):
        endpoint = f"/me/drive/items/{params['item_id']}"
    elif params.get("item_path"):
        endpoint = f"/me/drive/root:/{params['item_path'].strip('/')}"
    else:
        return {"error": "Provide item_path or item_id"}
    meta = await client.get(endpoint)
    download_url = meta.get("@microsoft.graph.downloadUrl", "")
    return {
        "name": meta.get("name"),
        "size": meta.get("size"),
        "downloadUrl": download_url,
        "mimeType": meta.get("file", {}).get("mimeType"),
        "webUrl": meta.get("webUrl"),
    }


async def _upload_file(params: dict) -> dict:
    import base64

    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    path = params["path"].strip("/")
    content_bytes = base64.b64decode(params["content"])
    endpoint = f"/me/drive/root:/{path}:/content"
    result = await client.put(endpoint, content=content_bytes)
    return {
        "name": result.get("name"),
        "id": result.get("id"),
        "size": result.get("size"),
        "webUrl": result.get("webUrl"),
    }


async def _delete_item(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    if params.get("item_id"):
        endpoint = f"/me/drive/items/{params['item_id']}"
    elif params.get("item_path"):
        endpoint = f"/me/drive/root:/{params['item_path'].strip('/')}"
    else:
        return {"error": "Provide item_path or item_id"}
    await client.delete(endpoint)
    return {"deleted": True}


async def _create_folder(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    parent = params.get("parent_path", "/").strip("/")
    parent_endpoint = (
        "/me/drive/root/children"
        if parent in ("", "/")
        else f"/me/drive/root:/{parent}:/children"
    )
    body = {
        "name": params["folder_name"],
        "folder": {},
        "@microsoft.graph.conflictBehavior": "rename",
    }
    result = await client.post(parent_endpoint, json=body)
    return {
        "name": result.get("name"),
        "id": result.get("id"),
        "webUrl": result.get("webUrl"),
    }


async def _share_item(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    if params.get("item_id"):
        base = f"/me/drive/items/{params['item_id']}"
    elif params.get("item_path"):
        base = f"/me/drive/root:/{params['item_path'].strip('/')}"
    else:
        return {"error": "Provide item_path or item_id"}
    body = {
        "type": params.get("type", "view"),
        "scope": params.get("scope", "organization"),
    }
    result = await client.post(f"{base}:/createLink", json=body)
    link = result.get("link", {})
    return {
        "webUrl": link.get("webUrl"),
        "type": link.get("type"),
        "scope": link.get("scope"),
    }


HANDLERS = {
    "onedrive_list_files": _list_files,
    "onedrive_download_file": _download_file,
    "onedrive_upload_file": _upload_file,
    "onedrive_delete_item": _delete_item,
    "onedrive_create_folder": _create_folder,
    "onedrive_share_item": _share_item,
}
