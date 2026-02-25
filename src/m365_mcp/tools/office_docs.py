"""Office Documents MCP tools.

Covers: get document content (Word/PPT), convert between formats.
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
        "name": "docs_get_content",
        "description": "Get the content / metadata of a Word or PowerPoint document in OneDrive.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {"type": "string", "description": "OneDrive path to the document"},
                "item_id": {"type": "string", "description": "OneDrive item ID (alternative)"},
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "docs_convert",
        "description": "Convert a document to another format (e.g. DOCX to PDF). Returns a download URL.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {"type": "string", "description": "OneDrive path"},
                "item_id": {"type": "string", "description": "OneDrive item ID (alternative)"},
                "format": {
                    "type": "string",
                    "description": "Target format",
                    "enum": ["pdf", "html", "jpg", "png"],
                    "default": "pdf",
                },
            },
            "required": ["user_id"],
        },
    },
]


def _item_endpoint(params: dict) -> str:
    if params.get("item_id"):
        return f"/me/drive/items/{params['item_id']}"
    path = params.get("item_path", "").strip("/")
    return f"/me/drive/root:/{path}"


async def _get_content(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    endpoint = _item_endpoint(params)
    meta = await client.get(endpoint)
    return {
        "name": meta.get("name"),
        "size": meta.get("size"),
        "mimeType": meta.get("file", {}).get("mimeType"),
        "lastModified": meta.get("lastModifiedDateTime"),
        "downloadUrl": meta.get("@microsoft.graph.downloadUrl", ""),
        "webUrl": meta.get("webUrl"),
        "createdBy": meta.get("createdBy", {}).get("user", {}).get("displayName"),
    }


async def _convert_doc(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    endpoint = _item_endpoint(params)
    fmt = params.get("format", "pdf")
    # Graph conversion endpoint: /content?format=pdf  returns a 302 redirect
    convert_url = f"{endpoint}/content?format={fmt}"
    # We use the client to get the redirect URL
    result = await client.get_redirect_url(convert_url)
    return {
        "format": fmt,
        "downloadUrl": result,
        "note": f"Document converted to {fmt}. Use the download URL to retrieve.",
    }


HANDLERS = {
    "docs_get_content": _get_content,
    "docs_convert": _convert_doc,
}
