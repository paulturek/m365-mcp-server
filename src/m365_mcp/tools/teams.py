"""Teams MCP tools.

Covers: list joined teams, list channels, send channel message,
        read channel messages, list chats, read chat messages, send chat message.
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
        "name": "teams_list_teams",
        "description": "List Microsoft Teams the user has joined.",
        "inputSchema": {
            "type": "object",
            "properties": {**_USER_ID_PROP},
            "required": ["user_id"],
        },
    },
    {
        "name": "teams_list_channels",
        "description": "List channels in a Microsoft Team.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "team_id": {"type": "string", "description": "Team ID"},
            },
            "required": ["user_id", "team_id"],
        },
    },
    {
        "name": "teams_send_message",
        "description": "Send a message to a Teams channel.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "team_id": {"type": "string"},
                "channel_id": {"type": "string"},
                "message": {"type": "string", "description": "Message body (HTML supported)"},
                "content_type": {
                    "type": "string",
                    "enum": ["text", "html"],
                    "default": "html",
                },
            },
            "required": ["user_id", "team_id", "channel_id", "message"],
        },
    },
    {
        "name": "teams_list_channel_messages",
        "description": "List recent messages in a Teams channel.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "team_id": {"type": "string", "description": "Team ID"},
                "channel_id": {"type": "string", "description": "Channel ID"},
                "top": {"type": "integer", "default": 20, "description": "Max messages to return"},
            },
            "required": ["user_id", "team_id", "channel_id"],
        },
    },
    {
        "name": "teams_list_chats",
        "description": "List the user's 1:1 and group chats.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "top": {"type": "integer", "default": 25, "description": "Max chats to return"},
                "include_members": {
                    "type": "boolean",
                    "default": False,
                    "description": "Include chat member names (requires ChatMember.Read.All permission)",
                },
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "teams_list_chat_messages",
        "description": "List recent messages in a 1:1 or group chat.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "chat_id": {"type": "string", "description": "Chat ID"},
                "top": {"type": "integer", "default": 20, "description": "Max messages to return"},
            },
            "required": ["user_id", "chat_id"],
        },
    },
    {
        "name": "teams_send_chat_message",
        "description": "Send a message in a 1:1 or group chat.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "chat_id": {"type": "string", "description": "Chat ID"},
                "message": {"type": "string", "description": "Message body (HTML supported)"},
                "content_type": {
                    "type": "string",
                    "enum": ["text", "html"],
                    "default": "html",
                },
            },
            "required": ["user_id", "chat_id", "message"],
        },
    },
]


# ---- Handlers -----------------------------------------------------------


async def _list_teams(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    data = await client.get("/me/joinedTeams")
    teams = data.get("value", [])
    return {
        "count": len(teams),
        "teams": [
            {
                "id": t.get("id"),
                "displayName": t.get("displayName"),
                "description": t.get("description", "")[:200],
            }
            for t in teams
        ],
    }


async def _list_channels(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    team_id = params["team_id"]
    data = await client.get(f"/teams/{team_id}/channels")
    channels = data.get("value", [])
    return {
        "count": len(channels),
        "channels": [
            {
                "id": c.get("id"),
                "displayName": c.get("displayName"),
                "description": c.get("description", ""),
                "membershipType": c.get("membershipType"),
                "webUrl": c.get("webUrl"),
            }
            for c in channels
        ],
    }


async def _send_message(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    team_id = params["team_id"]
    channel_id = params["channel_id"]
    body = {
        "body": {
            "contentType": params.get("content_type", "html"),
            "content": params["message"],
        }
    }
    result = await client.post(
        f"/teams/{team_id}/channels/{channel_id}/messages", json=body
    )
    return {
        "id": result.get("id"),
        "createdDateTime": result.get("createdDateTime"),
        "webUrl": result.get("webUrl"),
    }


async def _list_channel_messages(params: dict) -> dict:
    """GET /teams/{id}/channels/{id}/messages — list channel messages."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    team_id = params["team_id"]
    channel_id = params["channel_id"]
    top = params.get("top", 20)
    data = await client.get(
        f"/teams/{team_id}/channels/{channel_id}/messages?$top={top}"
    )
    messages = data.get("value", [])
    return {
        "count": len(messages),
        "messages": [
            {
                "id": m.get("id"),
                "createdDateTime": m.get("createdDateTime"),
                "from": (
                    m.get("from", {})
                    .get("user", {})
                    .get("displayName")
                ),
                "body": m.get("body", {}).get("content", "")[:500],
                "contentType": m.get("body", {}).get("contentType"),
                "webUrl": m.get("webUrl"),
            }
            for m in messages
        ],
    }


async def _list_chats(params: dict) -> dict:
    """GET /me/chats — list user's 1:1 and group chats.

    $expand=members requires ChatMember.Read.All — made optional via include_members param.
    """
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    top = params.get("top", 25)
    include_members = params.get("include_members", False)

    endpoint = f"/me/chats?$top={top}"
    if include_members:
        endpoint += "&$expand=members"

    data = await client.get(endpoint)
    chats = data.get("value", [])
    return {
        "count": len(chats),
        "chats": [
            {
                "id": c.get("id"),
                "topic": c.get("topic"),
                "chatType": c.get("chatType"),
                "createdDateTime": c.get("createdDateTime"),
                "lastUpdatedDateTime": c.get("lastUpdatedDateTime"),
                **(
                    {"members": [mb.get("displayName") for mb in c.get("members", [])]}
                    if include_members else {}
                ),
            }
            for c in chats
        ],
    }


async def _list_chat_messages(params: dict) -> dict:
    """GET /me/chats/{id}/messages — list messages in a chat."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    chat_id = params["chat_id"]
    top = params.get("top", 20)
    data = await client.get(f"/me/chats/{chat_id}/messages?$top={top}")
    messages = data.get("value", [])
    return {
        "count": len(messages),
        "messages": [
            {
                "id": m.get("id"),
                "createdDateTime": m.get("createdDateTime"),
                "from": (
                    m.get("from", {})
                    .get("user", {})
                    .get("displayName")
                ),
                "body": m.get("body", {}).get("content", "")[:500],
                "contentType": m.get("body", {}).get("contentType"),
            }
            for m in messages
        ],
    }


async def _send_chat_message(params: dict) -> dict:
    """POST /me/chats/{id}/messages — send a message in a chat."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    chat_id = params["chat_id"]
    body = {
        "body": {
            "contentType": params.get("content_type", "html"),
            "content": params["message"],
        }
    }
    result = await client.post(f"/me/chats/{chat_id}/messages", json=body)
    return {
        "id": result.get("id"),
        "createdDateTime": result.get("createdDateTime"),
    }


HANDLERS = {
    "teams_list_teams": _list_teams,
    "teams_list_channels": _list_channels,
    "teams_send_message": _send_message,
    "teams_list_channel_messages": _list_channel_messages,
    "teams_list_chats": _list_chats,
    "teams_list_chat_messages": _list_chat_messages,
    "teams_send_chat_message": _send_chat_message,
}
