"""Teams MCP tools.

Covers: list joined teams, list channels, send channel message.
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
]


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


HANDLERS = {
    "teams_list_teams": _list_teams,
    "teams_list_channels": _list_channels,
    "teams_send_message": _send_message,
}
