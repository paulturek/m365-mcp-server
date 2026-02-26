"""Outlook MCP tools.

Covers: mail listing, send, update, calendar events.
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
        "name": "outlook_list_mail",
        "description": "List recent email messages from the user's mailbox.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "folder": {"type": "string", "default": "inbox", "description": "Mail folder"},
                "top": {"type": "integer", "default": 10, "description": "Max messages to return"},
                "filter": {"type": "string", "description": "OData $filter expression"},
                "search": {"type": "string", "description": "$search keyword"},
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "outlook_send_mail",
        "description": "Send an email on behalf of the user.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "to": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Recipient email addresses",
                },
                "subject": {"type": "string"},
                "body": {"type": "string", "description": "Email body (HTML supported)"},
                "content_type": {
                    "type": "string",
                    "enum": ["Text", "HTML"],
                    "default": "HTML",
                },
                "cc": {"type": "array", "items": {"type": "string"}, "description": "CC addresses"},
            },
            "required": ["user_id", "to", "subject", "body"],
        },
    },
    {
        "name": "outlook_update_message",
        "description": (
            "Update properties of an existing email message. "
            "Supports marking read/unread, setting importance, "
            "managing categories, and toggling follow-up flags."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "message_id": {
                    "type": "string",
                    "description": "The Graph message ID to update",
                },
                "is_read": {
                    "type": "boolean",
                    "description": "Mark message as read (true) or unread (false)",
                },
                "importance": {
                    "type": "string",
                    "enum": ["low", "normal", "high"],
                    "description": "Message importance level",
                },
                "categories": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Outlook category labels to assign",
                },
                "flag": {
                    "type": "object",
                    "description": "Follow-up flag. Example: {\"flagStatus\": \"flagged\"}",
                    "properties": {
                        "flagStatus": {
                            "type": "string",
                            "enum": ["notFlagged", "flagged", "complete"],
                        }
                    },
                },
            },
            "required": ["user_id", "message_id"],
        },
    },
    {
        "name": "outlook_list_calendar_events",
        "description": "List upcoming calendar events.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "top": {"type": "integer", "default": 10},
                "start_datetime": {"type": "string", "description": "ISO datetime start bound"},
                "end_datetime": {"type": "string", "description": "ISO datetime end bound"},
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "outlook_create_event",
        "description": "Create a new calendar event.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "subject": {"type": "string"},
                "start": {"type": "string", "description": "ISO datetime for event start"},
                "end": {"type": "string", "description": "ISO datetime for event end"},
                "timezone": {"type": "string", "default": "America/Los_Angeles"},
                "body": {"type": "string", "description": "Event body/description"},
                "attendees": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Attendee email addresses",
                },
                "location": {"type": "string", "description": "Event location"},
                "is_online_meeting": {"type": "boolean", "default": False},
            },
            "required": ["user_id", "subject", "start", "end"],
        },
    },
]


async def _list_mail(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    folder = params.get("folder", "inbox")
    top = params.get("top", 10)
    qp = f"?$top={top}&$orderby=receivedDateTime desc"
    if params.get("filter"):
        qp += f"&$filter={params['filter']}"
    if params.get("search"):
        qp += f"&$search=\"{params['search']}\""
    data = await client.get(f"/me/mailFolders/{folder}/messages{qp}")
    messages = data.get("value", [])
    return {
        "count": len(messages),
        "messages": [
            {
                "id": m.get("id"),
                "subject": m.get("subject"),
                "from": m.get("from", {}).get("emailAddress", {}).get("address"),
                "receivedDateTime": m.get("receivedDateTime"),
                "isRead": m.get("isRead"),
                "hasAttachments": m.get("hasAttachments"),
                "bodyPreview": m.get("bodyPreview", "")[:200],
            }
            for m in messages
        ],
    }


async def _send_mail(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    to_recipients = [
        {"emailAddress": {"address": addr}} for addr in params["to"]
    ]
    cc_recipients = [
        {"emailAddress": {"address": addr}} for addr in params.get("cc", [])
    ]
    body = {
        "message": {
            "subject": params["subject"],
            "body": {
                "contentType": params.get("content_type", "HTML"),
                "content": params["body"],
            },
            "toRecipients": to_recipients,
        },
        "saveToSentItems": True,
    }
    if cc_recipients:
        body["message"]["ccRecipients"] = cc_recipients
    await client.post("/me/sendMail", json=body)
    return {"sent": True, "to": params["to"], "subject": params["subject"]}


async def _update_message(params: dict) -> dict:
    """PATCH /me/messages/{id} — update mutable message properties."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    message_id = params["message_id"]

    # Build patch body from only the supplied optional fields
    patch: dict = {}
    if "is_read" in params:
        patch["isRead"] = params["is_read"]
    if "importance" in params:
        patch["importance"] = params["importance"]
    if "categories" in params:
        patch["categories"] = params["categories"]
    if "flag" in params:
        patch["flag"] = params["flag"]

    if not patch:
        return {"error": "No updatable properties provided"}

    result = await client.patch(f"/me/messages/{message_id}", json=patch)
    return {
        "id": result.get("id"),
        "subject": result.get("subject"),
        "isRead": result.get("isRead"),
        "importance": result.get("importance"),
        "categories": result.get("categories"),
        "flag": result.get("flag"),
    }


async def _list_calendar_events(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    top = params.get("top", 10)
    start = params.get("start_datetime")
    end = params.get("end_datetime")
    if start and end:
        endpoint = f"/me/calendarView?startDateTime={start}&endDateTime={end}&$top={top}&$orderby=start/dateTime"
    else:
        endpoint = f"/me/events?$top={top}&$orderby=start/dateTime"
    data = await client.get(endpoint)
    events = data.get("value", [])
    return {
        "count": len(events),
        "events": [
            {
                "id": e.get("id"),
                "subject": e.get("subject"),
                "start": e.get("start"),
                "end": e.get("end"),
                "location": e.get("location", {}).get("displayName"),
                "organizer": e.get("organizer", {}).get("emailAddress", {}).get("address"),
                "isOnlineMeeting": e.get("isOnlineMeeting"),
                "webLink": e.get("webLink"),
            }
            for e in events
        ],
    }


async def _create_event(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    tz = params.get("timezone", "America/Los_Angeles")
    body = {
        "subject": params["subject"],
        "start": {"dateTime": params["start"], "timeZone": tz},
        "end": {"dateTime": params["end"], "timeZone": tz},
    }
    if params.get("body"):
        body["body"] = {"contentType": "HTML", "content": params["body"]}
    if params.get("location"):
        body["location"] = {"displayName": params["location"]}
    if params.get("attendees"):
        body["attendees"] = [
            {"emailAddress": {"address": a}, "type": "required"}
            for a in params["attendees"]
        ]
    if params.get("is_online_meeting"):
        body["isOnlineMeeting"] = True
        body["onlineMeetingProvider"] = "teamsForBusiness"
    result = await client.post("/me/events", json=body)
    return {
        "id": result.get("id"),
        "subject": result.get("subject"),
        "webLink": result.get("webLink"),
        "start": result.get("start"),
        "end": result.get("end"),
    }


HANDLERS = {
    "outlook_list_mail": _list_mail,
    "outlook_send_mail": _send_mail,
    "outlook_update_message": _update_message,
    "outlook_list_calendar_events": _list_calendar_events,
    "outlook_create_event": _create_event,
}
