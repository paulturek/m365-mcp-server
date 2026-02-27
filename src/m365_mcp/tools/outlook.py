"""Outlook MCP tools.

Covers: mail listing, reading, send, update, move, reply, forward, delete,
        calendar events (list, create, update, delete), mail folders.
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
                "filter": {
                    "type": "string",
                    "description": (
                        "OData $filter expression. Supported: eq, ne, startsWith(), "
                        "isRead eq true/false, hasAttachments eq true, "
                        "from/emailAddress/address eq '...'. "
                        "NOTE: contains() is NOT supported. "
                        "Cannot be combined with 'search' — use one or the other."
                    ),
                },
                "search": {
                    "type": "string",
                    "description": (
                        "Keyword search across subject, body, and sender. "
                        "Use this instead of $filter for keyword matching. "
                        "Cannot be combined with 'filter' or $orderby — "
                        "results use relevance ranking."
                    ),
                },
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "outlook_get_message",
        "description": "Get a single email message with full body content.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "message_id": {
                    "type": "string",
                    "description": "The Graph message ID",
                },
            },
            "required": ["user_id", "message_id"],
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
        "name": "outlook_delete_message",
        "description": "Delete an email message (moves to Deleted Items).",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "message_id": {
                    "type": "string",
                    "description": "The Graph message ID to delete",
                },
            },
            "required": ["user_id", "message_id"],
        },
    },
    {
        "name": "outlook_move_message",
        "description": "Move an email message to a different mail folder.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "message_id": {
                    "type": "string",
                    "description": "The Graph message ID to move",
                },
                "destination_folder": {
                    "type": "string",
                    "description": "Destination folder ID or well-known name (e.g. 'archive', 'deleteditems', 'drafts')",
                },
            },
            "required": ["user_id", "message_id", "destination_folder"],
        },
    },
    {
        "name": "outlook_reply_mail",
        "description": "Reply to an email message.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "message_id": {
                    "type": "string",
                    "description": "The Graph message ID to reply to",
                },
                "comment": {
                    "type": "string",
                    "description": "Reply body (HTML supported)",
                },
                "reply_all": {
                    "type": "boolean",
                    "default": False,
                    "description": "Reply to all recipients",
                },
            },
            "required": ["user_id", "message_id", "comment"],
        },
    },
    {
        "name": "outlook_forward_mail",
        "description": "Forward an email message to new recipients.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "message_id": {
                    "type": "string",
                    "description": "The Graph message ID to forward",
                },
                "to": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Recipient email addresses",
                },
                "comment": {
                    "type": "string",
                    "description": "Optional message to include with the forward",
                },
            },
            "required": ["user_id", "message_id", "to"],
        },
    },
    {
        "name": "outlook_list_mail_folders",
        "description": "List mail folders in the user's mailbox.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "top": {"type": "integer", "default": 25},
            },
            "required": ["user_id"],
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
    {
        "name": "outlook_update_event",
        "description": "Update an existing calendar event (reschedule, add attendees, etc.).",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "event_id": {
                    "type": "string",
                    "description": "The Graph event ID to update",
                },
                "subject": {"type": "string"},
                "start": {"type": "string", "description": "ISO datetime for new start"},
                "end": {"type": "string", "description": "ISO datetime for new end"},
                "timezone": {"type": "string", "default": "America/Los_Angeles"},
                "body": {"type": "string"},
                "location": {"type": "string"},
                "attendees": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Updated attendee list (replaces existing)",
                },
                "is_online_meeting": {"type": "boolean"},
            },
            "required": ["user_id", "event_id"],
        },
    },
    {
        "name": "outlook_delete_event",
        "description": "Delete a calendar event.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "event_id": {
                    "type": "string",
                    "description": "The Graph event ID to delete",
                },
            },
            "required": ["user_id", "event_id"],
        },
    },
]


# ---- Handlers -----------------------------------------------------------


async def _list_mail(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    folder = params.get("folder", "inbox")
    top = params.get("top", 10)
    has_search = bool(params.get("search"))

    # Graph API mutual exclusions:
    # - $search cannot be combined with $orderby (400 SearchWithOrderBy)
    # - $search cannot be combined with $filter (400)
    # When searching, Graph returns results by relevance ranking.
    if has_search:
        qp = f"?$top={top}&$search=\"{params['search']}\""
    else:
        qp = f"?$top={top}&$orderby=receivedDateTime desc"
        if params.get("filter"):
            qp += f"&$filter={params['filter']}"

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


async def _get_message(params: dict) -> dict:
    """GET /me/messages/{id} — full message with body."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    message_id = params["message_id"]
    data = await client.get(f"/me/messages/{message_id}")
    return {
        "id": data.get("id"),
        "subject": data.get("subject"),
        "from": data.get("from", {}).get("emailAddress", {}).get("address"),
        "toRecipients": [
            r.get("emailAddress", {}).get("address")
            for r in data.get("toRecipients", [])
        ],
        "ccRecipients": [
            r.get("emailAddress", {}).get("address")
            for r in data.get("ccRecipients", [])
        ],
        "receivedDateTime": data.get("receivedDateTime"),
        "isRead": data.get("isRead"),
        "importance": data.get("importance"),
        "hasAttachments": data.get("hasAttachments"),
        "body": {
            "contentType": data.get("body", {}).get("contentType"),
            "content": data.get("body", {}).get("content"),
        },
        "categories": data.get("categories"),
        "flag": data.get("flag"),
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


async def _delete_message(params: dict) -> dict:
    """DELETE /me/messages/{id} — moves message to Deleted Items."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    message_id = params["message_id"]
    await client.delete(f"/me/messages/{message_id}")
    return {"deleted": True, "message_id": message_id}


async def _move_message(params: dict) -> dict:
    """POST /me/messages/{id}/move — move to a different folder."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    message_id = params["message_id"]
    body = {"destinationId": params["destination_folder"]}
    result = await client.post(f"/me/messages/{message_id}/move", json=body)
    return {
        "id": result.get("id"),
        "subject": result.get("subject"),
        "parentFolderId": result.get("parentFolderId"),
    }


async def _reply_mail(params: dict) -> dict:
    """POST /me/messages/{id}/reply or /replyAll."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    message_id = params["message_id"]
    action = "replyAll" if params.get("reply_all") else "reply"
    body = {"comment": params["comment"]}
    await client.post(f"/me/messages/{message_id}/{action}", json=body)
    return {"replied": True, "action": action, "message_id": message_id}


async def _forward_mail(params: dict) -> dict:
    """POST /me/messages/{id}/forward."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    message_id = params["message_id"]
    to_recipients = [
        {"emailAddress": {"address": addr}} for addr in params["to"]
    ]
    body = {"toRecipients": to_recipients}
    if params.get("comment"):
        body["comment"] = params["comment"]
    await client.post(f"/me/messages/{message_id}/forward", json=body)
    return {"forwarded": True, "to": params["to"], "message_id": message_id}


async def _list_mail_folders(params: dict) -> dict:
    """GET /me/mailFolders — list mail folders."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    top = params.get("top", 25)
    data = await client.get(f"/me/mailFolders?$top={top}")
    folders = data.get("value", [])
    return {
        "count": len(folders),
        "folders": [
            {
                "id": f.get("id"),
                "displayName": f.get("displayName"),
                "totalItemCount": f.get("totalItemCount"),
                "unreadItemCount": f.get("unreadItemCount"),
                "parentFolderId": f.get("parentFolderId"),
            }
            for f in folders
        ],
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


async def _update_event(params: dict) -> dict:
    """PATCH /me/events/{id} — update event properties."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    event_id = params["event_id"]
    tz = params.get("timezone", "America/Los_Angeles")

    patch: dict = {}
    if "subject" in params:
        patch["subject"] = params["subject"]
    if "start" in params:
        patch["start"] = {"dateTime": params["start"], "timeZone": tz}
    if "end" in params:
        patch["end"] = {"dateTime": params["end"], "timeZone": tz}
    if "body" in params:
        patch["body"] = {"contentType": "HTML", "content": params["body"]}
    if "location" in params:
        patch["location"] = {"displayName": params["location"]}
    if "attendees" in params:
        patch["attendees"] = [
            {"emailAddress": {"address": a}, "type": "required"}
            for a in params["attendees"]
        ]
    if "is_online_meeting" in params:
        patch["isOnlineMeeting"] = params["is_online_meeting"]
        if params["is_online_meeting"]:
            patch["onlineMeetingProvider"] = "teamsForBusiness"

    if not patch:
        return {"error": "No updatable properties provided"}

    result = await client.patch(f"/me/events/{event_id}", json=patch)
    return {
        "id": result.get("id"),
        "subject": result.get("subject"),
        "start": result.get("start"),
        "end": result.get("end"),
        "webLink": result.get("webLink"),
    }


async def _delete_event(params: dict) -> dict:
    """DELETE /me/events/{id} — delete a calendar event."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    event_id = params["event_id"]
    await client.delete(f"/me/events/{event_id}")
    return {"deleted": True, "event_id": event_id}


HANDLERS = {
    "outlook_list_mail": _list_mail,
    "outlook_get_message": _get_message,
    "outlook_send_mail": _send_mail,
    "outlook_update_message": _update_message,
    "outlook_delete_message": _delete_message,
    "outlook_move_message": _move_message,
    "outlook_reply_mail": _reply_mail,
    "outlook_forward_mail": _forward_mail,
    "outlook_list_mail_folders": _list_mail_folders,
    "outlook_list_calendar_events": _list_calendar_events,
    "outlook_create_event": _create_event,
    "outlook_update_event": _update_event,
    "outlook_delete_event": _delete_event,
}
