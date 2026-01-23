"""Outlook Mail and Calendar service.

Provides operations for:
- Email: list, read, send, reply, search
- Calendar: list events, create, update, delete, find meeting times
- Mail folders: list, access different folders

Graph API Reference:
- Mail: https://learn.microsoft.com/graph/api/resources/mail-api-overview
- Calendar: https://learn.microsoft.com/graph/api/resources/calendar

"""

from datetime import datetime, timedelta
from typing import Any, Optional

from ..clients.graph_client import GraphClient


class OutlookService:
    """Outlook Mail and Calendar operations via Microsoft Graph.
    
    Example:
        >>> service = OutlookService(graph_client)
        >>> messages = await service.list_messages(count=10)
        >>> events = await service.list_events(days_ahead=7)
    """
    
    def __init__(self, client: GraphClient) -> None:
        """Initialize Outlook service.
        
        Args:
            client: GraphClient instance for API calls
        """
        self.client = client
    
    # =========================================================================
    # MAIL OPERATIONS
    # =========================================================================
    
    async def list_messages(
        self,
        folder: str = "inbox",
        count: int = 25,
        filter_query: Optional[str] = None,
        search: Optional[str] = None,
        select: Optional[list[str]] = None,
    ) -> list[dict[str, Any]]:
        """List email messages from a folder.
        
        Args:
            folder: Mail folder name (inbox, sentitems, drafts, deleteditems)
            count: Number of messages to retrieve (max 50)
            filter_query: OData filter expression
            search: Search query string
            select: Fields to return (default: common fields)
            
        Returns:
            List of message objects
        """
        default_select = [
            "id", "subject", "from", "receivedDateTime",
            "isRead", "hasAttachments", "bodyPreview", "importance"
        ]
        
        params: dict[str, Any] = {
            "$top": min(count, 50),
            "$orderby": "receivedDateTime desc",
            "$select": ",".join(select or default_select),
        }
        
        if filter_query:
            params["$filter"] = filter_query
        if search:
            params["$search"] = f'"{search}"'
        
        result = await self.client.get(
            f"/me/mailFolders/{folder}/messages",
            params=params
        )
        return result.get("value", [])
    
    async def get_message(
        self,
        message_id: str,
        include_body: bool = True
    ) -> dict[str, Any]:
        """Get a specific email message by ID.
        
        Args:
            message_id: The message ID
            include_body: Whether to include message body
            
        Returns:
            Full message object
        """
        select = [
            "id", "subject", "from", "toRecipients", "ccRecipients",
            "receivedDateTime", "sentDateTime", "isRead", "hasAttachments",
            "importance", "categories"
        ]
        if include_body:
            select.append("body")
        
        return await self.client.get(
            f"/me/messages/{message_id}",
            params={"$select": ",".join(select)}
        )
    
    async def send_message(
        self,
        to: list[str],
        subject: str,
        body: str,
        cc: Optional[list[str]] = None,
        bcc: Optional[list[str]] = None,
        is_html: bool = False,
        importance: str = "normal",
        save_to_sent: bool = True,
    ) -> None:
        """Send an email message.
        
        Args:
            to: List of recipient email addresses
            subject: Email subject
            body: Email body content
            cc: CC recipients
            bcc: BCC recipients
            is_html: Whether body is HTML (default: plain text)
            importance: 'low', 'normal', or 'high'
            save_to_sent: Whether to save to Sent Items
        """
        message: dict[str, Any] = {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body,
            },
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in to
            ],
            "importance": importance,
        }
        
        if cc:
            message["ccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in cc
            ]
        if bcc:
            message["bccRecipients"] = [
                {"emailAddress": {"address": addr}} for addr in bcc
            ]
        
        await self.client.post(
            "/me/sendMail",
            json={"message": message, "saveToSentItems": save_to_sent}
        )
    
    async def reply_to_message(
        self,
        message_id: str,
        body: str,
        reply_all: bool = False,
    ) -> None:
        """Reply to an email message.
        
        Args:
            message_id: ID of message to reply to
            body: Reply body content
            reply_all: Whether to reply to all recipients
        """
        action = "replyAll" if reply_all else "reply"
        await self.client.post(
            f"/me/messages/{message_id}/{action}",
            json={"comment": body}
        )
    
    async def forward_message(
        self,
        message_id: str,
        to: list[str],
        comment: Optional[str] = None,
    ) -> None:
        """Forward an email message.
        
        Args:
            message_id: ID of message to forward
            to: Recipient email addresses
            comment: Optional comment to add
        """
        body: dict[str, Any] = {
            "toRecipients": [
                {"emailAddress": {"address": addr}} for addr in to
            ]
        }
        if comment:
            body["comment"] = comment
        
        await self.client.post(
            f"/me/messages/{message_id}/forward",
            json=body
        )
    
    async def delete_message(self, message_id: str) -> None:
        """Delete an email message (moves to Deleted Items)."""
        await self.client.delete(f"/me/messages/{message_id}")
    
    async def mark_as_read(
        self,
        message_id: str,
        is_read: bool = True
    ) -> None:
        """Mark a message as read or unread."""
        await self.client.patch(
            f"/me/messages/{message_id}",
            json={"isRead": is_read}
        )
    
    async def list_folders(self) -> list[dict[str, Any]]:
        """List all mail folders.
        
        Returns:
            List of folder objects with id, displayName, etc.
        """
        result = await self.client.get(
            "/me/mailFolders",
            params={"$top": 100}
        )
        return result.get("value", [])
    
    # =========================================================================
    # CALENDAR OPERATIONS
    # =========================================================================
    
    async def list_events(
        self,
        days_ahead: int = 7,
        days_back: int = 0,
        count: int = 50,
        calendar_id: Optional[str] = None,
    ) -> list[dict[str, Any]]:
        """List calendar events in a time range.
        
        Args:
            days_ahead: Days to look ahead from today
            days_back: Days to look back from today
            count: Maximum events to return
            calendar_id: Specific calendar ID (default: primary)
            
        Returns:
            List of event objects
        """
        now = datetime.utcnow()
        start = (now - timedelta(days=days_back)).isoformat() + "Z"
        end = (now + timedelta(days=days_ahead)).isoformat() + "Z"
        
        params = {
            "startDateTime": start,
            "endDateTime": end,
            "$top": count,
            "$orderby": "start/dateTime",
            "$select": "id,subject,start,end,location,organizer,attendees,"
                       "isOnlineMeeting,onlineMeetingUrl,onlineMeeting,"
                       "bodyPreview,importance,showAs,isAllDay,isCancelled",
        }
        
        if calendar_id:
            endpoint = f"/me/calendars/{calendar_id}/calendarView"
        else:
            endpoint = "/me/calendarView"
        
        result = await self.client.get(endpoint, params=params)
        return result.get("value", [])
    
    async def get_event(self, event_id: str) -> dict[str, Any]:
        """Get a specific calendar event by ID."""
        return await self.client.get(f"/me/events/{event_id}")
    
    async def create_event(
        self,
        subject: str,
        start: datetime,
        end: datetime,
        timezone: str = "UTC",
        location: Optional[str] = None,
        body: Optional[str] = None,
        attendees: Optional[list[str]] = None,
        is_online_meeting: bool = False,
        is_all_day: bool = False,
        reminder_minutes: int = 15,
        show_as: str = "busy",
        importance: str = "normal",
        calendar_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Create a new calendar event.
        
        Args:
            subject: Event title
            start: Start datetime
            end: End datetime
            timezone: Timezone for start/end times
            location: Location name
            body: Event description
            attendees: List of attendee email addresses
            is_online_meeting: Create Teams meeting link
            is_all_day: All-day event
            reminder_minutes: Minutes before to remind (0 to disable)
            show_as: 'free', 'tentative', 'busy', 'oof', 'workingElsewhere'
            importance: 'low', 'normal', 'high'
            calendar_id: Specific calendar (default: primary)
            
        Returns:
            Created event object
        """
        event: dict[str, Any] = {
            "subject": subject,
            "start": {"dateTime": start.isoformat(), "timeZone": timezone},
            "end": {"dateTime": end.isoformat(), "timeZone": timezone},
            "isAllDay": is_all_day,
            "showAs": show_as,
            "importance": importance,
        }
        
        if location:
            event["location"] = {"displayName": location}
        
        if body:
            event["body"] = {"contentType": "Text", "content": body}
        
        if attendees:
            event["attendees"] = [
                {
                    "emailAddress": {"address": addr},
                    "type": "required"
                }
                for addr in attendees
            ]
        
        if is_online_meeting:
            event["isOnlineMeeting"] = True
            event["onlineMeetingProvider"] = "teamsForBusiness"
        
        if reminder_minutes > 0:
            event["reminderMinutesBeforeStart"] = reminder_minutes
            event["isReminderOn"] = True
        else:
            event["isReminderOn"] = False
        
        endpoint = f"/me/calendars/{calendar_id}/events" if calendar_id else "/me/events"
        return await self.client.post(endpoint, json=event)
    
    async def update_event(
        self,
        event_id: str,
        updates: dict[str, Any]
    ) -> dict[str, Any]:
        """Update an existing calendar event.
        
        Args:
            event_id: Event ID to update
            updates: Dict of fields to update
            
        Returns:
            Updated event object
        """
        return await self.client.patch(f"/me/events/{event_id}", json=updates)
    
    async def delete_event(self, event_id: str) -> None:
        """Delete a calendar event."""
        await self.client.delete(f"/me/events/{event_id}")
    
    async def respond_to_event(
        self,
        event_id: str,
        response: str,
        comment: Optional[str] = None,
        send_response: bool = True,
    ) -> None:
        """Respond to a meeting invitation.
        
        Args:
            event_id: Event ID to respond to
            response: 'accept', 'tentativelyAccept', or 'decline'
            comment: Optional comment with response
            send_response: Whether to send response to organizer
        """
        body: dict[str, Any] = {"sendResponse": send_response}
        if comment:
            body["comment"] = comment
        
        await self.client.post(f"/me/events/{event_id}/{response}", json=body)
    
    async def find_meeting_times(
        self,
        attendees: list[str],
        duration_minutes: int = 60,
        start: Optional[datetime] = None,
        end: Optional[datetime] = None,
        timezone: str = "UTC",
    ) -> list[dict[str, Any]]:
        """Find available meeting times for attendees.
        
        Args:
            attendees: List of attendee email addresses
            duration_minutes: Meeting duration in minutes
            start: Start of search window (default: now)
            end: End of search window (default: 7 days from now)
            timezone: Timezone for results
            
        Returns:
            List of meeting time suggestions
        """
        now = datetime.utcnow()
        search_start = start or now
        search_end = end or (now + timedelta(days=7))
        
        body = {
            "attendees": [
                {
                    "emailAddress": {"address": addr},
                    "type": "required"
                }
                for addr in attendees
            ],
            "timeConstraint": {
                "timeslots": [{
                    "start": {
                        "dateTime": search_start.isoformat(),
                        "timeZone": timezone
                    },
                    "end": {
                        "dateTime": search_end.isoformat(),
                        "timeZone": timezone
                    }
                }]
            },
            "meetingDuration": f"PT{duration_minutes}M",
        }
        
        result = await self.client.post("/me/findMeetingTimes", json=body)
        return result.get("meetingTimeSuggestions", [])
    
    async def list_calendars(self) -> list[dict[str, Any]]:
        """List all calendars.
        
        Returns:
            List of calendar objects
        """
        result = await self.client.get(
            "/me/calendars",
            params={"$select": "id,name,color,isDefaultCalendar,canEdit"}
        )
        return result.get("value", [])
