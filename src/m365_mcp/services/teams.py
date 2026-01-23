"""Microsoft Teams service.

Provides operations for:
- Teams: list, get info
- Channels: list, get, send messages
- Chats: list, send messages
- Members: list team members

Graph API Reference:
    https://learn.microsoft.com/graph/api/resources/teams-api-overview

Permissions:
    - Team.ReadBasic.All: List and read teams
    - Channel.ReadBasic.All: List and read channels
    - ChannelMessage.Send: Send channel messages
    - Chat.ReadWrite: Read and write chats

Note:
    Some Teams operations require the user to be a member of the team.
    Application permissions have additional restrictions.

"""

from typing import Any, Optional

from ..clients.graph_client import GraphClient


class TeamsService:
    """Microsoft Teams operations via Microsoft Graph.
    
    Example:
        >>> service = TeamsService(graph_client)
        >>> teams = await service.list_my_teams()
        >>> await service.send_channel_message(team_id, channel_id, "Hello!")
    """
    
    def __init__(self, client: GraphClient) -> None:
        """Initialize Teams service."""
        self.client = client
    
    # =========================================================================
    # TEAMS
    # =========================================================================
    
    async def list_my_teams(self) -> list[dict[str, Any]]:
        """List teams the user is a member of.
        
        Returns:
            List of team objects with id, displayName, description
        """
        result = await self.client.get(
            "/me/joinedTeams",
            params={"$select": "id,displayName,description,webUrl"}
        )
        return result.get("value", [])
    
    async def get_team(self, team_id: str) -> dict[str, Any]:
        """Get details of a specific team."""
        return await self.client.get(f"/teams/{team_id}")
    
    # =========================================================================
    # CHANNELS
    # =========================================================================
    
    async def list_channels(self, team_id: str) -> list[dict[str, Any]]:
        """List channels in a team.
        
        Returns:
            List of channel objects
        """
        result = await self.client.get(
            f"/teams/{team_id}/channels",
            params={"$select": "id,displayName,description,membershipType,webUrl"}
        )
        return result.get("value", [])
    
    async def get_channel(
        self,
        team_id: str,
        channel_id: str
    ) -> dict[str, Any]:
        """Get details of a specific channel."""
        return await self.client.get(
            f"/teams/{team_id}/channels/{channel_id}"
        )
    
    async def send_channel_message(
        self,
        team_id: str,
        channel_id: str,
        content: str,
        content_type: str = "html",
    ) -> dict[str, Any]:
        """Send a message to a Teams channel.
        
        Args:
            team_id: Team ID
            channel_id: Channel ID
            content: Message content
            content_type: 'text' or 'html'
            
        Returns:
            Created message object
        """
        return await self.client.post(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            json={
                "body": {
                    "contentType": content_type,
                    "content": content,
                }
            }
        )
    
    async def reply_to_channel_message(
        self,
        team_id: str,
        channel_id: str,
        message_id: str,
        content: str,
        content_type: str = "html",
    ) -> dict[str, Any]:
        """Reply to a channel message."""
        return await self.client.post(
            f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies",
            json={
                "body": {
                    "contentType": content_type,
                    "content": content,
                }
            }
        )
    
    async def list_channel_messages(
        self,
        team_id: str,
        channel_id: str,
        count: int = 25,
    ) -> list[dict[str, Any]]:
        """List recent messages in a channel.
        
        Note: Requires ChannelMessage.Read.All permission
        """
        result = await self.client.get(
            f"/teams/{team_id}/channels/{channel_id}/messages",
            params={"$top": count}
        )
        return result.get("value", [])
    
    # =========================================================================
    # CHATS
    # =========================================================================
    
    async def list_my_chats(self, count: int = 50) -> list[dict[str, Any]]:
        """List the user's chats (1:1 and group chats).
        
        Returns:
            List of chat objects
        """
        result = await self.client.get(
            "/me/chats",
            params={
                "$top": count,
                "$expand": "members",
                "$select": "id,topic,chatType,createdDateTime,lastUpdatedDateTime",
            }
        )
        return result.get("value", [])
    
    async def get_chat(self, chat_id: str) -> dict[str, Any]:
        """Get details of a specific chat."""
        return await self.client.get(
            f"/chats/{chat_id}",
            params={"$expand": "members"}
        )
    
    async def send_chat_message(
        self,
        chat_id: str,
        content: str,
        content_type: str = "html",
    ) -> dict[str, Any]:
        """Send a message to a chat.
        
        Args:
            chat_id: Chat ID
            content: Message content
            content_type: 'text' or 'html'
        """
        return await self.client.post(
            f"/chats/{chat_id}/messages",
            json={
                "body": {
                    "contentType": content_type,
                    "content": content,
                }
            }
        )
    
    async def list_chat_messages(
        self,
        chat_id: str,
        count: int = 25,
    ) -> list[dict[str, Any]]:
        """List messages in a chat."""
        result = await self.client.get(
            f"/chats/{chat_id}/messages",
            params={"$top": count}
        )
        return result.get("value", [])
    
    # =========================================================================
    # MEMBERS
    # =========================================================================
    
    async def list_team_members(self, team_id: str) -> list[dict[str, Any]]:
        """List members of a team.
        
        Returns:
            List of member objects with user info and roles
        """
        result = await self.client.get(f"/teams/{team_id}/members")
        return result.get("value", [])
    
    async def list_channel_members(
        self,
        team_id: str,
        channel_id: str
    ) -> list[dict[str, Any]]:
        """List members of a private channel."""
        result = await self.client.get(
            f"/teams/{team_id}/channels/{channel_id}/members"
        )
        return result.get("value", [])
