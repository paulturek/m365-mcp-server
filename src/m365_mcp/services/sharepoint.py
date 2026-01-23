"""SharePoint sites and lists service.

Provides operations for:
- Sites: search, get by ID or path
- Document libraries: list, browse, upload
- Lists: list, get items, create/update/delete items

Graph API Reference:
    https://learn.microsoft.com/graph/api/resources/sharepoint

Site ID Format:
    {hostname},{site-collection-id},{web-id}
    Example: contoso.sharepoint.com,abc123,def456

"""

from typing import Any, Optional

from ..clients.graph_client import GraphClient


class SharePointService:
    """SharePoint management via Microsoft Graph.
    
    Example:
        >>> service = SharePointService(graph_client)
        >>> sites = await service.search_sites("Marketing")
        >>> lists = await service.list_lists(site_id)
    """
    
    def __init__(self, client: GraphClient) -> None:
        """Initialize SharePoint service."""
        self.client = client
    
    # =========================================================================
    # SITES
    # =========================================================================
    
    async def search_sites(self, query: str) -> list[dict[str, Any]]:
        """Search for SharePoint sites.
        
        Args:
            query: Search query string
            
        Returns:
            List of matching site objects
        """
        result = await self.client.get(
            f"/sites?search={query}",
            params={"$select": "id,name,displayName,webUrl,description"}
        )
        return result.get("value", [])
    
    async def get_site(self, site_id: str) -> dict[str, Any]:
        """Get a specific site by ID."""
        return await self.client.get(f"/sites/{site_id}")
    
    async def get_site_by_path(
        self,
        hostname: str,
        site_path: str
    ) -> dict[str, Any]:
        """Get site by hostname and path.
        
        Args:
            hostname: SharePoint hostname (e.g., 'contoso.sharepoint.com')
            site_path: Site path (e.g., '/sites/Marketing')
            
        Returns:
            Site object
        """
        return await self.client.get(f"/sites/{hostname}:{site_path}")
    
    async def get_root_site(self, hostname: str) -> dict[str, Any]:
        """Get the root site for a hostname."""
        return await self.client.get(f"/sites/{hostname}")
    
    # =========================================================================
    # DOCUMENT LIBRARIES (DRIVES)
    # =========================================================================
    
    async def list_drives(self, site_id: str) -> list[dict[str, Any]]:
        """List document libraries in a site.
        
        Returns:
            List of drive objects
        """
        result = await self.client.get(
            f"/sites/{site_id}/drives",
            params={"$select": "id,name,webUrl,driveType"}
        )
        return result.get("value", [])
    
    async def get_drive(self, site_id: str, drive_id: str) -> dict[str, Any]:
        """Get a specific document library."""
        return await self.client.get(f"/sites/{site_id}/drives/{drive_id}")
    
    async def list_drive_items(
        self,
        site_id: str,
        drive_id: str,
        folder_id: Optional[str] = None,
        count: int = 100,
    ) -> list[dict[str, Any]]:
        """List items in a document library.
        
        Args:
            site_id: Site ID
            drive_id: Document library ID
            folder_id: Optional folder ID (root if not specified)
            count: Max items to return
        """
        if folder_id:
            endpoint = f"/sites/{site_id}/drives/{drive_id}/items/{folder_id}/children"
        else:
            endpoint = f"/sites/{site_id}/drives/{drive_id}/root/children"
        
        params = {
            "$top": count,
            "$select": "id,name,size,createdDateTime,lastModifiedDateTime,folder,file,webUrl",
        }
        
        result = await self.client.get(endpoint, params=params)
        return result.get("value", [])
    
    async def upload_to_site(
        self,
        site_id: str,
        drive_id: str,
        folder_path: str,
        filename: str,
        content: bytes,
    ) -> dict[str, Any]:
        """Upload file to SharePoint document library.
        
        Args:
            site_id: Site ID
            drive_id: Document library ID
            folder_path: Path within the library
            filename: Name for the file
            content: File content (< 4MB)
        """
        path = f"{folder_path}/{filename}" if folder_path else filename
        endpoint = f"/sites/{site_id}/drives/{drive_id}/root:/{path}:/content"
        
        return await self.client.put(
            endpoint,
            data=content,
            headers={"Content-Type": "application/octet-stream"},
        )
    
    async def download_from_site(
        self,
        site_id: str,
        drive_id: str,
        item_id: str,
    ) -> bytes:
        """Download file from SharePoint."""
        return await self.client.download_file(
            f"/sites/{site_id}/drives/{drive_id}/items/{item_id}/content"
        )
    
    # =========================================================================
    # LISTS
    # =========================================================================
    
    async def list_lists(self, site_id: str) -> list[dict[str, Any]]:
        """List all SharePoint lists in a site.
        
        Returns:
            List of list objects (not document libraries)
        """
        result = await self.client.get(
            f"/sites/{site_id}/lists",
            params={"$select": "id,name,displayName,webUrl,list"}
        )
        return result.get("value", [])
    
    async def get_list(self, site_id: str, list_id: str) -> dict[str, Any]:
        """Get a specific list."""
        return await self.client.get(f"/sites/{site_id}/lists/{list_id}")
    
    async def get_list_items(
        self,
        site_id: str,
        list_id: str,
        expand_fields: bool = True,
        count: int = 100,
        filter_query: Optional[str] = None,
    ) -> list[dict[str, Any]]:
        """Get items from a SharePoint list.
        
        Args:
            site_id: Site ID
            list_id: List ID
            expand_fields: Include field values
            count: Max items
            filter_query: OData filter
        """
        params: dict[str, Any] = {"$top": count}
        
        if expand_fields:
            params["$expand"] = "fields"
        if filter_query:
            params["$filter"] = filter_query
        
        result = await self.client.get(
            f"/sites/{site_id}/lists/{list_id}/items",
            params=params
        )
        return result.get("value", [])
    
    async def get_list_item(
        self,
        site_id: str,
        list_id: str,
        item_id: str,
    ) -> dict[str, Any]:
        """Get a specific list item."""
        return await self.client.get(
            f"/sites/{site_id}/lists/{list_id}/items/{item_id}",
            params={"$expand": "fields"}
        )
    
    async def create_list_item(
        self,
        site_id: str,
        list_id: str,
        fields: dict[str, Any],
    ) -> dict[str, Any]:
        """Create a new item in a SharePoint list.
        
        Args:
            site_id: Site ID
            list_id: List ID
            fields: Dictionary of field values
        """
        return await self.client.post(
            f"/sites/{site_id}/lists/{list_id}/items",
            json={"fields": fields}
        )
    
    async def update_list_item(
        self,
        site_id: str,
        list_id: str,
        item_id: str,
        fields: dict[str, Any],
    ) -> dict[str, Any]:
        """Update an existing list item.
        
        Args:
            site_id: Site ID
            list_id: List ID
            item_id: Item ID to update
            fields: Fields to update
        """
        return await self.client.patch(
            f"/sites/{site_id}/lists/{list_id}/items/{item_id}/fields",
            json=fields
        )
    
    async def delete_list_item(
        self,
        site_id: str,
        list_id: str,
        item_id: str,
    ) -> None:
        """Delete a list item."""
        await self.client.delete(
            f"/sites/{site_id}/lists/{list_id}/items/{item_id}"
        )
    
    async def get_list_columns(
        self,
        site_id: str,
        list_id: str,
    ) -> list[dict[str, Any]]:
        """Get column definitions for a list."""
        result = await self.client.get(
            f"/sites/{site_id}/lists/{list_id}/columns"
        )
        return result.get("value", [])
