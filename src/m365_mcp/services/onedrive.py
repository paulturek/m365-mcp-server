"""OneDrive file storage service.

Provides operations for:
- Listing files and folders
- Searching files
- Uploading and downloading
- Creating folders
- Sharing files
- Moving and copying

Graph API Reference:
    https://learn.microsoft.com/graph/api/resources/onedrive

File Size Limits:
    - Simple upload: < 4 MB
    - Resumable upload: up to 250 GB

"""

from typing import Any, Optional

from ..clients.graph_client import GraphClient


class OneDriveService:
    """OneDrive file management via Microsoft Graph.
    
    Example:
        >>> service = OneDriveService(graph_client)
        >>> files = await service.list_items(folder_path="Documents")
        >>> await service.upload_file("Documents", "test.txt", b"content")
    """
    
    # Default fields to return for items
    DEFAULT_SELECT = [
        "id", "name", "size", "createdDateTime", "lastModifiedDateTime",
        "folder", "file", "webUrl", "parentReference"
    ]
    
    def __init__(self, client: GraphClient) -> None:
        """Initialize OneDrive service."""
        self.client = client
    
    async def list_items(
        self,
        folder_id: Optional[str] = None,
        folder_path: Optional[str] = None,
        count: int = 100,
        order_by: str = "name",
    ) -> list[dict[str, Any]]:
        """List files and folders in OneDrive.
        
        Args:
            folder_id: Folder ID to list (overrides folder_path)
            folder_path: Folder path like 'Documents/Reports'
            count: Max items to return
            order_by: Sort order ('name', 'lastModifiedDateTime desc')
            
        Returns:
            List of file/folder objects
        """
        if folder_path:
            endpoint = f"/me/drive/root:/{folder_path}:/children"
        elif folder_id:
            endpoint = f"/me/drive/items/{folder_id}/children"
        else:
            endpoint = "/me/drive/root/children"
        
        params = {
            "$top": count,
            "$orderby": order_by,
            "$select": ",".join(self.DEFAULT_SELECT),
        }
        
        result = await self.client.get(endpoint, params=params)
        return result.get("value", [])
    
    async def search_files(
        self,
        query: str,
        count: int = 25,
    ) -> list[dict[str, Any]]:
        """Search for files in OneDrive.
        
        Args:
            query: Search query string
            count: Max results to return
            
        Returns:
            List of matching file objects
        """
        params = {
            "$top": count,
            "$select": ",".join(self.DEFAULT_SELECT),
        }
        
        result = await self.client.get(
            f"/me/drive/root/search(q='{query}')",
            params=params
        )
        return result.get("value", [])
    
    async def get_item(self, item_id: str) -> dict[str, Any]:
        """Get metadata for a specific item."""
        return await self.client.get(f"/me/drive/items/{item_id}")
    
    async def get_item_by_path(self, path: str) -> dict[str, Any]:
        """Get item by its path.
        
        Args:
            path: File path like 'Documents/report.docx'
        """
        return await self.client.get(f"/me/drive/root:/{path}")
    
    async def download_file(self, item_id: str) -> bytes:
        """Download file content.
        
        Args:
            item_id: File ID to download
            
        Returns:
            File content as bytes
        """
        return await self.client.download_file(
            f"/me/drive/items/{item_id}/content"
        )
    
    async def download_file_by_path(self, path: str) -> bytes:
        """Download file by path.
        
        Args:
            path: File path like 'Documents/report.docx'
        """
        return await self.client.download_file(
            f"/me/drive/root:/{path}:/content"
        )
    
    async def upload_file(
        self,
        folder_path: str,
        filename: str,
        content: bytes,
        conflict_behavior: str = "rename",
    ) -> dict[str, Any]:
        """Upload a file to OneDrive (< 4MB).
        
        For larger files, use upload_large_file (not yet implemented).
        
        Args:
            folder_path: Destination folder path
            filename: Name for the file
            content: File content as bytes
            conflict_behavior: 'rename', 'replace', or 'fail'
            
        Returns:
            Created file object
            
        Raises:
            ValueError: If file > 4MB
        """
        if len(content) > 4 * 1024 * 1024:
            raise ValueError(
                "File too large for simple upload. "
                "Files > 4MB require resumable upload."
            )
        
        path = f"{folder_path}/{filename}" if folder_path else filename
        endpoint = f"/me/drive/root:/{path}:/content"
        
        return await self.client.put(
            endpoint,
            data=content,
            headers={
                "Content-Type": "application/octet-stream",
                "@microsoft.graph.conflictBehavior": conflict_behavior,
            },
        )
    
    async def create_folder(
        self,
        parent_path: str,
        folder_name: str,
    ) -> dict[str, Any]:
        """Create a new folder.
        
        Args:
            parent_path: Parent folder path (empty for root)
            folder_name: Name for new folder
            
        Returns:
            Created folder object
        """
        if parent_path:
            endpoint = f"/me/drive/root:/{parent_path}:/children"
        else:
            endpoint = "/me/drive/root/children"
        
        return await self.client.post(
            endpoint,
            json={
                "name": folder_name,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "rename",
            }
        )
    
    async def delete_item(self, item_id: str) -> None:
        """Delete a file or folder."""
        await self.client.delete(f"/me/drive/items/{item_id}")
    
    async def move_item(
        self,
        item_id: str,
        new_parent_id: str,
        new_name: Optional[str] = None,
    ) -> dict[str, Any]:
        """Move an item to a different folder.
        
        Args:
            item_id: Item to move
            new_parent_id: Destination folder ID
            new_name: Optional new name
            
        Returns:
            Updated item object
        """
        update: dict[str, Any] = {
            "parentReference": {"id": new_parent_id}
        }
        if new_name:
            update["name"] = new_name
        
        return await self.client.patch(
            f"/me/drive/items/{item_id}",
            json=update
        )
    
    async def copy_item(
        self,
        item_id: str,
        dest_parent_id: str,
        new_name: Optional[str] = None,
    ) -> dict[str, Any]:
        """Copy an item to a different folder.
        
        Note: Copy is async. Returns a monitor URL to check status.
        
        Args:
            item_id: Item to copy
            dest_parent_id: Destination folder ID
            new_name: Optional new name for copy
        """
        body: dict[str, Any] = {
            "parentReference": {"id": dest_parent_id}
        }
        if new_name:
            body["name"] = new_name
        
        return await self.client.post(
            f"/me/drive/items/{item_id}/copy",
            json=body
        )
    
    async def rename_item(
        self,
        item_id: str,
        new_name: str,
    ) -> dict[str, Any]:
        """Rename a file or folder."""
        return await self.client.patch(
            f"/me/drive/items/{item_id}",
            json={"name": new_name}
        )
    
    async def create_sharing_link(
        self,
        item_id: str,
        link_type: str = "view",
        scope: str = "anonymous",
        expiration: Optional[str] = None,
    ) -> str:
        """Create a sharing link for an item.
        
        Args:
            item_id: Item to share
            link_type: 'view', 'edit', or 'embed'
            scope: 'anonymous' or 'organization'
            expiration: Optional expiration datetime ISO string
            
        Returns:
            Sharing URL string
        """
        body: dict[str, Any] = {
            "type": link_type,
            "scope": scope,
        }
        if expiration:
            body["expirationDateTime"] = expiration
        
        result = await self.client.post(
            f"/me/drive/items/{item_id}/createLink",
            json=body
        )
        return result.get("link", {}).get("webUrl", "")
    
    async def get_drive_info(self) -> dict[str, Any]:
        """Get information about the user's OneDrive.
        
        Returns:
            Drive object with quota, owner info, etc.
        """
        return await self.client.get("/me/drive")
