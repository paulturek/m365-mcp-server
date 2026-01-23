"""Word and PowerPoint document service.

Provides operations for:
- Download documents
- Upload documents
- Convert to PDF
- Get preview URLs
- Get thumbnails

Graph API Reference:
    https://learn.microsoft.com/graph/api/resources/driveitem

Note:
    Unlike Excel, Word and PowerPoint don't have rich editing APIs.
    Operations are primarily file-level (download, upload, convert).
    For editing, consider using the Microsoft 365 desktop apps or
    the Office JavaScript API in web contexts.

"""

from typing import Any, Optional

from ..clients.graph_client import GraphClient


class OfficeDocsService:
    """Word and PowerPoint operations via Microsoft Graph.
    
    Example:
        >>> service = OfficeDocsService(graph_client)
        >>> pdf_bytes = await service.download_as_pdf(item_id)
        >>> preview = await service.get_preview(item_id)
    """
    
    def __init__(self, client: GraphClient) -> None:
        """Initialize Office Docs service."""
        self.client = client
    
    def _item_endpoint(
        self,
        item_id: str,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
    ) -> str:
        """Build drive item endpoint."""
        if site_id and drive_id:
            return f"/sites/{site_id}/drives/{drive_id}/items/{item_id}"
        elif site_id:
            return f"/sites/{site_id}/drive/items/{item_id}"
        else:
            return f"/me/drive/items/{item_id}"
    
    async def download(
        self,
        item_id: str,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
    ) -> bytes:
        """Download a Word or PowerPoint document.
        
        Args:
            item_id: File ID
            site_id: Optional SharePoint site ID
            drive_id: Optional drive ID
            
        Returns:
            Document content as bytes
        """
        endpoint = f"{self._item_endpoint(item_id, site_id, drive_id)}/content"
        return await self.client.download_file(endpoint)
    
    async def download_as_pdf(
        self,
        item_id: str,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
    ) -> bytes:
        """Download document converted to PDF.
        
        Works for:
        - Word documents (.docx, .doc)
        - PowerPoint presentations (.pptx, .ppt)
        - Excel workbooks (.xlsx, .xls)
        
        Args:
            item_id: File ID
            site_id: Optional SharePoint site ID
            drive_id: Optional drive ID
            
        Returns:
            PDF content as bytes
        """
        endpoint = f"{self._item_endpoint(item_id, site_id, drive_id)}/content?format=pdf"
        return await self.client.download_file(endpoint)
    
    async def get_preview(
        self,
        item_id: str,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Get embeddable preview URLs for a document.
        
        Returns:
            Dict with getUrl (for viewing) and postUrl (for embedding)
        """
        endpoint = f"{self._item_endpoint(item_id, site_id, drive_id)}/preview"
        return await self.client.post(endpoint)
    
    async def upload(
        self,
        folder_path: str,
        filename: str,
        content: bytes,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Upload a Word or PowerPoint document.
        
        Args:
            folder_path: Destination folder path
            filename: File name (include extension)
            content: File content as bytes (< 4MB)
            site_id: Optional SharePoint site ID
            drive_id: Optional drive ID
            
        Returns:
            Created file object
        """
        if len(content) > 4 * 1024 * 1024:
            raise ValueError(
                "File too large. Use resumable upload for files > 4MB."
            )
        
        path = f"{folder_path}/{filename}" if folder_path else filename
        
        if site_id and drive_id:
            endpoint = f"/sites/{site_id}/drives/{drive_id}/root:/{path}:/content"
        elif site_id:
            endpoint = f"/sites/{site_id}/drive/root:/{path}:/content"
        else:
            endpoint = f"/me/drive/root:/{path}:/content"
        
        return await self.client.put(
            endpoint,
            data=content,
            headers={"Content-Type": "application/octet-stream"},
        )
    
    async def get_thumbnails(
        self,
        item_id: str,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
    ) -> list[dict[str, Any]]:
        """Get thumbnail images for a document.
        
        Returns:
            List of thumbnail sets, each containing:
            - small: 96x96 thumbnail
            - medium: 176x176 thumbnail  
            - large: 800x800 thumbnail
        """
        endpoint = f"{self._item_endpoint(item_id, site_id, drive_id)}/thumbnails"
        result = await self.client.get(endpoint)
        return result.get("value", [])
    
    async def get_item_info(
        self,
        item_id: str,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Get file metadata.
        
        Returns:
            File object with name, size, webUrl, etc.
        """
        return await self.client.get(
            self._item_endpoint(item_id, site_id, drive_id)
        )
