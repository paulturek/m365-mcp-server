"""Excel workbook operations service.

Provides operations for:
- Worksheets: list, read ranges, write ranges
- Tables: list, add rows, read data
- Named ranges and formulas

Graph API Reference:
    https://learn.microsoft.com/graph/api/resources/excel

Note:
    Excel API works with workbooks stored in OneDrive or SharePoint.
    The workbook must be identified by its file ID (drive item ID).
    
    Excel API has some limitations:
    - Some operations don't support application permissions
    - Large workbooks may have performance issues
    - Concurrent access requires session management

"""

from typing import Any, Optional

from ..clients.graph_client import GraphClient


class ExcelService:
    """Excel workbook operations via Microsoft Graph.
    
    Example:
        >>> service = ExcelService(graph_client)
        >>> sheets = await service.list_worksheets(file_id)
        >>> data = await service.get_range(file_id, "Sheet1", "A1:D10")
    """
    
    def __init__(self, client: GraphClient) -> None:
        """Initialize Excel service."""
        self.client = client
    
    def _workbook_endpoint(
        self,
        file_id: str,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
    ) -> str:
        """Build workbook API endpoint.
        
        Args:
            file_id: OneDrive/SharePoint item ID
            site_id: Optional SharePoint site ID
            drive_id: Optional drive ID (for SharePoint)
        """
        if site_id and drive_id:
            return f"/sites/{site_id}/drives/{drive_id}/items/{file_id}/workbook"
        elif site_id:
            return f"/sites/{site_id}/drive/items/{file_id}/workbook"
        else:
            return f"/me/drive/items/{file_id}/workbook"
    
    # =========================================================================
    # WORKSHEETS
    # =========================================================================
    
    async def list_worksheets(
        self,
        file_id: str,
        site_id: Optional[str] = None,
        drive_id: Optional[str] = None,
    ) -> list[dict[str, Any]]:
        """List all worksheets in a workbook.
        
        Returns:
            List of worksheet objects with id, name, position, visibility
        """
        endpoint = f"{self._workbook_endpoint(file_id, site_id, drive_id)}/worksheets"
        result = await self.client.get(endpoint)
        return result.get("value", [])
    
    async def get_worksheet(
        self,
        file_id: str,
        sheet_name: str,
        site_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Get a specific worksheet."""
        endpoint = (
            f"{self._workbook_endpoint(file_id, site_id)}"
            f"/worksheets('{sheet_name}')"
        )
        return await self.client.get(endpoint)
    
    async def add_worksheet(
        self,
        file_id: str,
        name: str,
        site_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Add a new worksheet to a workbook."""
        endpoint = f"{self._workbook_endpoint(file_id, site_id)}/worksheets/add"
        return await self.client.post(endpoint, json={"name": name})
    
    # =========================================================================
    # RANGES
    # =========================================================================
    
    async def get_range(
        self,
        file_id: str,
        sheet_name: str,
        range_address: str,
        site_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Read values from a range.
        
        Args:
            file_id: Workbook file ID
            sheet_name: Name of the worksheet
            range_address: A1 notation (e.g., 'A1:D10', 'A:A', '1:1')
            site_id: Optional SharePoint site ID
            
        Returns:
            Range object with:
            - values: 2D array of cell values
            - formulas: 2D array of formulas
            - text: 2D array of formatted text
            - address: Resolved range address
        """
        endpoint = (
            f"{self._workbook_endpoint(file_id, site_id)}"
            f"/worksheets('{sheet_name}')/range(address='{range_address}')"
        )
        return await self.client.get(endpoint)
    
    async def update_range(
        self,
        file_id: str,
        sheet_name: str,
        range_address: str,
        values: list[list[Any]],
        site_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Write values to a range.
        
        Args:
            file_id: Workbook file ID
            sheet_name: Worksheet name
            range_address: Target range in A1 notation
            values: 2D array of values to write
            site_id: Optional SharePoint site ID
            
        Returns:
            Updated range object
            
        Note:
            The values array dimensions must match the range dimensions.
        """
        endpoint = (
            f"{self._workbook_endpoint(file_id, site_id)}"
            f"/worksheets('{sheet_name}')/range(address='{range_address}')"
        )
        return await self.client.patch(endpoint, json={"values": values})
    
    async def get_used_range(
        self,
        file_id: str,
        sheet_name: str,
        site_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Get the used range of a worksheet (smallest range containing all data)."""
        endpoint = (
            f"{self._workbook_endpoint(file_id, site_id)}"
            f"/worksheets('{sheet_name}')/usedRange"
        )
        return await self.client.get(endpoint)
    
    async def clear_range(
        self,
        file_id: str,
        sheet_name: str,
        range_address: str,
        apply_to: str = "all",
        site_id: Optional[str] = None,
    ) -> None:
        """Clear a range.
        
        Args:
            apply_to: What to clear - 'all', 'formats', 'contents'
        """
        endpoint = (
            f"{self._workbook_endpoint(file_id, site_id)}"
            f"/worksheets('{sheet_name}')/range(address='{range_address}')/clear"
        )
        await self.client.post(endpoint, json={"applyTo": apply_to})
    
    # =========================================================================
    # TABLES
    # =========================================================================
    
    async def list_tables(
        self,
        file_id: str,
        site_id: Optional[str] = None,
    ) -> list[dict[str, Any]]:
        """List all tables in a workbook.
        
        Returns:
            List of table objects with id, name, showHeaders, etc.
        """
        endpoint = f"{self._workbook_endpoint(file_id, site_id)}/tables"
        result = await self.client.get(endpoint)
        return result.get("value", [])
    
    async def get_table(
        self,
        file_id: str,
        table_name: str,
        site_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Get a specific table."""
        endpoint = (
            f"{self._workbook_endpoint(file_id, site_id)}"
            f"/tables('{table_name}')"
        )
        return await self.client.get(endpoint)
    
    async def get_table_range(
        self,
        file_id: str,
        table_name: str,
        site_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Get the data range of a table (including headers)."""
        endpoint = (
            f"{self._workbook_endpoint(file_id, site_id)}"
            f"/tables('{table_name}')/range"
        )
        return await self.client.get(endpoint)
    
    async def get_table_data_range(
        self,
        file_id: str,
        table_name: str,
        site_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Get table data range (excluding headers)."""
        endpoint = (
            f"{self._workbook_endpoint(file_id, site_id)}"
            f"/tables('{table_name}')/dataBodyRange"
        )
        return await self.client.get(endpoint)
    
    async def add_table_rows(
        self,
        file_id: str,
        table_name: str,
        values: list[list[Any]],
        site_id: Optional[str] = None,
    ) -> dict[str, Any]:
        """Add rows to an Excel table.
        
        Args:
            file_id: Workbook file ID
            table_name: Table name or ID
            values: 2D array of row values
            site_id: Optional SharePoint site ID
        """
        endpoint = (
            f"{self._workbook_endpoint(file_id, site_id)}"
            f"/tables('{table_name}')/rows"
        )
        return await self.client.post(endpoint, json={"values": values})
    
    async def get_table_columns(
        self,
        file_id: str,
        table_name: str,
        site_id: Optional[str] = None,
    ) -> list[dict[str, Any]]:
        """Get table column definitions."""
        endpoint = (
            f"{self._workbook_endpoint(file_id, site_id)}"
            f"/tables('{table_name}')/columns"
        )
        result = await self.client.get(endpoint)
        return result.get("value", [])
    
    # =========================================================================
    # SESSIONS (for batch operations)
    # =========================================================================
    
    async def create_session(
        self,
        file_id: str,
        persist_changes: bool = True,
        site_id: Optional[str] = None,
    ) -> str:
        """Create a workbook session for batch operations.
        
        Sessions allow multiple operations to be grouped together
        and can improve performance for many sequential operations.
        
        Args:
            file_id: Workbook file ID
            persist_changes: Whether changes are saved automatically
            site_id: Optional SharePoint site ID
            
        Returns:
            Session ID to include in subsequent requests
        """
        endpoint = f"{self._workbook_endpoint(file_id, site_id)}/createSession"
        result = await self.client.post(
            endpoint,
            json={"persistChanges": persist_changes}
        )
        return result.get("id", "")
    
    async def close_session(
        self,
        file_id: str,
        session_id: str,
        site_id: Optional[str] = None,
    ) -> None:
        """Close a workbook session."""
        endpoint = f"{self._workbook_endpoint(file_id, site_id)}/closeSession"
        await self.client.post(
            endpoint,
            headers={"workbook-session-id": session_id}
        )
