"""Power BI service.

Provides operations for:
- Workspaces: list, get info
- Reports: list, get pages
- Datasets: list, refresh, get history
- Dashboards: list, get tiles

API Reference:
    https://learn.microsoft.com/rest/api/power-bi/

Note:
    Power BI uses a separate API endpoint (api.powerbi.com) and
    different OAuth scopes from Microsoft Graph. Users may need
    to authenticate separately for Power BI access.

Scopes:
    https://analysis.windows.net/powerbi/api/.default

"""

from typing import Any, Optional

from ..clients.powerbi_client import PowerBIClient


class PowerBIService:
    """Power BI operations via Power BI REST API.
    
    Example:
        >>> service = PowerBIService(powerbi_client)
        >>> workspaces = await service.list_workspaces()
        >>> reports = await service.list_reports(workspace_id)
    """
    
    def __init__(self, client: PowerBIClient) -> None:
        """Initialize Power BI service."""
        self.client = client
    
    # =========================================================================
    # WORKSPACES (GROUPS)
    # =========================================================================
    
    async def list_workspaces(self) -> list[dict[str, Any]]:
        """List Power BI workspaces the user has access to.
        
        Returns:
            List of workspace objects with id, name, isReadOnly, etc.
        """
        result = await self.client.get("/groups")
        return result.get("value", [])
    
    async def get_workspace(self, workspace_id: str) -> dict[str, Any]:
        """Get details of a specific workspace."""
        return await self.client.get(f"/groups/{workspace_id}")
    
    # =========================================================================
    # REPORTS
    # =========================================================================
    
    async def list_reports(
        self,
        workspace_id: Optional[str] = None
    ) -> list[dict[str, Any]]:
        """List reports in a workspace or 'My Workspace'.
        
        Args:
            workspace_id: Workspace ID. If None, lists from 'My Workspace'.
            
        Returns:
            List of report objects
        """
        if workspace_id:
            endpoint = f"/groups/{workspace_id}/reports"
        else:
            endpoint = "/reports"
        
        result = await self.client.get(endpoint)
        return result.get("value", [])
    
    async def get_report(
        self,
        report_id: str,
        workspace_id: Optional[str] = None
    ) -> dict[str, Any]:
        """Get details of a specific report."""
        if workspace_id:
            endpoint = f"/groups/{workspace_id}/reports/{report_id}"
        else:
            endpoint = f"/reports/{report_id}"
        
        return await self.client.get(endpoint)
    
    async def get_report_pages(
        self,
        report_id: str,
        workspace_id: Optional[str] = None
    ) -> list[dict[str, Any]]:
        """List pages in a report.
        
        Returns:
            List of page objects with name, displayName, order
        """
        if workspace_id:
            endpoint = f"/groups/{workspace_id}/reports/{report_id}/pages"
        else:
            endpoint = f"/reports/{report_id}/pages"
        
        result = await self.client.get(endpoint)
        return result.get("value", [])
    
    # =========================================================================
    # DATASETS
    # =========================================================================
    
    async def list_datasets(
        self,
        workspace_id: Optional[str] = None
    ) -> list[dict[str, Any]]:
        """List datasets in a workspace or 'My Workspace'.
        
        Returns:
            List of dataset objects
        """
        if workspace_id:
            endpoint = f"/groups/{workspace_id}/datasets"
        else:
            endpoint = "/datasets"
        
        result = await self.client.get(endpoint)
        return result.get("value", [])
    
    async def get_dataset(
        self,
        dataset_id: str,
        workspace_id: Optional[str] = None
    ) -> dict[str, Any]:
        """Get details of a specific dataset."""
        if workspace_id:
            endpoint = f"/groups/{workspace_id}/datasets/{dataset_id}"
        else:
            endpoint = f"/datasets/{dataset_id}"
        
        return await self.client.get(endpoint)
    
    async def refresh_dataset(
        self,
        dataset_id: str,
        workspace_id: Optional[str] = None
    ) -> None:
        """Trigger a dataset refresh.
        
        This starts an async refresh operation. Use get_refresh_history
        to check the status.
        """
        if workspace_id:
            endpoint = f"/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
        else:
            endpoint = f"/datasets/{dataset_id}/refreshes"
        
        await self.client.post(endpoint)
    
    async def get_refresh_history(
        self,
        dataset_id: str,
        workspace_id: Optional[str] = None,
        count: int = 10,
    ) -> list[dict[str, Any]]:
        """Get dataset refresh history.
        
        Returns:
            List of refresh objects with status, startTime, endTime
        """
        if workspace_id:
            endpoint = f"/groups/{workspace_id}/datasets/{dataset_id}/refreshes"
        else:
            endpoint = f"/datasets/{dataset_id}/refreshes"
        
        result = await self.client.get(endpoint, params={"$top": count})
        return result.get("value", [])
    
    async def get_dataset_tables(
        self,
        dataset_id: str,
        workspace_id: Optional[str] = None
    ) -> list[dict[str, Any]]:
        """Get tables in a dataset."""
        if workspace_id:
            endpoint = f"/groups/{workspace_id}/datasets/{dataset_id}/tables"
        else:
            endpoint = f"/datasets/{dataset_id}/tables"
        
        result = await self.client.get(endpoint)
        return result.get("value", [])
    
    # =========================================================================
    # DASHBOARDS
    # =========================================================================
    
    async def list_dashboards(
        self,
        workspace_id: Optional[str] = None
    ) -> list[dict[str, Any]]:
        """List dashboards in a workspace.
        
        Returns:
            List of dashboard objects
        """
        if workspace_id:
            endpoint = f"/groups/{workspace_id}/dashboards"
        else:
            endpoint = "/dashboards"
        
        result = await self.client.get(endpoint)
        return result.get("value", [])
    
    async def get_dashboard_tiles(
        self,
        dashboard_id: str,
        workspace_id: Optional[str] = None
    ) -> list[dict[str, Any]]:
        """List tiles in a dashboard.
        
        Returns:
            List of tile objects
        """
        if workspace_id:
            endpoint = f"/groups/{workspace_id}/dashboards/{dashboard_id}/tiles"
        else:
            endpoint = f"/dashboards/{dashboard_id}/tiles"
        
        result = await self.client.get(endpoint)
        return result.get("value", [])
