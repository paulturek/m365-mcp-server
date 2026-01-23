"""Power BI REST API HTTP client.

Power BI uses a separate API endpoint and authentication scope
from Microsoft Graph. This client handles Power BI-specific requests.

API Base: https://api.powerbi.com/v1.0/myorg
Scope: https://analysis.windows.net/powerbi/api/.default

"""

import logging
from typing import Any, Optional

import httpx

from ..auth.token_manager import TokenManager

logger = logging.getLogger(__name__)


class PowerBIAuthError(Exception):
    """Raised when Power BI authentication fails."""
    pass


class PowerBIClient:
    """Async HTTP client for Power BI REST API.
    
    Note: Power BI requires separate authentication from Microsoft Graph.
    Users may need to authenticate specifically for Power BI if they
    only authenticated for Graph initially.
    
    Attributes:
        BASE_URL: Power BI API base URL
        token_manager: Token manager for authentication
    """
    
    BASE_URL = "https://api.powerbi.com/v1.0/myorg"
    
    def __init__(self, token_manager: TokenManager) -> None:
        """Initialize Power BI client.
        
        Args:
            token_manager: TokenManager instance for authentication
        """
        self.token_manager = token_manager
        self._client: Optional[httpx.AsyncClient] = None
    
    async def _ensure_client(self) -> httpx.AsyncClient:
        """Get authenticated client for Power BI API.
        
        Raises:
            PowerBIAuthError: If no valid Power BI token available
        """
        token = self.token_manager.get_powerbi_token()
        
        if not token:
            raise PowerBIAuthError(
                "No Power BI token available. "
                "You may need to authenticate with Power BI scopes."
            )
        
        if self._client is None:
            self._client = httpx.AsyncClient(
                base_url=self.BASE_URL,
                timeout=30.0,
            )
        
        self._client.headers["Authorization"] = f"Bearer {token}"
        return self._client
    
    async def get(
        self,
        endpoint: str,
        params: Optional[dict[str, Any]] = None
    ) -> dict[str, Any]:
        """Make GET request to Power BI API."""
        client = await self._ensure_client()
        response = await client.get(endpoint, params=params)
        response.raise_for_status()
        return response.json() if response.content else {}
    
    async def post(
        self,
        endpoint: str,
        json: Optional[dict[str, Any]] = None
    ) -> dict[str, Any]:
        """Make POST request to Power BI API."""
        client = await self._ensure_client()
        response = await client.post(endpoint, json=json)
        response.raise_for_status()
        return response.json() if response.content else {}
    
    async def close(self) -> None:
        """Close the HTTP client."""
        if self._client:
            await self._client.aclose()
            self._client = None
    
    async def __aenter__(self) -> "PowerBIClient":
        return self
    
    async def __aexit__(self, *args: Any) -> None:
        await self.close()
