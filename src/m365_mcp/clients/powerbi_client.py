"""Power BI REST API HTTP client.

Provides an async HTTP client for the Power BI REST API with:
- Direct access token injection (per-user, per-request)
- Error handling consistent with GraphClient
- Report, dataset, and refresh operations

Usage:
    >>> async with PowerBIClient("eyJ0...") as client:
    ...     reports = await client.get("/reports")
"""

import logging
from typing import Any, Optional

import httpx

from .graph_client import AuthenticationRequiredError, GraphAPIError

logger = logging.getLogger(__name__)


class PowerBIClient:
    """Async HTTP client for Power BI REST API.

    Accepts a raw access_token string.  The caller (tool handler) is
    responsible for resolving the token via the token store for the
    appropriate user_id before constructing this client.

    Example:
        >>> async with PowerBIClient(access_token) as client:
        ...     reports = await client.get("/reports")
    """

    BASE_URL = "https://api.powerbi.com/v1.0/myorg"

    def __init__(self, access_token: str) -> None:
        """Initialize Power BI client with a pre-resolved access token.

        Args:
            access_token: Valid Power BI / Microsoft Graph access token

        Raises:
            AuthenticationRequiredError: If access_token is empty/None
        """
        if not access_token:
            raise AuthenticationRequiredError(
                "No access token provided. User must authenticate "
                "via /auth/login?user_id=<email>"
            )
        self._token = access_token
        self._client: Optional[httpx.AsyncClient] = None

    async def _ensure_client(self) -> httpx.AsyncClient:
        """Get or create the HTTP client with auth headers."""
        if self._client is None:
            self._client = httpx.AsyncClient(
                base_url=self.BASE_URL,
                headers={
                    "Authorization": f"Bearer {self._token}",
                    "Content-Type": "application/json",
                },
                timeout=httpx.Timeout(30.0, read=60.0),
            )
        return self._client

    async def get(
        self,
        endpoint: str,
        params: Optional[dict[str, Any]] = None,
    ) -> dict[str, Any]:
        """Make GET request to Power BI API.

        Args:
            endpoint: API endpoint (e.g., '/reports')
            params: Query parameters

        Returns:
            JSON response as dict
        """
        client = await self._ensure_client()
        response = await client.get(endpoint, params=params)
        return self._handle_response(response)

    async def post(
        self,
        endpoint: str,
        json: Optional[dict[str, Any]] = None,
    ) -> dict[str, Any]:
        """Make POST request to Power BI API.

        Args:
            endpoint: API endpoint
            json: JSON body

        Returns:
            JSON response as dict
        """
        client = await self._ensure_client()
        response = await client.post(endpoint, json=json)
        return self._handle_response(response)

    def _handle_response(
        self,
        response: httpx.Response,
    ) -> dict[str, Any]:
        """Handle API response and raise appropriate errors.

        Args:
            response: httpx Response object

        Returns:
            JSON response as dict

        Raises:
            GraphAPIError: For 4xx/5xx responses
        """
        if response.status_code in (200, 201, 202):
            if response.content:
                return response.json()
            return {}
        if response.status_code == 204:
            return {}

        try:
            error_data = response.json()
            error = error_data.get("error", {})
        except Exception:
            error = {"code": "unknown", "message": response.text}
        raise GraphAPIError(response.status_code, error)

    async def close(self) -> None:
        """Close the HTTP client."""
        if self._client:
            await self._client.aclose()
            self._client = None

    async def __aenter__(self) -> "PowerBIClient":
        """Async context manager entry."""
        return self

    async def __aexit__(self, *args: Any) -> None:
        """Async context manager exit."""
        await self.close()
