"""Microsoft Graph API HTTP client.

Provides an async HTTP client for Microsoft Graph API with:
- Automatic token injection and refresh
- Pagination support
- Error handling with typed exceptions
- File download/upload support

Usage:
    >>> client = GraphClient(token_manager)
    >>> profile = await client.get("/me")
    >>> await client.close()
"""

import logging
from typing import Any, AsyncIterator, Optional

import httpx

from ..auth.token_manager import TokenManager

logger = logging.getLogger(__name__)


class AuthenticationRequiredError(Exception):
    """Raised when authentication is needed but no valid token exists."""
    pass


class GraphAPIError(Exception):
    """Raised for Microsoft Graph API errors.
    
    Attributes:
        status_code: HTTP status code
        error_code: Graph API error code (e.g., 'itemNotFound')
        message: Human-readable error message
    """
    
    def __init__(self, status_code: int, error: dict[str, Any]) -> None:
        self.status_code = status_code
        self.error_code = error.get("code", "unknown")
        self.message = error.get("message", "Unknown error")
        super().__init__(f"[{status_code}] {self.error_code}: {self.message}")


class GraphClient:
    """Async HTTP client for Microsoft Graph API.
    
    Handles authentication, request formatting, and response parsing
    for all Microsoft Graph API calls.
    
    Attributes:
        BASE_URL: Microsoft Graph API base URL
        token_manager: Token manager for authentication
    
    Example:
        >>> async with GraphClient(token_manager) as client:
        ...     messages = await client.get("/me/messages")
        ...     await client.post("/me/sendMail", json=mail_data)
    """
    
    BASE_URL = "https://graph.microsoft.com/v1.0"
    
    def __init__(self, token_manager: TokenManager) -> None:
        """Initialize Graph client.
        
        Args:
            token_manager: TokenManager instance for authentication
        """
        self.token_manager = token_manager
        self._client: Optional[httpx.AsyncClient] = None
    
    async def _ensure_client(self) -> httpx.AsyncClient:
        """Get or create authenticated HTTP client.
        
        Raises:
            AuthenticationRequiredError: If no valid token available
        """
        token = self.token_manager.get_graph_token()
        
        if not token:
            raise AuthenticationRequiredError(
                "No valid token available. Please authenticate using m365_authenticate."
            )
        
        if self._client is None:
            self._client = httpx.AsyncClient(
                base_url=self.BASE_URL,
                timeout=httpx.Timeout(30.0, read=60.0),
                follow_redirects=True,
            )
        
        # Always update auth header (token may have been refreshed)
        self._client.headers["Authorization"] = f"Bearer {token}"
        self._client.headers["Content-Type"] = "application/json"
        
        return self._client
    
    async def get(
        self,
        endpoint: str,
        params: Optional[dict[str, Any]] = None,
        headers: Optional[dict[str, str]] = None,
    ) -> dict[str, Any]:
        """Make GET request to Graph API.
        
        Args:
            endpoint: API endpoint (e.g., '/me/messages')
            params: Query parameters
            headers: Additional headers
            
        Returns:
            JSON response as dict
        """
        client = await self._ensure_client()
        response = await client.get(endpoint, params=params, headers=headers)
        return self._handle_response(response)
    
    async def post(
        self,
        endpoint: str,
        json: Optional[dict[str, Any]] = None,
        data: Optional[bytes] = None,
        headers: Optional[dict[str, str]] = None,
    ) -> dict[str, Any]:
        """Make POST request to Graph API.
        
        Args:
            endpoint: API endpoint
            json: JSON body (for most requests)
            data: Raw bytes (for file uploads)
            headers: Additional headers
            
        Returns:
            JSON response as dict (empty dict for 204 responses)
        """
        client = await self._ensure_client()
        response = await client.post(
            endpoint,
            json=json,
            content=data,
            headers=headers,
        )
        return self._handle_response(response)
    
    async def patch(
        self,
        endpoint: str,
        json: Optional[dict[str, Any]] = None,
    ) -> dict[str, Any]:
        """Make PATCH request to Graph API."""
        client = await self._ensure_client()
        response = await client.patch(endpoint, json=json)
        return self._handle_response(response)
    
    async def put(
        self,
        endpoint: str,
        data: Optional[bytes] = None,
        json: Optional[dict[str, Any]] = None,
        headers: Optional[dict[str, str]] = None,
    ) -> dict[str, Any]:
        """Make PUT request to Graph API."""
        client = await self._ensure_client()
        response = await client.put(
            endpoint,
            content=data,
            json=json,
            headers=headers,
        )
        return self._handle_response(response)
    
    async def delete(self, endpoint: str) -> None:
        """Make DELETE request to Graph API."""
        client = await self._ensure_client()
        response = await client.delete(endpoint)
        
        if response.status_code not in (200, 204):
            self._handle_response(response)
    
    async def get_paginated(
        self,
        endpoint: str,
        params: Optional[dict[str, Any]] = None,
        max_pages: int = 10,
    ) -> AsyncIterator[dict[str, Any]]:
        """Iterate through paginated Graph API results.
        
        Yields individual items from the 'value' array across all pages.
        
        Args:
            endpoint: API endpoint
            params: Initial query parameters
            max_pages: Maximum pages to fetch (default 10)
            
        Yields:
            Individual items from response 'value' arrays
        """
        url: Optional[str] = endpoint
        page_count = 0
        
        while url and page_count < max_pages:
            # Only use params on first request
            result = await self.get(
                url,
                params=params if page_count == 0 else None
            )
            
            for item in result.get("value", []):
                yield item
            
            # Get next page URL
            url = result.get("@odata.nextLink")
            page_count += 1
    
    async def download_file(self, endpoint: str) -> bytes:
        """Download file content as bytes.
        
        Args:
            endpoint: File content endpoint
            
        Returns:
            File content as bytes
        """
        client = await self._ensure_client()
        response = await client.get(endpoint)
        
        if response.status_code != 200:
            self._handle_response(response)
        
        return response.content
    
    def _handle_response(
        self,
        response: httpx.Response
    ) -> dict[str, Any]:
        """Handle API response and raise appropriate errors.
        
        Args:
            response: httpx Response object
            
        Returns:
            JSON response as dict
            
        Raises:
            GraphAPIError: For 4xx/5xx responses
        """
        # Success responses
        if response.status_code in (200, 201, 202):
            if response.content:
                return response.json()
            return {}
        
        if response.status_code == 204:
            return {}
        
        # Error responses
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
    
    async def __aenter__(self) -> "GraphClient":
        """Async context manager entry."""
        return self
    
    async def __aexit__(self, *args: Any) -> None:
        """Async context manager exit."""
        await self.close()
