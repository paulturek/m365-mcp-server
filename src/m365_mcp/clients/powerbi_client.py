"""
Power BI REST API client.

Uses a separate token audience from Microsoft Graph:
    https://analysis.windows.net/powerbi/api/.default

MSAL issues a separate access token per audience from the shared cached
account, so no additional environment variables are required beyond the
existing AZURE_CLIENT_ID / AZURE_TENANT_ID.
"""

from __future__ import annotations

import httpx

from m365_mcp.auth import get_access_token

POWERBI_SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]
POWERBI_BASE = "https://api.powerbi.com/v1.0/myorg"


class PowerBIClient:
    """Thin async HTTP client for the Power BI REST API."""

    def __init__(self, user_email: str) -> None:
        self.user_email = user_email

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    async def _token(self) -> str:
        return await get_access_token(self.user_email, scopes=POWERBI_SCOPE)

    def _headers(self, token: str) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    # ------------------------------------------------------------------
    # HTTP verbs
    # ------------------------------------------------------------------

    async def get(self, path: str) -> dict:
        token = await self._token()
        async with httpx.AsyncClient() as client:
            resp = await client.get(
                f"{POWERBI_BASE}{path}",
                headers=self._headers(token),
                timeout=30.0,
            )
            resp.raise_for_status()
            return resp.json()

    async def post(self, path: str, body: dict | None = None) -> dict | None:
        token = await self._token()
        async with httpx.AsyncClient() as client:
            resp = await client.post(
                f"{POWERBI_BASE}{path}",
                headers=self._headers(token),
                json=body or {},
                timeout=30.0,
            )
            # Power BI refresh returns 202 Accepted with no body
            if resp.status_code == 202:
                return {"status": "accepted", "message": "Refresh triggered successfully"}
            resp.raise_for_status()
            return resp.json() if resp.content else None
