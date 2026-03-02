"""
Power BI REST API client.

Uses a separate token audience from Microsoft Graph:
    https://analysis.windows.net/powerbi/api/.default

MSAL issues a separate token per audience from the shared cached account,
so no additional environment variables are required beyond what is already
set for the Graph client (AZURE_CLIENT_ID, AZURE_TENANT_ID, DATABASE_URL).
"""

from __future__ import annotations

import httpx

POWERBI_SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]
POWERBI_BASE = "https://api.powerbi.com/v1.0/myorg"


class PowerBIClient:
    """Thin async HTTP client for the Power BI REST API."""

    def __init__(self, user_email: str, token_manager) -> None:
        """
        Args:
            user_email:    The authenticated user's email address.
            token_manager: Any object that exposes
                           ``await get_token(user_email, scopes)`` and returns
                           a plain access-token string.  Pass the module-level
                           ``token_manager`` from ``auth.oauth_web`` or
                           ``auth.device_code``.
        """
        self.user_email = user_email
        self._token_manager = token_manager

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    async def _token(self) -> str:
        return await self._token_manager.get_token(self.user_email, POWERBI_SCOPE)

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
        async with httpx.AsyncClient(timeout=30.0) as client:
            resp = await client.get(
                f"{POWERBI_BASE}{path}",
                headers=self._headers(token),
            )
            resp.raise_for_status()
            return resp.json()

    async def post(self, path: str, body: dict | None = None) -> dict | None:
        token = await self._token()
        async with httpx.AsyncClient(timeout=30.0) as client:
            resp = await client.post(
                f"{POWERBI_BASE}{path}",
                headers=self._headers(token),
                json=body or {},
            )
            # 202 Accepted — refresh was queued; no JSON body
            if resp.status_code == 202:
                return {
                    "status": "accepted",
                    "message": "Refresh triggered successfully. Poll refresh history to track completion.",
                }
            resp.raise_for_status()
            return resp.json() if resp.content else None
