"""
Power BI REST API client.

Uses a separate token scope from Microsoft Graph:
  https://analysis.windows.net/powerbi/api/.default

MSAL issues a separate token per audience from the shared cached account,
so no additional environment variables are required.
"""

import httpx
from app.auth.token_manager import get_access_token

POWERBI_SCOPE = ["https://analysis.windows.net/powerbi/api/.default"]
POWERBI_BASE = "https://api.powerbi.com/v1.0/myorg"


class PowerBIClient:
    def __init__(self, user_email: str):
        self.user_email = user_email
        self.base = POWERBI_BASE

    async def _get_token(self) -> str:
        token = await get_access_token(self.user_email, scopes=POWERBI_SCOPE)
        return token

    def _headers(self, token: str) -> dict:
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    async def get(self, path: str) -> dict:
        token = await self._get_token()
        async with httpx.AsyncClient() as client:
            resp = await client.get(
                f"{self.base}{path}",
                headers=self._headers(token),
                timeout=30.0,
            )
            resp.raise_for_status()
            return resp.json()

    async def post(self, path: str, body: dict | None = None) -> dict | None:
        token = await self._get_token()
        async with httpx.AsyncClient() as client:
            resp = await client.post(
                f"{self.base}{path}",
                headers=self._headers(token),
                json=body or {},
                timeout=30.0,
            )
            if resp.status_code == 202:
                return {"status": "accepted", "message": "Refresh triggered successfully"}
            resp.raise_for_status()
            return resp.json() if resp.content else None
