"""Device code authentication flow for M365 MCP Server."""

import logging
import os
from typing import Any

import msal

from ..token_store import PgTokenStore, FileTokenStore, TokenStore

logger = logging.getLogger("m365_mcp")

GRAPH_SCOPES = [
    "User.Read",
    "User.ReadBasic.All",
    "Mail.ReadWrite",
    "Mail.Send",
    "Calendars.ReadWrite",
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
    "Team.ReadBasic.All",
    "Channel.ReadBasic.All",
    "ChannelMessage.Send",
    "Tasks.ReadWrite",
]

CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "common")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

_ENCRYPTION_KEY = os.environ.get("TOKEN_ENCRYPTION_KEY", "")
_DATABASE_URL   = os.environ.get("DATABASE_URL", "")


def _get_token_store() -> TokenStore:
    if _DATABASE_URL:
        return PgTokenStore(_ENCRYPTION_KEY, _DATABASE_URL)
    return FileTokenStore(_ENCRYPTION_KEY)


# In-memory flow cache keyed by user_id
_pending_flows: dict[str, dict] = {}


async def _pg_store_token(user_id: str, token_data: dict) -> None:
    store = _get_token_store()
    await store.save_token(user_id, token_data)


async def start_device_login(user_id: str) -> dict[str, Any]:
    """Initiate a device-code flow. Returns user_code + verification_uri."""
    logger.info("start_device_login: user=%s scopes=%s", user_id, GRAPH_SCOPES)

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)

    if "error" in flow:
        logger.error("device_flow error: %s", flow)
        raise RuntimeError(f"Device flow error: {flow.get('error_description', flow)}")

    _pending_flows[user_id] = {"flow": flow, "app": app}

    logger.info("device_flow initiated for %s — user_code=%s", user_id, flow.get("user_code"))

    return {
        "user_code":        flow["user_code"],
        "verification_uri": flow["verification_uri"],
        "expires_in":       flow.get("expires_in", 900),
        "message":          flow.get("message", ""),
    }


async def check_device_login(user_id: str) -> dict[str, Any]:
    """Poll Azure AD for token. Returns status: pending | success | error."""
    if user_id not in _pending_flows:
        return {"status": "error", "message": "No pending flow. Call start_device_login first."}

    entry = _pending_flows[user_id]
    app: msal.PublicClientApplication = entry["app"]
    flow: dict = entry["flow"]

    result = app.acquire_token_by_device_flow(flow, timeout=5)

    if "error" in result:
        err  = result.get("error", "")
        desc = result.get("error_description", "")
        logger.info("device_flow poll: %s — %s", err, desc)

        if err == "authorization_pending":
            return {"status": "pending"}

        return {"status": "error", "message": f"{err}: {desc}"}

    # Success
    logger.info("device_flow success for %s — scopes: %s", user_id, result.get("scope", ""))

    token_data = {
        "access_token":  result["access_token"],
        "refresh_token": result.get("refresh_token", ""),
        "expires_in":    result.get("expires_in", 3600),
        "scope":         result.get("scope", ""),
        "token_type":    result.get("token_type", "Bearer"),
    }

    await _pg_store_token(user_id, token_data)
    del _pending_flows[user_id]

    return {
        "status":  "success",
        "message": f"Authenticated successfully. Scopes granted: {token_data['scope']}",
    }
