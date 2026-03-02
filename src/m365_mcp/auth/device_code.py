"""Device code authentication flow for M365 MCP Server.

Provides inline device-code auth via MCP tools:
  - auth_start_device_login  → returns user_code + verification_uri
  - auth_check_device_login  → polls Azure AD and stores token on success
"""

import asyncio
import logging
import os
from typing import Any

import msal

from ..token_store import PgTokenStore, FileTokenStore, TokenStore
from cryptography.fernet import Fernet

logger = logging.getLogger("m365_mcp")

# ── Canonical scope list ──────────────────────────────────────────────────────
# DIAGNOSTIC: logged verbatim before every Azure AD call.
# Must match auth/oauth_web.py SCOPES exactly.
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

# ── Azure AD app credentials ──────────────────────────────────────────────────
CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
TENANT_ID = os.environ.get("AZURE_TENANT_ID", "common")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# ── Token store ───────────────────────────────────────────────────────────────
_ENCRYPTION_KEY = os.environ.get("TOKEN_ENCRYPTION_KEY", "")
_DATABASE_URL   = os.environ.get("DATABASE_URL", "")

def _get_token_store() -> TokenStore:
    if _DATABASE_URL:
        return PgTokenStore(_ENCRYPTION_KEY, _DATABASE_URL)
    return FileTokenStore(_ENCRYPTION_KEY)

# ── In-memory flow cache (keyed by user_id) ───────────────────────────────────
_pending_flows: dict[str, dict] = {}


async def _pg_store_token(user_id: str, token_data: dict) -> None:
    """Persist token to PostgreSQL token store."""
    store = _get_token_store()
    await store.save_token(user_id, token_data)


async def start_device_login(user_id: str) -> dict[str, Any]:
    """Initiate a device-code flow for *user_id*.

    Returns a dict with ``user_code``, ``verification_uri``, and
    ``expires_in`` for display to the user.
    """
    # ── DIAGNOSTIC ────────────────────────────────────────────────────────────
    logger.info("DIAG device_code.py GRAPH_SCOPES (%d): %s",
                len(GRAPH_SCOPES), " ".join(GRAPH_SCOPES))
    logger.info("DIAG CLIENT_ID=%s  TENANT_ID=%s  AUTHORITY=%s",
                CLIENT_ID, TENANT_ID, AUTHORITY)
    # ─────────────────────────────────────────────────────────────────────────

    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY)

    logger.info("DIAG calling initiate_device_flow with scopes: %s", GRAPH_SCOPES)
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)

    if "error" in flow:
        logger.error("DIAG device_flow error response: %s", flow)
        raise RuntimeError(f"Device flow error: {flow.get('error_description', flow)}")

    _pending_flows[user_id] = {"flow": flow, "app": app}

    logger.info("DIAG device flow initiated OK for %s — user_code=%s",
                user_id, flow.get("user_code"))

    return {
        "user_code":        flow["user_code"],
        "verification_uri": flow["verification_uri"],
        "expires_in":       flow.get("expires_in", 900),
        "message":          flow.get("message", ""),
    }


async def check_device_login(user_id: str) -> dict[str, Any]:
    """Poll Azure AD for token completion.

    Returns ``{"status": "pending"}`` while the user hasn't authenticated,
    ``{"status": "success"}`` on completion, or raises on error.
    """
    if user_id not in _pending_flows:
        return {"status": "error", "message": "No pending flow. Call start_device_login first."}

    entry = _pending_flows[user_id]
    app: msal.PublicClientApplication = entry["app"]
    flow: dict = entry["flow"]

    logger.info("DIAG check_device_login polling for %s", user_id)

    result = app.acquire_token_by_device_flow(flow, timeout=5)

    if "error" in result:
        err = result.get("error", "")
        desc = result.get("error_description", "")

        logger.error("DIAG acquire_token_by_device_flow error: %s — %s", err, desc)

        if err == "authorization_pending":
            return {"status": "pending"}

        # Any other error — surface it clearly
        return {"status": "error", "message": f"{err}: {desc}"}

    # ── Success ───────────────────────────────────────────────────────────────
    logger.info("DIAG token acquired for %s — scopes granted: %s",
                user_id, result.get("scope", ""))

    token_data = {
        "access_token":  result["access_token"],
        "refresh_token": result.get("refresh_token", ""),
        "expires_in":    result.get("expires_in", 3600),
        "scope":         result.get("scope", ""),
        "token_type":    result.get("token_type", "Bearer"),
    }

    await _pg_store_token(user_id, token_data)
    del _pending_flows[user_id]

    logger.info("DIAG token stored for %s", user_id)

    return {
        "status":  "success",
        "message": f"Authenticated successfully. Scopes granted: {token_data['scope']}",
    }
