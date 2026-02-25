"""Multi-user OAuth2 web flow for M365 MCP Server.

Replaces the legacy token_manager.py + token_cache.py with a
browser-based consent redirect that stores per-user tokens.

Routes (mounted at /auth):
  GET  /auth/login?user_id=...   → redirect to Microsoft consent
  GET  /auth/callback            → exchange code, store token
  GET  /auth/status[?user_id=]   → check auth for one / all users
  DELETE /auth/revoke?user_id=   → remove stored tokens

Public API for tool handlers:
  token = await get_access_token(user_id)
"""
import os
import time
import secrets
import logging
from pathlib import Path
from typing import Optional

import msal
from fastapi import APIRouter, Request, HTTPException, Query
from fastapi.responses import RedirectResponse, HTMLResponse

from .token_store import TokenStore
from .token_store_file import FileTokenStore

logger = logging.getLogger(__name__)
router = APIRouter(prefix="/auth", tags=["auth"])

# ---------------------------------------------------------------------------
# Config (all from env)
# ---------------------------------------------------------------------------
CLIENT_ID: str = os.environ.get("AZURE_CLIENT_ID", "")
CLIENT_SECRET: str = os.environ.get("AZURE_CLIENT_SECRET", "")
TENANT_ID: str = os.environ.get("AZURE_TENANT_ID", "common")
REDIRECT_URI: str = os.environ.get("OAUTH_REDIRECT_URI", "")

SCOPES = [
    "User.Read",
    "Files.ReadWrite.All",
    "Mail.ReadWrite",
    "Calendars.ReadWrite",
    "Sites.ReadWrite.All",
    "ChannelMessage.Send",
    "Tasks.ReadWrite",
]

TOKEN_ENCRYPTION_KEY: str = os.environ.get("TOKEN_ENCRYPTION_KEY", "")
TOKEN_STORE_BACKEND: str = os.environ.get("TOKEN_STORE_BACKEND", "file")  # "file" | "pg"
TOKEN_STORE_PATH: str = os.environ.get("TOKEN_STORE_PATH", "/app/data/tokens.enc")
DATABASE_URL: str = os.environ.get("DATABASE_URL", "")

# CSRF: state → {"user_id": ..., "ts": ...}
_auth_states: dict[str, dict] = {}

# ---------------------------------------------------------------------------
# Token store singleton
# ---------------------------------------------------------------------------
_store: Optional[TokenStore] = None


def _get_store() -> TokenStore:
    global _store
    if _store is not None:
        return _store

    if TOKEN_STORE_BACKEND == "pg" and DATABASE_URL:
        from .token_store_pg import PgTokenStore

        _store = PgTokenStore(DATABASE_URL, TOKEN_ENCRYPTION_KEY)
        logger.info("Token store: PostgreSQL")
    else:
        _store = FileTokenStore(Path(TOKEN_STORE_PATH), TOKEN_ENCRYPTION_KEY)
        logger.info("Token store: encrypted file (%s)", TOKEN_STORE_PATH)
    return _store


def _get_msal_app() -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )


# ---------------------------------------------------------------------------
# Public: get a valid access token (auto-refresh)
# ---------------------------------------------------------------------------
async def get_access_token(user_id: str) -> str:
    """Return a valid MS Graph access token for *user_id*.

    Automatically refreshes expired tokens.  Raises HTTPException(401)
    if the user has not authorised or refresh fails.
    """
    if not user_id:
        raise HTTPException(
            status_code=400,
            detail="user_id is required for Microsoft 365 operations",
        )

    store = _get_store()
    cache = await store.get(user_id)

    if not cache:
        raise HTTPException(
            status_code=401,
            detail=(
                f"User '{user_id}' has not authorised. "
                f"Visit /auth/login?user_id={user_id} to connect."
            ),
        )

    # Still valid (5-min safety margin)?
    if cache.get("expires_at", 0) > time.time() + 300:
        return cache["access_token"]

    # --- Refresh ---
    refresh_token = cache.get("refresh_token")
    if not refresh_token:
        raise HTTPException(
            status_code=401,
            detail=f"No refresh token for '{user_id}'. Re-authorise at /auth/login?user_id={user_id}",
        )

    app = _get_msal_app()
    result = app.acquire_token_by_refresh_token(refresh_token, scopes=SCOPES)

    if "access_token" not in result:
        logger.error("Refresh failed for %s: %s", user_id, result.get("error_description"))
        await store.delete(user_id)
        raise HTTPException(
            status_code=401,
            detail=f"Token refresh failed for '{user_id}'. Re-authorise at /auth/login?user_id={user_id}",
        )

    updated = {
        "access_token": result["access_token"],
        "refresh_token": result.get("refresh_token", refresh_token),
        "expires_at": time.time() + result.get("expires_in", 3600),
        "scope": result.get("scope", ""),
        "display_name": cache.get("display_name", user_id),
        "authorized_at": cache.get("authorized_at", time.time()),
        "last_refreshed": time.time(),
    }
    await store.save(user_id, updated)
    logger.info("Token refreshed for %s", user_id)
    return updated["access_token"]


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------
@router.get("/login")
async def login(
    user_id: str = Query(..., description="Unique user identifier (email recommended)"),
):
    """Redirect the user to Microsoft's OAuth consent screen."""
    if not CLIENT_ID or not REDIRECT_URI:
        raise HTTPException(
            status_code=500,
            detail="OAuth not configured. Set AZURE_CLIENT_ID and OAUTH_REDIRECT_URI.",
        )

    app = _get_msal_app()
    state = secrets.token_urlsafe(32)
    _auth_states[state] = {"user_id": user_id, "ts": time.time()}

    auth_url = app.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        state=state,
        prompt="consent",
        login_hint=user_id if "@" in user_id else None,
    )
    return RedirectResponse(auth_url)


@router.get("/callback")
async def callback(request: Request):
    """Handle the OAuth redirect from Microsoft — exchange code for tokens."""
    code = request.query_params.get("code")
    state = request.query_params.get("state")
    error = request.query_params.get("error")

    if error:
        desc = request.query_params.get("error_description", error)
        return HTMLResponse(
            f"<h2>Authorisation Failed</h2><p>{desc}</p>", status_code=400
        )

    # CSRF validation + recover user_id
    state_data = _auth_states.pop(state, None)
    if not state_data:
        raise HTTPException(status_code=400, detail="Invalid or expired state parameter")
    user_id = state_data["user_id"]

    # Housekeep stale states (>10 min)
    cutoff = time.time() - 600
    for k in [k for k, v in _auth_states.items() if v["ts"] < cutoff]:
        del _auth_states[k]

    # Exchange authorisation code for tokens
    app = _get_msal_app()
    result = app.acquire_token_by_authorization_code(
        code=code, scopes=SCOPES, redirect_uri=REDIRECT_URI
    )

    if "access_token" not in result:
        desc = result.get("error_description", "Unknown error")
        return HTMLResponse(
            f"<h2>Token Exchange Failed</h2><p>{desc}</p>", status_code=400
        )

    id_claims = result.get("id_token_claims", {})
    display_name = (
        id_claims.get("name")
        or id_claims.get("preferred_username")
        or user_id
    )

    token_data = {
        "access_token": result["access_token"],
        "refresh_token": result.get("refresh_token", ""),
        "expires_at": time.time() + result.get("expires_in", 3600),
        "scope": result.get("scope", ""),
        "display_name": display_name,
        "microsoft_oid": id_claims.get("oid", ""),
        "authorized_at": time.time(),
        "last_refreshed": time.time(),
    }

    store = _get_store()
    await store.save(user_id, token_data)
    logger.info("Authorised: %s (%s)", display_name, user_id)

    return HTMLResponse(f"""
    <!DOCTYPE html>
    <html><head><title>M365 MCP — Authorised</title>
    <style>body{{font-family:system-ui;max-width:480px;margin:60px auto;text-align:center}}</style>
    </head><body>
        <h2>&#9989; Authorised</h2>
        <p><strong>{display_name}</strong><br><code>{user_id}</code></p>
        <p>Tokens stored &amp; auto-refresh enabled.<br>You can close this tab.</p>
    </body></html>
    """)


@router.get("/status")
async def auth_status(user_id: Optional[str] = Query(None)):
    """Check auth status for one user, or list all authorised users."""
    store = _get_store()

    if user_id:
        cache = await store.get(user_id)
        if not cache:
            return {"user_id": user_id, "authenticated": False}
        return {
            "user_id": user_id,
            "authenticated": True,
            "display_name": cache.get("display_name", ""),
            "token_expired": cache.get("expires_at", 0) < time.time(),
            "can_refresh": bool(cache.get("refresh_token")),
            "scopes": cache.get("scope", ""),
            "authorized_at": cache.get("authorized_at"),
        }

    users = await store.list_users()
    return {"total_users": len(users), "users": users}


@router.delete("/revoke")
async def revoke(user_id: str = Query(..., description="User to revoke")):
    """Remove all stored tokens for a user."""
    store = _get_store()
    if await store.delete(user_id):
        return {"revoked": True, "user_id": user_id}
    raise HTTPException(status_code=404, detail=f"No tokens found for '{user_id}'")
