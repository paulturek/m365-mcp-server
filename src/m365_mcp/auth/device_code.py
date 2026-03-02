"""Device Code Flow for M365 MCP Server.

Enables in-chat authentication: the assistant presents a code + URL,
the user enters the code at microsoft.com/devicelogin, and the server
polls in the background until authentication completes.

Prerequisites:
  - Azure Portal → App Registration → Authentication →
    "Allow public client flows" must be set to **Yes**
  - AZURE_CLIENT_ID and AZURE_TENANT_ID environment variables
"""

import os
import json
import logging
import asyncio
import inspect
from typing import Any

import msal

logger = logging.getLogger("m365_mcp.auth.device_code")

# ---------------------------------------------------------------------------
# Graph delegated scopes — must match those in the OAuth web flow
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# In-memory state for pending / completed device-code flows
# ---------------------------------------------------------------------------
_pending_flows: dict[str, dict] = {}
_flow_results: dict[str, dict] = {}
_polling_tasks: dict[str, asyncio.Task] = {}


# ---------------------------------------------------------------------------
# MSAL public client (no client_secret needed for device-code flow)
# ---------------------------------------------------------------------------
def _public_client_app() -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        client_id=os.environ["AZURE_CLIENT_ID"],
        authority=f"https://login.microsoftonline.com/{os.environ['AZURE_TENANT_ID']}",
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------
async def start_device_flow(user_id: str) -> dict:
    """Initiate a device-code flow and begin background polling.

    Returns a dict with user_code, verification_uri, and a human-readable
    message the assistant can present directly in chat.
    """
    # Cancel any existing flow for this user
    if user_id in _polling_tasks:
        _polling_tasks[user_id].cancel()
        _polling_tasks.pop(user_id, None)

    app = _public_client_app()
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)

    if "user_code" not in flow:
        return {
            "status": "error",
            "error": flow.get(
                "error_description", "Could not initiate device-code flow"
            ),
            "hint": (
                "Ensure 'Allow public client flows' is set to Yes in "
                "Azure Portal → App Registration → Authentication."
            ),
        }

    _pending_flows[user_id] = flow
    _flow_results.pop(user_id, None)

    # Start background polling (runs until user completes or flow expires)
    task = asyncio.create_task(_poll_for_token(user_id, app, flow))
    _polling_tasks[user_id] = task

    logger.info(
        "Device-code flow started for %s (code=%s, expires=%ss)",
        user_id,
        flow["user_code"],
        flow.get("expires_in", "?"),
    )

    return {
        "status": "pending",
        "user_code": flow["user_code"],
        "verification_uri": flow["verification_uri"],
        "message": (
            f"To sign in, visit **{flow['verification_uri']}** "
            f"and enter the code **{flow['user_code']}**"
        ),
        "expires_in_seconds": flow.get("expires_in", 900),
    }


def get_flow_status(user_id: str) -> dict:
    """Return the current device-code flow status for a user.

    Possible statuses: completed, pending, failed, cancelled, no_flow.
    Completed/failed results are consumed (returned once then cleared).
    """
    if user_id in _flow_results:
        return _flow_results.pop(user_id)

    if user_id in _pending_flows:
        return {
            "status": "pending",
            "user_code": _pending_flows[user_id].get("user_code", ""),
            "message": "Waiting for you to complete sign-in...",
        }

    return {"status": "no_flow"}


# ---------------------------------------------------------------------------
# Background polling
# ---------------------------------------------------------------------------
async def _poll_for_token(
    user_id: str,
    app: msal.PublicClientApplication,
    flow: dict,
) -> None:
    """Block-poll Microsoft until the user completes auth or the flow expires."""
    try:
        # acquire_token_by_device_flow is synchronous/blocking — run in thread
        result = await asyncio.to_thread(app.acquire_token_by_device_flow, flow)

        if "access_token" in result:
            logger.info("Device-code flow completed for %s", user_id)
            _flow_results[user_id] = {"status": "completed"}
            await _persist_token(user_id, result)
        else:
            error = result.get(
                "error_description", result.get("error", "Unknown error")
            )
            logger.warning("Device-code flow failed for %s: %s", user_id, error)
            _flow_results[user_id] = {"status": "failed", "error": error}

    except asyncio.CancelledError:
        logger.info("Device-code flow cancelled for %s", user_id)
        _flow_results[user_id] = {"status": "cancelled"}
    except Exception as exc:
        logger.exception("Device-code polling error for %s", user_id)
        _flow_results[user_id] = {"status": "failed", "error": str(exc)}
    finally:
        _pending_flows.pop(user_id, None)
        _polling_tasks.pop(user_id, None)


# ---------------------------------------------------------------------------
# Token persistence — delegates to the existing OAuth token store
# ---------------------------------------------------------------------------
async def _persist_token(user_id: str, token_result: dict) -> None:
    """Store the acquired token so existing M365 tools can use it.

    Primary path: import store_token from oauth_web (same store as code-flow).
    Fallback:     direct PostgreSQL upsert into oauth_tokens table.
    """
    # --- Primary: reuse oauth_web's store ---
    try:
        from .oauth_web import store_token

        if inspect.iscoroutinefunction(store_token):
            await store_token(user_id, token_result)
        else:
            store_token(user_id, token_result)
        logger.info("Token persisted for %s via oauth_web.store_token", user_id)
        return
    except (ImportError, AttributeError) as exc:
        logger.info(
            "oauth_web.store_token unavailable (%s); falling back to direct PG",
            exc,
        )

    # --- Fallback: direct PostgreSQL ---
    await _pg_store_token(user_id, token_result)


# ---------------------------------------------------------------------------
# Fallback: direct PostgreSQL token persistence
# ---------------------------------------------------------------------------
async def _pg_store_token(user_id: str, token_result: dict) -> None:
    """Direct PostgreSQL upsert into oauth_tokens table."""
    db_url = os.environ.get("DATABASE_URL", "")
    if not db_url:
        raise RuntimeError(
            "Cannot persist device-code token: "
            "oauth_web.store_token not importable and DATABASE_URL not set."
        )

    import psycopg2

    def _upsert():
        conn = psycopg2.connect(db_url)
        try:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    CREATE TABLE IF NOT EXISTS oauth_tokens (
                        user_id     TEXT PRIMARY KEY,
                        token_data  JSONB NOT NULL,
                        updated_at  TIMESTAMPTZ DEFAULT now()
                    )
                    """
                )
                cur.execute(
                    """
                    INSERT INTO oauth_tokens (user_id, token_data, updated_at)
                    VALUES (%s, %s, now())
                    ON CONFLICT (user_id) DO UPDATE
                      SET token_data  = EXCLUDED.token_data,
                          updated_at  = now()
                    """,
                    (user_id, json.dumps(token_result)),
                )
                conn.commit()
        finally:
            conn.close()

    await asyncio.to_thread(_upsert)
    logger.info("Token persisted for %s via direct PG fallback", user_id)
