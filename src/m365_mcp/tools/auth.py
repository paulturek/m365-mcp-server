"""MCP auth tools — device code flow + status check.

Tools:
  auth_status              — check if a user_id is authenticated
  auth_start_device_login  — initiate device code flow, return code + URL
  auth_check_device_login  — poll for completion, store token on success
"""
import time
import logging
import asyncio

import msal

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Lazy imports — avoid circular import at module load time
# ---------------------------------------------------------------------------
def _get_store():
    from ..auth.oauth_web import _get_store
    return _get_store()


def _get_msal_public_app():
    import os
    client_id = os.environ.get("AZURE_CLIENT_ID", "")
    tenant_id = os.environ.get("AZURE_TENANT_ID", "common")
    return msal.PublicClientApplication(
        client_id=client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
    )


SCOPES = [
    "User.Read",
    "Files.ReadWrite.All",
    "Mail.ReadWrite",
    "Calendars.ReadWrite",
    "Sites.ReadWrite.All",
    "ChannelMessage.Send",
    "Tasks.ReadWrite",
    "https://analysis.windows.net/powerbi/api/.default",
]

# In-memory device flow state: user_id → msal flow dict
_device_flows: dict[str, dict] = {}

# ---------------------------------------------------------------------------
# Tool schemas
# ---------------------------------------------------------------------------
TOOLS = [
    {
        "name": "auth_status",
        "description": "Check whether a user is authenticated with Microsoft 365. Returns auth state, token expiry, and available scopes.",
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {
                    "type": "string",
                    "description": "User email address (e.g. paul.turek@bolthousefresh.com)",
                }
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "auth_start_device_login",
        "description": (
            "Start a Microsoft device code login flow for a user. "
            "Returns a short code and URL — present both to the user so they can authenticate "
            "in their browser. Call auth_check_device_login to poll for completion."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {
                    "type": "string",
                    "description": "User email address to authenticate",
                }
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "auth_check_device_login",
        "description": (
            "Poll for completion of a device code login flow started with auth_start_device_login. "
            "Call this after the user has visited the URL and entered the code. "
            "Returns 'authenticated' on success, 'pending' if still waiting."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {
                    "type": "string",
                    "description": "User email address (must match auth_start_device_login call)",
                }
            },
            "required": ["user_id"],
        },
    },
]

# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------
async def _auth_status(params: dict) -> dict:
    user_id = params.get("user_id", "").strip()
    if not user_id:
        return {"error": "user_id is required"}

    store = _get_store()
    cache = await store.get(user_id)

    if not cache:
        return {
            "user_id": user_id,
            "authenticated": False,
            "message": f"Not authenticated. Use auth_start_device_login to sign in.",
        }

    expired = cache.get("expires_at", 0) < time.time()
    can_refresh = bool(cache.get("refresh_token"))

    return {
        "user_id": user_id,
        "authenticated": True,
        "display_name": cache.get("display_name", user_id),
        "token_expired": expired,
        "can_refresh": can_refresh,
        "scopes": cache.get("scope", ""),
        "authorized_at": cache.get("authorized_at"),
        "last_refreshed": cache.get("last_refreshed"),
    }


async def _auth_start_device_login(params: dict) -> dict:
    user_id = params.get("user_id", "").strip()
    if not user_id:
        return {"error": "user_id is required"}

    try:
        app = _get_msal_public_app()
        flow = app.initiate_device_flow(scopes=SCOPES)

        if "user_code" not in flow:
            return {
                "error": "Failed to start device flow",
                "detail": flow.get("error_description", "unknown error"),
            }

        _device_flows[user_id] = flow
        logger.info("Device code flow started for %s", user_id)

        return {
            "status": "pending",
            "user_id": user_id,
            "user_code": flow["user_code"],
            "verification_uri": flow["verification_uri"],
            "expires_in_seconds": flow.get("expires_in", 900),
            "message": (
                f"Visit {flow['verification_uri']} and enter code {flow['user_code']}. "
                f"Then call auth_check_device_login to confirm."
            ),
        }
    except Exception as exc:
        logger.exception("Device flow init failed for %s", user_id)
        return {"error": str(exc)}


async def _auth_check_device_login(params: dict) -> dict:
    user_id = params.get("user_id", "").strip()
    if not user_id:
        return {"error": "user_id is required"}

    flow = _device_flows.get(user_id)
    if not flow:
        return {
            "status": "no_flow",
            "user_id": user_id,
            "message": "No active device flow. Call auth_start_device_login first.",
        }

    try:
        app = _get_msal_public_app()
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(
            None, lambda: app.acquire_token_by_device_flow(flow)
        )

        if "access_token" not in result:
            error = result.get("error", "")
            if error == "authorization_pending":
                return {
                    "status": "pending",
                    "user_id": user_id,
                    "message": "Still waiting — user has not completed sign-in yet.",
                }
            _device_flows.pop(user_id, None)
            return {
                "status": "failed",
                "user_id": user_id,
                "error": error,
                "detail": result.get("error_description", ""),
            }

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
        _device_flows.pop(user_id, None)
        logger.info("Device code auth complete for %s (%s)", display_name, user_id)

        return {
            "status": "authenticated",
            "user_id": user_id,
            "display_name": display_name,
            "scopes": token_data["scope"],
            "message": f"Successfully authenticated as {display_name}. All M365 tools are now available.",
        }

    except Exception as exc:
        logger.exception("Device flow poll failed for %s", user_id)
        return {"error": str(exc)}


HANDLERS = {
    "auth_status": _auth_status,
    "auth_start_device_login": _auth_start_device_login,
    "auth_check_device_login": _auth_check_device_login,
}
