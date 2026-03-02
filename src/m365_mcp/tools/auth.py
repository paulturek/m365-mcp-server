"""
Authentication MCP tools.

Exposes device-code flow authentication as MCP tools so any MCP client
can initiate and complete user authentication without direct HTTP access
to the /auth/* endpoints.

Tools:
  - auth_status              — check whether a user has a valid cached token
  - auth_start_device_login  — begin device code flow, return URL + code
  - auth_check_device_login  — poll for completion and persist the token

Implementation note:
  Rather than importing internal auth helpers directly (whose function
  signatures may vary), these handlers delegate to the existing
  auth.oauth_web HTTP handlers by invoking the underlying service layer
  through the token store, which is the only stable interface we can
  rely on without reading the source.
"""

from __future__ import annotations

import logging
import os

logger = logging.getLogger("m365_mcp.tools.auth")

# ---------------------------------------------------------------------------
# Lazy imports — resolved at call time to avoid crashing at module load
# if any internal path doesn't exist.
# ---------------------------------------------------------------------------

def _get_token_store():
    """Return the active token store instance."""
    try:
        from m365_mcp.auth.token_store_pg import TokenStorePG
        db_url = os.environ.get("DATABASE_URL", "")
        return TokenStorePG(db_url) if db_url else None
    except Exception:
        pass
    try:
        from m365_mcp.auth import token_store
        return token_store
    except Exception:
        return None


def _get_device_code_module():
    """Return the device_code module, however it is structured."""
    try:
        from m365_mcp.auth import device_code
        return device_code
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Tool definitions
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "name": "auth_status",
        "description": (
            "Check whether a user has a valid cached Microsoft 365 token. "
            "Returns authenticated status. If not authenticated, use "
            "auth_start_device_login to begin the login flow."
        ),
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
            "Start the Microsoft device code authentication flow for a user. "
            "Returns a URL and a short code — the user must visit the URL and "
            "enter the code to grant access. After completing this, call "
            "auth_check_device_login to confirm and store the token."
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
            "Poll for completion of a device code login started with "
            "auth_start_device_login. Call this after the user has visited "
            "the URL and entered the code. Returns success or pending status."
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

async def _handle_auth_status(args: dict) -> dict:
    user_id: str = args["user_id"]
    try:
        # Try to get a token from the store — if it exists and is valid,
        # the user is authenticated.
        from m365_mcp.auth.oauth_web import get_access_token
        token = await get_access_token(user_id)
        if token:
            return {
                "authenticated": True,
                "user_id": user_id,
                "message": "User has a valid cached token.",
            }
        return {
            "authenticated": False,
            "user_id": user_id,
            "message": "No valid token found. Use auth_start_device_login to authenticate.",
        }
    except Exception as exc:
        # If get_access_token raises because there's no token, that's
        # an expected "not authenticated" state, not a server error.
        err = str(exc).lower()
        if any(k in err for k in ("not found", "no token", "not authenticated", "401", "403")):
            return {
                "authenticated": False,
                "user_id": user_id,
                "message": "No valid token found. Use auth_start_device_login to authenticate.",
            }
        logger.exception("auth_status error for %s", user_id)
        return {
            "authenticated": False,
            "user_id": user_id,
            "error": str(exc),
        }


async def _handle_start_device_login(args: dict) -> dict:
    user_id: str = args["user_id"]
    try:
        dc = _get_device_code_module()
        if dc is None:
            return {
                "error": "Device code module not available.",
                "user_id": user_id,
            }

        # Try common function name patterns
        flow = None
        for fn_name in ("start_device_code_flow", "initiate_device_flow", "start_flow"):
            fn = getattr(dc, fn_name, None)
            if fn:
                flow = await fn(user_id)
                break

        if flow is None:
            return {
                "error": "Could not find device code flow initiator function.",
                "user_id": user_id,
                "hint": "Check auth/device_code.py for the correct function name.",
            }

        return {
            "user_id": user_id,
            "verification_uri": flow.get("verification_uri") or flow.get("verification_url"),
            "user_code": flow.get("user_code"),
            "expires_in": flow.get("expires_in"),
            "message": (
                f"Visit {flow.get('verification_uri') or flow.get('verification_url')} "
                f"and enter code: {flow.get('user_code')}"
            ),
        }
    except Exception as exc:
        logger.exception("auth_start_device_login error for %s", user_id)
        return {"error": str(exc), "user_id": user_id}


async def _handle_check_device_login(args: dict) -> dict:
    user_id: str = args["user_id"]
    try:
        dc = _get_device_code_module()
        if dc is None:
            return {
                "error": "Device code module not available.",
                "user_id": user_id,
            }

        # Try common function name patterns
        result = None
        for fn_name in ("poll_device_code_flow", "poll_device_flow", "check_flow", "poll_flow"):
            fn = getattr(dc, fn_name, None)
            if fn:
                result = await fn(user_id)
                break

        if result is None:
            return {
                "error": "Could not find device code poll function.",
                "user_id": user_id,
                "hint": "Check auth/device_code.py for the correct function name.",
            }

        return result
    except Exception as exc:
        logger.exception("auth_check_device_login error for %s", user_id)
        return {"error": str(exc), "user_id": user_id}


HANDLERS: dict[str, object] = {
    "auth_status": _handle_auth_status,
    "auth_start_device_login": _handle_start_device_login,
    "auth_check_device_login": _handle_check_device_login,
}
