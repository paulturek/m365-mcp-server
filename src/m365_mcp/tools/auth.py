"""
Authentication MCP tools.

Exposes device-code flow authentication as MCP tools so any MCP client
can initiate and complete user authentication without direct HTTP access
to the /auth/* endpoints.

Tools:
  - auth_status              — check whether a user has a valid cached token
  - auth_start_device_login  — begin device code flow, return URL + code
  - auth_check_device_login  — poll for completion and persist the token
"""

from __future__ import annotations

import logging

from m365_mcp.auth.device_code import (
    start_device_code_flow,
    poll_device_code_flow,
)
from m365_mcp.auth.token_store_pg import get_token_store

logger = logging.getLogger("m365_mcp.tools.auth")

# ---------------------------------------------------------------------------
# Tool definitions
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "name": "auth_status",
        "description": (
            "Check whether a user has a valid cached Microsoft 365 token. "
            "Returns authenticated status and token expiry information."
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
            "Returns a URL and a short code — the user must visit the URL and enter the code "
            "to grant access. After the user completes this, call auth_check_device_login."
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
            "Poll for completion of a device code login that was started with "
            "auth_start_device_login. Call this after the user has visited the URL "
            "and entered the code. Returns success or pending status."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {
                    "type": "string",
                    "description": "User email address (must match the auth_start_device_login call)",
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
        store = await get_token_store()
        token_data = await store.get_token(user_id)
        if token_data:
            return {
                "authenticated": True,
                "user_id": user_id,
                "message": "User has a valid cached token.",
            }
        return {
            "authenticated": False,
            "user_id": user_id,
            "message": "No token found. Use auth_start_device_login to authenticate.",
        }
    except Exception as exc:
        logger.exception("auth_status failed for %s", user_id)
        return {
            "authenticated": False,
            "user_id": user_id,
            "error": str(exc),
        }


async def _handle_start_device_login(args: dict) -> dict:
    user_id: str = args["user_id"]
    try:
        result = await start_device_code_flow(user_id)
        return {
            "user_id": user_id,
            "verification_uri": result.get("verification_uri"),
            "user_code": result.get("user_code"),
            "message": result.get("message", (
                f"Visit {result.get('verification_uri')} and enter code: "
                f"{result.get('user_code')}"
            )),
            "expires_in": result.get("expires_in"),
        }
    except Exception as exc:
        logger.exception("auth_start_device_login failed for %s", user_id)
        return {"error": str(exc), "user_id": user_id}


async def _handle_check_device_login(args: dict) -> dict:
    user_id: str = args["user_id"]
    try:
        result = await poll_device_code_flow(user_id)
        return result
    except Exception as exc:
        logger.exception("auth_check_device_login failed for %s", user_id)
        return {"error": str(exc), "user_id": user_id}


HANDLERS: dict[str, object] = {
    "auth_status": _handle_auth_status,
    "auth_start_device_login": _handle_start_device_login,
    "auth_check_device_login": _handle_check_device_login,
}
