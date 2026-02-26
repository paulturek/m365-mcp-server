"""MCP tools for authentication — device code flow + status checks.

Tools:
  auth_status              — Check if a user has an active M365 token
  auth_start_device_login  — Start device-code flow (returns code + URL)
  auth_check_device_login  — Check if device-code flow has completed
"""

from ..auth.device_code import start_device_flow, get_flow_status

# ---------------------------------------------------------------------------
# Tool definitions (MCP JSON schema)
# ---------------------------------------------------------------------------
TOOL_DEFINITIONS = [
    {
        "name": "auth_status",
        "description": (
            "Check whether a user has an active Microsoft 365 authentication "
            "token. Call this before making M365 API requests to verify the "
            "user is connected. Returns authenticated: true/false."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {
                    "type": "string",
                    "description": "Microsoft 365 email address (e.g. user@company.com)",
                }
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "auth_start_device_login",
        "description": (
            "Start a device-code login flow for Microsoft 365. Returns a "
            "one-time code and a URL (https://microsoft.com/devicelogin) for "
            "the user to enter in their browser. The server polls automatically "
            "in the background — use auth_check_device_login to check whether "
            "the user has completed sign-in. If the user is already "
            "authenticated, returns immediately without starting a new flow."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {
                    "type": "string",
                    "description": "Microsoft 365 email address to authenticate",
                }
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "auth_check_device_login",
        "description": (
            "Check whether a previously started device-code login flow has "
            "completed. Returns status: completed | pending | failed | "
            "cancelled | no_flow. Once completed, the user's M365 tools "
            "are ready to use."
        ),
        "inputSchema": {
            "type": "object",
            "properties": {
                "user_id": {
                    "type": "string",
                    "description": "Microsoft 365 email address to check",
                }
            },
            "required": ["user_id"],
        },
    },
]


# ---------------------------------------------------------------------------
# Tool handlers
# ---------------------------------------------------------------------------
async def handle_auth_status(args: dict) -> dict:
    """Check if user has a valid M365 token."""
    user_id = args.get("user_id", "").strip()
    if not user_id:
        return {"error": "user_id is required"}

    try:
        from ..auth.oauth_web import get_access_token

        token = await get_access_token(user_id)
        if token:
            return {
                "authenticated": True,
                "user_id": user_id,
                "message": "User is authenticated and ready to use M365 tools.",
            }
    except Exception:
        pass

    return {
        "authenticated": False,
        "user_id": user_id,
        "message": (
            "User is not authenticated. "
            "Use auth_start_device_login to begin in-chat sign-in."
        ),
    }


async def handle_auth_start_device_login(args: dict) -> dict:
    """Start device-code flow (or report already authenticated)."""
    user_id = args.get("user_id", "").strip()
    if not user_id:
        return {"error": "user_id is required"}

    # Check if already authenticated — skip flow if so
    try:
        from ..auth.oauth_web import get_access_token

        token = await get_access_token(user_id)
        if token:
            return {
                "status": "already_authenticated",
                "user_id": user_id,
                "message": "User is already authenticated. No login needed.",
            }
    except Exception:
        pass  # Not authenticated — proceed with device-code flow

    try:
        result = await start_device_flow(user_id)
        return result
    except Exception as exc:
        return {"status": "error", "error": str(exc)}


async def handle_auth_check_device_login(args: dict) -> dict:
    """Poll for device-code flow completion."""
    user_id = args.get("user_id", "").strip()
    if not user_id:
        return {"error": "user_id is required"}

    status = get_flow_status(user_id)

    # Enrich the response based on status
    if status["status"] == "completed":
        status["user_id"] = user_id
        status["message"] = (
            f"Authentication complete for {user_id}! "
            "You can now use all M365 tools (mail, calendar, files, etc.)."
        )
    elif status["status"] == "no_flow":
        # No active flow — check if user is already authenticated
        try:
            from ..auth.oauth_web import get_access_token

            token = await get_access_token(user_id)
            if token:
                return {
                    "status": "already_authenticated",
                    "user_id": user_id,
                    "message": "User is already authenticated.",
                }
        except Exception:
            pass
        status["message"] = (
            "No active login flow for this user. "
            "Use auth_start_device_login to begin."
        )
    elif status["status"] == "failed":
        status["user_id"] = user_id

    return status


# ---------------------------------------------------------------------------
# Handler map (consumed by __main__.py for tool registration)
# ---------------------------------------------------------------------------
TOOL_HANDLERS = {
    "auth_status": handle_auth_status,
    "auth_start_device_login": handle_auth_start_device_login,
    "auth_check_device_login": handle_auth_check_device_login,
}
