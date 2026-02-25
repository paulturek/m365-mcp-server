"""M365 MCP Auth — multi-user OAuth2 web flow.

Usage:
    from m365_mcp.auth.oauth_web import router as auth_router, get_access_token

    app.include_router(auth_router)  # adds /auth/*

    # in tool handlers:
    token = await get_access_token(user_id)
"""
from .oauth_web import router, get_access_token  # noqa: F401
from .token_store import TokenStore  # noqa: F401

__all__ = ["router", "get_access_token", "TokenStore"]
