"""M365 MCP Server — FastAPI + MCP JSON-RPC 2.0 dispatcher.

v2.1.0: Multi-user OAuth, modular tool registry, Power BI integration.

All tool definitions and handlers are imported from the tools/ package.
Auth is handled by auth/oauth_web.py (multi-user, auto-refresh).
Token storage is pluggable: file (Option A) or PostgreSQL (Option B)
via TOKEN_STORE_BACKEND env var.

Endpoints:
  POST /mcp          — MCP JSON-RPC 2.0 handler
  GET  /mcp          — Server info + tool manifest
  GET  /health       — Health check
  /auth/*            — OAuth login, callback, status, revoke
"""

import os
import sys
import time
import logging
import asyncio
from typing import Any

import uvicorn
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware

# ---------------------------------------------------------------------------
# Auth router + token accessor
# ---------------------------------------------------------------------------
from .auth.oauth_web import router as auth_router

# ---------------------------------------------------------------------------
# Tool registry (auto-collected from tools/ sub-modules)
# ---------------------------------------------------------------------------
from .tools import TOOL_REGISTRY, TOOL_HANDLERS

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO").upper(),
    format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger("m365_mcp")

# ---------------------------------------------------------------------------
# MCP Protocol Constants
# ---------------------------------------------------------------------------
MCP_PROTOCOL_VERSION = "2024-11-05"
SERVER_NAME = "m365-mcp-server"
SERVER_VERSION = "2.1.0"

# ---------------------------------------------------------------------------
# FastAPI app
# ---------------------------------------------------------------------------
app = FastAPI(
    title=SERVER_NAME,
    version=SERVER_VERSION,
    docs_url=None,
    redoc_url=None,
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount auth routes: /auth/login, /auth/callback, /auth/status, /auth/revoke
app.include_router(auth_router)

# ---------------------------------------------------------------------------
# MCP bearer-token guard (server-level, independent of per-user OAuth)
# ---------------------------------------------------------------------------
MCP_BEARER_TOKEN = os.environ.get("MCP_BEARER_TOKEN", "")


def _verify_mcp_bearer(request: Request) -> None:
    """Enforce MCP_BEARER_TOKEN on POST /mcp if configured."""
    if not MCP_BEARER_TOKEN:
        return  # open / dev mode
    auth_header = request.headers.get("Authorization", "")
    if not auth_header.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Missing Bearer token")
    if auth_header[7:] != MCP_BEARER_TOKEN:
        raise HTTPException(status_code=403, detail="Invalid Bearer token")


# ---------------------------------------------------------------------------
# MCP JSON-RPC 2.0 dispatcher
# ---------------------------------------------------------------------------
def _jsonrpc_ok(result: Any, req_id: Any) -> dict:
    return {"jsonrpc": "2.0", "id": req_id, "result": result}


def _jsonrpc_error(code: int, message: str, req_id: Any = None) -> dict:
    return {"jsonrpc": "2.0", "id": req_id, "error": {"code": code, "message": message}}


async def _handle_initialize(params: dict) -> dict:
    return {
        "protocolVersion": MCP_PROTOCOL_VERSION,
        "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION},
        "capabilities": {"tools": {"listChanged": False}},
    }


async def _handle_tools_list(params: dict) -> dict:
    return {"tools": TOOL_REGISTRY}


async def _handle_tools_call(params: dict) -> dict:
    tool_name = params.get("name", "")
    arguments = params.get("arguments", {})

    handler = TOOL_HANDLERS.get(tool_name)
    if not handler:
        return {
            "content": [
                {"type": "text", "text": f"Unknown tool: {tool_name}"}
            ],
            "isError": True,
        }

    try:
        t0 = time.perf_counter()
        result = await handler(arguments)
        elapsed = time.perf_counter() - t0
        logger.info("Tool %s completed in %.2fs", tool_name, elapsed)

        import json
        return {
            "content": [
                {"type": "text", "text": json.dumps(result, indent=2, default=str)}
            ],
            "isError": False,
        }
    except HTTPException as exc:
        # Auth errors bubble up cleanly
        return {
            "content": [
                {"type": "text", "text": f"Auth error: {exc.detail}"}
            ],
            "isError": True,
        }
    except Exception as exc:
        logger.exception("Tool %s failed", tool_name)
        return {
            "content": [
                {"type": "text", "text": f"Error in {tool_name}: {exc}"}
            ],
            "isError": True,
        }


async def _handle_ping(params: dict) -> dict:
    """Respond to MCP ping requests."""
    return {}


# Method dispatch table
_MCP_METHODS = {
    "initialize": _handle_initialize,
    "tools/list": _handle_tools_list,
    "tools/call": _handle_tools_call,
    "ping": _handle_ping,
}

# Notifications the server should silently accept (no id, no response expected)
_MCP_NOTIFICATIONS = {
    "notifications/initialized",
    "notifications/cancelled",
    "notifications/progress",
}


@app.post("/mcp")
async def mcp_handler(request: Request):
    """MCP JSON-RPC 2.0 endpoint."""
    _verify_mcp_bearer(request)

    try:
        body = await request.json()
    except Exception:
        return JSONResponse(_jsonrpc_error(-32700, "Parse error"), status_code=400)

    method = body.get("method", "")
    params = body.get("params", {})
    req_id = body.get("id")

    # JSON-RPC notifications (no id) — accept silently
    if req_id is None and method in _MCP_NOTIFICATIONS:
        logger.debug("Notification received: %s", method)
        return JSONResponse({"jsonrpc": "2.0"}, status_code=202)

    # Any other notification without id — accept silently, don't 400
    if req_id is None and method not in _MCP_METHODS:
        logger.debug("Unknown notification ignored: %s", method)
        return JSONResponse({"jsonrpc": "2.0"}, status_code=202)

    handler = _MCP_METHODS.get(method)
    if not handler:
        return JSONResponse(
            _jsonrpc_error(-32601, f"Method not found: {method}", req_id),
            status_code=400,
        )

    result = await handler(params)
    return JSONResponse(_jsonrpc_ok(result, req_id))


@app.get("/mcp")
async def mcp_info():
    """Server info + full tool manifest."""
    return {
        "server": SERVER_NAME,
        "version": SERVER_VERSION,
        "protocol": MCP_PROTOCOL_VERSION,
        "tools_count": len(TOOL_REGISTRY),
        "tools": [t["name"] for t in TOOL_REGISTRY],
        "auth": {
            "mcp_bearer": "configured" if MCP_BEARER_TOKEN else "open",
            "oauth": "/auth/login?user_id=<email>",
            "status": "/auth/status",
        },
    }


@app.get("/health")
async def health():
    return {
        "status": "ok",
        "server": SERVER_NAME,
        "version": SERVER_VERSION,
        "tools_loaded": len(TOOL_REGISTRY),
    }


# ---------------------------------------------------------------------------
# Startup banner
# ---------------------------------------------------------------------------
@app.on_event("startup")
async def _startup():
    token_backend = os.environ.get("TOKEN_STORE_BACKEND", "file")
    logger.info("=" * 60)
    logger.info("%s v%s starting", SERVER_NAME, SERVER_VERSION)
    logger.info("MCP protocol: %s", MCP_PROTOCOL_VERSION)
    logger.info("Tools loaded: %d", len(TOOL_REGISTRY))
    logger.info("MCP bearer:   %s", "configured" if MCP_BEARER_TOKEN else "OPEN (dev mode)")
    logger.info("OAuth:        /auth/login?user_id=<email>")
    logger.info("Token store:  %s", token_backend)
    logger.info("Tool list:    %s", ", ".join(t["name"] for t in TOOL_REGISTRY))
    logger.info("=" * 60)


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------
def main():
    port = int(os.environ.get("PORT", 8080))
    uvicorn.run(
        "m365_mcp.__main__:app",
        host="0.0.0.0",
        port=port,
        log_level=os.environ.get("LOG_LEVEL", "info").lower(),
    )


if __name__ == "__main__":
    main()
