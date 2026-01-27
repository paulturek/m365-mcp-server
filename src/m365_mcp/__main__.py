"""M365 MCP Server entry point.

Runs the MCP server with HTTP transport for Railway deployment.
Supports both:
- Streamable HTTP transport (POST /mcp) - Modern standard
- Legacy SSE transport (GET /sse, POST /messages) - Backward compatibility

Requires bearer token authentication via MCP_BEARER_TOKEN env var.
Exposes /health endpoint for healthchecks.
Exposes /auth/device-code for initial authentication.
"""

import asyncio
import json
import logging
import os
import sys
import threading
import uuid
from datetime import datetime
from typing import Any, Optional

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

from starlette.applications import Starlette
from starlette.responses import JSONResponse, Response, StreamingResponse
from starlette.routing import Route
from starlette.requests import Request
import uvicorn

from .config import config, M365Config
from .auth.token_manager import TokenManager
from .clients.graph_client import GraphClient, AuthenticationRequiredError
from .clients.powerbi_client import PowerBIClient
from .services import (
    OutlookService,
    OneDriveService,
    SharePointService,
    ExcelService,
    OfficeDocsService,
    TeamsService,
    PowerBIService,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)

# Global instances - use imported config directly
token_manager: Optional[TokenManager] = None
graph_client: Optional[GraphClient] = None
powerbi_client: Optional[PowerBIClient] = None

# Services
outlook: Optional[OutlookService] = None
onedrive: Optional[OneDriveService] = None
sharepoint: Optional[SharePointService] = None
excel: Optional[ExcelService] = None
office_docs: Optional[OfficeDocsService] = None
teams: Optional[TeamsService] = None
powerbi: Optional[PowerBIService] = None

# Device code flow state
_device_flow_state = {
    "active": False,
    "info": None,
    "result": None,
    "error": None,
}
_device_flow_lock = threading.Lock()


def initialize_services() -> None:
    """Initialize all M365 services."""
    global token_manager, graph_client, powerbi_client
    global outlook, onedrive, sharepoint, excel, office_docs, teams, powerbi
    
    logger.info("Initializing M365 services...")
    
    token_manager = TokenManager(config)
    graph_client = GraphClient(token_manager)
    powerbi_client = PowerBIClient(token_manager)
    
    outlook = OutlookService(graph_client)
    onedrive = OneDriveService(graph_client)
    sharepoint = SharePointService(graph_client)
    excel = ExcelService(graph_client)
    office_docs = OfficeDocsService(graph_client)
    teams = TeamsService(graph_client)
    powerbi = PowerBIService(powerbi_client)
    
    logger.info("M365 services initialized")


# =============================================================================
# MCP Server Setup
# =============================================================================

server = Server("m365-mcp-server")


@server.list_tools()
async def list_tools() -> list[Tool]:
    """List all available M365 tools."""
    return [
        # Authentication
        Tool(
            name="m365_auth_status",
            description="Check M365 authentication status",
            inputSchema={"type": "object", "properties": {}},
        ),
        
        # Outlook - Mail
        Tool(
            name="outlook_list_messages",
            description="List email messages from inbox or other folder",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder": {"type": "string", "default": "inbox", "description": "Mail folder (inbox, sentitems, drafts, deleteditems)"},
                    "count": {"type": "integer", "default": 25, "description": "Number of messages (max 50)"},
                    "search": {"type": "string", "description": "Search query"},
                },
            },
        ),
        Tool(
            name="outlook_get_message",
            description="Get a specific email message by ID",
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {"type": "string", "description": "Message ID"},
                },
                "required": ["message_id"],
            },
        ),
        Tool(
            name="outlook_send_message",
            description="Send an email message",
            inputSchema={
                "type": "object",
                "properties": {
                    "to": {"type": "array", "items": {"type": "string"}, "description": "Recipient email addresses"},
                    "subject": {"type": "string", "description": "Email subject"},
                    "body": {"type": "string", "description": "Email body"},
                    "cc": {"type": "array", "items": {"type": "string"}, "description": "CC recipients"},
                    "is_html": {"type": "boolean", "default": False, "description": "Whether body is HTML"},
                },
                "required": ["to", "subject", "body"],
            },
        ),
        Tool(
            name="outlook_reply_message",
            description="Reply to an email message",
            inputSchema={
                "type": "object",
                "properties": {
                    "message_id": {"type": "string", "description": "Message ID to reply to"},
                    "body": {"type": "string", "description": "Reply body"},
                    "reply_all": {"type": "boolean", "default": False, "description": "Reply to all recipients"},
                },
                "required": ["message_id", "body"],
            },
        ),
        
        # Outlook - Calendar
        Tool(
            name="outlook_list_events",
            description="List calendar events",
            inputSchema={
                "type": "object",
                "properties": {
                    "days_ahead": {"type": "integer", "default": 7, "description": "Days to look ahead"},
                    "days_back": {"type": "integer", "default": 0, "description": "Days to look back"},
                    "count": {"type": "integer", "default": 50, "description": "Max events"},
                },
            },
        ),
        Tool(
            name="outlook_create_event",
            description="Create a calendar event",
            inputSchema={
                "type": "object",
                "properties": {
                    "subject": {"type": "string", "description": "Event title"},
                    "start": {"type": "string", "description": "Start datetime (ISO format)"},
                    "end": {"type": "string", "description": "End datetime (ISO format)"},
                    "location": {"type": "string", "description": "Location"},
                    "body": {"type": "string", "description": "Event description"},
                    "attendees": {"type": "array", "items": {"type": "string"}, "description": "Attendee emails"},
                    "is_online_meeting": {"type": "boolean", "default": False, "description": "Create Teams meeting"},
                    "timezone": {"type": "string", "default": "UTC", "description": "Timezone"},
                },
                "required": ["subject", "start", "end"],
            },
        ),
        
        # OneDrive
        Tool(
            name="onedrive_list_files",
            description="List files in OneDrive root or folder",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "default": "/", "description": "Folder path"},
                    "count": {"type": "integer", "default": 50, "description": "Max items"},
                },
            },
        ),
        Tool(
            name="onedrive_get_file_content",
            description="Get content of a text file from OneDrive",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "File path"},
                },
                "required": ["path"],
            },
        ),
        Tool(
            name="onedrive_upload_file",
            description="Upload a file to OneDrive",
            inputSchema={
                "type": "object",
                "properties": {
                    "path": {"type": "string", "description": "Destination path"},
                    "content": {"type": "string", "description": "File content"},
                },
                "required": ["path", "content"],
            },
        ),
        
        # SharePoint
        Tool(
            name="sharepoint_list_sites",
            description="List SharePoint sites",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="sharepoint_list_files",
            description="List files in a SharePoint document library",
            inputSchema={
                "type": "object",
                "properties": {
                    "site_id": {"type": "string", "description": "Site ID"},
                    "drive_id": {"type": "string", "description": "Drive/library ID (optional)"},
                    "path": {"type": "string", "default": "/", "description": "Folder path"},
                },
                "required": ["site_id"],
            },
        ),
        
        # Teams
        Tool(
            name="teams_list_teams",
            description="List Teams you are a member of",
            inputSchema={"type": "object", "properties": {}},
        ),
        Tool(
            name="teams_list_channels",
            description="List channels in a Team",
            inputSchema={
                "type": "object",
                "properties": {
                    "team_id": {"type": "string", "description": "Team ID"},
                },
                "required": ["team_id"],
            },
        ),
        Tool(
            name="teams_send_message",
            description="Send a message to a Teams channel",
            inputSchema={
                "type": "object",
                "properties": {
                    "team_id": {"type": "string", "description": "Team ID"},
                    "channel_id": {"type": "string", "description": "Channel ID"},
                    "message": {"type": "string", "description": "Message content"},
                },
                "required": ["team_id", "channel_id", "message"],
            },
        ),
        
        # Excel
        Tool(
            name="excel_read_range",
            description="Read data from an Excel workbook range",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "OneDrive path to Excel file"},
                    "sheet": {"type": "string", "description": "Sheet name"},
                    "range": {"type": "string", "description": "Range (e.g., A1:D10)"},
                },
                "required": ["file_path", "sheet", "range"],
            },
        ),
        
        # Power BI
        Tool(
            name="powerbi_list_reports",
            description="List Power BI reports",
            inputSchema={"type": "object", "properties": {}},
        ),
    ]


@server.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    """Execute an M365 tool."""
    try:
        result = await _execute_tool(name, arguments)
        return [TextContent(type="text", text=str(result))]
    except AuthenticationRequiredError as e:
        return [TextContent(type="text", text=f"Authentication required: {e}")]
    except Exception as e:
        logger.exception(f"Tool {name} failed")
        return [TextContent(type="text", text=f"Error: {e}")]


async def _execute_tool(name: str, args: dict[str, Any]) -> Any:
    """Route tool calls to appropriate service methods."""
    
    # Auth status
    if name == "m365_auth_status":
        if token_manager and token_manager.get_graph_token():
            return {"authenticated": True, "message": "Valid token available"}
        return {"authenticated": False, "message": "No valid token"}
    
    # Outlook - Mail
    if name == "outlook_list_messages":
        return await outlook.list_messages(
            folder=args.get("folder", "inbox"),
            count=args.get("count", 25),
            search=args.get("search"),
        )
    
    if name == "outlook_get_message":
        return await outlook.get_message(args["message_id"])
    
    if name == "outlook_send_message":
        await outlook.send_message(
            to=args["to"],
            subject=args["subject"],
            body=args["body"],
            cc=args.get("cc"),
            is_html=args.get("is_html", False),
        )
        return {"success": True, "message": "Email sent"}
    
    if name == "outlook_reply_message":
        await outlook.reply_to_message(
            message_id=args["message_id"],
            body=args["body"],
            reply_all=args.get("reply_all", False),
        )
        return {"success": True, "message": "Reply sent"}
    
    # Outlook - Calendar
    if name == "outlook_list_events":
        return await outlook.list_events(
            days_ahead=args.get("days_ahead", 7),
            days_back=args.get("days_back", 0),
            count=args.get("count", 50),
        )
    
    if name == "outlook_create_event":
        return await outlook.create_event(
            subject=args["subject"],
            start=datetime.fromisoformat(args["start"]),
            end=datetime.fromisoformat(args["end"]),
            timezone=args.get("timezone", "UTC"),
            location=args.get("location"),
            body=args.get("body"),
            attendees=args.get("attendees"),
            is_online_meeting=args.get("is_online_meeting", False),
        )
    
    # OneDrive
    if name == "onedrive_list_files":
        return await onedrive.list_items(
            path=args.get("path", "/"),
            count=args.get("count", 50),
        )
    
    if name == "onedrive_get_file_content":
        content = await onedrive.download_file(args["path"])
        return content.decode("utf-8")
    
    if name == "onedrive_upload_file":
        return await onedrive.upload_file(
            path=args["path"],
            content=args["content"].encode("utf-8"),
        )
    
    # SharePoint
    if name == "sharepoint_list_sites":
        return await sharepoint.list_sites()
    
    if name == "sharepoint_list_files":
        return await sharepoint.list_items(
            site_id=args["site_id"],
            drive_id=args.get("drive_id"),
            path=args.get("path", "/"),
        )
    
    # Teams
    if name == "teams_list_teams":
        return await teams.list_teams()
    
    if name == "teams_list_channels":
        return await teams.list_channels(args["team_id"])
    
    if name == "teams_send_message":
        return await teams.send_channel_message(
            team_id=args["team_id"],
            channel_id=args["channel_id"],
            content=args["message"],
        )
    
    # Excel
    if name == "excel_read_range":
        return await excel.read_range(
            file_path=args["file_path"],
            sheet=args["sheet"],
            range_address=args["range"],
        )
    
    # Power BI
    if name == "powerbi_list_reports":
        return await powerbi.list_reports()
    
    raise ValueError(f"Unknown tool: {name}")


# =============================================================================
# Device Code Authentication Flow
# =============================================================================

def _run_device_code_flow_background():
    """Run device code flow in background thread."""
    global _device_flow_state
    
    def callback(info):
        with _device_flow_lock:
            _device_flow_state["info"] = info
        # Log to Railway logs
        logger.info("=" * 60)
        logger.info("DEVICE CODE AUTHENTICATION REQUIRED")
        logger.info("=" * 60)
        logger.info(info.get("message", ""))
        logger.info("=" * 60)
    
    def run():
        global _device_flow_state
        try:
            with _device_flow_lock:
                _device_flow_state["active"] = True
                _device_flow_state["error"] = None
            
            result = token_manager.authenticate_device_code(callback=callback)
            
            with _device_flow_lock:
                _device_flow_state["result"] = result
                _device_flow_state["active"] = False
            
            if "refresh_token" in result:
                logger.info("=" * 60)
                logger.info("AUTHENTICATION SUCCESSFUL!")
                logger.info("=" * 60)
                logger.info("Add this to Railway environment variables for persistence:")
                logger.info(f"M365_REFRESH_TOKEN={result['refresh_token']}")
                logger.info("=" * 60)
            else:
                error = result.get("error_description", result.get("error", "Unknown error"))
                logger.error(f"Device code auth failed: {error}")
                with _device_flow_lock:
                    _device_flow_state["error"] = error
                    
        except Exception as e:
            logger.exception("Device code flow error")
            with _device_flow_lock:
                _device_flow_state["error"] = str(e)
                _device_flow_state["active"] = False
    
    thread = threading.Thread(target=run, daemon=True)
    thread.start()


async def handle_device_code_start(request: Request) -> Response:
    """Start device code authentication flow.
    
    GET /auth/device-code - Start flow and return device code info
    """
    # Check if already authenticated
    if token_manager and token_manager.is_authenticated():
        return JSONResponse({
            "status": "authenticated",
            "message": "Already authenticated. Use /auth/logout to re-authenticate.",
            "account": token_manager.get_current_account(),
        })
    
    # Check if flow already active
    with _device_flow_lock:
        if _device_flow_state["active"]:
            return JSONResponse({
                "status": "pending",
                "message": "Device code flow already in progress",
                "device_code_info": _device_flow_state["info"],
            })
        
        # Check if we have a result from previous flow
        if _device_flow_state["result"]:
            if "access_token" in _device_flow_state["result"]:
                return JSONResponse({
                    "status": "completed",
                    "message": "Authentication completed. Refresh the page or call /auth/status.",
                })
    
    # Start new flow
    _run_device_code_flow_background()
    
    # Wait briefly for flow to initialize
    await asyncio.sleep(2)
    
    with _device_flow_lock:
        info = _device_flow_state["info"]
    
    if info:
        return JSONResponse({
            "status": "pending",
            "message": "Device code flow started. Follow the instructions below.",
            "device_code_info": info,
            "instructions": [
                f"1. Go to: {info.get('verification_uri', 'https://microsoft.com/devicelogin')}",
                f"2. Enter code: {info.get('user_code', 'N/A')}",
                "3. Sign in with your Microsoft account",
                "4. Return here to verify authentication",
            ],
        })
    
    return JSONResponse({
        "status": "starting",
        "message": "Device code flow starting... Refresh in a few seconds.",
    })


async def handle_auth_status(request: Request) -> Response:
    """Check authentication status.
    
    GET /auth/status
    """
    if token_manager and token_manager.is_authenticated():
        account = token_manager.get_current_account()
        return JSONResponse({
            "status": "authenticated",
            "account": {
                "username": account.get("username") if account else None,
            } if account else None,
        })
    
    with _device_flow_lock:
        if _device_flow_state["active"]:
            return JSONResponse({
                "status": "pending",
                "message": "Device code flow in progress",
                "device_code_info": _device_flow_state["info"],
            })
        
        if _device_flow_state["error"]:
            return JSONResponse({
                "status": "error",
                "error": _device_flow_state["error"],
            })
    
    return JSONResponse({
        "status": "not_authenticated",
        "message": "Not authenticated. Visit /auth/device-code to start authentication.",
    })


async def handle_auth_logout(request: Request) -> Response:
    """Logout and clear tokens.
    
    POST /auth/logout
    """
    global _device_flow_state
    
    if token_manager:
        token_manager.logout()
    
    with _device_flow_lock:
        _device_flow_state = {
            "active": False,
            "info": None,
            "result": None,
            "error": None,
        }
    
    return JSONResponse({
        "status": "logged_out",
        "message": "Tokens cleared. Visit /auth/device-code to re-authenticate.",
    })


async def handle_auth_refresh_token(request: Request) -> Response:
    """Get current refresh token for manual backup.
    
    GET /auth/refresh-token
    
    Returns the current refresh token so you can save it to M365_REFRESH_TOKEN.
    Requires bearer auth.
    """
    if not check_bearer_token(request):
        return JSONResponse({"error": "Unauthorized"}, status_code=401)
    
    if not token_manager:
        return JSONResponse({"error": "Token manager not initialized"}, status_code=500)
    
    refresh_token = token_manager.get_current_refresh_token()
    
    if refresh_token:
        return JSONResponse({
            "status": "success",
            "message": "Copy this to Railway M365_REFRESH_TOKEN env var",
            "refresh_token": refresh_token,
        })
    
    return JSONResponse({
        "status": "not_available",
        "message": "No refresh token available. Authenticate first via /auth/device-code",
    })


# =============================================================================
# HTTP Server with Streamable HTTP Transport (Modern) + Legacy SSE
# =============================================================================

def check_bearer_token(request: Request) -> bool:
    """Validate bearer token from Authorization header."""
    expected_token = os.environ.get("MCP_BEARER_TOKEN")
    if not expected_token:
        logger.warning("MCP_BEARER_TOKEN not configured")
        return False
    
    auth_header = request.headers.get("Authorization", "")
    if not auth_header.startswith("Bearer "):
        return False
    
    provided_token = auth_header[7:]  # Strip "Bearer " prefix
    return provided_token == expected_token


async def health_check(request: Request) -> Response:
    """Health check endpoint for Railway."""
    return JSONResponse({
        "status": "healthy",
        "service": "m365-mcp-server",
        "timestamp": datetime.utcnow().isoformat() + "Z",
    })


async def status_check(request: Request) -> Response:
    """Status endpoint with more detail."""
    auth_status = "authenticated" if (token_manager and token_manager.get_graph_token()) else "not_authenticated"
    return JSONResponse({
        "status": "running",
        "auth_status": auth_status,
        "services": ["outlook", "onedrive", "sharepoint", "excel", "teams", "powerbi"],
        "mcp_endpoints": {
            "streamable_http": "/mcp",
            "legacy_sse": "/sse",
            "legacy_messages": "/messages",
        },
        "auth_endpoints": {
            "device_code": "/auth/device-code",
            "status": "/auth/status",
            "logout": "/auth/logout",
            "refresh_token": "/auth/refresh-token",
        },
    })


# =============================================================================
# Streamable HTTP Transport (Modern MCP Standard)
# =============================================================================

# Session storage for Streamable HTTP
sessions: dict[str, dict] = {}


async def handle_mcp_request(request: Request) -> Response:
    """
    Handle MCP Streamable HTTP transport.
    
    This is the modern MCP transport that uses a single /mcp endpoint.
    - POST: Handle JSON-RPC requests
    - GET: Optional SSE stream for server-initiated messages
    """
    # Check auth for all MCP requests
    if not check_bearer_token(request):
        return JSONResponse({"error": "Unauthorized"}, status_code=401)
    
    if request.method == "GET":
        # SSE stream for server-initiated notifications (optional)
        return await handle_mcp_sse_stream(request)
    
    elif request.method == "POST":
        # JSON-RPC request handling
        return await handle_mcp_post(request)
    
    return JSONResponse({"error": "Method not allowed"}, status_code=405)


async def handle_mcp_post(request: Request) -> Response:
    """Handle POST requests to /mcp endpoint (JSON-RPC)."""
    try:
        body = await request.json()
    except Exception as e:
        return JSONResponse({
            "jsonrpc": "2.0",
            "error": {"code": -32700, "message": f"Parse error: {e}"},
            "id": None
        }, status_code=400)
    
    # Get or create session
    session_id = request.headers.get("Mcp-Session-Id", str(uuid.uuid4()))
    
    # Handle the JSON-RPC request
    method = body.get("method", "")
    params = body.get("params", {})
    request_id = body.get("id")
    
    logger.info(f"MCP request: {method} (session: {session_id})")
    
    try:
        result = await process_jsonrpc_method(method, params, session_id)
        
        response_data = {
            "jsonrpc": "2.0",
            "result": result,
            "id": request_id
        }
        
        response = JSONResponse(response_data)
        response.headers["Mcp-Session-Id"] = session_id
        return response
        
    except Exception as e:
        logger.exception(f"MCP method {method} failed")
        return JSONResponse({
            "jsonrpc": "2.0",
            "error": {"code": -32603, "message": str(e)},
            "id": request_id
        }, status_code=500)


async def process_jsonrpc_method(method: str, params: dict, session_id: str) -> Any:
    """Process a JSON-RPC method call."""
    
    if method == "initialize":
        # MCP initialization
        sessions[session_id] = {"initialized": True}
        return {
            "protocolVersion": "2024-11-05",
            "capabilities": {
                "tools": {"listChanged": False},
                "resources": {"subscribe": False, "listChanged": False},
                "prompts": {"listChanged": False},
            },
            "serverInfo": {
                "name": "m365-mcp-server",
                "version": "1.0.0"
            }
        }
    
    elif method == "notifications/initialized":
        # Client confirming initialization - no response needed
        return {}
    
    elif method == "tools/list":
        # List available tools
        tools = await list_tools()
        return {
            "tools": [
                {
                    "name": t.name,
                    "description": t.description,
                    "inputSchema": t.inputSchema
                }
                for t in tools
            ]
        }
    
    elif method == "tools/call":
        # Execute a tool
        tool_name = params.get("name", "")
        tool_args = params.get("arguments", {})
        
        result = await call_tool(tool_name, tool_args)
        
        return {
            "content": [
                {"type": r.type, "text": r.text}
                for r in result
            ]
        }
    
    elif method == "resources/list":
        # No resources exposed currently
        return {"resources": []}
    
    elif method == "prompts/list":
        # No prompts exposed currently
        return {"prompts": []}
    
    elif method == "ping":
        return {}
    
    else:
        raise ValueError(f"Unknown method: {method}")


async def handle_mcp_sse_stream(request: Request) -> Response:
    """Handle GET requests for SSE streaming (server-initiated messages)."""
    
    async def event_generator():
        # Send initial connection event
        yield f"event: endpoint\ndata: /mcp\n\n"
        
        # Keep connection alive with periodic pings
        while True:
            await asyncio.sleep(30)
            yield f"event: ping\ndata: {{}}\n\n"
    
    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        }
    )


# =============================================================================
# Legacy SSE Transport (Backward Compatibility)
# =============================================================================

# Store for legacy SSE sessions
legacy_sessions: dict[str, asyncio.Queue] = {}


async def handle_legacy_sse(request: Request) -> Response:
    """Handle legacy SSE connection (GET /sse)."""
    if not check_bearer_token(request):
        return JSONResponse({"error": "Unauthorized"}, status_code=401)
    
    session_id = str(uuid.uuid4())
    legacy_sessions[session_id] = asyncio.Queue()
    
    logger.info(f"New legacy SSE connection: {session_id}")
    
    async def event_generator():
        # Send the endpoint event with session info
        endpoint_data = json.dumps({"uri": f"/messages?session_id={session_id}"})
        yield f"event: endpoint\ndata: {endpoint_data}\n\n"
        
        try:
            while True:
                try:
                    # Wait for messages to send
                    message = await asyncio.wait_for(
                        legacy_sessions[session_id].get(),
                        timeout=30
                    )
                    yield f"event: message\ndata: {json.dumps(message)}\n\n"
                except asyncio.TimeoutError:
                    # Send keepalive
                    yield f": keepalive\n\n"
        finally:
            legacy_sessions.pop(session_id, None)
    
    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        }
    )


async def handle_legacy_messages(request: Request) -> Response:
    """Handle legacy message POST (POST /messages)."""
    if not check_bearer_token(request):
        return JSONResponse({"error": "Unauthorized"}, status_code=401)
    
    session_id = request.query_params.get("session_id")
    if not session_id or session_id not in legacy_sessions:
        return JSONResponse({"error": "Invalid session"}, status_code=400)
    
    try:
        body = await request.json()
    except Exception as e:
        return JSONResponse({
            "jsonrpc": "2.0",
            "error": {"code": -32700, "message": f"Parse error: {e}"},
            "id": None
        }, status_code=400)
    
    method = body.get("method", "")
    params = body.get("params", {})
    request_id = body.get("id")
    
    logger.info(f"Legacy MCP request: {method}")
    
    try:
        result = await process_jsonrpc_method(method, params, session_id)
        
        response_data = {
            "jsonrpc": "2.0",
            "result": result,
            "id": request_id
        }
        
        # Queue response for SSE stream
        await legacy_sessions[session_id].put(response_data)
        
        return JSONResponse({"status": "accepted"}, status_code=202)
        
    except Exception as e:
        logger.exception(f"Legacy MCP method {method} failed")
        error_response = {
            "jsonrpc": "2.0",
            "error": {"code": -32603, "message": str(e)},
            "id": request_id
        }
        await legacy_sessions[session_id].put(error_response)
        return JSONResponse({"status": "error"}, status_code=500)


# =============================================================================
# App Factory
# =============================================================================

def create_http_app() -> Starlette:
    """Create Starlette app with health endpoints and MCP transports."""
    routes = [
        # Health/status endpoints (no auth required)
        Route("/health", health_check, methods=["GET"]),
        Route("/status", status_check, methods=["GET"]),
        Route("/", health_check, methods=["GET"]),
        
        # Auth endpoints (no bearer auth required for device code flow)
        Route("/auth/device-code", handle_device_code_start, methods=["GET"]),
        Route("/auth/status", handle_auth_status, methods=["GET"]),
        Route("/auth/logout", handle_auth_logout, methods=["POST"]),
        Route("/auth/refresh-token", handle_auth_refresh_token, methods=["GET"]),
        
        # Streamable HTTP transport (modern - recommended)
        Route("/mcp", handle_mcp_request, methods=["GET", "POST"]),
        
        # Legacy SSE transport (backward compatibility)
        Route("/sse", handle_legacy_sse, methods=["GET"]),
        Route("/messages", handle_legacy_messages, methods=["POST"]),
    ]
    return Starlette(routes=routes)


# =============================================================================
# Main Entry Points
# =============================================================================

async def run_stdio() -> None:
    """Run MCP server with stdio transport (for local use)."""
    initialize_services()
    async with stdio_server() as (read_stream, write_stream):
        await server.run(
            read_stream,
            write_stream,
            server.create_initialization_options(),
        )


def run_http() -> None:
    """Run HTTP server for Railway deployment."""
    initialize_services()
    
    port = int(os.environ.get("PORT", 8000))
    host = os.environ.get("HOST", "0.0.0.0")
    
    # Check for bearer token configuration
    if os.environ.get("MCP_BEARER_TOKEN"):
        logger.info("MCP_BEARER_TOKEN configured - MCP endpoints protected")
    else:
        logger.warning("MCP_BEARER_TOKEN not set - MCP endpoints will reject all requests")
    
    logger.info(f"Starting HTTP server on {host}:{port}")
    logger.info(f"Streamable HTTP endpoint: /mcp (modern)")
    logger.info(f"Legacy SSE endpoint: /sse")
    logger.info(f"Legacy messages endpoint: /messages")
    logger.info(f"Auth endpoint: /auth/device-code")
    
    app = create_http_app()
    uvicorn.run(app, host=host, port=port, log_level="info")


if __name__ == "__main__":
    # Check if running in Railway (has PORT env var) or locally
    if os.environ.get("PORT") or os.environ.get("RAILWAY_ENVIRONMENT"):
        logger.info("Railway environment detected, starting HTTP server")
        run_http()
    else:
        logger.info("Local environment detected, starting stdio server")
        asyncio.run(run_stdio())
