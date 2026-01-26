"""M365 MCP Server entry point.

Runs the MCP server with HTTP/SSE transport for Railway deployment.
Supports bearer token authentication via MCP_BEARER_TOKEN env var.
Exposes /health endpoint for healthchecks.
"""

import asyncio
import logging
import os
import sys
from datetime import datetime
from typing import Any, Optional

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.server.sse import SseServerTransport
from mcp.types import Tool, TextContent

from starlette.applications import Starlette
from starlette.responses import JSONResponse, Response
from starlette.routing import Route, Mount
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
# HTTP Server for Railway (with health endpoint and MCP SSE transport)
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
            "sse": "/sse",
            "messages": "/messages",
        },
    })


# SSE transport for MCP
sse_transport = SseServerTransport("/messages")


async def handle_sse(request: Request) -> Response:
    """Handle SSE connection for MCP with bearer token auth."""
    if not check_bearer_token(request):
        return JSONResponse({"error": "Unauthorized"}, status_code=401)
    
    logger.info("New MCP SSE connection established")
    
    async with sse_transport.connect_sse(
        request.scope, request.receive, request._send
    ) as streams:
        await server.run(
            streams[0],
            streams[1],
            server.create_initialization_options(),
        )
    
    return Response()


async def handle_messages(request: Request) -> Response:
    """Handle MCP messages with bearer token auth."""
    if not check_bearer_token(request):
        return JSONResponse({"error": "Unauthorized"}, status_code=401)
    
    return await sse_transport.handle_post_message(request.scope, request.receive, request._send)


def create_http_app() -> Starlette:
    """Create Starlette app with health endpoints and MCP SSE transport."""
    routes = [
        # Health/status endpoints (no auth required)
        Route("/health", health_check, methods=["GET"]),
        Route("/status", status_check, methods=["GET"]),
        Route("/", health_check, methods=["GET"]),
        
        # MCP SSE transport endpoints (bearer auth required)
        Route("/sse", handle_sse, methods=["GET"]),
        Route("/messages", handle_messages, methods=["POST"]),
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
    logger.info(f"MCP SSE endpoint: /sse")
    logger.info(f"MCP messages endpoint: /messages")
    
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
