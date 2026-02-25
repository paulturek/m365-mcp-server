"""Excel MCP tools.

Operations on cloud-hosted Excel workbooks via Microsoft Graph.
"""
import logging

from ..auth.oauth_web import get_access_token
from ..clients.graph_client import GraphClient

logger = logging.getLogger(__name__)

_USER_ID_PROP = {
    "user_id": {
        "type": "string",
        "description": "Your user identifier (email recommended)",
    }
}

TOOLS = [
    {
        "name": "excel_get_workbook_info",
        "description": "Get metadata about an Excel workbook (worksheets, named ranges).",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {
                    "type": "string",
                    "description": "OneDrive path to the workbook",
                },
                "item_id": {
                    "type": "string",
                    "description": "OneDrive item ID (alternative)",
                },
            },
            "required": ["user_id"],
        },
    },
    {
        "name": "excel_read_range",
        "description": "Read data from a range in an Excel worksheet.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {"type": "string", "description": "OneDrive path to workbook"},
                "item_id": {"type": "string", "description": "OneDrive item ID (alternative)"},
                "worksheet": {"type": "string", "description": "Worksheet name", "default": "Sheet1"},
                "range": {"type": "string", "description": "Cell range (e.g. 'A1:D10')"},
            },
            "required": ["user_id", "range"],
        },
    },
    {
        "name": "excel_write_range",
        "description": "Write data to a range in an Excel worksheet.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {"type": "string", "description": "OneDrive path to workbook"},
                "item_id": {"type": "string", "description": "OneDrive item ID (alternative)"},
                "worksheet": {"type": "string", "description": "Worksheet name", "default": "Sheet1"},
                "range": {"type": "string", "description": "Target range (e.g. 'A1:C3')"},
                "values": {
                    "type": "array",
                    "description": "2D array of values to write",
                    "items": {"type": "array", "items": {}},
                },
            },
            "required": ["user_id", "range", "values"],
        },
    },
    {
        "name": "excel_create_chart",
        "description": "Create a chart in an Excel worksheet from a data range.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {"type": "string", "description": "OneDrive path to workbook"},
                "item_id": {"type": "string", "description": "OneDrive item ID (alternative)"},
                "worksheet": {"type": "string", "description": "Worksheet name", "default": "Sheet1"},
                "chart_type": {
                    "type": "string",
                    "description": "Chart type (e.g. 'ColumnClustered', 'Pie', 'Line')",
                    "default": "ColumnClustered",
                },
                "source_range": {"type": "string", "description": "Data range for the chart"},
                "chart_name": {"type": "string", "description": "Display name for the chart"},
            },
            "required": ["user_id", "source_range"],
        },
    },
]


def _workbook_base(params: dict) -> str:
    if params.get("item_id"):
        return f"/me/drive/items/{params['item_id']}/workbook"
    path = params.get("item_path", "").strip("/")
    return f"/me/drive/root:/{path}:/workbook"


async def _get_workbook_info(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    base = _workbook_base(params)
    worksheets = await client.get(f"{base}/worksheets")
    names = await client.get(f"{base}/names")
    return {
        "worksheets": [
            {"name": ws.get("name"), "id": ws.get("id")}
            for ws in worksheets.get("value", [])
        ],
        "namedRanges": [
            {"name": n.get("name"), "value": n.get("value")}
            for n in names.get("value", [])
        ],
    }


async def _read_range(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    base = _workbook_base(params)
    ws = params.get("worksheet", "Sheet1")
    rng = params["range"]
    data = await client.get(f"{base}/worksheets/{ws}/range(address='{rng}')")
    return {
        "address": data.get("address"),
        "rowCount": data.get("rowCount"),
        "columnCount": data.get("columnCount"),
        "values": data.get("values", []),
    }


async def _write_range(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    base = _workbook_base(params)
    ws = params.get("worksheet", "Sheet1")
    rng = params["range"]
    body = {"values": params["values"]}
    data = await client.patch(f"{base}/worksheets/{ws}/range(address='{rng}')", json=body)
    return {
        "address": data.get("address"),
        "rowCount": data.get("rowCount"),
        "columnCount": data.get("columnCount"),
    }


async def _create_chart(params: dict) -> dict:
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    base = _workbook_base(params)
    ws = params.get("worksheet", "Sheet1")
    body = {
        "type": params.get("chart_type", "ColumnClustered"),
        "sourceData": params["source_range"],
        "seriesBy": "Auto",
    }
    result = await client.post(f"{base}/worksheets/{ws}/charts/add", json=body)
    chart_name = params.get("chart_name")
    if chart_name and result.get("name"):
        await client.patch(
            f"{base}/worksheets/{ws}/charts/{result['name']}",
            json={"name": chart_name},
        )
    return {
        "name": chart_name or result.get("name"),
        "id": result.get("id"),
    }


HANDLERS = {
    "excel_get_workbook_info": _get_workbook_info,
    "excel_read_range": _read_range,
    "excel_write_range": _write_range,
    "excel_create_chart": _create_chart,
}
