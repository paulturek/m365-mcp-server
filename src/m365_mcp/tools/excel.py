"""Excel MCP tools.

Operations on cloud-hosted Excel workbooks via Microsoft Graph.
Covers: workbook info, read range, write range, create chart,
        add table rows, create worksheet, delete worksheet.
"""
import logging
from urllib.parse import quote

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
                    "description": "OneDrive path to the workbook (e.g. '/Documents/Budget 2026.xlsx')",
                },
                "item_id": {
                    "type": "string",
                    "description": "OneDrive item ID (alternative to item_path)",
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
    {
        "name": "excel_add_table_rows",
        "description": "Add rows to an existing table in an Excel workbook.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {"type": "string", "description": "OneDrive path to workbook"},
                "item_id": {"type": "string", "description": "OneDrive item ID (alternative)"},
                "table_name": {
                    "type": "string",
                    "description": "Table name or ID in the workbook",
                },
                "values": {
                    "type": "array",
                    "description": "2D array of row values to append",
                    "items": {"type": "array", "items": {}},
                },
            },
            "required": ["user_id", "table_name", "values"],
        },
    },
    {
        "name": "excel_create_worksheet",
        "description": "Create a new worksheet in an Excel workbook.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {"type": "string", "description": "OneDrive path to workbook"},
                "item_id": {"type": "string", "description": "OneDrive item ID (alternative)"},
                "name": {
                    "type": "string",
                    "description": "Name for the new worksheet",
                },
            },
            "required": ["user_id", "name"],
        },
    },
    {
        "name": "excel_delete_worksheet",
        "description": "Delete a worksheet from an Excel workbook.",
        "inputSchema": {
            "type": "object",
            "properties": {
                **_USER_ID_PROP,
                "item_path": {"type": "string", "description": "OneDrive path to workbook"},
                "item_id": {"type": "string", "description": "OneDrive item ID (alternative)"},
                "worksheet": {
                    "type": "string",
                    "description": "Worksheet name to delete",
                },
            },
            "required": ["user_id", "worksheet"],
        },
    },
]


def _workbook_base(params: dict) -> str:
    """Build the workbook base URL, URL-encoding the path to handle spaces and special chars."""
    if params.get("item_id"):
        return f"/me/drive/items/{params['item_id']}/workbook"
    path = quote(params.get("item_path", "").strip("/"), safe="/")
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
    data = await client.patch(
        f"{base}/worksheets/{ws}/range(address='{rng}')",
        json={"values": params["values"]},
    )
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


async def _add_table_rows(params: dict) -> dict:
    """POST /workbook/tables/{name}/rows — append rows to a table."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    base = _workbook_base(params)
    result = await client.post(
        f"{base}/tables/{params['table_name']}/rows",
        json={"values": params["values"]},
    )
    return {
        "index": result.get("index"),
        "values": result.get("values"),
    }


async def _create_worksheet(params: dict) -> dict:
    """POST /workbook/worksheets — create a new worksheet."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    base = _workbook_base(params)
    result = await client.post(f"{base}/worksheets", json={"name": params["name"]})
    return {
        "name": result.get("name"),
        "id": result.get("id"),
        "position": result.get("position"),
    }


async def _delete_worksheet(params: dict) -> dict:
    """DELETE /workbook/worksheets/{name} — delete a worksheet."""
    token = await get_access_token(params["user_id"])
    client = GraphClient(token)
    base = _workbook_base(params)
    ws = params["worksheet"]
    await client.delete(f"{base}/worksheets/{ws}")
    return {"deleted": True, "worksheet": ws}


HANDLERS = {
    "excel_get_workbook_info": _get_workbook_info,
    "excel_read_range": _read_range,
    "excel_write_range": _write_range,
    "excel_create_chart": _create_chart,
    "excel_add_table_rows": _add_table_rows,
    "excel_create_worksheet": _create_worksheet,
    "excel_delete_worksheet": _delete_worksheet,
}
