"""Tool registry — auto-collects all TOOL_DEFINITIONS and TOOL_HANDLERS
from sub-modules in the tools/ package.

Each tool module exports:
  TOOL_DEFINITIONS: list[dict]   — MCP tool schemas
  TOOL_HANDLERS:    dict[str, Callable]  — name → async handler

Adding a new tool module:
  1. Create tools/<module>.py with TOOL_DEFINITIONS + TOOL_HANDLERS
  2. Add the module name to _TOOL_MODULES below
  3. Done — it will be registered automatically at startup
"""

import importlib
import logging

logger = logging.getLogger("m365_mcp.tools")

# ---------------------------------------------------------------------------
# Modules to load (order doesn't matter)
# ---------------------------------------------------------------------------
_TOOL_MODULES = [
    "onedrive",
    "excel",
    "outlook",
    "sharepoint",
    "teams",
    "todo",
    "users",
    "office_docs",
    "auth",
]

# ---------------------------------------------------------------------------
# Collected registries (populated at import time)
# ---------------------------------------------------------------------------
TOOL_REGISTRY: list[dict] = []
TOOL_HANDLERS: dict = {}

for _mod_name in _TOOL_MODULES:
    try:
        _mod = importlib.import_module(f".{_mod_name}", package=__name__)

        _defs = getattr(_mod, "TOOL_DEFINITIONS", [])
        _handlers = getattr(_mod, "TOOL_HANDLERS", {})

        TOOL_REGISTRY.extend(_defs)
        TOOL_HANDLERS.update(_handlers)

        logger.info(
            "Loaded %d tools from %s: %s",
            len(_defs),
            _mod_name,
            ", ".join(t["name"] for t in _defs),
        )
    except Exception as exc:
        logger.error("Failed to load tool module %s: %s", _mod_name, exc)

logger.info("Total tools registered: %d", len(TOOL_REGISTRY))
