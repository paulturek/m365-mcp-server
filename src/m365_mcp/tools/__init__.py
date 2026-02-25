"""M365 MCP Tools package.

Each sub-module defines:
  TOOLS  — list of MCP tool schema dicts
  HANDLERS — dict mapping tool name → async handler(params) -> dict

This __init__ merges them into two flat registries consumed
by __main__.py's MCP JSON-RPC dispatcher.
"""
from typing import Callable, Awaitable

from . import (
    onedrive,
    excel,
    outlook,
    sharepoint,
    teams,
    todo,
    users,
    office_docs,
)

_MODULES = [
    onedrive,
    excel,
    outlook,
    sharepoint,
    teams,
    todo,
    users,
    office_docs,
]

# Flat list of every tool schema dict
TOOL_REGISTRY: list[dict] = []
for mod in _MODULES:
    TOOL_REGISTRY.extend(getattr(mod, "TOOLS", []))

# tool_name → async handler(params) -> dict
TOOL_HANDLERS: dict[str, Callable[..., Awaitable[dict]]] = {}
for mod in _MODULES:
    TOOL_HANDLERS.update(getattr(mod, "HANDLERS", {}))

__all__ = ["TOOL_REGISTRY", "TOOL_HANDLERS"]
