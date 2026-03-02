"""HTTP clients for Microsoft APIs.

Provides async HTTP clients for:
- Microsoft Graph API (graph.microsoft.com)
"""

from .graph_client import GraphClient, GraphAPIError, AuthenticationRequiredError

__all__ = [
    "GraphClient",
    "GraphAPIError",
    "AuthenticationRequiredError",
]
