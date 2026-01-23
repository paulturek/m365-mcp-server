"""Authentication module for M365 MCP Server.

Provides MSAL-based OAuth 2.0 authentication with:
- Device code flow for interactive authentication
- Automatic token refresh
- Encrypted persistent token cache
"""

from .token_cache import EncryptedTokenCache
from .token_manager import TokenManager

__all__ = ["EncryptedTokenCache", "TokenManager"]
