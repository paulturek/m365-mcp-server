"""Configuration management for M365 MCP Server.

This module handles all configuration loading from environment variables
with sensible defaults. Configuration is validated on startup.

Environment Variables:
    M365_CLIENT_ID: Azure AD Application (client) ID (required)
    M365_TENANT_ID: Azure AD tenant ID (default: 'common')
    M365_CLIENT_SECRET: Client secret for confidential client flow (optional)
    M365_TOKEN_CACHE_PATH: Custom token cache location (optional)
    M365_CACHE_ENCRYPTION_KEY: Custom encryption key for cache (optional)
"""

import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

# Load .env file if present
load_dotenv()


@dataclass
class M365Config:
    """Configuration container for M365 MCP Server.
    
    Attributes:
        client_id: Azure AD Application (client) ID
        tenant_id: Azure AD tenant ID ('common' for multi-tenant)
        client_secret: Optional client secret for confidential client
        cache_path: Path to encrypted token cache file
        cache_encryption_key: Optional custom encryption key
        graph_scopes: OAuth scopes for Microsoft Graph API
        powerbi_scopes: OAuth scopes for Power BI API
    """
    
    # Azure AD App Registration
    client_id: str = field(
        default_factory=lambda: os.getenv("M365_CLIENT_ID", "")
    )
    tenant_id: str = field(
        default_factory=lambda: os.getenv("M365_TENANT_ID", "common")
    )
    client_secret: Optional[str] = field(
        default_factory=lambda: os.getenv("M365_CLIENT_SECRET")
    )
    
    # Token cache configuration
    cache_path: str = field(
        default_factory=lambda: os.getenv(
            "M365_TOKEN_CACHE_PATH",
            str(Path.home() / ".m365-mcp" / "token_cache.bin")
        )
    )
    cache_encryption_key: Optional[str] = field(
        default_factory=lambda: os.getenv("M365_CACHE_ENCRYPTION_KEY")
    )
    
    # Microsoft Graph API scopes (delegated permissions)
    graph_scopes: list[str] = field(default_factory=lambda: [
        # User profile
        "User.Read",
        
        # Outlook Mail
        "Mail.ReadWrite",
        "Mail.Send",
        
        # Calendar
        "Calendars.ReadWrite",
        
        # OneDrive & SharePoint files
        "Files.ReadWrite.All",
        
        # SharePoint sites and lists
        "Sites.ReadWrite.All",
        
        # Teams
        "Team.ReadBasic.All",
        "Channel.ReadBasic.All",
        "Chat.ReadWrite",
        "ChannelMessage.Send",
        
        # Required for refresh tokens
        "offline_access",
    ])
    

    
    def validate(self) -> None:
        """Validate required configuration values.
        
        Raises:
            ValueError: If required configuration is missing.
        """
        if not self.client_id:
            raise ValueError(
                "M365_CLIENT_ID environment variable is required. "
                "Get this from Azure Portal > App registrations."
            )
    
    @property
    def authority(self) -> str:
        """Get the Azure AD authority URL."""
        return f"https://login.microsoftonline.com/{self.tenant_id}"
    
    @property
    def is_confidential_client(self) -> bool:
        """Check if configured for confidential client flow."""
        return self.client_secret is not None


# Global configuration instance
config = M365Config()
