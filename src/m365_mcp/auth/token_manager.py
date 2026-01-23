"""MSAL-based token manager with automatic refresh.

This module provides the TokenManager class that handles:
- Device code flow authentication
- Silent token acquisition with automatic refresh
- Separate token management for Graph API and Power BI API

Token Lifecycle:
    1. User authenticates via device code flow
    2. Access token (~1 hour) and refresh token (~90 days) cached
    3. On subsequent calls, MSAL automatically refreshes if needed
    4. Refresh tokens extended on each use

"""

import logging
from typing import Any, Callable, Optional

import msal

from ..config import M365Config
from .token_cache import EncryptedTokenCache

logger = logging.getLogger(__name__)


class TokenManager:
    """Manages M365 authentication using MSAL.
    
    Handles OAuth 2.0 authentication flows and automatic token refresh.
    Supports both Microsoft Graph API and Power BI API tokens.
    
    Attributes:
        config: M365 configuration instance
        cache: Encrypted token cache
        app: MSAL application instance
    
    Example:
        >>> config = M365Config()
        >>> manager = TokenManager(config)
        >>> 
        >>> # Check if authenticated
        >>> if not manager.is_authenticated():
        ...     manager.authenticate_device_code()
        >>> 
        >>> # Get token (auto-refreshes if needed)
        >>> token = manager.get_graph_token()
    """
    
    def __init__(self, config: M365Config) -> None:
        """Initialize token manager.
        
        Args:
            config: M365 configuration instance
            
        Raises:
            ValueError: If required configuration is missing
        """
        self.config = config
        config.validate()
        
        # Initialize encrypted cache
        self.cache = EncryptedTokenCache(
            cache_path=config.cache_path,
            encryption_key=config.cache_encryption_key
        )
        
        # Create MSAL application
        # Use ConfidentialClientApplication if client_secret is provided
        if config.is_confidential_client:
            self.app: msal.ClientApplication = msal.ConfidentialClientApplication(
                client_id=config.client_id,
                authority=config.authority,
                client_credential=config.client_secret,
                token_cache=self.cache,
            )
            logger.info("Initialized confidential client application")
        else:
            self.app = msal.PublicClientApplication(
                client_id=config.client_id,
                authority=config.authority,
                token_cache=self.cache,
            )
            logger.info("Initialized public client application")
        
        logger.info(f"TokenManager initialized for tenant: {config.tenant_id}")
    
    def get_graph_token(self) -> Optional[str]:
        """Get valid access token for Microsoft Graph API.
        
        MSAL automatically handles token refresh if the current token
        is expired or about to expire (~5 minutes before expiry).
        
        Returns:
            Access token string, or None if authentication required
        """
        return self._acquire_token_silent(self.config.graph_scopes)
    
    def get_powerbi_token(self) -> Optional[str]:
        """Get valid access token for Power BI API.
        
        Power BI uses a separate API endpoint and scope from Graph.
        
        Returns:
            Access token string, or None if authentication required
        """
        return self._acquire_token_silent(self.config.powerbi_scopes)
    
    def _acquire_token_silent(
        self,
        scopes: list[str]
    ) -> Optional[str]:
        """Attempt silent token acquisition.
        
        Tries to get token from cache or via refresh token.
        MSAL handles the refresh logic internally.
        
        Args:
            scopes: OAuth scopes to request
            
        Returns:
            Access token string, or None if interactive auth needed
        """
        accounts = self.app.get_accounts()
        
        if not accounts:
            logger.debug("No cached accounts found")
            return None
        
        # Try silent acquisition (uses refresh token if needed)
        result = self.app.acquire_token_silent(
            scopes=scopes,
            account=accounts[0],
        )
        
        if result:
            if "access_token" in result:
                logger.debug("Token acquired silently")
                return result["access_token"]
            
            if "error" in result:
                logger.warning(
                    f"Silent auth failed: {result.get('error_description', result['error'])}"
                )
        
        return None
    
    def authenticate_device_code(
        self,
        scopes: Optional[list[str]] = None,
        callback: Optional[Callable[[dict[str, Any]], None]] = None,
    ) -> dict[str, Any]:
        """Authenticate using device code flow.
        
        Device code flow is ideal for CLI/headless environments where
        you can't open a browser directly. User visits a URL and enters
        a code to complete authentication.
        
        Args:
            scopes: OAuth scopes to request (defaults to graph_scopes)
            callback: Optional function to receive device code info.
                     Called with dict containing 'user_code', 
                     'verification_uri', and 'message'.
        
        Returns:
            MSAL token result dict. Contains 'access_token' on success,
            or 'error' and 'error_description' on failure.
        
        Example:
            >>> def show_code(info):
            ...     print(f"Go to {info['verification_uri']}")
            ...     print(f"Enter code: {info['user_code']}")
            >>> result = manager.authenticate_device_code(callback=show_code)
        """
        scopes = scopes or self.config.graph_scopes
        
        # Initiate device code flow
        flow = self.app.initiate_device_flow(scopes=scopes)
        
        if "user_code" not in flow:
            error_msg = flow.get("error_description", "Unknown error")
            logger.error(f"Failed to initiate device flow: {error_msg}")
            return {
                "error": "device_flow_failed",
                "error_description": error_msg
            }
        
        # Notify caller with device code info
        if callback:
            callback({
                "user_code": flow["user_code"],
                "verification_uri": flow["verification_uri"],
                "message": flow["message"],
                "expires_in": flow.get("expires_in", 900),
            })
        else:
            # Default: print to console
            print(f"\n{flow['message']}\n")
        
        # Block until user completes authentication (or timeout)
        result = self.app.acquire_token_by_device_flow(flow)
        
        if "access_token" in result:
            logger.info("Device code authentication successful")
        else:
            logger.error(
                f"Device code auth failed: "
                f"{result.get('error_description', result.get('error'))}"
            )
        
        return result
    
    def is_authenticated(self, for_powerbi: bool = False) -> bool:
        """Check if valid tokens are available.
        
        Args:
            for_powerbi: If True, check Power BI token instead of Graph
            
        Returns:
            True if valid (or refreshable) tokens exist
        """
        if for_powerbi:
            return self.get_powerbi_token() is not None
        return self.get_graph_token() is not None
    
    def get_current_account(self) -> Optional[dict[str, Any]]:
        """Get info about the currently authenticated account.
        
        Returns:
            Account dict with 'username', 'name', etc., or None
        """
        accounts = self.app.get_accounts()
        return accounts[0] if accounts else None
    
    def logout(self) -> None:
        """Clear all cached tokens and accounts."""
        accounts = self.app.get_accounts()
        for account in accounts:
            self.app.remove_account(account)
        
        self.cache.clear()
        logger.info("Logged out and cleared token cache")
