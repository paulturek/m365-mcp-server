"""MSAL-based token manager with automatic refresh.

This module provides the TokenManager class that handles:
- Device code flow authentication
- Silent token acquisition with automatic refresh
- Separate token management for Graph API and Power BI API
- Environment variable-based refresh token for Railway deployment

Token Lifecycle:
    1. User authenticates via device code flow
    2. Access token (~1 hour) and refresh token (~90 days) cached
    3. On subsequent calls, MSAL automatically refreshes if needed
    4. Refresh tokens extended on each use

Railway Deployment:
    Set M365_REFRESH_TOKEN env var to persist auth across deploys.
    The refresh token is used to bootstrap the MSAL cache on startup.

Note on MSAL Application Types:
    - ConfidentialClientApplication: Has client_secret, used for token refresh
    - PublicClientApplication: No client_secret, required for device code flow
    
    Device code flow ONLY works with PublicClientApplication, so we maintain
    both types when client_secret is provided.

Note on Scopes:
    MSAL automatically handles reserved scopes (offline_access, openid, profile).
    Do NOT include these in scope lists passed to initiate_device_flow() or
    acquire_token methods - MSAL will raise ValueError.

"""

import logging
import os
from typing import Any, Callable, Optional

import msal

from ..config import M365Config
from .token_cache import EncryptedTokenCache

logger = logging.getLogger(__name__)

# Scopes that MSAL handles automatically - never include in requests
RESERVED_SCOPES = {"offline_access", "openid", "profile"}


class TokenManager:
    """Manages M365 authentication using MSAL.
    
    Handles OAuth 2.0 authentication flows and automatic token refresh.
    Supports both Microsoft Graph API and Power BI API tokens.
    
    Supports M365_REFRESH_TOKEN env var for Railway deployment where
    the token cache is ephemeral.
    
    Attributes:
        config: M365 configuration instance
        cache: Encrypted token cache
        app: MSAL application instance (ConfidentialClient if secret provided)
        public_app: MSAL PublicClientApplication for device code flow
    
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
        
        # Always create a PublicClientApplication for device code flow
        # Device code flow ONLY works with PublicClientApplication
        self.public_app = msal.PublicClientApplication(
            client_id=config.client_id,
            authority=config.authority,
            token_cache=self.cache,
        )
        logger.info("Initialized public client application (for device code flow)")
        
        # Create MSAL application for token operations
        # Use ConfidentialClientApplication if client_secret is provided
        if config.is_confidential_client:
            self.app: msal.ClientApplication = msal.ConfidentialClientApplication(
                client_id=config.client_id,
                authority=config.authority,
                client_credential=config.client_secret,
                token_cache=self.cache,
            )
            logger.info("Initialized confidential client application (for token refresh)")
        else:
            # No client secret - use public app for everything
            self.app = self.public_app
            logger.info("Using public client application (no client secret)")
        
        logger.info(f"TokenManager initialized for tenant: {config.tenant_id}")
        
        # Bootstrap from M365_REFRESH_TOKEN env var if available
        # This enables token persistence across Railway deploys
        self._bootstrap_from_env_refresh_token()
    
    @staticmethod
    def _filter_reserved_scopes(scopes: list[str]) -> list[str]:
        """Remove reserved scopes that MSAL handles automatically.
        
        MSAL raises ValueError if you include offline_access, openid, or profile
        in scope lists. These are handled internally by MSAL.
        
        Args:
            scopes: List of OAuth scopes
            
        Returns:
            Filtered list without reserved scopes
        """
        filtered = [s for s in scopes if s.lower() not in RESERVED_SCOPES]
        removed = set(scopes) - set(filtered)
        if removed:
            logger.debug(f"Filtered reserved scopes: {removed}")
        return filtered
    
    def _bootstrap_from_env_refresh_token(self) -> None:
        """Bootstrap MSAL cache from M365_REFRESH_TOKEN env var.
        
        This is critical for Railway deployment where the container
        filesystem is ephemeral. The refresh token from the env var
        is used to acquire a fresh access token and populate the cache.
        """
        refresh_token = os.environ.get("M365_REFRESH_TOKEN")
        
        if not refresh_token:
            logger.debug("M365_REFRESH_TOKEN not set, skipping bootstrap")
            return
        
        # Check if we already have valid cached accounts
        accounts = self.app.get_accounts()
        if accounts:
            logger.info("Existing cached accounts found, skipping refresh token bootstrap")
            return
        
        logger.info("Bootstrapping from M365_REFRESH_TOKEN environment variable...")
        
        # Filter out reserved scopes before making the request
        scopes = self._filter_reserved_scopes(self.config.graph_scopes)
        
        # Use the refresh token to acquire new tokens
        # This populates the MSAL cache with the account and tokens
        try:
            result = self.app.acquire_token_by_refresh_token(
                refresh_token=refresh_token,
                scopes=scopes,
            )
            
            if "access_token" in result:
                logger.info("Successfully bootstrapped auth from M365_REFRESH_TOKEN")
                
                # Log the new refresh token if it changed (for manual update)
                new_refresh_token = result.get("refresh_token")
                if new_refresh_token and new_refresh_token != refresh_token:
                    logger.info(
                        "Refresh token was rotated. Consider updating M365_REFRESH_TOKEN "
                        "env var with the new value for extended validity."
                    )
                    # Log first/last few chars for identification (not the full token)
                    logger.debug(
                        f"New refresh token: {new_refresh_token[:10]}...{new_refresh_token[-10:]}"
                    )
            else:
                error = result.get("error", "unknown_error")
                error_desc = result.get("error_description", "No description")
                logger.error(f"Failed to bootstrap from refresh token: {error} - {error_desc}")
                
                if "invalid_grant" in error or "expired" in error_desc.lower():
                    logger.error(
                        "The M365_REFRESH_TOKEN has expired or been revoked. "
                        "You need to re-authenticate and obtain a new refresh token."
                    )
                    
        except Exception as e:
            logger.exception(f"Error bootstrapping from refresh token: {e}")
    
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
        
        # Filter reserved scopes
        filtered_scopes = self._filter_reserved_scopes(scopes)
        
        # Try silent acquisition (uses refresh token if needed)
        result = self.app.acquire_token_silent(
            scopes=filtered_scopes,
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
        
        NOTE: Device code flow requires PublicClientApplication, not
        ConfidentialClientApplication. This method uses self.public_app.
        
        Args:
            scopes: OAuth scopes to request (defaults to graph_scopes)
            callback: Optional function to receive device code info.
                     Called with dict containing 'user_code', 
                     'verification_uri', and 'message'.
        
        Returns:
            MSAL token result dict. Contains 'access_token' on success,
            or 'error' and 'error_description' on failure.
            
            IMPORTANT: If you need to persist auth across Railway deploys,
            save the 'refresh_token' from this result as M365_REFRESH_TOKEN
            environment variable.
        
        Example:
            >>> def show_code(info):
            ...     print(f"Go to {info['verification_uri']}")
            ...     print(f"Enter code: {info['user_code']}")
            >>> result = manager.authenticate_device_code(callback=show_code)
            >>> if 'refresh_token' in result:
            ...     print(f"Save this as M365_REFRESH_TOKEN: {result['refresh_token']}")
        """
        scopes = scopes or self.config.graph_scopes
        
        # CRITICAL: Filter out reserved scopes (offline_access, openid, profile)
        # MSAL handles these automatically and raises ValueError if included
        filtered_scopes = self._filter_reserved_scopes(scopes)
        
        logger.info(f"Initiating device flow with scopes: {filtered_scopes}")
        
        # IMPORTANT: Device code flow ONLY works with PublicClientApplication
        # ConfidentialClientApplication does NOT have initiate_device_flow method
        flow = self.public_app.initiate_device_flow(scopes=filtered_scopes)
        
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
        result = self.public_app.acquire_token_by_device_flow(flow)
        
        if "access_token" in result:
            logger.info("Device code authentication successful")
            
            # Log refresh token info for Railway deployment
            if "refresh_token" in result:
                rt = result["refresh_token"]
                logger.info(
                    f"Refresh token obtained. For Railway deployment, set "
                    f"M365_REFRESH_TOKEN env var. Token preview: {rt[:10]}...{rt[-10:]}"
                )
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
    
    def get_current_refresh_token(self) -> Optional[str]:
        """Get the current refresh token from cache for manual backup.
        
        This is useful for obtaining the refresh token to store as
        M365_REFRESH_TOKEN env var for Railway deployment.
        
        Returns:
            Refresh token string, or None if not available
            
        Note:
            This accesses internal MSAL cache structure and may break
            with future MSAL updates. Use with caution.
        """
        try:
            # Access the internal cache state
            cache_state = self.cache.serialize()
            if cache_state:
                import json
                cache_data = json.loads(cache_state)
                refresh_tokens = cache_data.get("RefreshToken", {})
                if refresh_tokens:
                    # Get the first refresh token
                    first_rt = next(iter(refresh_tokens.values()), {})
                    return first_rt.get("secret")
        except Exception as e:
            logger.warning(f"Could not extract refresh token from cache: {e}")
        
        return None
    
    def logout(self) -> None:
        """Clear all cached tokens and accounts."""
        accounts = self.app.get_accounts()
        for account in accounts:
            self.app.remove_account(account)
        
        self.cache.clear()
        logger.info("Logged out and cleared token cache")
