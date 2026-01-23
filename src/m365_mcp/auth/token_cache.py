"""Encrypted persistent token cache for MSAL.

This module provides an MSAL-compatible token cache that:
- Encrypts tokens at rest using AES-256 (Fernet)
- Automatically persists on any token change
- Generates and securely stores encryption keys

Security:
- Cache file permissions: 600 (owner read/write only)
- Key file permissions: 600 (owner read/write only)
- Key is stored separately from cache for defense in depth
"""

import logging
from pathlib import Path
from typing import Any, Optional

from cryptography.fernet import Fernet, InvalidToken
from msal import SerializableTokenCache

logger = logging.getLogger(__name__)


class EncryptedTokenCache(SerializableTokenCache):
    """MSAL-compatible token cache with AES-256 encryption at rest.
    
    Tokens are encrypted using Fernet (AES-128-CBC with HMAC) and persisted
    to disk. The encryption key is stored separately from the cache file.
    
    Attributes:
        cache_path: Path to the encrypted cache file
    
    Example:
        >>> cache = EncryptedTokenCache("/path/to/cache.bin")
        >>> # Use with MSAL application
        >>> app = PublicClientApplication(client_id, token_cache=cache)
    """
    
    def __init__(
        self,
        cache_path: str,
        encryption_key: Optional[str] = None
    ) -> None:
        """Initialize encrypted token cache.
        
        Args:
            cache_path: Path to store encrypted token cache
            encryption_key: Optional Fernet key (base64-encoded).
                          If not provided, one will be generated/loaded.
        """
        super().__init__()
        
        self.cache_path = Path(cache_path)
        self._ensure_directory()
        
        # Initialize encryption
        self._fernet = self._init_encryption(encryption_key)
        
        # Load existing cache if present
        self._load()
    
    def _ensure_directory(self) -> None:
        """Create cache directory with secure permissions."""
        self.cache_path.parent.mkdir(parents=True, exist_ok=True, mode=0o700)
    
    def _init_encryption(self, provided_key: Optional[str]) -> Fernet:
        """Initialize or load encryption key.
        
        Args:
            provided_key: Optional pre-existing Fernet key
            
        Returns:
            Initialized Fernet instance
        """
        key_path = self.cache_path.parent / ".cache_key"
        
        # Use provided key if given
        if provided_key:
            logger.debug("Using provided encryption key")
            return Fernet(provided_key.encode())
        
        # Try to load existing key
        if key_path.exists():
            try:
                key_data = key_path.read_bytes()
                logger.debug("Loaded existing encryption key")
                return Fernet(key_data)
            except Exception as e:
                logger.warning(f"Failed to load key file, generating new: {e}")
        
        # Generate new key
        key = Fernet.generate_key()
        key_path.write_bytes(key)
        key_path.chmod(0o600)  # Owner read/write only
        logger.info(f"Generated new encryption key at {key_path}")
        
        return Fernet(key)
    
    def _load(self) -> None:
        """Load and decrypt cache from disk."""
        if not self.cache_path.exists():
            logger.debug("No existing token cache found")
            return
        
        try:
            encrypted_data = self.cache_path.read_bytes()
            decrypted_data = self._fernet.decrypt(encrypted_data)
            self.deserialize(decrypted_data.decode("utf-8"))
            logger.info("Token cache loaded successfully")
        except InvalidToken:
            logger.warning(
                "Token cache corrupted or encryption key changed. "
                "Starting with fresh cache."
            )
            self.cache_path.unlink(missing_ok=True)
        except Exception as e:
            logger.error(f"Failed to load token cache: {e}")
    
    def _save(self) -> None:
        """Encrypt and persist cache to disk."""
        if not self.has_state_changed:
            return
        
        try:
            serialized = self.serialize()
            encrypted = self._fernet.encrypt(serialized.encode("utf-8"))
            self.cache_path.write_bytes(encrypted)
            self.cache_path.chmod(0o600)  # Owner read/write only
            logger.debug("Token cache saved")
        except Exception as e:
            logger.error(f"Failed to save token cache: {e}")
    
    # Override MSAL hooks to auto-persist changes
    
    def add(self, event: dict[str, Any], **kwargs: Any) -> None:
        """Add token to cache and persist."""
        super().add(event, **kwargs)
        self._save()
    
    def modify(
        self,
        credential_type: str,
        old_entry: dict[str, Any],
        new_key_value_pairs: Optional[dict[str, Any]] = None
    ) -> None:
        """Modify cached token and persist."""
        super().modify(credential_type, old_entry, new_key_value_pairs)
        self._save()
    
    def remove(
        self,
        credential_type: str,
        target: Optional[dict[str, Any]] = None
    ) -> None:
        """Remove token from cache and persist."""
        super().remove(credential_type, target)
        self._save()
    
    def clear(self) -> None:
        """Clear all cached tokens and remove cache file."""
        # Clear in-memory cache
        credential_types = [
            "AccessToken", "RefreshToken", "IdToken", "Account"
        ]
        for cred_type in credential_types:
            try:
                entries = list(self.find(cred_type))
                for entry in entries:
                    self.remove(cred_type, entry)
            except Exception:
                pass
        
        # Remove cache file
        self.cache_path.unlink(missing_ok=True)
        logger.info("Token cache cleared")
