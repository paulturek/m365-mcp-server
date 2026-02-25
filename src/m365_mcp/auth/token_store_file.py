"""Option A — File-based encrypted token store.

All user tokens live in a single Fernet-encrypted JSON file.
Suitable for single-instance Railway deploys with a persistent volume.

Config env vars:
  TOKEN_ENCRYPTION_KEY  — Fernet key (generate: python -c "from cryptography.fernet import Fernet; print(Fernet.generate_key().decode())")
  TOKEN_STORE_PATH      — File path (default /app/data/tokens.enc)
"""
import json
import time
import logging
import asyncio
from pathlib import Path
from typing import Optional

from cryptography.fernet import Fernet

from .token_store import TokenStore

logger = logging.getLogger(__name__)


class FileTokenStore(TokenStore):
    """Encrypted JSON file keyed by user_id."""

    def __init__(self, path: Path, encryption_key: str):
        self._path = path
        self._fernet = Fernet(
            encryption_key.encode() if isinstance(encryption_key, str) else encryption_key
        )
        self._cache: dict[str, dict] = {}
        self._loaded = False
        self._lock = asyncio.Lock()

    # -- internal --------------------------------------------------------

    def _load(self) -> dict[str, dict]:
        if self._loaded:
            return self._cache
        if not self._path.exists():
            self._loaded = True
            return self._cache
        try:
            decrypted = self._fernet.decrypt(self._path.read_bytes())
            self._cache = json.loads(decrypted)
            logger.info("Loaded tokens for %d user(s) from %s", len(self._cache), self._path)
        except Exception as exc:
            logger.warning("Token file %s unreadable, starting fresh: %s", self._path, exc)
            self._cache = {}
        self._loaded = True
        return self._cache

    def _flush(self) -> None:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        encrypted = self._fernet.encrypt(json.dumps(self._cache).encode())
        self._path.write_bytes(encrypted)

    # -- public ----------------------------------------------------------

    async def get(self, user_id: str) -> Optional[dict]:
        async with self._lock:
            return self._load().get(user_id)

    async def save(self, user_id: str, token_data: dict) -> None:
        async with self._lock:
            self._load()
            self._cache[user_id] = token_data
            self._flush()

    async def delete(self, user_id: str) -> bool:
        async with self._lock:
            self._load()
            if user_id in self._cache:
                del self._cache[user_id]
                self._flush()
                return True
            return False

    async def list_users(self) -> list[dict]:
        async with self._lock:
            tokens = self._load()
            now = time.time()
            return [
                {
                    "user_id": uid,
                    "display_name": d.get("display_name", ""),
                    "token_expired": d.get("expires_at", 0) < now,
                    "authorized_at": d.get("authorized_at"),
                }
                for uid, d in tokens.items()
            ]
