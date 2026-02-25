"""Abstract token store interface.

Implementations:
  - FileTokenStore  (Option A) — Fernet-encrypted JSON file
  - PgTokenStore    (Option B) — Fernet-encrypted rows in PostgreSQL

Switch via TOKEN_STORE_BACKEND env var ("file" | "pg").
"""
from abc import ABC, abstractmethod
from typing import Optional


class TokenStore(ABC):
    """Multi-user token persistence interface."""

    @abstractmethod
    async def get(self, user_id: str) -> Optional[dict]:
        """Retrieve token data for a user. Returns None if not found."""
        ...

    @abstractmethod
    async def save(self, user_id: str, token_data: dict) -> None:
        """Persist token data for a user (upsert)."""
        ...

    @abstractmethod
    async def delete(self, user_id: str) -> bool:
        """Remove token data. Returns True if the user existed."""
        ...

    @abstractmethod
    async def list_users(self) -> list[dict]:
        """Return summary dicts for every authorised user."""
        ...
