"""Option B — PostgreSQL-backed encrypted token store.

Survives container redeployments. Uses asyncpg for async I/O
and Fernet for at-rest encryption of token blobs.

Config env vars:
  DATABASE_URL          — Postgres connection string
  TOKEN_ENCRYPTION_KEY  — Fernet key
"""
import json
import time
import logging
import asyncio
from datetime import datetime, timezone
from typing import Optional

from cryptography.fernet import Fernet

from .token_store import TokenStore

logger = logging.getLogger(__name__)

CREATE_TABLE_SQL = """
CREATE TABLE IF NOT EXISTS mcp_user_tokens (
    user_id         TEXT PRIMARY KEY,
    token_data      BYTEA NOT NULL,
    display_name    TEXT DEFAULT '',
    authorized_at   TIMESTAMPTZ,
    updated_at      TIMESTAMPTZ DEFAULT now()
);
"""


class PgTokenStore(TokenStore):
    """PostgreSQL + Fernet token store keyed by user_id."""

    def __init__(self, database_url: str, encryption_key: str):
        self._dsn = database_url
        self._fernet = Fernet(
            encryption_key.encode() if isinstance(encryption_key, str) else encryption_key
        )
        self._pool = None
        self._init_lock = asyncio.Lock()

    async def _ensure_pool(self):
        if self._pool is not None:
            return self._pool
        async with self._init_lock:
            if self._pool is not None:
                return self._pool
            import asyncpg

            self._pool = await asyncpg.create_pool(
                dsn=self._dsn, min_size=1, max_size=5
            )
            async with self._pool.acquire() as conn:
                await conn.execute(CREATE_TABLE_SQL)
            logger.info("PgTokenStore: pool ready, table ensured")
            return self._pool

    async def close(self):
        if self._pool:
            await self._pool.close()
            self._pool = None

    def _encrypt(self, data: dict) -> bytes:
        return self._fernet.encrypt(json.dumps(data).encode())

    def _decrypt(self, blob: bytes) -> dict:
        return json.loads(self._fernet.decrypt(bytes(blob)))

    async def get(self, user_id: str) -> Optional[dict]:
        pool = await self._ensure_pool()
        async with pool.acquire() as conn:
            row = await conn.fetchrow(
                "SELECT token_data FROM mcp_user_tokens WHERE user_id = $1",
                user_id,
            )
        if not row:
            return None
        try:
            return self._decrypt(row["token_data"])
        except Exception as exc:
            logger.warning("Decrypt failed for '%s': %s", user_id, exc)
            return None

    async def save(self, user_id: str, token_data: dict) -> None:
        pool = await self._ensure_pool()
        encrypted = self._encrypt(token_data)
        display_name = token_data.get("display_name", "")
        authorized_at = datetime.fromtimestamp(
            token_data.get("authorized_at", time.time()), tz=timezone.utc
        )
        async with pool.acquire() as conn:
            await conn.execute(
                """
                INSERT INTO mcp_user_tokens
                    (user_id, token_data, display_name, authorized_at, updated_at)
                VALUES ($1, $2, $3, $4, now())
                ON CONFLICT (user_id) DO UPDATE SET
                    token_data   = EXCLUDED.token_data,
                    display_name = EXCLUDED.display_name,
                    updated_at   = now()
                """,
                user_id,
                encrypted,
                display_name,
                authorized_at,
            )

    async def delete(self, user_id: str) -> bool:
        pool = await self._ensure_pool()
        async with pool.acquire() as conn:
            result = await conn.execute(
                "DELETE FROM mcp_user_tokens WHERE user_id = $1", user_id
            )
        return result != "DELETE 0"

    async def list_users(self) -> list[dict]:
        pool = await self._ensure_pool()
        async with pool.acquire() as conn:
            rows = await conn.fetch(
                """SELECT user_id, display_name, authorized_at, updated_at
                   FROM mcp_user_tokens ORDER BY updated_at DESC"""
            )
        return [
            {
                "user_id": r["user_id"],
                "display_name": r["display_name"] or "",
                "authorized_at": r["authorized_at"].isoformat() if r["authorized_at"] else None,
                "updated_at": r["updated_at"].isoformat() if r["updated_at"] else None,
            }
            for r in rows
        ]
