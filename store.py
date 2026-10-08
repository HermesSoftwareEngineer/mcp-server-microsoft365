"""Persistência em SQLite: clientes OAuth, códigos, tokens do MCP e tokens da Microsoft."""

import hashlib
import json
import os
import sqlite3
import threading
import time
from typing import Any

import config

_lock = threading.Lock()


def _connect() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(config.DB_PATH) or ".", exist_ok=True)
    conn = sqlite3.connect(config.DB_PATH, check_same_thread=False)
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS kv (
            kind TEXT NOT NULL,
            key TEXT NOT NULL,
            value TEXT NOT NULL,
            expires_at REAL,
            client_id TEXT,
            subject TEXT,
            PRIMARY KEY (kind, key)
        )
        """
    )
    conn.execute("CREATE INDEX IF NOT EXISTS kv_owner ON kv (client_id, subject)")
    conn.commit()
    return conn


_conn = _connect()


def hash_secret(secret: str) -> str:
    return hashlib.sha256(secret.encode()).hexdigest()


def put(
    kind: str,
    key: str,
    value: dict[str, Any],
    expires_at: float | None = None,
    client_id: str | None = None,
    subject: str | None = None,
) -> None:
    with _lock:
        _conn.execute(
            "INSERT OR REPLACE INTO kv (kind, key, value, expires_at, client_id, subject) VALUES (?, ?, ?, ?, ?, ?)",
            (kind, key, json.dumps(value), expires_at, client_id, subject),
        )
        _conn.commit()


def get(kind: str, key: str) -> dict[str, Any] | None:
    with _lock:
        row = _conn.execute(
            "SELECT value, expires_at FROM kv WHERE kind = ? AND key = ?", (kind, key)
        ).fetchone()
    if row is None:
        return None
    value, expires_at = row
    if expires_at is not None and expires_at < time.time():
        delete(kind, key)
        return None
    return json.loads(value)


def pop(kind: str, key: str) -> dict[str, Any] | None:
    value = get(kind, key)
    if value is not None:
        delete(kind, key)
    return value


def delete(kind: str, key: str) -> None:
    with _lock:
        _conn.execute("DELETE FROM kv WHERE kind = ? AND key = ?", (kind, key))
        _conn.commit()


def delete_owner_tokens(subject: str, client_id: str | None = None) -> None:
    """Apaga tokens de acesso/refresh do MCP de um usuário (opcionalmente só de um cliente)."""
    sql = "DELETE FROM kv WHERE kind IN ('access_token', 'refresh_token') AND subject = ?"
    params: list[Any] = [subject]
    if client_id is not None:
        sql += " AND client_id = ?"
        params.append(client_id)
    with _lock:
        _conn.execute(sql, params)
        _conn.commit()


def purge_expired() -> None:
    with _lock:
        _conn.execute("DELETE FROM kv WHERE expires_at IS NOT NULL AND expires_at < ?", (time.time(),))
        _conn.commit()
