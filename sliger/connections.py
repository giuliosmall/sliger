"""Named SQL connections loaded from sliger TOML config."""

from __future__ import annotations

import os
import re
import sqlite3
from dataclasses import dataclass
from typing import Any
from urllib.parse import unquote, urlparse

from sliger.exceptions import ConfigError
from sliger.results import TableResult

_PLACEHOLDER = re.compile(r":([A-Za-z_][A-Za-z0-9_]*)")


@dataclass(frozen=True)
class Connection:
    name: str
    kind: str
    url: str


def resolve_secret(value: str) -> str:
    """Resolve ``env:NAME`` to an environment variable; pass other strings through."""
    if value.startswith("env:"):
        name = value[4:]
        if name not in os.environ:
            raise ConfigError(f"Environment variable {name} is not set")
        return os.environ[name]
    return value


def parse_connections(raw: Any) -> dict[str, Connection]:
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise ConfigError("'connections' must be a table")
    parsed: dict[str, Connection] = {}
    for name, spec in raw.items():
        if not isinstance(spec, dict):
            raise ConfigError(f"connections.{name} must be a table")
        kind = str(spec.get("type") or spec.get("kind") or "sqlite")
        url = spec.get("url")
        if not isinstance(url, str) or not url:
            raise ConfigError(f"connections.{name}.url is required")
        parsed[name] = Connection(name=name, kind=kind.lower(), url=resolve_secret(url))
    return parsed


def sqlite_path(url: str) -> str:
    if url in {":memory:", "sqlite:///:memory:", "sqlite://memory"}:
        return ":memory:"
    parsed = urlparse(url)
    if parsed.scheme in {"", "sqlite", "file"}:
        path = unquote(parsed.path or "")
        if parsed.scheme == "sqlite" and path.startswith("//"):
            path = path[1:]
        if not path or path in {":memory:", "/:memory:"}:
            return ":memory:"
        return path
    raise ConfigError(f"Unsupported SQL URL scheme: {parsed.scheme}")


def run_sql(
    connection: Connection, query: str, params: dict[str, Any] | None = None
) -> TableResult:
    if connection.kind not in {"sqlite", "sqlite3"}:
        raise ConfigError(
            f"Connection {connection.name!r} uses type {connection.kind!r}; "
            "only sqlite is built in. Wrap another warehouse in a custom function."
        )
    values = params or {}
    unknown = [name for name in _PLACEHOLDER.findall(query) if name not in values]
    if unknown:
        raise ConfigError(f"SQL is missing parameters: {', '.join(unknown)}")
    path = sqlite_path(connection.url)
    conn = sqlite3.connect(path)
    try:
        cur = conn.execute(query, values)
        fetched = cur.fetchall()
        conn.commit()
        headers = tuple(description[0] for description in cur.description or [])
        data = tuple(tuple("" if cell is None else str(cell) for cell in row) for row in fetched)
        return TableResult(headers=headers, rows=data)
    except sqlite3.Error as exc:
        raise ConfigError(f"SQL failed: {exc}") from exc
    finally:
        conn.close()
