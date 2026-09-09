"""Built-in sqlite connector (no extra)."""

from __future__ import annotations

import sqlite3
from collections.abc import Mapping
from typing import Any
from urllib.parse import unquote, urlparse

from sliger.connectors.base import Connection, Connector, cells_from_rows, require_params
from sliger.exceptions import ConfigError
from sliger.results import TableResult


class SQLiteConnector(Connector):
    name = "sqlite"
    aliases = ("sqlite3",)
    extra = None

    def parse(self, name: str, spec: Mapping[str, Any]) -> Connection:
        url = spec.get("url")
        if not isinstance(url, str) or not url:
            raise ConfigError(f"connections.{name}.url is required")
        return Connection(name=name, kind=self.name, url=url)

    def execute(self, connection: Connection, query: str, params: Mapping[str, Any]) -> TableResult:
        require_params(query, params)
        path = sqlite_path(connection.url)
        conn = sqlite3.connect(path)
        try:
            cur = conn.execute(query, dict(params))
            fetched = cur.fetchall()
            conn.commit()
            headers = tuple(description[0] for description in cur.description or [])
            return cells_from_rows(headers, fetched)
        except sqlite3.Error as exc:
            raise ConfigError(f"SQL failed: {exc}") from exc
        finally:
            conn.close()


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
