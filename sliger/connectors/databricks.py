"""Databricks SQL connector. Requires the ``databricks`` extra."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse

from sliger.connectors.base import (
    Connection,
    Connector,
    cells_from_rows,
    first_text,
    missing_extra,
    named_placeholders,
    require_params,
)
from sliger.exceptions import ConfigError
from sliger.results import TableResult


def load_databricks_sdk() -> Any:
    """Import databricks.sql. Isolated so tests can mock a missing extra."""
    from databricks import sql

    return sql


class DatabricksConnector(Connector):
    name = "databricks"
    aliases = ("dbsql",)
    extra = "databricks"

    def parse(self, name: str, spec: Mapping[str, Any]) -> Connection:
        url = spec.get("url") if isinstance(spec.get("url"), str) else ""
        parsed_url = _parse_url(url) if url else {}
        host = first_text(
            spec.get("host"),
            spec.get("server_hostname"),
            spec.get("hostname"),
            parsed_url.get("host"),
        )
        http_path = first_text(
            spec.get("http_path"), spec.get("http-path"), parsed_url.get("http_path")
        )
        token = first_text(
            spec.get("token"),
            spec.get("access_token"),
            parsed_url.get("token"),
        )
        if not host:
            raise ConfigError(
                f"connections.{name} needs host= or a databricks://token:TOKEN@HOST/HTTP_PATH URL"
            )
        if not http_path:
            raise ConfigError(f"connections.{name} needs http_path=")
        if not token:
            raise ConfigError(f"connections.{name} needs token=")
        options: dict[str, Any] = {
            "host": host,
            "http_path": http_path,
            "token": token,
        }
        catalog = first_text(spec.get("catalog"), parsed_url.get("catalog"))
        schema = first_text(spec.get("schema"), parsed_url.get("schema"))
        if catalog:
            options["catalog"] = catalog
        if schema:
            options["schema"] = schema
        return Connection(
            name=name,
            kind=self.name,
            url=url or f"databricks://{host}{_slash_path(http_path)}",
            options=options,
        )

    def execute(self, connection: Connection, query: str, params: Mapping[str, Any]) -> TableResult:
        require_params(query, params)
        try:
            sql = load_databricks_sdk()
        except ImportError as exc:
            raise missing_extra(self) from exc

        bound = {name: params[name] for name in named_placeholders(query)}
        conn = None
        try:
            conn = sql.connect(**self._connect_kwargs(connection))
            cur = conn.cursor()
            try:
                if bound:
                    cur.execute(query, bound)
                else:
                    cur.execute(query)
                fetched = cur.fetchall() or ()
                headers = tuple(col[0] for col in (cur.description or ()))
            finally:
                cur.close()
        except ConfigError:
            raise
        except Exception as exc:
            raise ConfigError(f"Databricks failed: {exc}") from exc
        finally:
            if conn is not None:
                conn.close()
        return cells_from_rows(headers, fetched)

    def _connect_kwargs(self, connection: Connection) -> dict[str, Any]:
        kwargs: dict[str, Any] = {
            "server_hostname": connection.option("host"),
            "http_path": connection.option("http_path"),
            "access_token": connection.option("token"),
        }
        catalog = connection.option("catalog")
        schema = connection.option("schema")
        if catalog:
            kwargs["catalog"] = catalog
        if schema:
            kwargs["schema"] = schema
        return kwargs


def _slash_path(path: str) -> str:
    return path if path.startswith("/") else f"/{path}"


def _parse_url(url: str) -> dict[str, str]:
    parsed = urlparse(url)
    if parsed.scheme not in {"", "databricks", "dbsql"}:
        raise ConfigError(f"Unsupported Databricks URL scheme: {parsed.scheme}")
    out: dict[str, str] = {}
    host = unquote(parsed.hostname or "")
    if host:
        out["host"] = host
    user = unquote(parsed.username) if parsed.username else ""
    password = unquote(parsed.password) if parsed.password else ""
    if password:
        out["token"] = password
    elif user and user.lower() != "token":
        out["token"] = user
    path = parsed.path or ""
    if path and path != "/":
        out["http_path"] = path if path.startswith("/") else f"/{path}"
    query = {key: values[-1] for key, values in parse_qs(parsed.query).items() if values}
    for key, alias in (
        ("http_path", "http_path"),
        ("http-path", "http_path"),
        ("token", "token"),
        ("access_token", "token"),
        ("catalog", "catalog"),
        ("schema", "schema"),
        ("host", "host"),
    ):
        if key in query:
            out[alias] = query[key]
    return out
