"""Snowflake connector. Requires the ``snowflake`` extra."""

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
    rewrite_placeholders,
)
from sliger.exceptions import ConfigError
from sliger.results import TableResult

_OPTIONAL = ("warehouse", "database", "schema", "role", "authenticator")


def load_snowflake_sdk() -> Any:
    """Import snowflake.connector. Isolated so tests can mock a missing extra."""
    import snowflake.connector

    return snowflake.connector


class SnowflakeConnector(Connector):
    name = "snowflake"
    aliases = ("sf",)
    extra = "snowflake"

    def parse(self, name: str, spec: Mapping[str, Any]) -> Connection:
        url = spec.get("url") if isinstance(spec.get("url"), str) else ""
        parsed_url = _parse_url(url) if url else {}
        account = first_text(spec.get("account"), parsed_url.get("account"))
        user = first_text(spec.get("user"), spec.get("username"), parsed_url.get("user"))
        password = first_text(spec.get("password"), parsed_url.get("password"))
        token = first_text(spec.get("token"), parsed_url.get("token"))
        if not account:
            raise ConfigError(
                f"connections.{name} needs account= or a snowflake://USER@ACCOUNT URL"
            )
        if not user:
            raise ConfigError(f"connections.{name} needs user=")
        if not password and not token:
            raise ConfigError(f"connections.{name} needs password= or token=")
        options: dict[str, Any] = {"account": account, "user": user}
        if password:
            options["password"] = password
        if token:
            options["token"] = token
        for key in _OPTIONAL:
            value = first_text(spec.get(key), parsed_url.get(key))
            if value:
                options[key] = value
        return Connection(
            name=name,
            kind=self.name,
            url=url or f"snowflake://{user}@{account}",
            options=options,
        )

    def execute(self, connection: Connection, query: str, params: Mapping[str, Any]) -> TableResult:
        require_params(query, params)
        try:
            snowflake = load_snowflake_sdk()
        except ImportError as exc:
            raise missing_extra(self) from exc

        sql = rewrite_placeholders(query, "%")
        bound = {name: params[name] for name in named_placeholders(query)}
        conn = None
        try:
            conn = snowflake.connect(**self._connect_kwargs(connection))
            cur = conn.cursor()
            try:
                cur.execute(sql, bound)
                fetched = cur.fetchall() or ()
                headers = tuple(col[0] for col in (cur.description or ()))
            finally:
                cur.close()
        except ConfigError:
            raise
        except Exception as exc:
            raise ConfigError(f"Snowflake failed: {exc}") from exc
        finally:
            if conn is not None:
                conn.close()
        return cells_from_rows(headers, fetched)

    def _connect_kwargs(self, connection: Connection) -> dict[str, Any]:
        kwargs: dict[str, Any] = {
            "account": connection.option("account"),
            "user": connection.option("user"),
        }
        password = connection.option("password")
        token = connection.option("token")
        authenticator = connection.option("authenticator")
        if password:
            kwargs["password"] = password
        if token:
            kwargs["token"] = token
            if not authenticator and not password:
                kwargs["authenticator"] = "PROGRAMMATIC_ACCESS_TOKEN"
        if authenticator:
            kwargs["authenticator"] = authenticator
        for key in ("warehouse", "database", "schema", "role"):
            value = connection.option(key)
            if value:
                kwargs[key] = value
        return kwargs


def _parse_url(url: str) -> dict[str, str]:
    parsed = urlparse(url)
    if parsed.scheme not in {"", "snowflake", "sf"}:
        raise ConfigError(f"Unsupported Snowflake URL scheme: {parsed.scheme}")
    out: dict[str, str] = {}
    host = unquote(parsed.hostname or "")
    if host.endswith(".snowflakecomputing.com"):
        host = host[: -len(".snowflakecomputing.com")]
    if host:
        out["account"] = host
    if parsed.username:
        out["user"] = unquote(parsed.username)
    if parsed.password:
        out["password"] = unquote(parsed.password)
    path_parts = [unquote(part) for part in parsed.path.split("/") if part]
    if path_parts:
        out["database"] = path_parts[0]
    if len(path_parts) > 1:
        out["schema"] = path_parts[1]
    query = {key: values[-1] for key, values in parse_qs(parsed.query).items() if values}
    for key in (*_OPTIONAL, "token", "account", "user", "password"):
        if key in query:
            out[key] = query[key]
    return out
