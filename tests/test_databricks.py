from __future__ import annotations

import os
from collections.abc import Iterator
from typing import Any

import pytest

from sliger.connectors import parse_connections, reset_registry, run_sql
from sliger.connectors.base import Connection
from sliger.connectors.databricks import DatabricksConnector, load_databricks_sdk
from sliger.exceptions import ConfigError
from sliger.results import TableResult


@pytest.fixture(autouse=True)
def _isolate_registry() -> Iterator[None]:
    reset_registry()
    yield
    reset_registry()


class _FakeCursor:
    def __init__(self) -> None:
        self.calls: list[tuple[str, Any]] = []
        self.description = (("id",), ("name",))
        self._rows: tuple[tuple[Any, ...], ...] = (("7", "Ada"),)

    def execute(self, sql: str, params: Any = None) -> None:
        self.calls.append((sql, params))

    def fetchall(self) -> tuple[tuple[Any, ...], ...]:
        return self._rows

    def close(self) -> None:
        return None


class _FakeConn:
    def __init__(self, cursor: _FakeCursor) -> None:
        self._cursor = cursor
        self.closed = False

    def cursor(self) -> _FakeCursor:
        return self._cursor

    def close(self) -> None:
        self.closed = True


class _FakeSDK:
    def __init__(self, conn: _FakeConn) -> None:
        self.conn = conn
        self.kwargs: dict[str, Any] = {}

    def connect(self, **kwargs: Any) -> _FakeConn:
        self.kwargs = kwargs
        return self.conn


def test_parse_fields() -> None:
    parsed = parse_connections(
        {
            "wh": {
                "type": "databricks",
                "host": "adb.example.net",
                "http_path": "/sql/1.0/warehouses/abc",
                "token": "dapi123",
                "catalog": "main",
                "schema": "metrics",
            }
        }
    )
    conn = parsed["wh"]
    assert conn.kind == "databricks"
    assert conn.option("host") == "adb.example.net"
    assert conn.option("http_path") == "/sql/1.0/warehouses/abc"
    assert conn.option("token") == "dapi123"
    assert conn.option("catalog") == "main"


def test_parse_url() -> None:
    parsed = parse_connections(
        {
            "wh": {
                "url": "databricks://token:dapi123@adb.example.net/sql/1.0/warehouses/abc"
                "?catalog=main&schema=metrics"
            }
        }
    )
    conn = parsed["wh"]
    assert conn.kind == "databricks"
    assert conn.option("host") == "adb.example.net"
    assert conn.option("http_path") == "/sql/1.0/warehouses/abc"
    assert conn.option("token") == "dapi123"
    assert conn.option("catalog") == "main"
    assert conn.option("schema") == "metrics"


def test_parse_dbsql_alias() -> None:
    parsed = parse_connections(
        {
            "wh": {
                "type": "dbsql",
                "host": "h",
                "http_path": "/sql/1.0/warehouses/x",
                "token": "t",
            }
        }
    )
    assert parsed["wh"].kind == "databricks"


def test_parse_missing_host() -> None:
    with pytest.raises(ConfigError, match="needs host="):
        parse_connections(
            {"wh": {"type": "databricks", "http_path": "/sql/1.0/warehouses/x", "token": "t"}}
        )


def test_parse_missing_token() -> None:
    with pytest.raises(ConfigError, match="needs token="):
        parse_connections(
            {
                "wh": {
                    "type": "databricks",
                    "host": "h",
                    "http_path": "/sql/1.0/warehouses/x",
                }
            }
        )


def test_parse_env_token(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("DATABRICKS_TOKEN", "from-env")
    parsed = parse_connections(
        {
            "wh": {
                "type": "databricks",
                "host": "h",
                "http_path": "/sql/1.0/warehouses/x",
                "token": "env:DATABRICKS_TOKEN",
            }
        }
    )
    assert parsed["wh"].option("token") == "from-env"


def test_execute_keeps_named_placeholders(monkeypatch: pytest.MonkeyPatch) -> None:
    cursor = _FakeCursor()
    sdk = _FakeSDK(_FakeConn(cursor))
    monkeypatch.setattr("sliger.connectors.databricks.load_databricks_sdk", lambda: sdk)
    connection = Connection(
        name="wh",
        kind="databricks",
        options={
            "host": "adb.example.net",
            "http_path": "/sql/1.0/warehouses/abc",
            "token": "dapi",
            "catalog": "main",
        },
    )
    result = run_sql(
        connection,
        "select :id as id, :name as name",
        {"id": 7, "name": "Ada"},
    )
    sql, bound = cursor.calls[0]
    assert sql == "select :id as id, :name as name"
    assert bound == {"id": 7, "name": "Ada"}
    assert sdk.kwargs["server_hostname"] == "adb.example.net"
    assert sdk.kwargs["http_path"] == "/sql/1.0/warehouses/abc"
    assert sdk.kwargs["access_token"] == "dapi"
    assert sdk.kwargs["catalog"] == "main"
    assert result == TableResult(headers=("id", "name"), rows=(("7", "Ada"),))


def test_execute_missing_extra(monkeypatch: pytest.MonkeyPatch) -> None:
    def _boom() -> Any:
        raise ImportError("no sdk")

    monkeypatch.setattr("sliger.connectors.databricks.load_databricks_sdk", _boom)
    connection = Connection(
        name="wh",
        kind="databricks",
        options={"host": "h", "http_path": "/p", "token": "t"},
    )
    with pytest.raises(ConfigError, match=r"sliger\[databricks\]"):
        run_sql(connection, "select 1")


def test_execute_wraps_sdk_errors(monkeypatch: pytest.MonkeyPatch) -> None:
    class _BadSDK:
        @staticmethod
        def connect(**kwargs: Any) -> Any:
            raise RuntimeError("warehouse stopped")

    monkeypatch.setattr("sliger.connectors.databricks.load_databricks_sdk", lambda: _BadSDK)
    connection = Connection(
        name="wh",
        kind="databricks",
        options={"host": "h", "http_path": "/p", "token": "t"},
    )
    with pytest.raises(ConfigError, match="Databricks failed: warehouse stopped"):
        run_sql(connection, "select 1")


def test_connector_class_and_sdk_helper_exist() -> None:
    assert DatabricksConnector.extra == "databricks"
    assert callable(load_databricks_sdk)


@pytest.mark.skipif(
    os.environ.get("SLIGER_DATABRICKS_LIVE") != "1", reason="set SLIGER_DATABRICKS_LIVE=1"
)
def test_live_databricks_select_one() -> None:
    pytest.importorskip("databricks.sql")
    host = os.environ.get("DATABRICKS_HOST") or os.environ.get("DATABRICKS_SERVER_HOSTNAME")
    http_path = os.environ.get("DATABRICKS_HTTP_PATH")
    token = os.environ.get("DATABRICKS_TOKEN")
    if not host or not http_path or not token:
        pytest.skip("needs DATABRICKS_HOST, DATABRICKS_HTTP_PATH, DATABRICKS_TOKEN")
    spec = {
        "type": "databricks",
        "host": host,
        "http_path": http_path,
        "token": token,
    }
    catalog = os.environ.get("DATABRICKS_CATALOG")
    schema = os.environ.get("DATABRICKS_SCHEMA")
    if catalog:
        spec["catalog"] = catalog
    if schema:
        spec["schema"] = schema
    parsed = parse_connections({"wh": spec})
    result = run_sql(parsed["wh"], "select 1 as n")
    assert result.rows[0][0] == "1"
