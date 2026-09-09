from __future__ import annotations

import os
from collections.abc import Iterator
from typing import Any

import pytest

from sliger.connectors import parse_connections, reset_registry, run_sql
from sliger.connectors.base import Connection
from sliger.connectors.snowflake import SnowflakeConnector, load_snowflake_sdk
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
                "type": "snowflake",
                "account": "acme-prod",
                "user": "analyst",
                "password": "s3cret",
                "warehouse": "COMPUTE_WH",
                "database": "ANALYTICS",
                "schema": "PUBLIC",
                "role": "READER",
            }
        }
    )
    conn = parsed["wh"]
    assert conn.kind == "snowflake"
    assert conn.option("account") == "acme-prod"
    assert conn.option("user") == "analyst"
    assert conn.option("warehouse") == "COMPUTE_WH"
    assert conn.option("database") == "ANALYTICS"


def test_parse_url() -> None:
    parsed = parse_connections(
        {"wh": {"url": "snowflake://ada:pw@acme-east/ANALYTICS/PUBLIC?warehouse=WH&role=R"}}
    )
    conn = parsed["wh"]
    assert conn.kind == "snowflake"
    assert conn.option("account") == "acme-east"
    assert conn.option("user") == "ada"
    assert conn.option("password") == "pw"
    assert conn.option("database") == "ANALYTICS"
    assert conn.option("schema") == "PUBLIC"
    assert conn.option("warehouse") == "WH"


def test_parse_strips_computing_host() -> None:
    parsed = parse_connections(
        {"wh": {"url": "snowflake://ada:pw@acme-east.snowflakecomputing.com/DB"}}
    )
    assert parsed["wh"].option("account") == "acme-east"


def test_parse_sf_alias_and_token() -> None:
    parsed = parse_connections(
        {"wh": {"type": "sf", "account": "a", "user": "u", "token": "pat-1"}}
    )
    assert parsed["wh"].kind == "snowflake"
    assert parsed["wh"].option("token") == "pat-1"


def test_parse_missing_account() -> None:
    with pytest.raises(ConfigError, match="needs account="):
        parse_connections({"wh": {"type": "snowflake", "user": "u", "password": "p"}})


def test_parse_missing_secret() -> None:
    with pytest.raises(ConfigError, match="needs password= or token="):
        parse_connections({"wh": {"type": "snowflake", "account": "a", "user": "u"}})


def test_parse_env_password(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SLIGER_SF_PASSWORD", "from-env")
    parsed = parse_connections(
        {
            "wh": {
                "type": "snowflake",
                "account": "a",
                "user": "u",
                "password": "env:SLIGER_SF_PASSWORD",
            }
        }
    )
    assert parsed["wh"].option("password") == "from-env"


def test_execute_rewrites_placeholders(monkeypatch: pytest.MonkeyPatch) -> None:
    cursor = _FakeCursor()
    sdk = _FakeSDK(_FakeConn(cursor))
    monkeypatch.setattr("sliger.connectors.snowflake.load_snowflake_sdk", lambda: sdk)
    connection = Connection(
        name="wh",
        kind="snowflake",
        url="snowflake://u@a",
        options={
            "account": "a",
            "user": "u",
            "password": "p",
            "warehouse": "WH",
            "database": "DB",
        },
    )
    result = run_sql(
        connection,
        "select :id as id, :name as name",
        {"id": 7, "name": "Ada"},
    )
    sql, bound = cursor.calls[0]
    assert sql == "select %(id)s as id, %(name)s as name"
    assert bound == {"id": 7, "name": "Ada"}
    assert sdk.kwargs["account"] == "a"
    assert sdk.kwargs["warehouse"] == "WH"
    assert result == TableResult(headers=("id", "name"), rows=(("7", "Ada"),))


def test_execute_uses_programmatic_token(monkeypatch: pytest.MonkeyPatch) -> None:
    cursor = _FakeCursor()
    sdk = _FakeSDK(_FakeConn(cursor))
    monkeypatch.setattr("sliger.connectors.snowflake.load_snowflake_sdk", lambda: sdk)
    connection = Connection(
        name="wh",
        kind="snowflake",
        options={"account": "a", "user": "u", "token": "pat"},
    )
    run_sql(connection, "select 1")
    assert sdk.kwargs["token"] == "pat"
    assert sdk.kwargs["authenticator"] == "PROGRAMMATIC_ACCESS_TOKEN"


def test_execute_missing_extra(monkeypatch: pytest.MonkeyPatch) -> None:
    def _boom() -> Any:
        raise ImportError("no sdk")

    monkeypatch.setattr("sliger.connectors.snowflake.load_snowflake_sdk", _boom)
    connection = Connection(
        name="wh",
        kind="snowflake",
        options={"account": "a", "user": "u", "password": "p"},
    )
    with pytest.raises(ConfigError, match=r"sliger\[snowflake\]"):
        run_sql(connection, "select 1")


def test_execute_wraps_sdk_errors(monkeypatch: pytest.MonkeyPatch) -> None:
    class _BadSDK:
        @staticmethod
        def connect(**kwargs: Any) -> Any:
            raise RuntimeError("warehouse suspended")

    monkeypatch.setattr("sliger.connectors.snowflake.load_snowflake_sdk", lambda: _BadSDK)
    connection = Connection(
        name="wh",
        kind="snowflake",
        options={"account": "a", "user": "u", "password": "p"},
    )
    with pytest.raises(ConfigError, match="Snowflake failed: warehouse suspended"):
        run_sql(connection, "select 1")


def test_connector_class_and_sdk_helper_exist() -> None:
    assert SnowflakeConnector.extra == "snowflake"
    assert callable(load_snowflake_sdk)


@pytest.mark.skipif(
    os.environ.get("SLIGER_SNOWFLAKE_LIVE") != "1", reason="set SLIGER_SNOWFLAKE_LIVE=1"
)
def test_live_snowflake_select_one() -> None:
    pytest.importorskip("snowflake.connector")
    account = os.environ.get("SNOWFLAKE_ACCOUNT")
    user = os.environ.get("SNOWFLAKE_USER")
    password = os.environ.get("SNOWFLAKE_PASSWORD")
    token = os.environ.get("SNOWFLAKE_TOKEN")
    if not account or not user or not (password or token):
        pytest.skip("needs SNOWFLAKE_ACCOUNT, SNOWFLAKE_USER, and password or token")
    spec: dict[str, str] = {"type": "snowflake", "account": account, "user": user}
    if password:
        spec["password"] = password
    if token:
        spec["token"] = token
    for key, env in (
        ("warehouse", "SNOWFLAKE_WAREHOUSE"),
        ("database", "SNOWFLAKE_DATABASE"),
        ("schema", "SNOWFLAKE_SCHEMA"),
        ("role", "SNOWFLAKE_ROLE"),
    ):
        value = os.environ.get(env)
        if value:
            spec[key] = value
    parsed = parse_connections({"wh": spec})
    result = run_sql(parsed["wh"], "select 1 as n")
    assert result.rows[0][0] == "1"
