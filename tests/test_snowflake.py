from __future__ import annotations

import os
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path
from typing import Any

import pytest

from live_flags import live_flag
from sliger.connectors import get_connector, parse_connections, reset_registry, run_sql
from sliger.connectors.base import Connection
from sliger.connectors.snowflake import SnowflakeConnector, load_snowflake_sdk
from sliger.context import FunctionContext, reset_context, set_context
from sliger.exceptions import ConfigError
from sliger.jinja_utils import load_connections, load_jinja_environment, render_box
from sliger.results import TableResult


@pytest.fixture(autouse=True)
def _isolate_registry() -> Iterator[None]:
    reset_registry()
    yield
    reset_registry()


class _FakeCursor:
    def __init__(
        self,
        *,
        rows: tuple[tuple[Any, ...], ...] | None = (("7", "Ada"),),
        description: tuple[tuple[str, ...], ...] | None = (("id",), ("name",)),
        fail: Exception | None = None,
    ) -> None:
        self.calls: list[tuple[str, Any]] = []
        self.description = description
        self._rows = rows
        self._fail = fail
        self.closed = False

    def execute(self, sql: str, params: Any = None) -> None:
        self.calls.append((sql, params))
        if self._fail is not None:
            raise self._fail

    def fetchall(self) -> tuple[tuple[Any, ...], ...] | None:
        return self._rows

    def close(self) -> None:
        self.closed = True


class _FakeConn:
    def __init__(self, cursor: _FakeCursor) -> None:
        self._cursor = cursor
        self.closed = False

    def cursor(self) -> _FakeCursor:
        return self._cursor

    def close(self) -> None:
        self.closed = True


class _FakeSDK:
    def __init__(
        self, conn: _FakeConn | None = None, *, connect_error: Exception | None = None
    ) -> None:
        self.conn = conn
        self.connect_error = connect_error
        self.kwargs: dict[str, Any] = {}
        self.connect_calls = 0

    def connect(self, **kwargs: Any) -> _FakeConn:
        self.connect_calls += 1
        self.kwargs = kwargs
        if self.connect_error is not None:
            raise self.connect_error
        assert self.conn is not None
        return self.conn


def _password_conn(**options: Any) -> Connection:
    base = {"account": "a", "user": "u", "password": "p"}
    base.update(options)
    return Connection(name="wh", kind="snowflake", options=base)


def _install_sdk(monkeypatch: pytest.MonkeyPatch, sdk: _FakeSDK) -> _FakeSDK:
    monkeypatch.setattr("sliger.connectors.snowflake.load_snowflake_sdk", lambda: sdk)
    return sdk


# --- parse ---


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
    assert conn.option("schema") == "PUBLIC"
    assert conn.option("role") == "READER"
    assert conn.url == "snowflake://analyst@acme-prod"


def test_parse_username_alias() -> None:
    parsed = parse_connections(
        {"wh": {"type": "snowflake", "account": "a", "username": "ada", "password": "p"}}
    )
    assert parsed["wh"].option("user") == "ada"


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
    assert conn.option("role") == "R"


def test_parse_url_scheme_sf() -> None:
    parsed = parse_connections({"wh": {"url": "sf://ada:pw@acct"}})
    assert parsed["wh"].kind == "snowflake"
    assert parsed["wh"].option("account") == "acct"


def test_parse_strips_computing_host() -> None:
    parsed = parse_connections(
        {"wh": {"url": "snowflake://ada:pw@acme-east.snowflakecomputing.com/DB"}}
    )
    assert parsed["wh"].option("account") == "acme-east"
    assert parsed["wh"].option("database") == "DB"


def test_parse_url_encoded_password() -> None:
    parsed = parse_connections({"wh": {"url": "snowflake://ada:p%40ss@acct"}})
    assert parsed["wh"].option("password") == "p@ss"


def test_parse_token_in_query() -> None:
    parsed = parse_connections({"wh": {"url": "snowflake://ada@acct?token=pat-1"}})
    assert parsed["wh"].option("token") == "pat-1"
    assert parsed["wh"].option("user") == "ada"


def test_parse_spec_overrides_url() -> None:
    parsed = parse_connections(
        {
            "wh": {
                "url": "snowflake://urluser:urlpw@urlacct/URLDB",
                "account": "specacct",
                "user": "specuser",
                "password": "specpw",
                "database": "SPECDB",
            }
        }
    )
    conn = parsed["wh"]
    assert conn.option("account") == "specacct"
    assert conn.option("user") == "specuser"
    assert conn.option("password") == "specpw"
    assert conn.option("database") == "SPECDB"


def test_parse_sf_alias_and_token() -> None:
    parsed = parse_connections(
        {"wh": {"type": "sf", "account": "a", "user": "u", "token": "pat-1"}}
    )
    assert parsed["wh"].kind == "snowflake"
    assert get_connector("sf").name == "snowflake"
    assert parsed["wh"].option("token") == "pat-1"


def test_parse_authenticator() -> None:
    parsed = parse_connections(
        {
            "wh": {
                "type": "snowflake",
                "account": "a",
                "user": "u",
                "token": "pat",
                "authenticator": "oauth",
            }
        }
    )
    assert parsed["wh"].option("authenticator") == "oauth"


def test_parse_missing_account() -> None:
    with pytest.raises(ConfigError, match="needs account="):
        parse_connections({"wh": {"type": "snowflake", "user": "u", "password": "p"}})


def test_parse_missing_user() -> None:
    with pytest.raises(ConfigError, match="needs user="):
        parse_connections({"wh": {"type": "snowflake", "account": "a", "password": "p"}})


def test_parse_missing_secret() -> None:
    with pytest.raises(ConfigError, match="needs password= or token="):
        parse_connections({"wh": {"type": "snowflake", "account": "a", "user": "u"}})


def test_parse_unsupported_scheme() -> None:
    with pytest.raises(ConfigError, match="Unsupported Snowflake URL scheme: postgres"):
        parse_connections({"wh": {"type": "snowflake", "url": "postgres://u:p@host/db"}})


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


def test_parse_env_token(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SLIGER_SF_TOKEN", "pat-env")
    parsed = parse_connections(
        {
            "wh": {
                "type": "snowflake",
                "account": "a",
                "user": "u",
                "token": "env:SLIGER_SF_TOKEN",
            }
        }
    )
    assert parsed["wh"].option("token") == "pat-env"


def test_load_connections_from_toml(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SF_TOKEN", "pat-file")
    path = tmp_path / "config.toml"
    path.write_text(
        "\n".join(
            [
                "[connections.snow]",
                'type = "snowflake"',
                'account = "LEOJRZQ-LE93180"',
                'user = "GIULIOPICCOLO"',
                'token = "env:SF_TOKEN"',
                'warehouse = "COMPUTE_WH"',
                "",
            ]
        ),
        encoding="utf-8",
    )
    loaded = load_connections(path)
    assert loaded["snow"].kind == "snowflake"
    assert loaded["snow"].option("token") == "pat-file"
    assert loaded["snow"].option("warehouse") == "COMPUTE_WH"


# --- connect kwargs ---


def test_connect_kwargs_password_only() -> None:
    kwargs = SnowflakeConnector()._connect_kwargs(_password_conn(warehouse="WH", role="R"))
    assert kwargs == {
        "account": "a",
        "user": "u",
        "password": "p",
        "warehouse": "WH",
        "role": "R",
    }
    assert "authenticator" not in kwargs
    assert "database" not in kwargs


def test_connect_kwargs_token_sets_pat_authenticator() -> None:
    kwargs = SnowflakeConnector()._connect_kwargs(
        Connection(
            name="wh",
            kind="snowflake",
            options={"account": "a", "user": "u", "token": "pat"},
        )
    )
    assert kwargs["token"] == "pat"
    assert kwargs["authenticator"] == "PROGRAMMATIC_ACCESS_TOKEN"
    assert "password" not in kwargs


def test_connect_kwargs_password_and_token_skips_pat_default() -> None:
    kwargs = SnowflakeConnector()._connect_kwargs(_password_conn(token="pat"))
    assert kwargs["password"] == "p"
    assert kwargs["token"] == "pat"
    assert "authenticator" not in kwargs


def test_connect_kwargs_explicit_authenticator_wins() -> None:
    kwargs = SnowflakeConnector()._connect_kwargs(
        Connection(
            name="wh",
            kind="snowflake",
            options={"account": "a", "user": "u", "token": "pat", "authenticator": "oauth"},
        )
    )
    assert kwargs["authenticator"] == "oauth"


# --- execute ---


def test_execute_rewrites_placeholders(monkeypatch: pytest.MonkeyPatch) -> None:
    cursor = _FakeCursor()
    sdk = _install_sdk(monkeypatch, _FakeSDK(_FakeConn(cursor)))
    result = run_sql(
        _password_conn(warehouse="WH", database="DB"),
        "select :id as id, :name as name",
        {"id": 7, "name": "Ada"},
    )
    sql, bound = cursor.calls[0]
    assert sql == "select %(id)s as id, %(name)s as name"
    assert bound == {"id": 7, "name": "Ada"}
    assert sdk.kwargs["account"] == "a"
    assert sdk.kwargs["warehouse"] == "WH"
    assert sdk.kwargs["database"] == "DB"
    assert result == TableResult(headers=("id", "name"), rows=(("7", "Ada"),))
    assert cursor.closed is True
    assert sdk.conn is not None and sdk.conn.closed is True


def test_execute_leaves_casts_and_binds_only_placeholders(monkeypatch: pytest.MonkeyPatch) -> None:
    cursor = _FakeCursor()
    _install_sdk(monkeypatch, _FakeSDK(_FakeConn(cursor)))
    run_sql(
        _password_conn(),
        "select :id::int as n, :name as name",
        {"id": 7, "name": "Ada", "unused": "nope"},
    )
    sql, bound = cursor.calls[0]
    assert sql == "select %(id)s::int as n, %(name)s as name"
    assert bound == {"id": 7, "name": "Ada"}


def test_execute_missing_params_never_connects(monkeypatch: pytest.MonkeyPatch) -> None:
    sdk = _install_sdk(monkeypatch, _FakeSDK(_FakeConn(_FakeCursor())))
    with pytest.raises(ConfigError, match="SQL is missing parameters: id"):
        run_sql(_password_conn(), "select :id as id")
    assert sdk.connect_calls == 0


def test_execute_empty_result(monkeypatch: pytest.MonkeyPatch) -> None:
    cursor = _FakeCursor(rows=None, description=None)
    _install_sdk(monkeypatch, _FakeSDK(_FakeConn(cursor)))
    result = run_sql(_password_conn(), "select 1")
    assert result == TableResult(headers=(), rows=())


def test_execute_null_cells(monkeypatch: pytest.MonkeyPatch) -> None:
    cursor = _FakeCursor(rows=((None, 1),), description=(("a",), ("b",)))
    _install_sdk(monkeypatch, _FakeSDK(_FakeConn(cursor)))
    result = run_sql(_password_conn(), "select 1")
    assert result.rows == (("", "1"),)


def test_execute_uses_programmatic_token(monkeypatch: pytest.MonkeyPatch) -> None:
    cursor = _FakeCursor()
    sdk = _install_sdk(monkeypatch, _FakeSDK(_FakeConn(cursor)))
    run_sql(
        Connection(
            name="wh",
            kind="snowflake",
            options={"account": "a", "user": "u", "token": "pat"},
        ),
        "select 1",
    )
    assert sdk.kwargs["token"] == "pat"
    assert sdk.kwargs["authenticator"] == "PROGRAMMATIC_ACCESS_TOKEN"


def test_execute_missing_extra(monkeypatch: pytest.MonkeyPatch) -> None:
    def _boom() -> Any:
        raise ImportError("no sdk")

    monkeypatch.setattr("sliger.connectors.snowflake.load_snowflake_sdk", _boom)
    with pytest.raises(ConfigError, match=r"sliger\[snowflake\]"):
        run_sql(_password_conn(), "select 1")


def test_execute_wraps_connect_errors(monkeypatch: pytest.MonkeyPatch) -> None:
    _install_sdk(monkeypatch, _FakeSDK(connect_error=RuntimeError("warehouse suspended")))
    with pytest.raises(ConfigError, match="Snowflake failed: warehouse suspended"):
        run_sql(_password_conn(), "select 1")


def test_execute_wraps_query_errors_and_closes(monkeypatch: pytest.MonkeyPatch) -> None:
    cursor = _FakeCursor(fail=RuntimeError("bad sql"))
    conn = _FakeConn(cursor)
    _install_sdk(monkeypatch, _FakeSDK(conn))
    with pytest.raises(ConfigError, match="Snowflake failed: bad sql"):
        run_sql(_password_conn(), "select 1")
    assert cursor.closed is True
    assert conn.closed is True


def test_connector_class_and_sdk_helper_exist() -> None:
    assert SnowflakeConnector.extra == "snowflake"
    assert SnowflakeConnector.aliases == ("sf",)
    assert callable(load_snowflake_sdk)


# --- live ---


def _live_spec() -> dict[str, str]:
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
    return spec


@pytest.fixture
def sf_live() -> Connection:
    if not live_flag("SLIGER_SNOWFLAKE_LIVE"):
        pytest.skip("set SLIGER_LIVE=1 or SLIGER_SNOWFLAKE_LIVE=1")
    pytest.importorskip("snowflake.connector")
    return parse_connections({"wh": _live_spec()})["wh"]


@contextmanager
def _jinja_ctx(connection: Connection, data: dict[str, Any] | None = None):
    ctx = FunctionContext(
        data=data or {},
        connections={"wh": connection},
        default_connection="wh",
    )
    token = set_context(ctx)
    try:
        yield load_jinja_environment()
    finally:
        reset_context(token)


def test_live_snowflake_select_one(sf_live: Connection) -> None:
    result = run_sql(sf_live, "select 1 as n")
    assert result.rows[0][0] == "1"


def test_live_snowflake_named_params(sf_live: Connection) -> None:
    result = run_sql(sf_live, "select :name as name, :n as n", {"name": "Ada", "n": 7})
    assert result.rows == (("Ada", "7"),)


def test_live_snowflake_quoted_lowercase_alias(sf_live: Connection) -> None:
    result = run_sql(sf_live, 'select 1 as "n"')
    assert result.headers[0] == "n"


def test_live_snowflake_union_rows(sf_live: Connection) -> None:
    result = run_sql(
        sf_live,
        "select :left as name, 1 as n union all select :right as name, 2 as n",
        {"left": "Ada", "right": "Bob"},
    )
    names = {row[0] for row in result.rows}
    assert names == {"Ada", "Bob"}
    assert len(result.rows) == 2


def test_live_snowflake_null_and_types(sf_live: Connection) -> None:
    result = run_sql(
        sf_live,
        "select :n as n, :flag as flag, null as missing",
        {"n": 3, "flag": True},
    )
    row = result.rows[0]
    assert row[0] == "3"
    assert row[1] in {"True", "true", "1"}
    assert row[2] == ""


def test_live_snowflake_current_user(sf_live: Connection) -> None:
    result = run_sql(sf_live, "select current_user() as u")
    assert result.rows[0][0].upper() == os.environ["SNOWFLAKE_USER"].upper()


def test_live_snowflake_current_account(sf_live: Connection) -> None:
    result = run_sql(
        sf_live,
        "select current_organization_name() as org, current_account_name() as name",
    )
    org, name = result.rows[0][0].upper(), result.rows[0][1].upper()
    configured = os.environ["SNOWFLAKE_ACCOUNT"].upper()
    assert org in configured
    assert name in configured


def test_live_snowflake_missing_param(sf_live: Connection) -> None:
    with pytest.raises(ConfigError, match="SQL is missing parameters: n"):
        run_sql(sf_live, "select :n as n")


def test_live_snowflake_bad_sql(sf_live: Connection) -> None:
    with pytest.raises(ConfigError, match="Snowflake failed"):
        run_sql(sf_live, "selct broken")


def test_live_snowflake_jinja_sql(sf_live: Connection) -> None:
    with _jinja_ctx(sf_live, {"name": "Ada"}) as env:
        boxed = render_box(env, "{{ sql('select :name as name') }}")
    assert isinstance(boxed, TableResult)
    assert boxed.rows[0][0] == "Ada"


def test_live_snowflake_jinja_union(sf_live: Connection) -> None:
    with _jinja_ctx(sf_live, {"left": "Ada", "right": "Bob"}) as env:
        boxed = render_box(
            env,
            "{{ sql('select :left as name union all select :right as name') }}",
        )
    assert isinstance(boxed, TableResult)
    assert {row[0] for row in boxed.rows} == {"Ada", "Bob"}


def test_live_snowflake_sample_tpch(sf_live: Connection) -> None:
    try:
        result = run_sql(
            sf_live,
            "select c_custkey, c_name from snowflake_sample_data.tpch_sf1.customer "
            "where c_custkey = :id",
            {"id": 1},
        )
    except ConfigError as exc:
        pytest.skip(f"sample data or warehouse unavailable: {exc}")
    assert result.rows
    assert result.rows[0][0] == "1"
    assert result.rows[0][1]


def test_live_snowflake_toml_config(tmp_path: Path, sf_live: Connection) -> None:
    token = os.environ.get("SNOWFLAKE_TOKEN")
    password = os.environ.get("SNOWFLAKE_PASSWORD")
    path = tmp_path / "config.toml"
    lines = [
        "[connections.snow]",
        'type = "snowflake"',
        f'account = "{sf_live.option("account")}"',
        f'user = "{sf_live.option("user")}"',
    ]
    if token:
        lines.append('token = "env:SNOWFLAKE_TOKEN"')
    if password:
        lines.append('password = "env:SNOWFLAKE_PASSWORD"')
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    loaded = load_connections(path)
    result = run_sql(loaded["snow"], "select 1 as n")
    assert result.rows[0][0] == "1"
