from __future__ import annotations

from pathlib import Path

import pytest

from sliger.connections import Connection, parse_connections, resolve_secret, run_sql, sqlite_path
from sliger.exceptions import ConfigError
from sliger.results import TableResult


def test_resolve_secret_from_env(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SLIGER_TEST_DB", "sqlite:///:memory:")
    assert resolve_secret("env:SLIGER_TEST_DB") == "sqlite:///:memory:"


def test_resolve_secret_passthrough() -> None:
    assert resolve_secret("sqlite:///:memory:") == "sqlite:///:memory:"


def test_resolve_secret_missing_env() -> None:
    with pytest.raises(ConfigError, match="Environment variable SLIGER_MISSING_DB_XYZ is not set"):
        resolve_secret("env:SLIGER_MISSING_DB_XYZ")


def test_parse_connections_none_is_empty() -> None:
    assert parse_connections(None) == {}


def test_parse_connections_not_a_table() -> None:
    with pytest.raises(ConfigError, match="'connections' must be a table"):
        parse_connections(["not", "a", "table"])


def test_parse_connections_entry_not_a_table() -> None:
    with pytest.raises(ConfigError, match="connections.main must be a table"):
        parse_connections({"main": "sqlite:///:memory:"})


def test_parse_connections_missing_url() -> None:
    with pytest.raises(ConfigError, match="connections.main.url is required"):
        parse_connections({"main": {"type": "sqlite"}})


def test_parse_connections_empty_url() -> None:
    with pytest.raises(ConfigError, match="connections.main.url is required"):
        parse_connections({"main": {"url": ""}})


def test_parse_connections_resolves_env_url(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SLIGER_TEST_DB", ":memory:")
    parsed = parse_connections({"main": {"type": "SQLite", "url": "env:SLIGER_TEST_DB"}})
    assert parsed["main"] == Connection(name="main", kind="sqlite", url=":memory:")


def test_sqlite_path_memory_aliases() -> None:
    assert sqlite_path(":memory:") == ":memory:"
    assert sqlite_path("sqlite:///:memory:") == ":memory:"
    assert sqlite_path("sqlite://memory") == ":memory:"


def test_sqlite_path_file_paths(tmp_path: Path) -> None:
    db = tmp_path / "app.db"
    assert sqlite_path(str(db)) == str(db)
    assert sqlite_path(f"sqlite:///{db}") == str(db)
    assert sqlite_path(f"file://{db}") == str(db)


def test_sqlite_path_unsupported_scheme() -> None:
    with pytest.raises(ConfigError, match="Unsupported SQL URL scheme: postgres"):
        sqlite_path("postgres://localhost/db")


def test_run_sql_memory_named_params() -> None:
    connection = Connection(name="mem", kind="sqlite", url="sqlite:///:memory:")
    result = run_sql(connection, "SELECT :x AS x, :y AS y", {"x": 1, "y": "ada"})
    assert result == TableResult(headers=("x", "y"), rows=(("1", "ada"),))


def test_run_sql_missing_params() -> None:
    connection = Connection(name="mem", kind="sqlite", url=":memory:")
    with pytest.raises(ConfigError, match="SQL is missing parameters: n"):
        run_sql(connection, "SELECT :n AS n")


def test_run_sql_unknown_type() -> None:
    connection = Connection(name="pg", kind="postgres", url="postgres://localhost/db")
    with pytest.raises(ConfigError, match="Unknown connection type 'postgres'"):
        run_sql(connection, "SELECT 1")


def test_run_sql_bad_syntax() -> None:
    connection = Connection(name="mem", kind="sqlite", url=":memory:")
    with pytest.raises(ConfigError, match="SQL failed"):
        run_sql(connection, "SELCT broken")


def test_run_sql_temp_file_create_insert_select(tmp_path: Path) -> None:
    db = tmp_path / "deals.db"
    connection = Connection(name="file", kind="sqlite", url=str(db))
    run_sql(connection, "CREATE TABLE deals (id INTEGER, name TEXT)")
    run_sql(
        connection, "INSERT INTO deals (id, name) VALUES (:id, :name)", {"id": 1, "name": "Ada"}
    )
    result = run_sql(connection, "SELECT id, name FROM deals WHERE id = :id", {"id": 1})
    assert result == TableResult(headers=("id", "name"), rows=(("1", "Ada"),))
