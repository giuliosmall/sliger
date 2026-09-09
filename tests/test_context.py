from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

import pytest

from sliger.connections import Connection, run_sql
from sliger.context import (
    FunctionContext,
    bind_function,
    get_context,
    overlay_item,
    reset_context,
    set_context,
    sliger_repeat,
    sql_global,
)
from sliger.exceptions import ConfigError
from sliger.results import RepeatDirective, TableResult


@contextmanager
def active_context(ctx: FunctionContext) -> Iterator[FunctionContext]:
    token = set_context(ctx)
    try:
        yield ctx
    finally:
        reset_context(token)


def _sqlite_connection(tmp_path: Path, name: str = "main") -> Connection:
    db = tmp_path / f"{name}.db"
    connection = Connection(name=name, kind="sqlite", url=str(db))
    run_sql(connection, "CREATE TABLE deals (id INTEGER, name TEXT)")
    run_sql(
        connection,
        "INSERT INTO deals (id, name) VALUES (:id, :name)",
        {"id": 1, "name": "Ada"},
    )
    return connection


def test_bind_function_injects_data_keys() -> None:
    def greet(name: str, title: str) -> str:
        return f"{title} {name}"

    wrapped = bind_function("greet", greet)
    ctx = FunctionContext(data={"name": "Ada", "title": "Dr"})
    with active_context(ctx):
        assert wrapped() == "Dr Ada"


def test_bind_function_injects_ctx_parameter() -> None:
    def uses_ctx(ctx: FunctionContext) -> str:
        return ctx.shape_id

    wrapped = bind_function("uses_ctx", uses_ctx)
    ctx = FunctionContext(data={}, shape_id="shape-1")
    with active_context(ctx):
        assert wrapped() == "shape-1"


def test_bind_function_missing_required_arg() -> None:
    def needs(foo: str) -> str:
        return foo

    wrapped = bind_function("needs", needs)
    with (
        active_context(FunctionContext(data={})),
        pytest.raises(ConfigError, match="Function needs\\(\\) requires foo"),
    ):
        wrapped()


def test_bind_function_without_context_calls_original() -> None:
    def greet(name: str) -> str:
        return f"hi {name}"

    wrapped = bind_function("greet", greet)
    assert wrapped("Zed") == "hi Zed"


def test_bind_function_cache_returns_same_object() -> None:
    def factory() -> dict[str, object]:
        return {"token": object()}

    wrapped = bind_function("factory", factory)
    with active_context(FunctionContext(data={})):
        first = wrapped()
        second = wrapped()
    assert first is second


def test_sql_with_one_connection(tmp_path: Path) -> None:
    connection = _sqlite_connection(tmp_path)
    ctx = FunctionContext(data={"id": 1}, connections={"main": connection})
    result = ctx.sql("SELECT name FROM deals WHERE id = :id")
    assert result == TableResult(headers=("name",), rows=(("Ada",),))


def test_sql_with_named_connection(tmp_path: Path) -> None:
    main = _sqlite_connection(tmp_path, "main")
    other = Connection(name="other", kind="sqlite", url=":memory:")
    ctx = FunctionContext(
        data={"id": 1},
        connections={"main": main, "other": other},
    )
    result = ctx.sql("SELECT name FROM deals WHERE id = :id", connection="main")
    assert result.rows == (("Ada",),)


def test_sql_unknown_connection(tmp_path: Path) -> None:
    ctx = FunctionContext(data={}, connections={"main": _sqlite_connection(tmp_path)})
    with pytest.raises(ConfigError, match="Unknown connection 'nope'"):
        ctx.sql("SELECT 1", connection="nope")


def test_sql_no_connections() -> None:
    ctx = FunctionContext(data={})
    with pytest.raises(ConfigError, match="sql\\(\\) needs a connection name"):
        ctx.sql("SELECT 1")


def test_sql_cache_returns_same_object(tmp_path: Path) -> None:
    ctx = FunctionContext(data={"id": 1}, connections={"main": _sqlite_connection(tmp_path)})
    first = ctx.sql("SELECT name FROM deals WHERE id = :id")
    second = ctx.sql("SELECT name FROM deals WHERE id = :id")
    assert first is second


def test_overlay_item_merges_dict() -> None:
    item = {"name": "Ada", "qty": 3}
    merged = overlay_item({"name": "old", "extra": 1}, item)
    assert merged == {"name": "Ada", "extra": 1, "item": item, "qty": 3}


def test_overlay_item_non_dict() -> None:
    assert overlay_item({"a": 1}, "row") == {"a": 1, "item": "row"}


def test_sliger_repeat() -> None:
    assert sliger_repeat("deals") == RepeatDirective(key="deals")


def test_get_context_without_context() -> None:
    with pytest.raises(ConfigError, match="No sliger function context is active"):
        get_context()


def test_sql_global_without_context() -> None:
    with pytest.raises(ConfigError, match="No sliger function context is active"):
        sql_global("SELECT 1")


def test_sql_global_with_context(tmp_path: Path) -> None:
    ctx = FunctionContext(data={"id": 1}, connections={"main": _sqlite_connection(tmp_path)})
    with active_context(ctx):
        result = sql_global("SELECT name FROM deals WHERE id = :id")
    assert result.rows == (("Ada",),)
