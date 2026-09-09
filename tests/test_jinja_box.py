from __future__ import annotations

import sqlite3
from collections.abc import Iterator
from contextlib import contextmanager
from pathlib import Path

import pytest

from sliger.connections import Connection
from sliger.context import FunctionContext, reset_context, set_context
from sliger.exceptions import ConfigError
from sliger.jinja_utils import load_connections, load_jinja_environment, render_box
from sliger.results import ErrorResult, RepeatDirective, ScalarResult, TableResult


@contextmanager
def active_context(ctx: FunctionContext) -> Iterator[FunctionContext]:
    token = set_context(ctx)
    try:
        yield ctx
    finally:
        reset_context(token)


def test_render_box_whole_box_expression() -> None:
    env = load_jinja_environment()
    result = render_box(env, "{{ 1+1 }}")
    assert result == ScalarResult("2")


def test_render_box_sliger_repeat() -> None:
    env = load_jinja_environment()
    result = render_box(env, '{{ sliger_repeat("deals") }}')
    assert result == RepeatDirective(key="deals")


def test_render_box_string_template() -> None:
    env = load_jinja_environment()
    result = render_box(env, "Hi {{ name }}", {"name": "Ada"})
    assert result == ScalarResult("Hi Ada")


def test_render_box_bad_template_returns_error() -> None:
    env = load_jinja_environment()
    result = render_box(env, "{{ unterminated")
    assert isinstance(result, ErrorResult)
    assert result.formula == "{{ unterminated"


def test_load_connections_from_toml(tmp_path: Path) -> None:
    path = tmp_path / "config.toml"
    path.write_text('[connections.warehouse]\ntype = "sqlite"\nurl = "sqlite:///:memory:"\n')
    connections = load_connections(path)
    assert set(connections) == {"warehouse"}
    assert connections["warehouse"].kind == "sqlite"
    assert connections["warehouse"].url == "sqlite:///:memory:"


def test_load_connections_none() -> None:
    assert load_connections(None) == {}


def test_sql_global_with_function_context(tmp_path: Path) -> None:
    db = tmp_path / "app.db"
    raw = sqlite3.connect(db)
    raw.execute("CREATE TABLE deals (name TEXT)")
    raw.execute("INSERT INTO deals VALUES ('Ada')")
    raw.commit()
    raw.close()

    env = load_jinja_environment()
    ctx = FunctionContext(
        data={},
        connections={"main": Connection(name="main", kind="sqlite", url=str(db))},
    )
    with active_context(ctx):
        result = render_box(env, "{{ sql('SELECT name FROM deals') }}")
    assert isinstance(result, TableResult)
    assert result.rows == (("Ada",),)


def test_sql_global_without_context_raises() -> None:
    env = load_jinja_environment()
    with pytest.raises(ConfigError, match="No sliger function context is active"):
        render_box(env, "{{ sql('SELECT 1') }}")
