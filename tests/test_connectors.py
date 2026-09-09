from __future__ import annotations

from collections.abc import Iterator, Mapping
from typing import Any

import pytest

from sliger.connectors import (
    Connection,
    Connector,
    get_connector,
    known_types,
    parse_connections,
    register,
    reset_registry,
    run_sql,
)
from sliger.connectors.base import named_placeholders, rewrite_placeholders
from sliger.exceptions import ConfigError
from sliger.results import TableResult


@pytest.fixture(autouse=True)
def _isolate_registry() -> Iterator[None]:
    reset_registry()
    yield
    reset_registry()


class _FakeConnector(Connector):
    name = "fake"
    aliases = ("faker",)
    extra = None

    def parse(self, name: str, spec: Mapping[str, Any]) -> Connection:
        return Connection(name=name, kind=self.name, url=str(spec.get("url") or "fake://"))

    def execute(self, connection: Connection, query: str, params: Mapping[str, Any]) -> TableResult:
        return TableResult(headers=("q", "n"), rows=((query, connection.name),))


def test_known_types_include_sqlite() -> None:
    assert "sqlite" in known_types()


def test_get_sqlite_alias() -> None:
    assert get_connector("sqlite3").name == "sqlite"


def test_unknown_type_lists_known() -> None:
    with pytest.raises(ConfigError, match="Unknown connection type 'athena'"):
        get_connector("athena")


def test_register_runtime_connector() -> None:
    register("fake", _FakeConnector, aliases=("faker",))
    assert "fake" in known_types()
    parsed = parse_connections({"lab": {"type": "faker", "url": "fake://lab"}})
    assert parsed["lab"].kind == "fake"
    result = run_sql(parsed["lab"], "select 1")
    assert result.rows == (("select 1", "lab"),)


def test_infer_kind_from_url_scheme() -> None:
    register("fake", _FakeConnector)
    parsed = parse_connections({"lab": {"url": "fake://warehouse"}})
    assert parsed["lab"].kind == "fake"


def test_named_placeholders_skip_cast() -> None:
    assert named_placeholders("select :id::int as n, :name") == ("id", "name")


def test_rewrite_placeholders() -> None:
    sql = "select :id as id, :name as name"
    assert rewrite_placeholders(sql, "@") == "select @id as id, @name as name"
    assert rewrite_placeholders(sql, "%") == "select %(id)s as id, %(name)s as name"
    assert rewrite_placeholders(sql, ":") == sql
    with pytest.raises(ConfigError, match="Unknown SQL placeholder style"):
        rewrite_placeholders(sql, "?")
