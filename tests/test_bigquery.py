from __future__ import annotations

import os
from collections.abc import Iterator
from types import SimpleNamespace
from typing import Any

import pytest

from live_flags import live_flag
from sliger.connectors import parse_connections, reset_registry, run_sql
from sliger.connectors.base import Connection
from sliger.connectors.bigquery import BigQueryConnector, _bq_value, load_bigquery_sdk
from sliger.exceptions import ConfigError
from sliger.results import TableResult


@pytest.fixture(autouse=True)
def _isolate_registry() -> Iterator[None]:
    reset_registry()
    yield
    reset_registry()


class _FakeJob:
    def __init__(self, headers: tuple[str, ...], rows: tuple[tuple[Any, ...], ...]) -> None:
        self.schema = [SimpleNamespace(name=name) for name in headers]
        self._rows = rows

    def result(self) -> _FakeJob:
        return self

    def __iter__(self):
        return iter(self._rows)


class _FakeClient:
    def __init__(self) -> None:
        self.calls: list[tuple[str, Any]] = []

    def query(self, sql: str, job_config: Any = None) -> _FakeJob:
        self.calls.append((sql, job_config))
        params = getattr(job_config, "query_parameters", [])
        if params:
            row = tuple(p.value for p in params)
            headers = tuple(p.name for p in params)
        else:
            headers = ("n",)
            row = (1,)
        return _FakeJob(headers, (row,))


class _FakeParam:
    def __init__(self, name: str, type_: str, value: Any) -> None:
        self.name = name
        self.type_ = type_
        self.value = value


class _FakeBQ:
    def __init__(self, client: _FakeClient) -> None:
        self._client = client
        self.Client = lambda **kwargs: self._record_client(kwargs)
        self.QueryJobConfig = lambda: SimpleNamespace(query_parameters=[], default_dataset=None)
        self.ScalarQueryParameter = _FakeParam
        self.client_kwargs: dict[str, Any] = {}

    def _record_client(self, kwargs: dict[str, Any]) -> _FakeClient:
        self.client_kwargs = kwargs
        return self._client


def test_parse_project_and_optional_fields() -> None:
    parsed = parse_connections(
        {
            "wh": {
                "type": "bigquery",
                "project": "acme-prod",
                "dataset": "metrics",
                "location": "EU",
                "credentials_path": "/tmp/sa.json",
            }
        }
    )
    conn = parsed["wh"]
    assert conn.kind == "bigquery"
    assert conn.option("project") == "acme-prod"
    assert conn.option("dataset") == "metrics"
    assert conn.option("location") == "EU"
    assert conn.option("credentials_path") == "/tmp/sa.json"


def test_parse_url() -> None:
    parsed = parse_connections({"wh": {"url": "bigquery://acme-prod/metrics?location=US"}})
    conn = parsed["wh"]
    assert conn.kind == "bigquery"
    assert conn.option("project") == "acme-prod"
    assert conn.option("dataset") == "metrics"
    assert conn.option("location") == "US"


def test_parse_bq_alias() -> None:
    parsed = parse_connections({"wh": {"type": "bq", "project": "p"}})
    assert parsed["wh"].kind == "bigquery"


def test_parse_missing_project() -> None:
    with pytest.raises(ConfigError, match="needs project="):
        parse_connections({"wh": {"type": "bigquery"}})


def test_parse_env_project(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SLIGER_BQ_PROJECT", "from-env")
    parsed = parse_connections({"wh": {"type": "bigquery", "project": "env:SLIGER_BQ_PROJECT"}})
    assert parsed["wh"].option("project") == "from-env"


def test_bq_value_types() -> None:
    assert _bq_value(None) == ("STRING", None)
    assert _bq_value(True) == ("BOOL", True)
    assert _bq_value(3) == ("INT64", 3)
    assert _bq_value(1.5) == ("FLOAT64", 1.5)
    assert _bq_value(b"x") == ("BYTES", b"x")
    assert _bq_value("ada") == ("STRING", "ada")


def test_execute_rewrites_placeholders(monkeypatch: pytest.MonkeyPatch) -> None:
    client = _FakeClient()
    sdk = _FakeBQ(client)
    monkeypatch.setattr("sliger.connectors.bigquery.load_bigquery_sdk", lambda: sdk)
    connection = Connection(
        name="wh",
        kind="bigquery",
        url="bigquery://acme",
        options={"project": "acme", "dataset": "metrics", "location": "EU"},
    )
    result = run_sql(
        connection,
        "select :id as id, :name as name",
        {"id": 7, "name": "Ada"},
    )
    sql, job_config = client.calls[0]
    assert sql == "select @id as id, @name as name"
    assert job_config.default_dataset == "acme.metrics"
    assert sdk.client_kwargs["project"] == "acme"
    assert sdk.client_kwargs["location"] == "EU"
    assert result == TableResult(headers=("id", "name"), rows=(("7", "Ada"),))


def test_execute_missing_extra(monkeypatch: pytest.MonkeyPatch) -> None:
    def _boom() -> Any:
        raise ImportError("no sdk")

    monkeypatch.setattr("sliger.connectors.bigquery.load_bigquery_sdk", _boom)
    connection = Connection(
        name="wh", kind="bigquery", url="bigquery://acme", options={"project": "acme"}
    )
    with pytest.raises(ConfigError, match=r"sliger\[bigquery\]"):
        run_sql(connection, "select 1")


def test_execute_wraps_sdk_errors(monkeypatch: pytest.MonkeyPatch) -> None:
    class _BadClient:
        def query(self, sql: str, job_config: Any = None) -> Any:
            raise RuntimeError("quota")

    class _BadSDK:
        ScalarQueryParameter = _FakeParam

        @staticmethod
        def Client(**kwargs: Any) -> _BadClient:
            return _BadClient()

        @staticmethod
        def QueryJobConfig() -> Any:
            return SimpleNamespace(query_parameters=[], default_dataset=None)

    monkeypatch.setattr("sliger.connectors.bigquery.load_bigquery_sdk", lambda: _BadSDK)
    connection = Connection(
        name="wh", kind="bigquery", url="bigquery://acme", options={"project": "acme"}
    )
    with pytest.raises(ConfigError, match="BigQuery failed: quota"):
        run_sql(connection, "select 1")


def test_connector_class_and_sdk_helper_exist() -> None:
    assert BigQueryConnector.extra == "bigquery"
    assert callable(load_bigquery_sdk)


@pytest.mark.skipif(not live_flag("SLIGER_BQ_LIVE"), reason="set SLIGER_LIVE=1 or SLIGER_BQ_LIVE=1")
def test_live_bigquery_select_one() -> None:
    pytest.importorskip("google.cloud.bigquery")
    project = (
        os.environ.get("SLIGER_BQ_PROJECT")
        or os.environ.get("GOOGLE_CLOUD_PROJECT")
        or os.environ.get("GOOGLE_CLOUD_QUOTA_PROJECT")
    )
    if not project:
        pytest.skip("needs SLIGER_BQ_PROJECT or GOOGLE_CLOUD_PROJECT")
    parsed = parse_connections({"wh": {"type": "bigquery", "project": project}})
    result = run_sql(parsed["wh"], "select 1 as n")
    assert result.headers[0] == "n"
    assert result.rows[0][0] == "1"
    named = run_sql(
        parsed["wh"],
        "select :name as name, :n as n",
        {"name": "Ada", "n": 7},
    )
    assert named == TableResult(headers=("name", "n"), rows=(("Ada", "7"),))
    public = run_sql(
        parsed["wh"],
        "select word from `bigquery-public-data.samples.shakespeare` where word = :word limit 1",
        {"word": "the"},
    )
    assert public.rows == (("the",),)
