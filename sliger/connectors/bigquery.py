"""BigQuery connector. Requires the ``bigquery`` extra."""

from __future__ import annotations

from collections.abc import Mapping
from typing import Any
from urllib.parse import parse_qs, unquote, urlparse

from sliger.connectors.base import (
    Connection,
    Connector,
    cells_from_rows,
    missing_extra,
    named_placeholders,
    require_params,
    rewrite_placeholders,
)
from sliger.exceptions import ConfigError
from sliger.results import TableResult


def load_bigquery_sdk() -> Any:
    """Import google.cloud.bigquery. Isolated so tests can mock a missing extra."""
    from google.cloud import bigquery

    return bigquery


class BigQueryConnector(Connector):
    name = "bigquery"
    aliases = ("bq",)
    extra = "bigquery"

    def parse(self, name: str, spec: Mapping[str, Any]) -> Connection:
        url = spec.get("url") if isinstance(spec.get("url"), str) else ""
        parsed_url = _parse_url(url) if url else {}
        project = spec.get("project") or parsed_url.get("project") or spec.get("project_id") or ""
        dataset = spec.get("dataset") or parsed_url.get("dataset") or ""
        location = spec.get("location") or parsed_url.get("location") or ""
        credentials_path = (
            spec.get("credentials_path")
            or spec.get("credentials")
            or spec.get("creds")
            or parsed_url.get("credentials_path")
            or ""
        )
        if not isinstance(project, str) or not project:
            raise ConfigError(f"connections.{name} needs project= or a bigquery://PROJECT URL")
        options: dict[str, Any] = {"project": project}
        if dataset:
            options["dataset"] = dataset
        if location:
            options["location"] = location
        if credentials_path:
            options["credentials_path"] = credentials_path
        return Connection(
            name=name,
            kind=self.name,
            url=url or f"bigquery://{project}",
            options=options,
        )

    def execute(self, connection: Connection, query: str, params: Mapping[str, Any]) -> TableResult:
        require_params(query, params)
        try:
            bigquery = load_bigquery_sdk()
        except ImportError as exc:
            raise missing_extra(self) from exc

        client = self._client(connection, bigquery)
        job_config = bigquery.QueryJobConfig()
        dataset = connection.option("dataset")
        project = connection.option("project")
        if dataset:
            job_config.default_dataset = f"{project}.{dataset}"
        job_config.query_parameters = [
            _parameter(bigquery, name, params[name]) for name in named_placeholders(query)
        ]
        sql = rewrite_placeholders(query, "@")
        try:
            job = client.query(sql, job_config=job_config)
            result = job.result()
        except ConfigError:
            raise
        except Exception as exc:
            raise ConfigError(f"BigQuery failed: {exc}") from exc
        schema = getattr(result, "schema", None) or getattr(job, "schema", None) or []
        headers = tuple(field.name for field in schema)
        rows = tuple(tuple(row) for row in result)
        return cells_from_rows(headers, rows)

    def _client(self, connection: Connection, bigquery: Any) -> Any:
        kwargs: dict[str, Any] = {}
        project = connection.option("project")
        if project:
            kwargs["project"] = project
        location = connection.option("location")
        if location:
            kwargs["location"] = location
        creds_path = connection.option("credentials_path")
        if creds_path:
            from google.oauth2 import service_account

            kwargs["credentials"] = service_account.Credentials.from_service_account_file(
                str(creds_path)
            )
        return bigquery.Client(**kwargs)


def _parse_url(url: str) -> dict[str, str]:
    parsed = urlparse(url)
    if parsed.scheme not in {"", "bigquery", "bq"}:
        raise ConfigError(f"Unsupported BigQuery URL scheme: {parsed.scheme}")
    out: dict[str, str] = {}
    host = unquote(parsed.hostname or parsed.netloc or "")
    path_parts = [unquote(part) for part in parsed.path.split("/") if part]
    if host:
        out["project"] = host
        if path_parts:
            out["dataset"] = path_parts[0]
    elif path_parts:
        out["project"] = path_parts[0]
        if len(path_parts) > 1:
            out["dataset"] = path_parts[1]
    query = {key: values[-1] for key, values in parse_qs(parsed.query).items() if values}
    if "location" in query:
        out["location"] = query["location"]
    for key in ("credentials_path", "credentials", "creds"):
        if key in query:
            out["credentials_path"] = query[key]
            break
    return out


def _parameter(bigquery: Any, name: str, value: Any) -> Any:
    bq_type, coerced = _bq_value(value)
    return bigquery.ScalarQueryParameter(name, bq_type, coerced)


def _bq_value(value: Any) -> tuple[str, Any]:
    if value is None:
        return "STRING", None
    if isinstance(value, bool):
        return "BOOL", value
    if isinstance(value, int):
        return "INT64", value
    if isinstance(value, float):
        return "FLOAT64", value
    if isinstance(value, bytes):
        return "BYTES", value
    return "STRING", str(value)
