"""Shared types for SQL warehouse connectors."""

from __future__ import annotations

import os
import re
from abc import ABC, abstractmethod
from collections.abc import Mapping
from dataclasses import dataclass, field
from typing import Any, ClassVar

from sliger.exceptions import ConfigError
from sliger.results import TableResult

# Named bind parameters in the sliger SQL dialect: :account_uuid
# Negative lookbehind keeps PostgreSQL-style ::casts intact.
PLACEHOLDER = re.compile(r"(?<!:):([A-Za-z_][A-Za-z0-9_]*)")


@dataclass(frozen=True)
class Connection:
    """A named, parsed connection from TOML ``[connections.<name>]``."""

    name: str
    kind: str
    url: str = ""
    options: Mapping[str, Any] = field(default_factory=dict)

    def option(self, key: str, default: Any = None) -> Any:
        if key in self.options:
            return self.options[key]
        return default


class Connector(ABC):
    """One warehouse backend.

    Adding a connector is: implement this class, list it in
    ``sliger.connectors.registry.BUILTIN``, and (if the SDK is heavy) add a
    ``[project.optional-dependencies]`` extra with the same name as ``extra``.
    """

    name: ClassVar[str]
    aliases: ClassVar[tuple[str, ...]] = ()
    extra: ClassVar[str | None] = None

    @abstractmethod
    def parse(self, name: str, spec: Mapping[str, Any]) -> Connection:
        """Validate a TOML table (secrets already resolved) into a Connection."""

    @abstractmethod
    def execute(self, connection: Connection, query: str, params: Mapping[str, Any]) -> TableResult:
        """Run ``query`` with sliger ``:named`` parameters."""


def resolve_secret(value: str) -> str:
    """Resolve ``env:NAME`` to an environment variable; pass other strings through."""
    if value.startswith("env:"):
        name = value[4:]
        if name not in os.environ:
            raise ConfigError(f"Environment variable {name} is not set")
        return os.environ[name]
    return value


def resolve_spec(spec: Mapping[str, Any]) -> dict[str, Any]:
    """Copy a connection table, resolving ``env:`` on string values."""
    resolved: dict[str, Any] = {}
    for key, value in spec.items():
        if isinstance(value, str):
            resolved[key] = resolve_secret(value)
        else:
            resolved[key] = value
    return resolved


def named_placeholders(query: str) -> tuple[str, ...]:
    seen: dict[str, None] = {}
    for name in PLACEHOLDER.findall(query):
        seen.setdefault(name, None)
    return tuple(seen)


def require_params(query: str, params: Mapping[str, Any]) -> None:
    missing = [name for name in named_placeholders(query) if name not in params]
    if missing:
        raise ConfigError(f"SQL is missing parameters: {', '.join(missing)}")


def rewrite_placeholders(query: str, style: str) -> str:
    """Rewrite ``:name`` placeholders to another bind style (``@name`` or ``%(name)s``)."""
    if style == ":":
        return query
    if style == "@":
        return PLACEHOLDER.sub(r"@\1", query)
    if style == "%":
        return PLACEHOLDER.sub(r"%(\1)s", query)
    raise ConfigError(f"Unknown SQL placeholder style {style!r}")


def cells_from_rows(
    headers: list[str] | tuple[str, ...],
    rows: list[tuple[Any, ...]] | tuple[tuple[Any, ...], ...],
) -> TableResult:
    data = tuple(tuple("" if cell is None else str(cell) for cell in row) for row in rows)
    return TableResult(headers=tuple(headers), rows=data)


def missing_extra(connector: Connector) -> ConfigError:
    extra = connector.extra or connector.name
    return ConfigError(
        f"{connector.name} support requires the {extra!r} extra. "
        f"Install with: pip install 'sliger[{extra}]'  (or uv add sliger --extra {extra})"
    )
