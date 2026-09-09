"""Pluggable SQL warehouse connectors.

Built-in: sqlite. Optional extras: see ``[project.optional-dependencies]``.
"""

from sliger.connectors.base import Connection, Connector, resolve_secret
from sliger.connectors.registry import (
    get_connector,
    known_types,
    parse_connections,
    register,
    reset_registry,
    run_sql,
)

__all__ = [
    "Connection",
    "Connector",
    "get_connector",
    "known_types",
    "parse_connections",
    "register",
    "reset_registry",
    "resolve_secret",
    "run_sql",
]
