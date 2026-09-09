"""Map connection types to connector classes.

First-party warehouses go in ``BUILTIN`` (and ``ALIASES`` if needed). Third-party
packages can expose ``sliger.connectors`` entry points or call :func:`register`.
"""

from __future__ import annotations

import importlib
from collections.abc import Iterable, Mapping
from typing import Any
from urllib.parse import urlparse

from sliger.connectors.base import Connection, Connector, resolve_spec
from sliger.exceptions import ConfigError
from sliger.results import TableResult

# kind -> "module:Class". Add one line here when you add a warehouse.
BUILTIN: dict[str, str] = {
    "sqlite": "sliger.connectors.sqlite:SQLiteConnector",
}

# alias -> canonical kind
ALIASES: dict[str, str] = {
    "sqlite3": "sqlite",
}

_runtime: dict[str, Connector | type[Connector] | str] = {}
_alias_runtime: dict[str, str] = {}
_instances: dict[str, Connector] = {}


def register(
    kind: str,
    connector: Connector | type[Connector] | str,
    *,
    aliases: Iterable[str] = (),
) -> None:
    """Register a connector for this process (tests and in-process plugins)."""
    canonical = kind.lower()
    _runtime[canonical] = connector
    _instances.pop(canonical, None)
    for alias in aliases:
        _alias_runtime[alias.lower()] = canonical
        _instances.pop(alias.lower(), None)


def reset_registry() -> None:
    """Drop runtime registrations and cached instances. Builtins stay."""
    _runtime.clear()
    _alias_runtime.clear()
    _instances.clear()


def canonical_kind(kind: str) -> str:
    lowered = kind.lower()
    if lowered in _alias_runtime:
        return _alias_runtime[lowered]
    if lowered in ALIASES:
        return ALIASES[lowered]
    return lowered


def known_types() -> tuple[str, ...]:
    kinds = set(BUILTIN)
    kinds.update(_runtime)
    kinds.update(_entry_point_targets())
    return tuple(sorted(kinds))


def get_connector(kind: str) -> Connector:
    canonical = canonical_kind(kind)
    cached = _instances.get(canonical)
    if cached is not None:
        return cached
    target = _lookup(canonical)
    if target is None:
        known = ", ".join(known_types()) or "(none)"
        raise ConfigError(
            f"Unknown connection type {kind!r}. Known types: {known}. "
            "Add sliger/connectors/<name>.py and list it in "
            "sliger.connectors.registry.BUILTIN, or register() it, or expose a "
            "sliger.connectors entry point."
        )
    connector = _instantiate(target)
    _instances[canonical] = connector
    _instances[connector.name] = connector
    for alias in connector.aliases:
        _instances[alias] = connector
    return connector


def infer_kind(spec: Mapping[str, Any]) -> str:
    explicit = spec.get("type") or spec.get("kind")
    if explicit:
        return str(explicit).lower()
    url = spec.get("url")
    if isinstance(url, str) and "://" in url:
        scheme = urlparse(url).scheme.lower()
        if scheme in {"", "file"}:
            return "sqlite"
        if scheme:
            return scheme
    return "sqlite"


def parse_connections(raw: Any) -> dict[str, Connection]:
    if raw is None:
        return {}
    if not isinstance(raw, dict):
        raise ConfigError("'connections' must be a table")
    parsed: dict[str, Connection] = {}
    for name, spec in raw.items():
        if not isinstance(spec, dict):
            raise ConfigError(f"connections.{name} must be a table")
        resolved = resolve_spec(spec)
        kind = infer_kind(resolved)
        parsed[name] = get_connector(kind).parse(str(name), resolved)
    return parsed


def run_sql(
    connection: Connection, query: str, params: dict[str, Any] | None = None
) -> TableResult:
    return get_connector(connection.kind).execute(connection, query, params or {})


def _lookup(kind: str) -> Connector | type[Connector] | str | None:
    if kind in _runtime:
        return _runtime[kind]
    if kind in BUILTIN:
        return BUILTIN[kind]
    entry_points = _entry_point_targets()
    if kind in entry_points:
        return entry_points[kind]
    return None


def _instantiate(target: Connector | type[Connector] | str) -> Connector:
    if isinstance(target, str):
        return _load_class(target)()
    if isinstance(target, type):
        return target()
    return target


def _load_class(target: str) -> type[Connector]:
    module_name, sep, attr = target.partition(":")
    if not sep or not attr:
        raise ConfigError(f"Invalid connector target {target!r}; expected 'module:Class'")
    try:
        module = importlib.import_module(module_name)
    except ImportError as exc:
        raise ConfigError(f"Could not import connector {target!r}: {exc}") from exc
    try:
        cls = getattr(module, attr)
    except AttributeError as exc:
        raise ConfigError(f"Could not load connector {target!r}: {exc}") from exc
    return cls


def _entry_point_targets() -> dict[str, str]:
    try:
        from importlib.metadata import entry_points
    except ImportError:  # pragma: no cover
        return {}
    try:
        eps = entry_points()
    except Exception:  # pragma: no cover - importlib.metadata edge cases
        return {}
    if hasattr(eps, "select"):
        group = eps.select(group="sliger.connectors")
    else:  # pragma: no cover - Python 3.10 dict API
        group = eps.get("sliger.connectors", [])  # type: ignore[union-attr]
    found: dict[str, str] = {}
    for ep in group:
        found[ep.name.lower()] = f"{ep.module}:{ep.attr}"
    return found
