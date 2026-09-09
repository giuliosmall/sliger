"""Per-render function context, cache, and signature binding."""

from __future__ import annotations

import hashlib
import inspect
import json
from collections.abc import Callable, Mapping
from contextvars import ContextVar
from dataclasses import dataclass, field
from typing import Any

from sliger.connections import Connection, run_sql
from sliger.exceptions import ConfigError
from sliger.results import RepeatDirective, TableResult

_current: ContextVar[FunctionContext | None] = ContextVar("sliger_ctx", default=None)


class ResultCache:
    def __init__(self) -> None:
        self._store: dict[str, Any] = {}

    def key(self, name: str, args: tuple[Any, ...], kwargs: dict[str, Any]) -> str:
        payload = json.dumps({"n": name, "a": args, "k": kwargs}, default=str, sort_keys=True)
        return hashlib.sha256(payload.encode()).hexdigest()

    def get(self, key: str) -> Any | None:
        return self._store.get(key)

    def set(self, key: str, value: Any) -> None:
        self._store[key] = value


@dataclass
class FunctionContext:
    data: dict[str, Any]
    connections: dict[str, Connection] = field(default_factory=dict)
    cache: ResultCache = field(default_factory=ResultCache)
    slide_index: int = 0
    slide_id: str = ""
    shape_id: str = ""
    default_connection: str | None = None

    def sql(self, query: str, connection: str | None = None, **params: Any) -> TableResult:
        name = connection or self.default_connection
        if not name:
            if len(self.connections) == 1:
                name = next(iter(self.connections))
            else:
                raise ConfigError("sql() needs a connection name or a single [connections] entry")
        if name not in self.connections:
            raise ConfigError(f"Unknown connection {name!r}")
        bound = {key: value for key, value in self.data.items() if isinstance(key, str)}
        bound.update(params)
        cache_key = self.cache.key(f"sql:{name}", (query,), bound)
        cached = self.cache.get(cache_key)
        if isinstance(cached, TableResult):
            return cached
        result = run_sql(self.connections[name], query, bound)
        self.cache.set(cache_key, result)
        return result


def get_context() -> FunctionContext:
    ctx = _current.get()
    if ctx is None:
        raise ConfigError("No sliger function context is active")
    return ctx


def set_context(ctx: FunctionContext | None):
    return _current.set(ctx)


def reset_context(token: Any) -> None:
    _current.reset(token)


def sliger_repeat(key: str) -> RepeatDirective:
    return RepeatDirective(key=key)


def sql_global(query: str, connection: str | None = None, **params: Any) -> TableResult:
    return get_context().sql(query, connection=connection, **params)


def bind_function(name: str, func: Callable[..., Any]) -> Callable[..., Any]:
    """Fill missing arguments from FunctionContext.data; inject ctx if requested."""
    signature = inspect.signature(func)

    def wrapped(*args: Any, **kwargs: Any) -> Any:
        ctx = _current.get()
        if ctx is None:
            return func(*args, **kwargs)
        bound = signature.bind_partial(*args, **kwargs)
        for param_name, param in signature.parameters.items():
            if param_name in bound.arguments:
                continue
            if param.kind in (param.VAR_POSITIONAL, param.VAR_KEYWORD):
                continue
            if param_name in {"ctx", "context"}:
                bound.arguments[param_name] = ctx
            elif param_name in ctx.data:
                bound.arguments[param_name] = ctx.data[param_name]
            elif param.default is inspect.Parameter.empty:
                raise ConfigError(
                    f"Function {name}() requires {param_name} (pass it in --data / --data-file)"
                )
        bound.apply_defaults()
        cache_key = ctx.cache.key(name, bound.args, bound.kwargs)
        cached = ctx.cache.get(cache_key)
        if cached is not None:
            return cached
        result = func(*bound.args, **bound.kwargs)
        ctx.cache.set(cache_key, result)
        return result

    wrapped.__name__ = getattr(func, "__name__", name)
    wrapped.__doc__ = func.__doc__
    return wrapped


def overlay_item(data: Mapping[str, Any], item: Any) -> dict[str, Any]:
    merged = dict(data)
    merged["item"] = item
    if isinstance(item, dict):
        merged.update(item)
    return merged
