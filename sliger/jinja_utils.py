"""Jinja2 environment setup and template rendering helpers."""

from __future__ import annotations

import importlib
import logging
import re
import sys
import time
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any

from jinja2 import BaseLoader, Environment, TemplateError

from sliger.connections import Connection, parse_connections
from sliger.context import bind_function, sliger_repeat, sql_global
from sliger.exceptions import ConfigError
from sliger.results import ErrorResult, ScalarResult, normalize_result

try:
    import tomllib
except ModuleNotFoundError:  # pragma: no cover - Python 3.10
    import tomli as tomllib  # type: ignore[no-redef]

logger = logging.getLogger(__name__)

_SINGLE_EXPR = re.compile(r"^\{\{(.*)\}\}$", re.DOTALL)


class StringLoader(BaseLoader):
    """Jinja loader that treats the template name as the template source."""

    def get_source(
        self, environment: Environment, template: str
    ) -> tuple[str, None, Callable[[], bool]]:
        return template, None, lambda: True


def ordinal(n: int) -> str:
    """Return the English ordinal suffix for a day of the month."""
    return f"{n:d}{'tsnrhtdd'[(n // 10 % 10 != 1) * (n % 10 < 4) * n % 10 :: 4]}"


def strftime_with_ordinal(fmt: str, t: time.struct_time | None = None) -> str:
    """Format a time tuple, supporting ``%O`` as an ordinal day (1st, 2nd, …)."""
    if t is None:
        t = time.localtime()
    return time.strftime(fmt.replace("%O", ordinal(t.tm_mday)), t)


def _import_function(dotted_path: str) -> Callable[..., Any]:
    if "." not in dotted_path:
        raise ImportError(f"Function path '{dotted_path}' must be of the form 'module.function'")
    mod_name, func_name = dotted_path.rsplit(".", 1)
    module = importlib.import_module(mod_name)
    func = getattr(module, func_name)
    if not callable(func):
        raise ImportError(f"'{dotted_path}' is not callable")
    return func


def load_toml_config(config_path: Path) -> dict[str, Any]:
    try:
        with config_path.open("rb") as fh:
            return tomllib.load(fh)
    except FileNotFoundError as exc:
        raise ConfigError(f"Jinja config file not found: {config_path}") from exc
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"Invalid TOML in {config_path}: {exc}") from exc


def load_function_map(config_path: Path) -> dict[str, Callable[..., Any]]:
    """Load ``[function_map]`` from a TOML file and import the named callables."""
    config = load_toml_config(config_path)
    function_map = config.get("function_map", {})
    if not isinstance(function_map, dict):
        raise ConfigError("'function_map' in config must be a table")

    config_dir = str(config_path.parent.resolve())
    if config_dir not in sys.path:
        sys.path.insert(0, config_dir)

    loaded: dict[str, Callable[..., Any]] = {}
    for name, function_name in function_map.items():
        if not isinstance(function_name, str):
            raise ConfigError(f"function_map.{name} must be a string of the form 'module.function'")
        try:
            loaded[name] = bind_function(name, _import_function(function_name))
        except (ImportError, AttributeError) as exc:
            raise ConfigError(
                f"Could not load Jinja function '{name}' from '{function_name}': {exc}"
            ) from exc
        logger.debug("Loaded Jinja function %s <- %s", name, function_name)
    return loaded


def load_connections(config_path: Path | None) -> dict[str, Connection]:
    if config_path is None:
        return {}
    config = load_toml_config(Path(config_path))
    return parse_connections(config.get("connections"))


def load_jinja_environment(config_path: Path | None = None) -> Environment:
    """Create a Jinja environment, optionally with custom functions from TOML."""
    function_map: dict[str, Callable[..., Any]] = {}
    if config_path is not None:
        function_map = load_function_map(Path(config_path))

    env = Environment(loader=StringLoader(), autoescape=False)
    env.globals.update(function_map)
    env.globals.setdefault("strftime", strftime_with_ordinal)
    env.globals["sql"] = sql_global
    env.globals["sliger_repeat"] = sliger_repeat
    return env


def _render_data(data: Mapping[str, Any] | None) -> dict[str, Any]:
    render_data: dict[str, Any] = {
        "now": time.localtime(),
        "strftime": strftime_with_ordinal,
    }
    if data:
        render_data.update(data)
    return render_data


def render_jinja_in_string(
    jinja_env: Environment,
    template_string: str,
    data: Mapping[str, Any] | None = None,
) -> str:
    """Render a template string. ``now`` and ``strftime`` are always available."""
    try:
        return jinja_env.from_string(template_string).render(**_render_data(data))
    except TemplateError as exc:
        raise ConfigError(f"Failed to render Jinja template: {exc}") from exc


def render_box(
    jinja_env: Environment,
    template_string: str,
    data: Mapping[str, Any] | None = None,
) -> Any:
    """Render a text box. A whole-box ``{{ expr }}`` keeps a typed result."""
    stripped = template_string.strip()
    match = _SINGLE_EXPR.fullmatch(stripped)
    render_data = _render_data(data)
    if match and stripped.count("{{") == 1:
        expression = match.group(1).strip()
        if expression and not expression.startswith("%"):
            try:
                value = jinja_env.compile_expression(expression)(**render_data)
                return normalize_result(value)
            except ConfigError:
                raise
            except Exception as exc:
                return ErrorResult(str(exc), formula=stripped)
    try:
        text = jinja_env.from_string(template_string).render(**render_data)
        return ScalarResult(text)
    except TemplateError as exc:
        return ErrorResult(str(exc), formula=template_string)
