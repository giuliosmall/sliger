"""Jinja2 environment setup and template rendering helpers."""

from __future__ import annotations

import importlib
import logging
import sys
import time
from collections.abc import Callable, Mapping
from pathlib import Path
from typing import Any

from jinja2 import BaseLoader, Environment, TemplateError

from sliger.exceptions import ConfigError

try:
    import tomllib
except ModuleNotFoundError:  # Python 3.10
    import tomli as tomllib  # type: ignore[no-redef]

logger = logging.getLogger(__name__)


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


def load_function_map(config_path: Path) -> dict[str, Callable[..., Any]]:
    """Load ``[function_map]`` from a TOML file and import the named callables."""
    try:
        with config_path.open("rb") as fh:
            config = tomllib.load(fh)
    except FileNotFoundError as exc:
        raise ConfigError(f"Jinja config file not found: {config_path}") from exc
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"Invalid TOML in {config_path}: {exc}") from exc

    function_map = config.get("function_map", {})
    if not isinstance(function_map, dict):
        raise ConfigError(f"'function_map' in {config_path} must be a table")

    config_dir = str(config_path.parent.resolve())
    if config_dir not in sys.path:
        sys.path.insert(0, config_dir)

    loaded: dict[str, Callable[..., Any]] = {}
    for name, dotted_path in function_map.items():
        if not isinstance(dotted_path, str):
            raise ConfigError(f"function_map.{name} must be a string of the form 'module.function'")
        try:
            loaded[name] = _import_function(dotted_path)
        except (ImportError, AttributeError) as exc:
            raise ConfigError(
                f"Could not load Jinja function '{name}' from '{dotted_path}': {exc}"
            ) from exc
        logger.debug("Loaded Jinja function %s <- %s", name, dotted_path)
    return loaded


def load_jinja_environment(config_path: Path | None = None) -> Environment:
    """Create a Jinja environment, optionally with custom functions from TOML."""
    function_map: dict[str, Callable[..., Any]] = {}
    if config_path is not None:
        function_map = load_function_map(Path(config_path))

    env = Environment(loader=StringLoader(), autoescape=False)
    env.globals.update(function_map)
    env.globals.setdefault("strftime", strftime_with_ordinal)
    return env


def render_jinja_in_string(
    jinja_env: Environment,
    template_string: str,
    data: Mapping[str, Any] | None = None,
) -> str:
    """Render a template string. ``now`` and ``strftime`` are always available."""
    render_data: dict[str, Any] = {
        "now": time.localtime(),
        "strftime": strftime_with_ordinal,
    }
    if data:
        render_data.update(data)
    try:
        return jinja_env.from_string(template_string).render(**render_data)
    except TemplateError as exc:
        raise ConfigError(f"Failed to render Jinja template: {exc}") from exc
