import importlib
import time
from typing import Callable, Optional, Any

import toml
from jinja2 import BaseLoader, Environment


class StringLoader(BaseLoader):
    """Custom Jinja2 template loader for rendering templates from strings."""
    def get_source(self, environment: Environment, template: str) -> tuple[str, None, Callable[[], bool]]:
        """Returns template source, no filename, and always up-to-date status.""" 
        return template, None, lambda: True


def load_jinja_environment(config_path: Optional[str] = None) -> Environment:
    """Loads Jinja2 environment with custom functions from a TOML config."""
    function_map = {}
    if config_path:
        try:
            with open(config_path) as f:
                config = toml.load(f)

            def import_function(function_name: str) -> Callable:
                if "." in function_name:
                    mod_name, func_name = function_name.rsplit(".", 1)
                    mod = importlib.import_module(mod_name)
                    res = getattr(mod, func_name)
                else:
                    # Attempt to import from global/built-in scope if not a module path
                    # This part might need to be more robust or restricted for security
                    res = getattr(globals(), function_name, None)
                    if res is None:
                        # Try builtins as a last resort
                        res = getattr(__builtins__, function_name, None)
                    if res is None:
                        raise ImportError(f"Function {function_name} not found in global or built-in scope.")
                return res

            loaded_function_map = config.get("function_map", {})
            for name, function_name in loaded_function_map.items():
                try:
                    function_map[name] = import_function(function_name)
                except (ImportError, AttributeError) as e:
                    print(f"Warning: Could not load Jinja function '{name}' ('{function_name}'): {e}")
        except FileNotFoundError:
            print(f"Warning: Jinja config file not found at {config_path}")
        except Exception as e:
            print(f"Warning: Error loading Jinja config from {config_path}: {e}")

    env = Environment(loader=StringLoader())
    env.globals.update(function_map)
    # Add built-in strftime_with_ordinal to the environment
    env.globals["strftime"] = strftime_with_ordinal
    return env


def strftime_with_ordinal(fmt: str, t: Optional[time.struct_time] = None) -> str:
    """Formats time with ordinal day (e.g., 1st, 2nd, 3rd).
    Uses current local time if t is None.
    """
    if t is None:
        t = time.localtime()

    def ordinal(n: int) -> str:
        """Derive the ordinal numeral for a given number n."""
        return f"{n:d}{'tsnrhtdd'[(n // 10 % 10 != 1) * (n % 10 < 4) * n % 10::4]}"

    # Python's strftime doesn't have a direct ordinal day, so we handle %O specially
    # Other format codes are passed to time.strftime
    # This simple replacement might not cover all edge cases of strftime, but is common.
    return time.strftime(fmt.replace("%O", ordinal(t.tm_mday)), t)


def render_jinja_in_string(jinja_env: Environment, template_string: str, data: Optional[dict] = None) -> str:
    """Renders a Jinja template string with the given data."""
    if data is None:
        data = {}
    
    # Add 'now' to the data context if not already present
    # This ensures 'now' is always available as in the original jinjify command
    render_data = {"now": time.localtime(), **data}

    rtemplate = jinja_env.from_string(template_string)
    return rtemplate.render(**render_data) 