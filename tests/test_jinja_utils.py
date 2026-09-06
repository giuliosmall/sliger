from __future__ import annotations

import time
from pathlib import Path

import pytest
from jinja2 import Environment

from sliger.exceptions import ConfigError
from sliger.jinja_utils import (
    StringLoader,
    load_jinja_environment,
    ordinal,
    render_jinja_in_string,
    strftime_with_ordinal,
)


def test_string_loader() -> None:
    source, filename, uptodate = StringLoader().get_source(Environment(), "hi {{ x }}")
    assert source == "hi {{ x }}"
    assert filename is None
    assert uptodate() is True


@pytest.mark.parametrize(
    ("day", "expected"),
    [
        (1, "1st"),
        (2, "2nd"),
        (3, "3rd"),
        (4, "4th"),
        (11, "11th"),
        (12, "12th"),
        (13, "13th"),
        (21, "21st"),
        (22, "22nd"),
        (23, "23rd"),
        (31, "31st"),
    ],
)
def test_ordinal(day: int, expected: str) -> None:
    assert ordinal(day) == expected


def test_strftime_with_ordinal() -> None:
    t = time.strptime("2023-09-02", "%Y-%m-%d")
    assert strftime_with_ordinal("%A, %O %B", t) == "Saturday, 2nd September"


def test_strftime_defaults_to_localtime() -> None:
    assert strftime_with_ordinal("%Y")


def test_render_injects_now_and_strftime() -> None:
    env = load_jinja_environment()
    rendered = render_jinja_in_string(
        env, '{{ strftime("%O", now) }}', {"now": time.strptime("2023-09-02", "%Y-%m-%d")}
    )
    assert rendered == "2nd"


def test_render_with_custom_data() -> None:
    env = load_jinja_environment()
    assert render_jinja_in_string(env, "Hi {{ name }}", {"name": "Ada"}) == "Hi Ada"


def test_load_custom_functions(config_file: Path) -> None:
    env = load_jinja_environment(config_file)
    assert env.globals["greet"]("PyCon") == "hi PyCon"
    rendered = render_jinja_in_string(env, "{{ greet('there') }}")
    assert rendered == "hi there"


def test_missing_config_raises(tmp_path: Path) -> None:
    with pytest.raises(ConfigError, match="not found"):
        load_jinja_environment(tmp_path / "missing.toml")


def test_invalid_toml_raises(tmp_path: Path) -> None:
    path = tmp_path / "bad.toml"
    path.write_text("this is not = toml [")
    with pytest.raises(ConfigError, match="Invalid TOML"):
        load_jinja_environment(path)


def test_bare_function_name_rejected(tmp_path: Path) -> None:
    path = tmp_path / "config.toml"
    path.write_text("[function_map]\ngreet = 'greet'\n")
    with pytest.raises(ConfigError, match="module.function"):
        load_jinja_environment(path)


def test_invalid_template_raises() -> None:
    env = load_jinja_environment()
    with pytest.raises(ConfigError, match="Failed to render"):
        render_jinja_in_string(env, "{{ unterminated")


def test_function_map_must_be_table(tmp_path: Path) -> None:
    path = tmp_path / "config.toml"
    path.write_text("function_map = 1\n")
    with pytest.raises(ConfigError, match="must be a table"):
        load_jinja_environment(path)


def test_non_callable_function_rejected(tmp_path: Path) -> None:
    (tmp_path / "constants.py").write_text("greet = 123\n")
    path = tmp_path / "config.toml"
    path.write_text("[function_map]\ngreet = 'constants.greet'\n")
    with pytest.raises(ConfigError, match="not callable"):
        load_jinja_environment(path)


def test_function_map_value_must_be_string(tmp_path: Path) -> None:
    path = tmp_path / "config.toml"
    path.write_text("[function_map]\ngreet = 1\n")
    with pytest.raises(ConfigError, match="must be a string"):
        load_jinja_environment(path)
