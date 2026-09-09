"""Fill remaining branches in the new plugin surface."""

from __future__ import annotations

from pathlib import Path

import pytest
from typer.testing import CliRunner

from sliger.cli import app
from sliger.client import Sliger
from sliger.connections import sqlite_path
from sliger.context import FunctionContext, bind_function, set_context
from sliger.exceptions import ConfigError, SligerError
from sliger.jinja_utils import load_jinja_environment, render_box
from sliger.results import RepeatDirective, ScalarResult, TableResult, normalize_result
from sliger.templates import notes_shape_id, speaker_notes_insert_requests


def test_sqlite_path_memory_variants() -> None:
    assert sqlite_path("sqlite:///:memory:") == ":memory:"
    assert sqlite_path("file:foo.db")
    assert sqlite_path("file:///:memory:") == ":memory:"


def test_normalize_plain_dict_and_list() -> None:
    assert isinstance(normalize_result({"a": 1}), ScalarResult)
    assert isinstance(normalize_result(["a", "b"]), ScalarResult)
    empty = TableResult.from_sequences([])
    assert empty.headers == ()


def test_bind_varargs_and_kwargs() -> None:
    def grab(*args, **kwargs):
        return (args, kwargs)

    wrapped = bind_function("grab", grab)
    ctx = FunctionContext(data={"x": 1})
    token = set_context(ctx)
    try:
        assert wrapped(1, 2, a=3)[1]["a"] == 3
    finally:
        from sliger.context import reset_context

        reset_context(token)


def test_notes_shape_id_body_placeholder() -> None:
    slide = {
        "slideProperties": {
            "notesPage": {
                "pageElements": [
                    {
                        "objectId": "notes-body",
                        "shape": {"placeholder": {"type": "BODY"}},
                    }
                ]
            }
        }
    }
    assert notes_shape_id(slide) == "notes-body"
    assert speaker_notes_insert_requests("", "x") == []


def test_expand_repeaters_non_list(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("rep", "{{ sliger_repeat('deals') }}")],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    changes = sliger_client.expand_repeaters({"deals": "nope"})
    assert changes and changes[0].kind == "error"


def test_write_provenance_no_slides(sliger_client: Sliger, monkeypatch) -> None:
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [])
    sliger_client.write_provenance()


def test_render_box_reraises_config_error() -> None:
    env = load_jinja_environment()

    def boom():
        raise ConfigError("nope")

    env.globals["boom"] = boom
    with pytest.raises(ConfigError, match="nope"):
        render_box(env, "{{ boom() }}")


def test_cli_render_and_inspect_errors(tmp_path: Path, monkeypatch) -> None:
    runner = CliRunner()
    creds = tmp_path / "c.json"
    creds.write_text("{}")

    class Boom:
        def render(self, *a, **k):
            raise SligerError("render-fail")

        def inspect(self, *a, **k):
            raise SligerError("inspect-fail")

        def eval_template(self, *a, **k):
            raise SligerError("repl-fail")

    monkeypatch.setattr("sliger.cli.Sliger", lambda *a, **k: Boom())
    base = ["--creds-file", str(creds), "--presentation-id", "p"]
    assert runner.invoke(app, [*base, "render", "--copy-title", "T"]).exit_code == 1
    assert runner.invoke(app, [*base, "inspect"]).exit_code == 1
    assert runner.invoke(app, ["--creds-file", str(creds), "repl", "{{ 1 }}"]).exit_code == 1


def test_cli_inspect_missing_and_smart_quotes(tmp_path: Path, monkeypatch) -> None:
    from sliger.inspect import InspectReport

    runner = CliRunner()
    creds = tmp_path / "c.json"
    creds.write_text("{}")

    class Fake:
        def inspect(self, data):
            return InspectReport(
                formulas=(),
                functions=(),
                variables=("name",),
                smart_quotes=True,
                missing_data=("name",),
            )

        def eval_template(self, expression, data):
            return TableResult(headers=("a",), rows=(("1",),))

        def render(self, *a, **k):
            from sliger.client import ImagifyResult, JinjifyResult, RenderResult

            return RenderResult(
                presentation_id="new",
                jinjify=JinjifyResult(changes=(), dry_run=False),
                imagify=ImagifyResult(changes=(), dry_run=False),
                dry_run=False,
            )

    monkeypatch.setattr("sliger.cli.Sliger", lambda *a, **k: Fake())
    creds_args = ["--creds-file", str(creds), "--presentation-id", "p"]
    inspect = runner.invoke(app, [*creds_args, "inspect"])
    assert inspect.exit_code == 1
    assert "Missing data" in inspect.output
    repl = runner.invoke(app, ["--creds-file", str(creds), "repl", "{{ sql('select 1') }}"])
    assert repl.exit_code == 0
    assert "a" in repl.stdout
    rendered = runner.invoke(app, [*creds_args, "render", "--copy-title", "X"])
    assert rendered.exit_code == 0
    assert "new" in rendered.stdout


def test_cli_repl_other_result(tmp_path: Path, monkeypatch) -> None:
    runner = CliRunner()
    creds = tmp_path / "c.json"
    creds.write_text("{}")

    class Fake:
        def eval_template(self, expression, data):
            return RepeatDirective("items")

    monkeypatch.setattr("sliger.cli.Sliger", lambda *a, **k: Fake())
    result = runner.invoke(
        app, ["--creds-file", str(creds), "repl", "{{ sliger_repeat('items') }}"]
    )
    assert result.exit_code == 0
    assert "items" in result.stdout


def test_bind_positional_only() -> None:
    def grab(a, /, ctx=None):
        return a, ctx is not None

    wrapped = bind_function("grab", grab)
    ctx = FunctionContext(data={})
    token = set_context(ctx)
    try:
        value, has_ctx = wrapped(7)
        assert value == 7
        assert has_ctx is True
    finally:
        from sliger.context import reset_context

        reset_context(token)


def test_jinjify_smart_quotes_and_empty_repeat(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [
            text_element_factory("t", "Hi {{ name }}".replace("{{", "\u201c{{\u201c")),
            text_element_factory("r", "{{ sliger_repeat('xs') }}"),
        ],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    captured = []
    monkeypatch.setattr(
        "sliger.slides_utils.batch_update", lambda *a, **k: captured.append(a) or {}
    )
    sliger_client.jinjify({"name": "Ada", "xs": []})


def test_jinjify_image_result_skipped_for_imagify(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    from sliger.results import ImageResult

    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("img", "{{ pic() }}")],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    monkeypatch.setattr(
        "sliger.jinja_utils.render_box",
        lambda *a, **k: ImageResult(url="https://example.com/a.png"),
    )
    result = sliger_client.jinjify()
    assert all(c.kind != "image" for c in result.changes)
