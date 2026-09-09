from __future__ import annotations

import json
import re
from pathlib import Path

import pytest
from typer import BadParameter
from typer.testing import CliRunner

from sliger.cli import _parse_data, app
from sliger.client import ImageChange, ImagifyResult, JinjifyResult, RenderResult, TextChange
from sliger.exceptions import SlideNotFoundError, SligerError
from sliger.inspect import InspectReport
from sliger.results import ErrorResult, ScalarResult, TableResult

runner = CliRunner()
_ANSI = re.compile(r"\x1b\[[0-9;]*m")


def _visible(text: str) -> str:
    """Help text on GitHub Actions is colorized; Typer splits ``--flag`` with ANSI."""
    return _ANSI.sub("", text)


def test_parse_data_json() -> None:
    assert _parse_data('{"name": "Ada"}') == {"name": "Ada"}


def test_parse_data_python_literal() -> None:
    assert _parse_data("{'name': 'Ada'}") == {"name": "Ada"}


def test_parse_data_rejects_non_object() -> None:
    with pytest.raises(BadParameter):
        _parse_data("[1, 2]")


def test_parse_data_rejects_garbage() -> None:
    with pytest.raises(BadParameter):
        _parse_data("not-json-or-dict")


def test_help() -> None:
    result = runner.invoke(app, ["--help"])
    assert result.exit_code == 0
    help_text = _visible(result.stdout)
    assert "jinjify" in help_text
    assert "imagify" in help_text
    assert "duplicate-presentation" in help_text
    assert "render" in help_text
    assert "inspect" in help_text
    assert "repl" in help_text
    assert "--client-secrets" in help_text
    assert "--full-drive" in help_text


def test_oauth_flags_forwarded(tmp_path: Path, monkeypatch) -> None:
    captured: dict = {}

    def fake_sliger(*args, **kwargs):
        captured["args"] = args
        captured["kwargs"] = kwargs
        return _FakeClient()

    secrets = tmp_path / "client_secret.json"
    secrets.write_text('{"installed": {"client_id": "x"}}')
    monkeypatch.setattr("sliger.cli.Sliger", fake_sliger)
    result = runner.invoke(
        app,
        [
            "--client-secrets",
            str(secrets),
            "--presentation-id",
            "pres-1",
            "--full-drive",
            "jinjify",
        ],
    )
    assert result.exit_code == 0
    assert captured["kwargs"]["full_drive"] is True
    assert captured["kwargs"]["client_secrets"] == secrets.resolve()


def _invoke(tmp_path: Path, *args: str, monkeypatch, fake) -> object:
    creds = tmp_path / "creds.json"
    creds.write_text("{}")
    monkeypatch.setattr("sliger.cli.Sliger", lambda *a, **k: fake)
    return runner.invoke(
        app,
        ["--creds-file", str(creds), "--presentation-id", "pres-1", *args],
    )


class _FakeClient:
    def __init__(
        self,
        *,
        eval_result: ScalarResult | TableResult | ErrorResult | None = None,
    ) -> None:
        self.calls: list[tuple] = []
        self._eval_result = eval_result

    def duplicate_presentation(self, title, *, anyone_can_edit=False, anyone_can_view=False):
        self.calls.append(("dup", title, anyone_can_edit, anyone_can_view))
        return "new-id"

    def delete_slide(self, number: int) -> None:
        self.calls.append(("delete", number))

    def duplicate_slide(self, number: int) -> None:
        self.calls.append(("dup-slide", number))

    def jinjify(self, data, *, dry_run: bool = False) -> JinjifyResult:
        self.calls.append(("jinjify", dict(data), dry_run))
        return JinjifyResult(
            changes=(
                TextChange(
                    slide_number=1,
                    object_id="shape-1",
                    original="Hello {{ name }}",
                    rendered="Hello Ada",
                ),
                TextChange(
                    slide_number=2,
                    object_id="shape-2",
                    original="Hi {{ name }}",
                    rendered="Hi Ada",
                ),
            ),
            dry_run=dry_run,
        )

    def imagify(self, data, *, dry_run: bool = False) -> ImagifyResult:
        self.calls.append(("imagify", dict(data), dry_run))
        return ImagifyResult(
            changes=(
                ImageChange(
                    slide_number=1,
                    object_id="img-1",
                    image_ref="graph.jpg",
                    resolved="graph.jpg",
                ),
            ),
            dry_run=dry_run,
        )

    def render(self, copy_title, data, *, dry_run: bool = False, provenance: bool = True):
        self.calls.append(("render", copy_title, dict(data), dry_run, provenance))
        return RenderResult(
            presentation_id="pres-1",
            jinjify=JinjifyResult(changes=(), dry_run=dry_run),
            imagify=ImagifyResult(changes=(), dry_run=dry_run),
            dry_run=dry_run,
        )

    def inspect(self, data=None) -> InspectReport:
        self.calls.append(("inspect", dict(data or {})))
        return InspectReport(
            formulas=(),
            functions=(),
            variables=("name",),
            smart_quotes=False,
            missing_data=(),
        )

    def eval_template(self, template, data=None):
        self.calls.append(("eval", template, dict(data or {})))
        if isinstance(self._eval_result, (ErrorResult, ScalarResult, TableResult)):
            return self._eval_result
        return ScalarResult("1")


def test_subcommand_help_lists_slide_flags(tmp_path: Path, monkeypatch) -> None:
    result = _invoke(
        tmp_path, "delete-slide", "--help", monkeypatch=monkeypatch, fake=_FakeClient()
    )
    assert result.exit_code == 0
    help_text = _visible(result.stdout)
    assert "--id" in help_text
    assert "--slide-number" in help_text


def test_duplicate_presentation_command(tmp_path: Path, monkeypatch) -> None:
    fake = _FakeClient()
    result = _invoke(
        tmp_path,
        "duplicate-presentation",
        "--copy-title",
        "Copy",
        "--anyone-can-view",
        monkeypatch=monkeypatch,
        fake=fake,
    )
    assert result.exit_code == 0
    assert "new-id" in result.stdout
    assert fake.calls == [("dup", "Copy", False, True)]


def test_delete_and_duplicate_slide_commands(tmp_path: Path, monkeypatch) -> None:
    fake = _FakeClient()
    deleted = _invoke(tmp_path, "delete-slide", "--id", "3", monkeypatch=monkeypatch, fake=fake)
    duplicated = _invoke(
        tmp_path, "duplicate-slide", "--slide-number", "2", monkeypatch=monkeypatch, fake=fake
    )
    assert deleted.exit_code == 0
    assert duplicated.exit_code == 0
    assert fake.calls == [("delete", 3), ("dup-slide", 2)]


def test_jinjify_and_imagify_commands(tmp_path: Path, monkeypatch) -> None:
    fake = _FakeClient()
    data_file = tmp_path / "vars.json"
    data_file.write_text('{"name": "Ada"}')
    jinja = _invoke(
        tmp_path,
        "jinjify",
        "--data-file",
        str(data_file),
        monkeypatch=monkeypatch,
        fake=fake,
    )
    images = _invoke(
        tmp_path,
        "imagify",
        "--data",
        '{"account": "1"}',
        monkeypatch=monkeypatch,
        fake=fake,
    )
    assert jinja.exit_code == 0
    assert "4" in jinja.stdout
    assert images.exit_code == 0
    assert "1" in images.stdout
    assert fake.calls == [
        ("jinjify", {"name": "Ada"}, False),
        ("imagify", {"account": "1"}, False),
    ]


def test_jinjify_and_imagify_dry_run(tmp_path: Path, monkeypatch) -> None:
    fake = _FakeClient()
    jinja = _invoke(
        tmp_path,
        "jinjify",
        "--data",
        '{"name": "Ada"}',
        "--dry-run",
        monkeypatch=monkeypatch,
        fake=fake,
    )
    images = _invoke(
        tmp_path,
        "imagify",
        "--dry-run",
        monkeypatch=monkeypatch,
        fake=fake,
    )
    assert jinja.exit_code == 0
    assert "Slide 1: Hello {{ name }} → Hello Ada" in jinja.stdout
    assert "Dry run: 4 text update(s)" in jinja.stdout
    assert "Applied" not in jinja.stdout
    assert images.exit_code == 0
    assert "Slide 1: graph.jpg → graph.jpg" in images.stdout
    assert "Dry run: 1 image placeholder(s)" in images.stdout
    assert "Replaced" not in images.stdout
    assert fake.calls == [
        ("jinjify", {"name": "Ada"}, True),
        ("imagify", {}, True),
    ]


@pytest.mark.parametrize(
    ("args", "method"),
    [
        (("delete-slide", "--id", "9"), "delete_slide"),
        (("duplicate-slide", "--id", "9"), "duplicate_slide"),
        (("jinjify",), "jinjify"),
        (("imagify",), "imagify"),
        (("duplicate-presentation", "--copy-title", "X"), "duplicate_presentation"),
    ],
)
def test_command_error_becomes_exit_1(
    tmp_path: Path, monkeypatch, args: tuple[str, ...], method: str
) -> None:
    class Boom(_FakeClient):
        pass

    def fail(*a, **k):
        raise SlideNotFoundError(9, 1)

    setattr(Boom, method, fail)
    result = _invoke(tmp_path, *args, monkeypatch=monkeypatch, fake=Boom())
    assert result.exit_code == 1
    assert "does not exist" in result.output


def test_client_init_error(tmp_path: Path, monkeypatch) -> None:
    creds = tmp_path / "creds.json"
    creds.write_text("{}")

    def boom(*args, **kwargs):
        raise SligerError("bad creds")

    monkeypatch.setattr("sliger.cli.Sliger", boom)
    result = runner.invoke(
        app,
        ["--creds-file", str(creds), "--presentation-id", "p", "jinjify"],
    )
    assert result.exit_code == 1
    assert "bad creds" in result.output


def test_jinjify_data_file_must_be_object(tmp_path: Path, monkeypatch) -> None:
    data_file = tmp_path / "vars.json"
    data_file.write_text("[1, 2]")
    result = _invoke(
        tmp_path,
        "jinjify",
        "--data-file",
        str(data_file),
        monkeypatch=monkeypatch,
        fake=_FakeClient(),
    )
    assert result.exit_code == 1


def test_render_dry_run_command(tmp_path: Path, monkeypatch) -> None:
    fake = _FakeClient()
    result = _invoke(
        tmp_path,
        "render",
        "--copy-title",
        "Copy",
        "--dry-run",
        monkeypatch=monkeypatch,
        fake=fake,
    )
    assert result.exit_code == 0
    assert "Dry run render on pres-1" in result.stdout
    assert "text changes: 0" in result.stdout
    assert fake.calls == [("render", "Copy", {}, True, True)]


def test_inspect_command_prints_json(tmp_path: Path, monkeypatch) -> None:
    fake = _FakeClient()
    result = _invoke(tmp_path, "inspect", monkeypatch=monkeypatch, fake=fake)
    assert result.exit_code == 0
    payload = json.loads(result.stdout)
    assert payload["formulas"] == []
    assert payload["variables"] == ["name"]
    assert payload["smart_quotes"] is False
    assert fake.calls == [("inspect", {})]


def test_repl_scalar_expression(tmp_path: Path, monkeypatch) -> None:
    fake = _FakeClient()
    result = _invoke(tmp_path, "repl", "{{ 1 }}", monkeypatch=monkeypatch, fake=fake)
    assert result.exit_code == 0
    assert result.stdout.strip() == "1"
    assert fake.calls == [("eval", "{{ 1 }}", {})]


def test_repl_error_result_exits_1(tmp_path: Path, monkeypatch) -> None:
    fake = _FakeClient(eval_result=ErrorResult("boom"))
    result = _invoke(tmp_path, "repl", "{{ 1 }}", monkeypatch=monkeypatch, fake=fake)
    assert result.exit_code == 1
    assert "boom" in result.output


def test_jinjify_requires_presentation_id(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.delenv("SLIGER_PRESENTATION_ID", raising=False)
    creds = tmp_path / "creds.json"
    creds.write_text("{}")
    monkeypatch.setattr("sliger.cli.Sliger", lambda *a, **k: _FakeClient())
    result = runner.invoke(app, ["--creds-file", str(creds), "jinjify"])
    assert result.exit_code == 1
    assert "--presentation-id is required" in result.output


def test_render_requires_presentation_id(tmp_path: Path, monkeypatch) -> None:
    monkeypatch.delenv("SLIGER_PRESENTATION_ID", raising=False)
    creds = tmp_path / "creds.json"
    creds.write_text("{}")
    monkeypatch.setattr("sliger.cli.Sliger", lambda *a, **k: _FakeClient())
    result = runner.invoke(app, ["--creds-file", str(creds), "render", "--copy-title", "X"])
    assert result.exit_code == 1
    assert "--presentation-id is required" in result.output
