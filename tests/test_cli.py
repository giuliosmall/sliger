from __future__ import annotations

from pathlib import Path

import pytest
from typer import BadParameter
from typer.testing import CliRunner

from sliger.cli import _parse_data, app
from sliger.exceptions import SlideNotFoundError, SligerError

runner = CliRunner()


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
    assert "jinjify" in result.stdout
    assert "imagify" in result.stdout
    assert "duplicate-presentation" in result.stdout


def _invoke(tmp_path: Path, *args: str, monkeypatch, fake) -> object:
    creds = tmp_path / "creds.json"
    creds.write_text("{}")
    monkeypatch.setattr("sliger.cli.Sliger", lambda *a, **k: fake)
    return runner.invoke(
        app,
        ["--creds-file", str(creds), "--presentation-id", "pres-1", *args],
    )


class _FakeClient:
    def __init__(self) -> None:
        self.calls: list[tuple] = []

    def duplicate_presentation(self, title, *, anyone_can_edit=False, anyone_can_view=False):
        self.calls.append(("dup", title, anyone_can_edit, anyone_can_view))
        return "new-id"

    def delete_slide(self, number: int) -> None:
        self.calls.append(("delete", number))

    def duplicate_slide(self, number: int) -> None:
        self.calls.append(("dup-slide", number))

    def jinjify(self, data) -> int:
        self.calls.append(("jinjify", dict(data)))
        return 4

    def imagify(self, data) -> int:
        self.calls.append(("imagify", dict(data)))
        return 1


def test_subcommand_help_lists_slide_flags(tmp_path: Path, monkeypatch) -> None:
    result = _invoke(
        tmp_path, "delete-slide", "--help", monkeypatch=monkeypatch, fake=_FakeClient()
    )
    assert result.exit_code == 0
    assert "--id" in result.stdout
    assert "--slide-number" in result.stdout


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
    assert fake.calls == [("jinjify", {"name": "Ada"}), ("imagify", {"account": "1"})]


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
