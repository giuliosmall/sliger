"""Command-line interface for sliger."""

from __future__ import annotations

import ast
import json
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Annotated, Any

import typer

from sliger.client import Sliger
from sliger.exceptions import SligerError

app = typer.Typer(
    name="sliger",
    help="Slide the power of Python (and Jinja2) into Google Slides.",
    no_args_is_help=True,
)

logger = logging.getLogger("sliger")


def _configure_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.WARNING,
        format="%(levelname)s %(name)s: %(message)s",
    )


def _parse_data(value: str) -> dict[str, Any]:
    try:
        parsed: Any = json.loads(value)
    except json.JSONDecodeError:
        try:
            parsed = ast.literal_eval(value)
        except (ValueError, SyntaxError) as exc:
            raise typer.BadParameter("data must be a JSON object or a Python dict literal") from exc
    if not isinstance(parsed, dict):
        raise typer.BadParameter("data must be an object/dict")
    return parsed


@dataclass
class _CliConfig:
    creds_file: Path
    presentation_id: str
    config_path: Path | None


@app.callback()
def main(
    ctx: typer.Context,
    creds_file: Annotated[
        Path,
        typer.Option(
            "--creds-file",
            envvar="SLIGER_CREDS_FILE",
            exists=True,
            file_okay=True,
            dir_okay=False,
            readable=True,
            resolve_path=True,
            help="GCP service account JSON key.",
        ),
    ],
    presentation_id: Annotated[
        str,
        typer.Option(
            "--presentation-id",
            envvar="SLIGER_PRESENTATION_ID",
            help="Google Slides presentation ID.",
        ),
    ],
    config_path: Annotated[
        Path | None,
        typer.Option(
            "--config-path",
            envvar="SLIGER_CONFIG_PATH",
            exists=True,
            file_okay=True,
            dir_okay=False,
            readable=True,
            resolve_path=True,
            help="TOML file mapping Jinja names to module.function paths.",
        ),
    ] = None,
    verbose: Annotated[
        bool,
        typer.Option("--verbose", "-v", help="Enable debug logging."),
    ] = False,
) -> None:
    """Sliger: automate Google Slides with Python and Jinja2."""
    _configure_logging(verbose)
    ctx.obj = _CliConfig(creds_file, presentation_id, config_path)


def _client(ctx: typer.Context) -> Sliger:
    cfg: _CliConfig = ctx.obj
    try:
        return Sliger(cfg.creds_file, cfg.presentation_id, cfg.config_path)
    except SligerError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=1) from exc


@app.command("duplicate-presentation")
def duplicate_presentation(
    ctx: typer.Context,
    copy_title: Annotated[str, typer.Option(help="Title for the copied presentation.")],
    anyone_can_edit: Annotated[
        bool,
        typer.Option(
            "--anyone-can-edit",
            help="Grant anyone-with-the-link writer access (insecure; opt-in).",
        ),
    ] = False,
    anyone_can_view: Annotated[
        bool,
        typer.Option(
            "--anyone-can-view",
            help="Grant anyone-with-the-link reader access.",
        ),
    ] = False,
) -> None:
    """Duplicate the presentation and print the new presentation ID."""
    try:
        new_id = _client(ctx).duplicate_presentation(
            copy_title,
            anyone_can_edit=anyone_can_edit,
            anyone_can_view=anyone_can_view,
        )
    except SligerError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=1) from exc
    typer.echo(new_id)
    typer.echo(
        f"https://docs.google.com/presentation/d/{new_id}/edit",
        err=True,
    )


@app.command("delete-slide")
def delete_slide(
    ctx: typer.Context,
    slide_id: Annotated[
        int,
        typer.Option(
            "--id",
            "--slide-number",
            help="1-based index of the slide to delete.",
        ),
    ],
) -> None:
    """Delete a slide by its 1-based index."""
    try:
        _client(ctx).delete_slide(slide_id)
    except SligerError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=1) from exc
    typer.echo(f"Deleted slide #{slide_id}.")


@app.command("duplicate-slide")
def duplicate_slide(
    ctx: typer.Context,
    slide_id: Annotated[
        int,
        typer.Option(
            "--id",
            "--slide-number",
            help="1-based index of the slide to duplicate.",
        ),
    ],
) -> None:
    """Duplicate a slide by its 1-based index."""
    try:
        _client(ctx).duplicate_slide(slide_id)
    except SligerError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=1) from exc
    typer.echo(f"Duplicated slide #{slide_id}.")


@app.command()
def jinjify(
    ctx: typer.Context,
    data: Annotated[
        str,
        typer.Option(help='JSON object of extra Jinja variables, e.g. \'{"name": "Ada"}\'.'),
    ] = "{}",
    data_file: Annotated[
        Path | None,
        typer.Option(
            "--data-file",
            exists=True,
            file_okay=True,
            dir_okay=False,
            readable=True,
            help="JSON file of extra Jinja variables (overrides --data).",
        ),
    ] = None,
) -> None:
    """Render Jinja templates inside the presentation's text boxes."""
    payload = json.loads(data_file.read_text()) if data_file else _parse_data(data)
    if not isinstance(payload, dict):
        typer.echo("--data-file must contain a JSON object.", err=True)
        raise typer.Exit(code=1)
    try:
        updates = _client(ctx).jinjify(payload)
    except SligerError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=1) from exc
    typer.echo(f"Applied {updates} text update request(s).")


@app.command()
def imagify(
    ctx: typer.Context,
    data: Annotated[
        str,
        typer.Option(help='JSON object of extra Jinja variables, e.g. \'{"name": "Ada"}\'.'),
    ] = "{}",
) -> None:
    """Replace markdown image placeholders with real images."""
    try:
        replaced = _client(ctx).imagify(_parse_data(data))
    except SligerError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=1) from exc
    typer.echo(f"Replaced {replaced} image placeholder(s).")
