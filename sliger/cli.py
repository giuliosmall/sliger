"""Command-line interface for sliger."""

from __future__ import annotations

import ast
import json
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Annotated, Any

import typer

from sliger.auth import default_token_path
from sliger.client import ImagifyResult, JinjifyResult, RenderResult, Sliger
from sliger.exceptions import SligerError
from sliger.results import ErrorResult, ScalarResult, TableResult

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
    creds_file: Path | None
    presentation_id: str | None
    config_path: Path | None
    client_secrets: Path | None
    token_file: Path | None
    full_drive: bool
    use_adc: bool


@app.callback()
def main(
    ctx: typer.Context,
    presentation_id: Annotated[
        str | None,
        typer.Option(
            "--presentation-id",
            envvar="SLIGER_PRESENTATION_ID",
            help="Google Slides presentation ID.",
        ),
    ] = None,
    creds_file: Annotated[
        Path | None,
        typer.Option(
            "--creds-file",
            envvar="SLIGER_CREDS_FILE",
            exists=True,
            file_okay=True,
            dir_okay=False,
            readable=True,
            resolve_path=True,
            help="Service-account JSON, OAuth client secrets, or a saved user token.",
        ),
    ] = None,
    client_secrets: Annotated[
        Path | None,
        typer.Option(
            "--client-secrets",
            envvar="SLIGER_CLIENT_SECRETS",
            exists=True,
            file_okay=True,
            dir_okay=False,
            readable=True,
            resolve_path=True,
            help="OAuth installed-app client secrets JSON.",
        ),
    ] = None,
    token_file: Annotated[
        Path | None,
        typer.Option(
            "--token-file",
            envvar="SLIGER_TOKEN_FILE",
            help=f"Cached OAuth token (default: {default_token_path()}).",
        ),
    ] = None,
    full_drive: Annotated[
        bool,
        typer.Option(
            "--full-drive",
            envvar="SLIGER_FULL_DRIVE",
            help="Request full Drive scope (needed to copy presentations the app did not create).",
        ),
    ] = False,
    adc: Annotated[
        bool,
        typer.Option(
            "--adc",
            envvar="SLIGER_ADC",
            help="Use Application Default Credentials (gcloud auth application-default login).",
        ),
    ] = False,
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
    ctx.obj = _CliConfig(
        creds_file=creds_file,
        presentation_id=presentation_id,
        config_path=config_path,
        client_secrets=client_secrets,
        token_file=token_file,
        full_drive=full_drive,
        use_adc=adc,
    )


def _require_presentation(ctx: typer.Context) -> str:
    presentation_id = ctx.obj.presentation_id
    if not presentation_id:
        typer.echo("--presentation-id is required for this command.", err=True)
        raise typer.Exit(code=1)
    return presentation_id


def _payload(data: str, data_file: Path | None) -> dict[str, Any]:
    payload = json.loads(data_file.read_text()) if data_file else _parse_data(data)
    if not isinstance(payload, dict):
        typer.echo("--data-file must contain a JSON object.", err=True)
        raise typer.Exit(code=1)
    return payload


def _client(ctx: typer.Context, presentation_id: str | None = None) -> Sliger:
    cfg: _CliConfig = ctx.obj
    pid = presentation_id if presentation_id is not None else cfg.presentation_id or "repl"
    try:
        return Sliger(
            cfg.creds_file,
            pid,
            cfg.config_path,
            client_secrets=cfg.client_secrets,
            token_file=cfg.token_file,
            full_drive=cfg.full_drive,
            use_adc=cfg.use_adc,
        )
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
    _require_presentation(ctx)
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
    _require_presentation(ctx)
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
    _require_presentation(ctx)
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
    dry_run: Annotated[
        bool,
        typer.Option("--dry-run", help="Preview text changes without updating the presentation."),
    ] = False,
) -> None:
    """Render Jinja templates inside the presentation's text boxes."""
    _require_presentation(ctx)
    payload = _payload(data, data_file)
    try:
        result: JinjifyResult = _client(ctx).jinjify(payload, dry_run=dry_run)
    except SligerError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=1) from exc
    if dry_run:
        for change in result.changes:
            typer.echo(f"Slide {change.slide_number}: {change.original} → {change.rendered}")
        typer.echo(f"Dry run: {result.updates} text update(s)")
        return
    typer.echo(f"Applied {result.updates} text update request(s).")


@app.command()
def imagify(
    ctx: typer.Context,
    data: Annotated[
        str,
        typer.Option(help='JSON object of extra Jinja variables, e.g. \'{"name": "Ada"}\'.'),
    ] = "{}",
    dry_run: Annotated[
        bool,
        typer.Option("--dry-run", help="Preview image replacements without uploading or updating."),
    ] = False,
) -> None:
    """Replace markdown image placeholders with real images."""
    _require_presentation(ctx)
    try:
        result: ImagifyResult = _client(ctx).imagify(_parse_data(data), dry_run=dry_run)
    except SligerError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=1) from exc
    if dry_run:
        for change in result.changes:
            typer.echo(f"Slide {change.slide_number}: {change.image_ref} → {change.resolved}")
        typer.echo(f"Dry run: {result.replaced} image placeholder(s)")
        return
    typer.echo(f"Replaced {result.replaced} image placeholder(s).")


@app.command()
def render(
    ctx: typer.Context,
    copy_title: Annotated[str, typer.Option(help="Title for the copied presentation.")],
    data: Annotated[str, typer.Option(help="JSON object of Jinja variables.")] = "{}",
    data_file: Annotated[
        Path | None,
        typer.Option("--data-file", exists=True, file_okay=True, dir_okay=False, readable=True),
    ] = None,
    dry_run: Annotated[bool, typer.Option("--dry-run")] = False,
    no_provenance: Annotated[bool, typer.Option("--no-provenance")] = False,
) -> None:
    """Copy the template deck, then jinjify and imagify the copy."""
    _require_presentation(ctx)
    payload = _payload(data, data_file)
    try:
        result: RenderResult = _client(ctx).render(
            copy_title, payload, dry_run=dry_run, provenance=not no_provenance
        )
    except SligerError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=1) from exc
    if dry_run:
        typer.echo(f"Dry run render on {result.presentation_id}")
        typer.echo(f"  text changes: {len(result.jinjify.changes)}")
        typer.echo(f"  images: {result.imagify.replaced}")
        return
    typer.echo(result.presentation_id)
    typer.echo(result.url, err=True)


@app.command("inspect")
def inspect_cmd(
    ctx: typer.Context,
    data: Annotated[str, typer.Option(help="JSON object of Jinja variables.")] = "{}",
    data_file: Annotated[
        Path | None,
        typer.Option("--data-file", exists=True, file_okay=True, dir_okay=False, readable=True),
    ] = None,
) -> None:
    """List formulas, functions, and missing --data keys in the presentation."""
    _require_presentation(ctx)
    payload = _payload(data, data_file)
    try:
        report = _client(ctx).inspect(payload)
    except SligerError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=1) from exc
    typer.echo(json.dumps(report.to_dict(), indent=2))
    if report.smart_quotes:
        typer.echo("Warning: smart quotes found. Turn off Use smart quotes in Slides.", err=True)
    if report.missing_data:
        typer.echo("Missing data keys: " + ", ".join(report.missing_data), err=True)
        raise typer.Exit(code=1)


@app.command()
def repl(
    ctx: typer.Context,
    expression: Annotated[str, typer.Argument(help="Jinja box, e.g. '{{ sql(\"select 1\") }}'.")],
    data: Annotated[str, typer.Option(help="JSON object of Jinja variables.")] = "{}",
    data_file: Annotated[
        Path | None,
        typer.Option("--data-file", exists=True, file_okay=True, dir_okay=False, readable=True),
    ] = None,
) -> None:
    """Evaluate a Jinja box locally (no Slides writes)."""
    payload = _payload(data, data_file)
    try:
        boxed = _client(ctx, presentation_id="repl").eval_template(expression, payload)
    except SligerError as exc:
        typer.echo(str(exc), err=True)
        raise typer.Exit(code=1) from exc
    if isinstance(boxed, ErrorResult):
        typer.echo(boxed.message, err=True)
        raise typer.Exit(code=1)
    if isinstance(boxed, TableResult):
        typer.echo("\t".join(boxed.headers))
        for row in boxed.rows:
            typer.echo("\t".join(row))
        return
    if isinstance(boxed, ScalarResult):
        typer.echo(boxed.text)
        return
    typer.echo(str(boxed))
