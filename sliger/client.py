"""Programmatic client for automating Google Slides with Jinja templates."""

from __future__ import annotations

import logging
import re
import uuid
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from sliger import auth, drive_utils, jinja_utils, slides_utils
from sliger.context import FunctionContext, overlay_item, reset_context, set_context
from sliger.exceptions import ConfigError, ImageNotFoundError, SlideNotFoundError
from sliger.images import materialize_image
from sliger.inspect import InspectReport, inspect_slides
from sliger.provenance import format_provenance
from sliger.quotes import has_smart_quotes, replace_smart_quotes
from sliger.results import (
    ErrorResult,
    ImageResult,
    RepeatDirective,
    ScalarResult,
    TableResult,
)
from sliger.tables import create_table_requests
from sliger.templates import (
    alt_text_update_request,
    formula_from_element,
    looks_like_formula,
    notes_shape_id,
    speaker_notes_insert_requests,
)

logger = logging.getLogger(__name__)

IMAGE_PLACEHOLDER = re.compile(r"!\[image\]\((.*)\)")


@dataclass(frozen=True)
class TextChange:
    slide_number: int
    object_id: str
    original: str
    rendered: str
    formula: str = ""
    kind: str = "scalar"
    error: str | None = None


@dataclass(frozen=True)
class JinjifyResult:
    changes: tuple[TextChange, ...]
    dry_run: bool
    errors: tuple[str, ...] = ()

    @property
    def updates(self) -> int:
        return sum(
            (1 if change.original else 0) + (1 if change.rendered else 0)
            for change in self.changes
            if change.kind == "scalar"
        ) + sum(1 for change in self.changes if change.kind != "scalar")


@dataclass(frozen=True)
class ImageChange:
    slide_number: int
    object_id: str
    image_ref: str
    resolved: str


@dataclass(frozen=True)
class ImagifyResult:
    changes: tuple[ImageChange, ...]
    dry_run: bool

    @property
    def replaced(self) -> int:
        return len(self.changes)


@dataclass(frozen=True)
class RenderResult:
    presentation_id: str
    jinjify: JinjifyResult
    imagify: ImagifyResult
    dry_run: bool
    url: str = field(init=False)

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "url",
            f"https://docs.google.com/presentation/d/{self.presentation_id}/edit",
        )


class Sliger:
    """Operate on a single Google Slides presentation."""

    def __init__(
        self,
        creds_file: str | Path | None,
        presentation_id: str,
        config_path: str | Path | None = None,
        *,
        client_secrets: str | Path | None = None,
        token_file: str | Path | None = None,
        full_drive: bool = False,
        use_adc: bool = False,
    ) -> None:
        self.presentation_id = presentation_id
        self._config_path = Path(config_path) if config_path else None
        self._creds = auth.load_credentials(
            creds_file=Path(creds_file) if creds_file else None,
            client_secrets=Path(client_secrets) if client_secrets else None,
            token_file=Path(token_file) if token_file else None,
            full_drive=full_drive,
            use_adc=use_adc,
        )
        self._jinja_env = jinja_utils.load_jinja_environment(self._config_path)
        self._connections = jinja_utils.load_connections(self._config_path)
        self._slides_service: Any | None = None
        self._drive_service: Any | None = None
        self._item_by_slide: dict[str, Any] = {}

    def with_presentation(self, presentation_id: str) -> Sliger:
        clone = Sliger.__new__(Sliger)
        clone.presentation_id = presentation_id
        clone._config_path = self._config_path
        clone._creds = self._creds
        clone._jinja_env = self._jinja_env
        clone._connections = self._connections
        clone._slides_service = self._slides_service
        clone._drive_service = self._drive_service
        clone._item_by_slide = {}
        return clone

    @property
    def slides_service(self) -> Any:
        if self._slides_service is None:
            self._slides_service = slides_utils.get_slides_service(self._creds)
        return self._slides_service

    @property
    def drive_service(self) -> Any:
        if self._drive_service is None:
            self._drive_service = drive_utils.get_drive_service(self._creds)
        return self._drive_service

    def _context(self, data: Mapping[str, Any] | None, **kwargs: Any) -> FunctionContext:
        payload = dict(data or {})
        default = next(iter(self._connections), None)
        return FunctionContext(
            data=payload,
            connections=self._connections,
            default_connection=default,
            **kwargs,
        )

    def duplicate_presentation(
        self,
        copy_title: str,
        *,
        anyone_can_edit: bool = False,
        anyone_can_view: bool = False,
    ) -> str:
        new_id = drive_utils.copy_file(self.drive_service, self.presentation_id, copy_title)
        if anyone_can_edit:
            drive_utils.set_anyone_permission(self.drive_service, new_id, "writer")
        elif anyone_can_view:
            drive_utils.set_anyone_permission(self.drive_service, new_id, "reader")
        logger.info("Duplicated presentation %s -> %s", self.presentation_id, new_id)
        return new_id

    def delete_slide(self, slide_number: int) -> None:
        slide_id = self._require_slide_id(slide_number)
        slides_utils.delete_slide_by_id(self.slides_service, self.presentation_id, slide_id)
        logger.info("Deleted slide #%s (%s)", slide_number, slide_id)

    def duplicate_slide(self, slide_number: int) -> None:
        slide_id = self._require_slide_id(slide_number)
        slides_utils.duplicate_slide_by_id(self.slides_service, self.presentation_id, slide_id)
        logger.info("Duplicated slide #%s (%s)", slide_number, slide_id)

    def inspect(self, data: Mapping[str, Any] | None = None) -> InspectReport:
        slides = slides_utils.get_presentation_slides(self.slides_service, self.presentation_id)
        return inspect_slides(slides, self._jinja_env, dict(data or {}))

    def eval_template(self, template: str, data: Mapping[str, Any] | None = None) -> Any:
        ctx = self._context(data)
        token = set_context(ctx)
        try:
            return jinja_utils.render_box(self._jinja_env, template, ctx.data)
        finally:
            reset_context(token)

    def expand_repeaters(
        self, data: Mapping[str, Any] | None = None, *, dry_run: bool = False
    ) -> list[TextChange]:
        payload = dict(data or {})
        slides = slides_utils.get_presentation_slides(self.slides_service, self.presentation_id)
        changes: list[TextChange] = []
        for index, slide in enumerate(slides, start=1):
            page_id = slide.get("objectId") or ""
            directive, shape_id, formula = self._repeat_on_slide(slide, page_id, payload)
            if directive is None:
                continue
            items = payload.get(directive.key)
            if not isinstance(items, list):
                changes.append(
                    TextChange(
                        slide_number=index,
                        object_id=shape_id,
                        original=formula,
                        rendered="",
                        formula=formula,
                        kind="error",
                        error=f"sliger_repeat({directive.key!r}) needs a list in data",
                    )
                )
                continue
            changes.append(
                TextChange(
                    slide_number=index,
                    object_id=shape_id,
                    original=formula,
                    rendered=f"[repeat {directive.key} x{len(items)}]",
                    formula=formula,
                    kind="repeat",
                )
            )
            if dry_run:
                continue
            extra = max(0, len(items) - 1)
            replies: list[dict[str, Any]] = []
            if extra:
                response = slides_utils.batch_update(
                    self.slides_service,
                    self.presentation_id,
                    [{"duplicateObject": {"objectId": page_id}} for _ in range(extra)],
                )
                replies = response.get("replies") or []
            self._item_by_slide[page_id] = items[0] if items else None
            for reply, item in zip(replies, items[1:], strict=False):
                new_id = (reply.get("duplicateObject") or {}).get("objectId")
                if new_id:
                    self._item_by_slide[new_id] = item
            if shape_id:
                slides_utils.batch_update(
                    self.slides_service,
                    self.presentation_id,
                    [{"deleteObject": {"objectId": shape_id}}],
                )
        return changes

    def jinjify(
        self,
        data: Mapping[str, Any] | None = None,
        *,
        dry_run: bool = False,
        provenance: bool = False,
    ) -> JinjifyResult:
        payload = dict(data or {})
        changes: list[TextChange] = list(self.expand_repeaters(payload, dry_run=dry_run))
        slides = slides_utils.get_presentation_slides(self.slides_service, self.presentation_id)
        errors: list[str] = []
        for index, slide in enumerate(slides, start=1):
            page_id = slide.get("objectId") or ""
            slide_data = payload
            if page_id in self._item_by_slide:
                slide_data = overlay_item(payload, self._item_by_slide[page_id])
            requests: list[dict[str, Any]] = []
            for element in slides_utils.get_text_elements_from_slide(slide):
                parsed = slides_utils.gslides_element_to_text(element, page_id)
                formula = formula_from_element(element, parsed["text"])
                if has_smart_quotes(formula):
                    logger.warning("Smart quotes in formula on slide %s; converting", index)
                source = replace_smart_quotes(formula)
                ctx = self._context(
                    slide_data,
                    slide_index=index - 1,
                    slide_id=page_id,
                    shape_id=parsed["object_id"],
                )
                token = set_context(ctx)
                try:
                    result = jinja_utils.render_box(self._jinja_env, source, ctx.data)
                except ConfigError as exc:
                    result = ErrorResult(str(exc), formula=formula)
                finally:
                    reset_context(token)

                change, extra = self._apply_result(
                    result,
                    index=index,
                    page_id=page_id,
                    element=element,
                    parsed=parsed,
                    formula=formula,
                    dry_run=dry_run,
                )
                if change is None:
                    continue
                changes.append(change)
                if change.error:
                    errors.append(change.error)
                requests.extend(extra)
                # Table replacement deletes this shape; alt-text on a gone id 400s the batch.
                if looks_like_formula(formula) and not dry_run and change.kind != "table":
                    requests.append(alt_text_update_request(parsed["object_id"], formula))
            if requests and not dry_run:
                slides_utils.batch_update(self.slides_service, self.presentation_id, requests)
        if provenance and not dry_run:
            self.write_provenance({"changes": len(changes)})
        result = JinjifyResult(changes=tuple(changes), dry_run=dry_run, errors=tuple(errors))
        logger.info("jinjify %s: %s change(s)", "dry run" if dry_run else "applied", len(changes))
        return result

    def imagify(
        self,
        data: Mapping[str, Any] | None = None,
        *,
        dry_run: bool = False,
    ) -> ImagifyResult:
        payload = dict(data or {})
        slides = slides_utils.get_presentation_slides(self.slides_service, self.presentation_id)
        changes: list[ImageChange] = []
        for index, slide in enumerate(slides):
            page_id = slide.get("objectId") or ""
            slide_data = payload
            if page_id in self._item_by_slide:
                slide_data = overlay_item(payload, self._item_by_slide[page_id])
            for element in slides_utils.get_text_elements_from_slide(slide):
                parsed = slides_utils.gslides_element_to_text(element, page_id)
                formula = formula_from_element(element, parsed["text"])
                ctx = self._context(
                    slide_data, slide_index=index, slide_id=page_id, shape_id=parsed["object_id"]
                )
                token = set_context(ctx)
                try:
                    boxed = jinja_utils.render_box(
                        self._jinja_env, replace_smart_quotes(formula), ctx.data
                    )
                finally:
                    reset_context(token)
                image_ref: str | None = None
                if isinstance(boxed, ImageResult):
                    image_ref = (
                        materialize_image(boxed)
                        if not dry_run
                        else (boxed.url or boxed.path or "image")
                    )
                elif isinstance(boxed, ScalarResult):
                    match = IMAGE_PLACEHOLDER.fullmatch(boxed.text.strip())
                    if match:
                        image_ref = match.group(1).strip()
                if not image_ref:
                    continue
                shape_id = element.get("objectId") or ""
                if dry_run:
                    resolved = (
                        image_ref
                        if image_ref.startswith("http")
                        else self._preview_image_ref(image_ref)
                    )
                    changes.append(
                        ImageChange(
                            slide_number=index + 1,
                            object_id=shape_id,
                            image_ref=image_ref,
                            resolved=resolved,
                        )
                    )
                    continue
                resolved = self._replace_placeholder_with_image(
                    page_id, element, shape_id, image_ref
                )
                changes.append(
                    ImageChange(
                        slide_number=index + 1,
                        object_id=shape_id,
                        image_ref=image_ref,
                        resolved=resolved,
                    )
                )
        result = ImagifyResult(changes=tuple(changes), dry_run=dry_run)
        logger.info("imagify %s: %s image(s)", "dry run" if dry_run else "applied", result.replaced)
        return result

    def render(
        self,
        copy_title: str,
        data: Mapping[str, Any] | None = None,
        *,
        dry_run: bool = False,
        provenance: bool = True,
    ) -> RenderResult:
        payload = dict(data or {})
        if dry_run:
            return RenderResult(
                presentation_id=self.presentation_id,
                jinjify=self.jinjify(payload, dry_run=True),
                imagify=self.imagify(payload, dry_run=True),
                dry_run=True,
            )
        new_id = self.duplicate_presentation(copy_title)
        child = self.with_presentation(new_id)
        jinja = child.jinjify(payload, provenance=provenance)
        images = child.imagify(payload)
        return RenderResult(presentation_id=new_id, jinjify=jinja, imagify=images, dry_run=False)

    def write_provenance(self, extra: Mapping[str, Any] | None = None) -> None:
        slides = slides_utils.get_presentation_slides(self.slides_service, self.presentation_id)
        if not slides:
            return
        notes_id = notes_shape_id(slides[0])
        if not notes_id:
            return
        text = format_provenance(extra=dict(extra or {}))
        existing = ""
        notes_page = (slides[0].get("slideProperties") or {}).get("notesPage") or {}
        for element in notes_page.get("pageElements") or []:
            if element.get("objectId") == notes_id:
                existing = slides_utils.gslides_element_to_text(element, "")["text"]
                break
        slides_utils.batch_update(
            self.slides_service,
            self.presentation_id,
            speaker_notes_insert_requests(notes_id, text, replace=bool(existing)),
        )

    def _repeat_on_slide(
        self, slide: dict[str, Any], page_id: str, data: Mapping[str, Any]
    ) -> tuple[RepeatDirective | None, str, str]:
        ctx = self._context(data, slide_id=page_id)
        token = set_context(ctx)
        try:
            for element in slides_utils.get_text_elements_from_slide(slide):
                parsed = slides_utils.gslides_element_to_text(element, page_id)
                formula = formula_from_element(element, parsed["text"])
                boxed = jinja_utils.render_box(
                    self._jinja_env, replace_smart_quotes(formula), ctx.data
                )
                if isinstance(boxed, RepeatDirective):
                    return boxed, parsed["object_id"], formula
        finally:
            reset_context(token)
        return None, "", ""

    def _apply_result(
        self,
        result: Any,
        *,
        index: int,
        page_id: str,
        element: dict[str, Any],
        parsed: dict[str, str],
        formula: str,
        dry_run: bool,
    ) -> tuple[TextChange | None, list[dict[str, Any]]]:
        shape_id = parsed["object_id"]
        original = parsed["text"]
        if isinstance(result, RepeatDirective):
            return None, []
        if isinstance(result, ErrorResult):
            rendered = f"[sliger error: {result.message}]"
            change = TextChange(
                slide_number=index,
                object_id=shape_id,
                original=original,
                rendered=rendered,
                formula=formula,
                kind="error",
                error=result.message,
            )
            extra = (
                [] if dry_run else slides_utils.text_replace_requests(shape_id, original, rendered)
            )
            return change, extra
        if isinstance(result, TableResult):
            rendered = f"[table {len(result.headers)}x{len(result.rows)}]"
            change = TextChange(
                slide_number=index,
                object_id=shape_id,
                original=original,
                rendered=rendered,
                formula=formula,
                kind="table",
            )
            if dry_run:
                return change, []
            table_id = "sTbl" + uuid.uuid4().hex[:12]
            return change, create_table_requests(
                table_id=table_id,
                page_id=page_id,
                element=element,
                table=result,
                replace_shape_id=shape_id,
            )
        if isinstance(result, ImageResult):
            return None, []
        if isinstance(result, ScalarResult):
            if result.text == original and not looks_like_formula(formula):
                return None, []
            change = TextChange(
                slide_number=index,
                object_id=shape_id,
                original=original,
                rendered=result.text,
                formula=formula,
                kind="scalar",
            )
            extra: list[dict[str, Any]] = []
            if not dry_run and result.text != original:
                extra = slides_utils.text_replace_requests(shape_id, original, result.text)
            return change, extra
        return None, []

    def _require_slide_id(self, slide_number: int) -> str:
        slides = slides_utils.get_presentation_slides(self.slides_service, self.presentation_id)
        slide_id = slides_utils.get_slide_id_by_index(slides, slide_number)
        if slide_id is None:
            raise SlideNotFoundError(slide_number, len(slides))
        return slide_id

    def _preview_image_ref(self, image_ref: str) -> str:
        if image_ref.startswith(("http://", "https://")):
            return image_ref
        path = Path(image_ref)
        if not path.is_file():
            raise ImageNotFoundError(f"Image path does not exist: {path}")
        return image_ref

    def _replace_placeholder_with_image(
        self,
        page_id: str,
        element: dict[str, Any],
        shape_id: str,
        image_ref: str,
    ) -> str:
        if image_ref.startswith(("http://", "https://")):
            self._create_image(page_id, element, shape_id, image_ref)
            return image_ref
        with drive_utils.temporary_public_image(self.drive_service, image_ref) as url:
            self._create_image(page_id, element, shape_id, url)
            return url

    def _create_image(
        self,
        page_id: str,
        element: dict[str, Any],
        shape_id: str,
        image_url: str,
    ) -> None:
        requests = [
            {
                "createImage": {
                    "url": image_url,
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": element.get("size"),
                        "transform": element.get("transform"),
                    },
                }
            },
            {"deleteObject": {"objectId": shape_id}},
        ]
        slides_utils.batch_update(self.slides_service, self.presentation_id, requests)
