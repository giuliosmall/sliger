"""Programmatic client for automating Google Slides with Jinja templates."""

from __future__ import annotations

import logging
import re
from collections.abc import Mapping
from pathlib import Path
from typing import Any

from google.oauth2 import service_account

from sliger import drive_utils, jinja_utils, slides_utils
from sliger.exceptions import CredentialsError, SlideNotFoundError

logger = logging.getLogger(__name__)

SCOPES = (
    "https://www.googleapis.com/auth/presentations",
    "https://www.googleapis.com/auth/drive",
)

IMAGE_PLACEHOLDER = re.compile(r"!\[image\]\((.*)\)")


class Sliger:
    """Operate on a single Google Slides presentation.

    Parameters
    ----------
    creds_file:
        Path to a GCP service-account JSON key. The presentation must be shared
        with the service account email.
    presentation_id:
        The target presentation ID (the value in the Slides URL after ``/d/``).
    config_path:
        Optional TOML file whose ``[function_map]`` table maps Jinja names to
        ``module.function`` dotted paths.
    """

    def __init__(
        self,
        creds_file: str | Path,
        presentation_id: str,
        config_path: str | Path | None = None,
    ) -> None:
        self.presentation_id = presentation_id
        self._creds = self._load_credentials(Path(creds_file))
        self._jinja_env = jinja_utils.load_jinja_environment(
            Path(config_path) if config_path else None
        )
        self._slides_service: Any | None = None
        self._drive_service: Any | None = None

    @staticmethod
    def _load_credentials(creds_file: Path) -> Any:
        try:
            return service_account.Credentials.from_service_account_file(
                str(creds_file), scopes=SCOPES
            )
        except OSError as exc:
            raise CredentialsError(f"Could not read credentials file {creds_file}: {exc}") from exc
        except ValueError as exc:
            raise CredentialsError(f"Invalid credentials file {creds_file}: {exc}") from exc

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

    def duplicate_presentation(
        self,
        copy_title: str,
        *,
        anyone_can_edit: bool = False,
        anyone_can_view: bool = False,
    ) -> str:
        """Copy the current presentation. Returns the new presentation ID.

        Sharing is off by default. The service account already owns the copy,
        so it can keep editing without a public permission. Pass
        ``anyone_can_view`` or ``anyone_can_edit`` only when you intentionally
        want a world-readable/writable link.
        """
        new_id = drive_utils.copy_file(self.drive_service, self.presentation_id, copy_title)
        if anyone_can_edit:
            drive_utils.set_anyone_permission(self.drive_service, new_id, "writer")
        elif anyone_can_view:
            drive_utils.set_anyone_permission(self.drive_service, new_id, "reader")
        logger.info("Duplicated presentation %s -> %s", self.presentation_id, new_id)
        return new_id

    def delete_slide(self, slide_number: int) -> None:
        """Delete a slide by 1-based index."""
        slide_id = self._require_slide_id(slide_number)
        slides_utils.delete_slide_by_id(self.slides_service, self.presentation_id, slide_id)
        logger.info("Deleted slide #%s (%s)", slide_number, slide_id)

    def duplicate_slide(self, slide_number: int) -> None:
        """Duplicate a slide by 1-based index."""
        slide_id = self._require_slide_id(slide_number)
        slides_utils.duplicate_slide_by_id(self.slides_service, self.presentation_id, slide_id)
        logger.info("Duplicated slide #%s (%s)", slide_number, slide_id)

    def jinjify(self, data: Mapping[str, Any] | None = None) -> int:
        """Render Jinja templates in every text box. Returns the number of updates."""
        slides = slides_utils.get_presentation_slides(self.slides_service, self.presentation_id)
        total = 0
        for index, slide in enumerate(slides, start=1):
            page_id = slide.get("objectId") or ""
            requests: list[dict[str, Any]] = []
            for element in slides_utils.get_text_elements_from_slide(slide):
                parsed = slides_utils.gslides_element_to_text(element, page_id)
                original = parsed["text"]
                rendered = jinja_utils.render_jinja_in_string(self._jinja_env, original, data)
                if rendered == original:
                    continue
                logger.debug(
                    "Slide #%s shape %s: %r -> %r",
                    index,
                    parsed["object_id"],
                    original,
                    rendered,
                )
                requests.extend(
                    slides_utils.text_replace_requests(parsed["object_id"], original, rendered)
                )
            if requests:
                slides_utils.batch_update(self.slides_service, self.presentation_id, requests)
                total += len(requests)
        logger.info("jinjify applied %s update request(s)", total)
        return total

    def imagify(self, data: Mapping[str, Any] | None = None) -> int:
        """Replace ``![image](path-or-url)`` text boxes with images.

        ``path-or-url`` is Jinja-rendered first. Local paths are uploaded to
        Drive (link-readable, required by the Slides API). HTTP(S) URLs are
        used as-is.
        """
        slides = slides_utils.get_presentation_slides(self.slides_service, self.presentation_id)
        replaced = 0
        for index, slide in enumerate(slides):
            page_id = slide.get("objectId") or ""
            for element in slides_utils.get_text_elements_from_slide(slide):
                parsed = slides_utils.gslides_element_to_text(element, page_id)
                context = {
                    "slide_index": index,
                    "slide_id": page_id,
                    **(data or {}),
                }
                rendered = jinja_utils.render_jinja_in_string(
                    self._jinja_env, parsed["text"], context
                )
                match = IMAGE_PLACEHOLDER.fullmatch(rendered.strip())
                if not match:
                    continue
                image_ref = match.group(1).strip()
                image_url = self._resolve_image_url(image_ref)
                shape_id = element.get("objectId")
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
                replaced += 1
                logger.debug(
                    "Replaced shape %s on slide #%s with %s",
                    shape_id,
                    index + 1,
                    image_ref,
                )
        logger.info("imagify replaced %s image placeholder(s)", replaced)
        return replaced

    def _require_slide_id(self, slide_number: int) -> str:
        slides = slides_utils.get_presentation_slides(self.slides_service, self.presentation_id)
        slide_id = slides_utils.get_slide_id_by_index(slides, slide_number)
        if slide_id is None:
            raise SlideNotFoundError(slide_number, len(slides))
        return slide_id

    def _resolve_image_url(self, image_ref: str) -> str:
        if image_ref.startswith(("http://", "https://")):
            return image_ref
        file_id = drive_utils.upload_image_to_drive(self.drive_service, image_ref)
        return drive_utils.drive_image_url(file_id)
