"""Thin wrappers around the Google Slides API."""

from __future__ import annotations

import logging
from typing import Any

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from sliger.exceptions import GoogleAPIError

logger = logging.getLogger(__name__)


def get_slides_service(creds: Any) -> Any:
    return build("slides", "v1", credentials=creds, cache_discovery=False)


def _execute(request: Any, action: str) -> Any:
    try:
        return request.execute()
    except HttpError as err:
        raise GoogleAPIError(f"{action}: {err}") from err


def get_presentation_slides(service: Any, presentation_id: str) -> list[dict[str, Any]]:
    presentation = _execute(
        service.presentations().get(presentationId=presentation_id),
        f"fetch presentation {presentation_id}",
    )
    slides = presentation.get("slides") or []
    logger.debug("Presentation %s contains %s slides", presentation_id, len(slides))
    return slides


def get_slide_id_by_index(slides: list[dict[str, Any]], index: int) -> str | None:
    """Return the objectId of a slide given its 1-based index."""
    if 1 <= index <= len(slides):
        return slides[index - 1].get("objectId")
    return None


def get_text_elements_from_slide(slide: dict[str, Any]) -> list[dict[str, Any]]:
    """Return TEXT_BOX page elements from a slide."""
    elements = slide.get("pageElements") or []
    return [el for el in elements if el.get("shape") and el["shape"].get("shapeType") == "TEXT_BOX"]


def gslides_element_to_text(el: dict[str, Any], page_object_id: str) -> dict[str, str]:
    """Flatten a text-box page element into object id, page id, and text."""
    shape_object_id = el.get("objectId") or ""
    shape = el.get("shape") or {}
    if "text" not in shape:
        return {
            "object_id": shape_object_id,
            "page_object_id": page_object_id,
            "text": "",
        }

    text_contents = [
        text_el["textRun"]["content"] if "textRun" in text_el else ""
        for text_el in shape["text"].get("textElements", [])
    ]
    return {
        "object_id": shape_object_id,
        "page_object_id": page_object_id,
        "text": "".join(text_contents).strip(),
    }


def text_replace_requests(shape_id: str, original: str, rendered: str) -> list[dict[str, Any]]:
    """Build batchUpdate requests that replace all text in a single shape."""
    requests: list[dict[str, Any]] = []
    if original:
        requests.append(
            {
                "deleteText": {
                    "objectId": shape_id,
                    "textRange": {"type": "ALL"},
                }
            }
        )
    if rendered:
        requests.append(
            {
                "insertText": {
                    "objectId": shape_id,
                    "insertionIndex": 0,
                    "text": rendered,
                }
            }
        )
    return requests


def batch_update(
    service: Any, presentation_id: str, requests: list[dict[str, Any]]
) -> dict[str, Any]:
    if not requests:
        return {}
    return _execute(
        service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": requests},
        ),
        f"batch update presentation {presentation_id}",
    )


def delete_slide_by_id(service: Any, presentation_id: str, slide_id: str) -> dict[str, Any]:
    return batch_update(service, presentation_id, [{"deleteObject": {"objectId": slide_id}}])


def duplicate_slide_by_id(service: Any, presentation_id: str, slide_id: str) -> dict[str, Any]:
    return batch_update(service, presentation_id, [{"duplicateObject": {"objectId": slide_id}}])
