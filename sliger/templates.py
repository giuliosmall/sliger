"""Persist Jinja formulas on shape alt-text so templates survive jinjify."""

from __future__ import annotations

from typing import Any

FORMULA_PREFIX = "sliger:"


def formula_from_element(element: dict[str, Any], visible_text: str) -> str:
    """Prefer a stored formula in alt-text; otherwise the visible text."""
    description = element.get("description") or ""
    if description.startswith(FORMULA_PREFIX):
        return description[len(FORMULA_PREFIX) :]
    return visible_text


def looks_like_formula(text: str) -> bool:
    stripped = text.strip()
    return "{{" in stripped or "{%" in stripped


def alt_text_update_request(object_id: str, formula: str) -> dict[str, Any]:
    return {
        "updatePageElementAltText": {
            "objectId": object_id,
            "description": f"{FORMULA_PREFIX}{formula}",
        }
    }


def speaker_notes_insert_requests(
    notes_object_id: str, text: str, *, replace: bool = True
) -> list[dict[str, Any]]:
    if not notes_object_id or not text:
        return []
    requests: list[dict[str, Any]] = []
    if replace:
        requests.append({"deleteText": {"objectId": notes_object_id, "textRange": {"type": "ALL"}}})
    requests.append(
        {"insertText": {"objectId": notes_object_id, "insertionIndex": 0, "text": text}}
    )
    return requests


def notes_shape_id(slide: dict[str, Any]) -> str | None:
    notes = (slide.get("slideProperties") or {}).get("notesPage") or {}
    for element in notes.get("pageElements") or []:
        shape = element.get("shape") or {}
        if shape.get("shapeType") == "TEXT_BOX" or element.get("objectId"):
            placeholder = (shape.get("placeholder") or {}).get("type")
            if placeholder == "BODY" or shape.get("shapeType") == "TEXT_BOX":
                return element.get("objectId")
    return None
