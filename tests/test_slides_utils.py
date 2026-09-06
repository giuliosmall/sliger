from __future__ import annotations

from unittest.mock import Mock

import pytest
from googleapiclient.errors import HttpError

from sliger import slides_utils
from sliger.exceptions import GoogleAPIError
from sliger.slides_utils import (
    batch_update,
    delete_slide_by_id,
    duplicate_slide_by_id,
    get_presentation_slides,
    get_slide_id_by_index,
    get_text_elements_from_slide,
    gslides_element_to_text,
    text_replace_requests,
)


def test_get_slide_id_by_index() -> None:
    slides = [{"objectId": "a"}, {"objectId": "b"}]
    assert get_slide_id_by_index(slides, 1) == "a"
    assert get_slide_id_by_index(slides, 2) == "b"
    assert get_slide_id_by_index(slides, 0) is None
    assert get_slide_id_by_index(slides, 3) is None


def test_get_text_elements_from_slide(sample_slide: dict) -> None:
    elements = get_text_elements_from_slide(sample_slide)
    assert [el["objectId"] for el in elements] == ["shape-1", "shape-2"]


def test_gslides_element_to_text_joins_runs(text_element_factory) -> None:
    el = text_element_factory("s1", "Hello world\n")
    parsed = gslides_element_to_text(el, "page-1")
    assert parsed == {
        "object_id": "s1",
        "page_object_id": "page-1",
        "text": "Hello world",
    }


def test_gslides_element_empty_shape(text_element_factory) -> None:
    el = text_element_factory("s1", "")
    parsed = gslides_element_to_text(el, "page-1")
    assert parsed["text"] == ""


def test_gslides_element_skips_non_textrun() -> None:
    el = {
        "objectId": "s1",
        "shape": {
            "shapeType": "TEXT_BOX",
            "text": {
                "textElements": [
                    {"paragraphMarker": {}},
                    {"textRun": {"content": "Hi"}},
                ]
            },
        },
    }
    assert gslides_element_to_text(el, "p")["text"] == "Hi"


def test_text_replace_requests_full_swap() -> None:
    requests = text_replace_requests("shape-1", "old", "new")
    assert requests == [
        {"deleteText": {"objectId": "shape-1", "textRange": {"type": "ALL"}}},
        {"insertText": {"objectId": "shape-1", "insertionIndex": 0, "text": "new"}},
    ]


def test_text_replace_requests_empty_original() -> None:
    requests = text_replace_requests("shape-1", "", "new")
    assert requests == [
        {"insertText": {"objectId": "shape-1", "insertionIndex": 0, "text": "new"}},
    ]


def test_text_replace_requests_clear() -> None:
    requests = text_replace_requests("shape-1", "old", "")
    assert requests == [
        {"deleteText": {"objectId": "shape-1", "textRange": {"type": "ALL"}}},
    ]


def test_batch_update_skips_empty_requests() -> None:
    assert batch_update(object(), "pres", []) == {}


def test_batch_update_sends_requests() -> None:
    service = Mock()
    service.presentations.return_value.batchUpdate.return_value.execute.return_value = {"ok": True}
    assert batch_update(service, "pres", [{"x": 1}]) == {"ok": True}


def test_get_slides_service(monkeypatch) -> None:
    monkeypatch.setattr(slides_utils, "build", lambda *a, **k: "service")
    assert slides_utils.get_slides_service("creds") == "service"


def test_get_presentation_slides_wraps_http_error() -> None:
    response = Mock(status=404, reason="Not Found")
    response.getheaders.return_value = []
    request = Mock()
    request.execute.side_effect = HttpError(response, b"{}")
    service = Mock()
    service.presentations.return_value.get.return_value = request
    with pytest.raises(GoogleAPIError, match="fetch presentation"):
        get_presentation_slides(service, "pres")


def test_get_presentation_slides_returns_empty_list() -> None:
    request = Mock()
    request.execute.return_value = {}
    service = Mock()
    service.presentations.return_value.get.return_value = request
    assert get_presentation_slides(service, "pres") == []


def test_delete_and_duplicate_slide_by_id() -> None:
    service = Mock()
    service.presentations.return_value.batchUpdate.return_value.execute.return_value = {
        "replies": []
    }
    assert delete_slide_by_id(service, "pres", "s1") == {"replies": []}
    assert duplicate_slide_by_id(service, "pres", "s1") == {"replies": []}
