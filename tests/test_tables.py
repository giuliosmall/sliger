from __future__ import annotations

from sliger.results import TableResult
from sliger.tables import create_table_requests

_ELEMENT = {
    "size": {"height": {"magnitude": 100, "unit": "PT"}, "width": {"magnitude": 200, "unit": "PT"}},
    "transform": {"scaleX": 1, "scaleY": 1, "translateX": 0, "translateY": 0, "unit": "PT"},
}


def _create_table(requests: list[dict]) -> dict:
    return next(item["createTable"] for item in requests if "createTable" in item)


def _inserts(requests: list[dict]) -> list[dict]:
    return [item["insertText"] for item in requests if "insertText" in item]


def test_create_table_with_headers() -> None:
    table = TableResult(headers=("name", "qty"), rows=(("Ada", "3"), ("Bob", "1")))
    requests = create_table_requests(
        table_id="tbl-1",
        page_id="page-1",
        element=_ELEMENT,
        table=table,
    )
    created = _create_table(requests)
    assert created["objectId"] == "tbl-1"
    assert created["rows"] == 3
    assert created["columns"] == 2
    assert created["elementProperties"]["pageObjectId"] == "page-1"
    assert created["elementProperties"]["size"] == _ELEMENT["size"]
    assert created["elementProperties"]["transform"] == _ELEMENT["transform"]
    assert _inserts(requests) == [
        {
            "objectId": "tbl-1",
            "cellLocation": {"rowIndex": 0, "columnIndex": 0},
            "text": "name",
        },
        {
            "objectId": "tbl-1",
            "cellLocation": {"rowIndex": 0, "columnIndex": 1},
            "text": "qty",
        },
        {
            "objectId": "tbl-1",
            "cellLocation": {"rowIndex": 1, "columnIndex": 0},
            "text": "Ada",
        },
        {
            "objectId": "tbl-1",
            "cellLocation": {"rowIndex": 1, "columnIndex": 1},
            "text": "3",
        },
        {
            "objectId": "tbl-1",
            "cellLocation": {"rowIndex": 2, "columnIndex": 0},
            "text": "Bob",
        },
        {
            "objectId": "tbl-1",
            "cellLocation": {"rowIndex": 2, "columnIndex": 1},
            "text": "1",
        },
    ]


def test_create_table_without_headers() -> None:
    table = TableResult(headers=(), rows=(("Ada", "3"), ("Bob", "1")))
    requests = create_table_requests(
        table_id="tbl-1",
        page_id="page-1",
        element=_ELEMENT,
        table=table,
    )
    created = _create_table(requests)
    assert created["rows"] == 2
    assert created["columns"] == 2
    assert _inserts(requests)[0]["cellLocation"] == {"rowIndex": 0, "columnIndex": 0}
    assert _inserts(requests)[0]["text"] == "Ada"
    assert _inserts(requests)[-1]["cellLocation"] == {"rowIndex": 1, "columnIndex": 1}


def test_create_table_empty_still_one_by_one() -> None:
    table = TableResult(headers=(), rows=())
    requests = create_table_requests(
        table_id="tbl-empty",
        page_id="page-1",
        element=_ELEMENT,
        table=table,
    )
    created = _create_table(requests)
    assert created["rows"] >= 1
    assert created["columns"] >= 1
    assert created["rows"] == 1
    assert created["columns"] == 1
    assert _inserts(requests) == []


def test_create_table_replace_shape_id_deletes() -> None:
    table = TableResult(headers=("name",), rows=())
    requests = create_table_requests(
        table_id="tbl-1",
        page_id="page-1",
        element=_ELEMENT,
        table=table,
        replace_shape_id="old-shape",
    )
    assert requests[0] == {"deleteObject": {"objectId": "old-shape"}}
    assert "createTable" in requests[1]


def test_create_table_forces_unit_scale() -> None:
    element = {
        "size": _ELEMENT["size"],
        "transform": {"translateX": 40, "translateY": 80, "unit": "PT"},
    }
    requests = create_table_requests(
        table_id="tbl-1",
        page_id="page-1",
        element=element,
        table=TableResult(headers=("name",), rows=()),
    )
    transform = _create_table(requests)["elementProperties"]["transform"]
    assert transform["scaleX"] == 1
    assert transform["scaleY"] == 1
    assert transform["translateX"] == 40
    assert transform["translateY"] == 80
    assert transform["unit"] == "PT"
    assert "shearX" not in transform


def test_create_table_skips_empty_cells() -> None:
    table = TableResult(headers=("name", ""), rows=(("Ada", ""),))
    requests = create_table_requests(
        table_id="tbl-1",
        page_id="page-1",
        element=_ELEMENT,
        table=table,
    )
    texts = [item["text"] for item in _inserts(requests)]
    assert texts == ["name", "Ada"]
    for insert in _inserts(requests):
        assert "cellLocation" in insert
        assert "rowIndex" in insert["cellLocation"]
        assert "columnIndex" in insert["cellLocation"]
