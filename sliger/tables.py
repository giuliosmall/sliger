"""Build Slides API requests that turn a TableResult into a table."""

from __future__ import annotations

from typing import Any

from sliger.results import TableResult


def _table_transform(element: dict[str, Any]) -> dict[str, Any]:
    # GET omits scale=1; createTable treats missing scale as 0 and 400s.
    raw = element.get("transform") or {}
    return {
        "scaleX": 1,
        "scaleY": 1,
        "translateX": raw.get("translateX", 0),
        "translateY": raw.get("translateY", 0),
        "unit": raw.get("unit") or "EMU",
    }


def create_table_requests(
    *,
    table_id: str,
    page_id: str,
    element: dict[str, Any],
    table: TableResult,
    replace_shape_id: str | None = None,
) -> list[dict[str, Any]]:
    headers = table.headers
    body = table.rows
    row_count = max(1, 1 + len(body) if headers else len(body) or 1)
    col_count = max(1, len(headers) or (len(body[0]) if body else 1))
    requests: list[dict[str, Any]] = []
    if replace_shape_id:
        requests.append({"deleteObject": {"objectId": replace_shape_id}})
    requests.append(
        {
            "createTable": {
                "objectId": table_id,
                "elementProperties": {
                    "pageObjectId": page_id,
                    "size": element.get("size"),
                    "transform": _table_transform(element),
                },
                "rows": row_count,
                "columns": col_count,
            }
        }
    )
    cells: list[tuple[int, int, str]] = []
    if headers:
        for column, header in enumerate(headers):
            cells.append((0, column, header))
        for row_index, row in enumerate(body, start=1):
            for column, value in enumerate(row):
                cells.append((row_index, column, value))
    else:
        for row_index, row in enumerate(body):
            for column, value in enumerate(row):
                cells.append((row_index, column, value))
    for row_index, column, value in cells:
        if not value:
            continue
        requests.append(
            {
                "insertText": {
                    "objectId": table_id,
                    "cellLocation": {"rowIndex": row_index, "columnIndex": column},
                    "text": value,
                }
            }
        )
    return requests
