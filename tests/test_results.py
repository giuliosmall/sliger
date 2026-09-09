from __future__ import annotations

from pathlib import Path

import pytest

from sliger.results import (
    ErrorResult,
    ImageResult,
    RepeatDirective,
    ScalarResult,
    TableResult,
    normalize_result,
)


@pytest.mark.parametrize(
    "value",
    [
        ScalarResult("already"),
        TableResult(headers=("a",), rows=(("1",),)),
        ImageResult(url="https://example.com/chart.png"),
        ErrorResult("nope"),
        RepeatDirective(key="deals"),
    ],
)
def test_normalize_passthrough_typed_results(value: object) -> None:
    assert normalize_result(value) is value


def test_normalize_exception_to_error_result() -> None:
    result = normalize_result(ValueError("boom"))
    assert result == ErrorResult("boom")


def test_normalize_path_to_image_result(tmp_path: Path) -> None:
    path = tmp_path / "chart.png"
    result = normalize_result(path)
    assert result == ImageResult(path=str(path))


def test_normalize_bytes_to_image_result() -> None:
    payload = b"\x89PNG\r\n"
    assert normalize_result(payload) == ImageResult(data=payload)


def test_normalize_list_of_dicts_to_table() -> None:
    result = normalize_result(
        [
            {"name": "Ada", "qty": 3},
            {"name": "Bob", "qty": 1},
        ]
    )
    assert result == TableResult(
        headers=("name", "qty"),
        rows=(("Ada", "3"), ("Bob", "1")),
    )


def test_normalize_list_of_lists_to_table() -> None:
    result = normalize_result([["name", "qty"], ["Ada", 3], ["Bob", 1]])
    assert result == TableResult(
        headers=("name", "qty"),
        rows=(("Ada", "3"), ("Bob", "1")),
    )


def test_normalize_dict_with_headers_and_rows() -> None:
    result = normalize_result({"headers": ["name", "qty"], "rows": [["Ada", 3]]})
    assert result == TableResult(headers=("name", "qty"), rows=(("Ada", "3"),))


def test_normalize_none_to_empty_scalar() -> None:
    assert normalize_result(None) == ScalarResult("")


def test_table_from_dicts_empty_list() -> None:
    assert TableResult.from_dicts([]) == TableResult(headers=(), rows=())


def test_table_from_sequences_with_explicit_headers() -> None:
    result = TableResult.from_sequences([["Ada", 3], ["Bob", 1]], headers=["name", "qty"])
    assert result == TableResult(
        headers=("name", "qty"),
        rows=(("Ada", "3"), ("Bob", "1")),
    )


def test_table_from_dicts_uses_first_record_keys() -> None:
    result = TableResult.from_dicts(
        [
            {"name": "Ada", "qty": 3},
            {"name": "Bob"},
        ]
    )
    assert result.rows == (("Ada", "3"), ("Bob", ""))
