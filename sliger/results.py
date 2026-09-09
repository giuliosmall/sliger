"""Typed values custom Jinja functions may return."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class ScalarResult:
    text: str


@dataclass(frozen=True)
class TableResult:
    headers: tuple[str, ...]
    rows: tuple[tuple[str, ...], ...]

    @classmethod
    def from_dicts(cls, records: list[dict[str, Any]]) -> TableResult:
        if not records:
            return cls(headers=(), rows=())
        headers = tuple(str(key) for key in records[0])
        rows = tuple(tuple(str(record.get(key, "")) for key in headers) for record in records)
        return cls(headers=headers, rows=rows)

    @classmethod
    def from_sequences(
        cls, rows: list[list[Any]] | list[tuple[Any, ...]], *, headers: list[str] | None = None
    ) -> TableResult:
        if headers is None:
            if not rows:
                return cls(headers=(), rows=())
            headers = [str(value) for value in rows[0]]
            data = rows[1:]
        else:
            data = rows
        return cls(
            headers=tuple(headers),
            rows=tuple(tuple(str(value) for value in row) for row in data),
        )


@dataclass(frozen=True)
class ImageResult:
    path: str | None = None
    url: str | None = None
    data: bytes | None = None
    mime: str = "image/png"
    figure: Any | None = None


@dataclass(frozen=True)
class ErrorResult:
    message: str
    formula: str = ""


@dataclass(frozen=True)
class RepeatDirective:
    key: str


def normalize_result(
    value: Any,
) -> ScalarResult | TableResult | ImageResult | ErrorResult | RepeatDirective:
    """Coerce a function return value into a typed result."""
    if isinstance(value, (ScalarResult, TableResult, ImageResult, ErrorResult, RepeatDirective)):
        return value
    if isinstance(value, Exception):
        return ErrorResult(str(value))
    if isinstance(value, Path):
        return ImageResult(path=str(value))
    if isinstance(value, bytes):
        return ImageResult(data=value)
    if isinstance(value, dict):
        if {"headers", "rows"} <= value.keys():
            return TableResult.from_sequences(list(value["rows"]), headers=list(value["headers"]))
        return ScalarResult(str(value))
    if isinstance(value, list):
        if value and isinstance(value[0], dict):
            return TableResult.from_dicts(value)
        if value and isinstance(value[0], (list, tuple)):
            return TableResult.from_sequences(value)
        return ScalarResult("\n".join(str(item) for item in value))
    return ScalarResult("" if value is None else str(value))
