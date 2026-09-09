from __future__ import annotations

from pathlib import Path

import pytest

from sliger.client import Sliger


@pytest.fixture
def text_element_factory():
    def factory(
        object_id: str,
        text: str,
        *,
        size: dict | None = None,
        transform: dict | None = None,
    ) -> dict:
        shape: dict = {"shapeType": "TEXT_BOX"}
        if text:
            shape["text"] = {
                "textElements": [
                    {"textRun": {"content": text}},
                ]
            }
        return {
            "objectId": object_id,
            "size": size
            or {
                "height": {"magnitude": 100, "unit": "PT"},
                "width": {"magnitude": 200, "unit": "PT"},
            },
            "transform": transform
            or {
                "scaleX": 1,
                "scaleY": 1,
                "translateX": 0,
                "translateY": 0,
                "unit": "PT",
            },
            "shape": shape,
        }

    return factory


@pytest.fixture
def sample_slide(text_element_factory):
    return {
        "objectId": "slide-1",
        "pageElements": [
            text_element_factory("shape-1", "Hello {{ name }}"),
            text_element_factory("shape-2", "![image](graph.jpg)"),
            {"objectId": "image-el", "image": {}},
        ],
    }


@pytest.fixture
def config_file(tmp_path: Path) -> Path:
    (tmp_path / "helpers.py").write_text("def greet(name: str) -> str:\n    return f'hi {name}'\n")
    path = tmp_path / "config.toml"
    path.write_text("[function_map]\ngreet = 'helpers.greet'\n")
    return path


@pytest.fixture
def sliger_client(monkeypatch, tmp_path: Path) -> Sliger:
    creds_file = tmp_path / "creds.json"
    creds_file.write_text("{}")

    monkeypatch.setattr(
        "sliger.auth.load_credentials",
        lambda **kwargs: "fake-creds",
    )
    client = Sliger(creds_file, "pres-123")
    client._slides_service = object()
    client._drive_service = object()
    return client
