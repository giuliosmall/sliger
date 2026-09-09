from __future__ import annotations

from pathlib import Path

import pytest

from sliger.exceptions import ImageNotFoundError
from sliger.images import materialize_image
from sliger.results import ImageResult


def test_materialize_image_url() -> None:
    url = "https://example.com/chart.png"
    assert materialize_image(ImageResult(url=url)) == url


def test_materialize_image_existing_path(tmp_path: Path) -> None:
    path = tmp_path / "chart.png"
    path.write_bytes(b"png-bytes")
    assert materialize_image(ImageResult(path=str(path))) == str(path)


def test_materialize_image_missing_path() -> None:
    with pytest.raises(ImageNotFoundError, match="Image path does not exist"):
        materialize_image(ImageResult(path="/definitely/missing/sliger-chart.png"))


def test_materialize_image_data_writes_temp_file() -> None:
    payload = b"\x89PNG\r\ndata"
    path = Path(materialize_image(ImageResult(data=payload, mime="image/png")))
    try:
        assert path.is_file()
        assert path.suffix == ".png"
        assert path.read_bytes() == payload
    finally:
        path.unlink(missing_ok=True)


def test_materialize_image_figure_savefig() -> None:
    saved: dict[str, str] = {}

    class Figure:
        def savefig(self, path: str) -> None:
            saved["path"] = path
            Path(path).write_bytes(b"figure-png")

    path = Path(materialize_image(ImageResult(figure=Figure())))
    try:
        assert saved["path"] == str(path)
        assert path.read_bytes() == b"figure-png"
    finally:
        path.unlink(missing_ok=True)


def test_materialize_empty_image_result_raises() -> None:
    with pytest.raises(ImageNotFoundError, match="no path, url, data, or figure"):
        materialize_image(ImageResult())
