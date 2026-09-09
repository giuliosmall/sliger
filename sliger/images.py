"""Turn ImageResult values into a path or HTTP URL."""

from __future__ import annotations

import tempfile
from pathlib import Path

from sliger.exceptions import ImageNotFoundError
from sliger.results import ImageResult


def materialize_image(result: ImageResult) -> str:
    """Return a local path or http(s) URL the Slides API can fetch."""
    if result.url:
        return result.url
    if result.path:
        path = Path(result.path)
        if not path.is_file():
            raise ImageNotFoundError(f"Image path does not exist: {path}")
        return str(path)
    if result.figure is not None and hasattr(result.figure, "savefig"):
        with tempfile.NamedTemporaryFile(prefix="sliger-", suffix=".png", delete=False) as handle:
            path = handle.name
        result.figure.savefig(path)
        return path
    if result.data:
        suffix = ".png" if "png" in result.mime else ".img"
        with tempfile.NamedTemporaryFile(prefix="sliger-", suffix=suffix, delete=False) as handle:
            handle.write(result.data)
            return handle.name
    raise ImageNotFoundError("ImageResult has no path, url, data, or figure")
