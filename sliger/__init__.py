"""Sliger: slide the power of Python (and Jinja2) into Google Slides."""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError, version

from sliger.client import (
    ImageChange,
    ImagifyResult,
    JinjifyResult,
    RenderResult,
    Sliger,
    TextChange,
)
from sliger.exceptions import (
    ConfigError,
    CredentialsError,
    GoogleAPIError,
    ImageNotFoundError,
    SlideNotFoundError,
    SligerError,
)
from sliger.results import (
    ErrorResult,
    ImageResult,
    RepeatDirective,
    ScalarResult,
    TableResult,
)

try:
    __version__ = version("sliger")
except PackageNotFoundError:  # pragma: no cover - editable checkouts without install
    __version__ = "0.0.0"

__all__ = [
    "Sliger",
    "TextChange",
    "JinjifyResult",
    "ImageChange",
    "ImagifyResult",
    "RenderResult",
    "ScalarResult",
    "TableResult",
    "ImageResult",
    "ErrorResult",
    "RepeatDirective",
    "SligerError",
    "CredentialsError",
    "GoogleAPIError",
    "SlideNotFoundError",
    "ConfigError",
    "ImageNotFoundError",
    "__version__",
]
