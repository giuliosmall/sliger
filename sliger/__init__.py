"""Sliger: slide the power of Python (and Jinja2) into Google Slides."""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError, version

from sliger.client import Sliger
from sliger.exceptions import (
    ConfigError,
    CredentialsError,
    GoogleAPIError,
    ImageNotFoundError,
    SlideNotFoundError,
    SligerError,
)

try:
    __version__ = version("sliger")
except PackageNotFoundError:  # pragma: no cover - editable checkouts without install
    __version__ = "0.0.0"

__all__ = [
    "Sliger",
    "SligerError",
    "CredentialsError",
    "GoogleAPIError",
    "SlideNotFoundError",
    "ConfigError",
    "ImageNotFoundError",
    "__version__",
]
