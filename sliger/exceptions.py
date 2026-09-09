"""Public exception types for the sliger library."""

from __future__ import annotations


class SligerError(Exception):
    """Base class for all sliger errors."""


class CredentialsError(SligerError):
    """Raised when Google API credentials cannot be loaded."""


class GoogleAPIError(SligerError):
    """Raised when a Google API request fails."""


class SlideNotFoundError(SligerError):
    """Raised when a 1-based slide index is out of range."""

    def __init__(self, slide_number: int, slide_count: int) -> None:
        self.slide_number = slide_number
        self.slide_count = slide_count
        super().__init__(
            f"Slide number {slide_number} does not exist. Valid range is 1 to {slide_count}."
            if slide_count
            else f"Slide number {slide_number} does not exist. The presentation has no slides."
        )


class ConfigError(SligerError):
    """Raised when a Jinja config file cannot be loaded."""


class ImageNotFoundError(SligerError):
    """Raised when a local image path from a placeholder does not exist."""
