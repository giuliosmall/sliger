"""Detect curly quotes that break Jinja in Google Slides."""

from __future__ import annotations

SMART_QUOTES = {
    "\u2018": "'",
    "\u2019": "'",
    "\u201c": '"',
    "\u201d": '"',
}


def find_smart_quotes(text: str) -> list[str]:
    return sorted({char for char in text if char in SMART_QUOTES})


def replace_smart_quotes(text: str) -> str:
    for curly, straight in SMART_QUOTES.items():
        text = text.replace(curly, straight)
    return text


def has_smart_quotes(text: str) -> bool:
    return bool(find_smart_quotes(text))
