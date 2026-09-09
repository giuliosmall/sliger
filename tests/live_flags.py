"""Shared live-test environment flags."""

from __future__ import annotations

import os


def live_flag(*names: str) -> bool:
    """True when SLIGER_LIVE=1 or any of the named flags is 1."""
    if os.environ.get("SLIGER_LIVE") == "1":
        return True
    return any(os.environ.get(name) == "1" for name in names)
