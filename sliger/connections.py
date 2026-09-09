"""Named SQL connections loaded from sliger TOML config.

Implementation lives in ``sliger.connectors``. This module re-exports the
public surface so existing imports keep working.
"""

from sliger.connectors import (
    Connection,
    parse_connections,
    resolve_secret,
    run_sql,
)
from sliger.connectors.sqlite import sqlite_path

__all__ = [
    "Connection",
    "parse_connections",
    "resolve_secret",
    "run_sql",
    "sqlite_path",
]
