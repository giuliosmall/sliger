"""Import-only smoke test for an installed sliger wheel (no Google APIs)."""

from __future__ import annotations


def test_wheel_import() -> None:
    import sliger
    from sliger.cli import app

    version = sliger.__version__
    assert isinstance(version, str) and version
    assert app is not None
    print(f"sliger {version}")


if __name__ == "__main__":
    test_wheel_import()
    print("ok")
