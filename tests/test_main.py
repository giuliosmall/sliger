from __future__ import annotations

import runpy


def test_module_entrypoint(monkeypatch) -> None:
    called = []
    monkeypatch.setattr("sliger.cli.app", lambda: called.append(True))
    runpy.run_module("sliger", run_name="__main__")
    assert called == [True]
