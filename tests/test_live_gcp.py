"""Live E2E against Google Slides/Drive.

Skipped unless SLIGER_LIVE=1.

Uses SLIGER_CREDS_FILE if set, otherwise Application Default Credentials.
Optional SLIGER_SHARED_DRIVE_ID lets a service account create files (SAs have
no My Drive quota). Imagify uses a public HTTPS URL so it does not need Drive
upload.
"""

from __future__ import annotations

import os
import sqlite3
import time
import uuid
from contextlib import suppress
from datetime import datetime, timezone
from pathlib import Path

import pytest
from googleapiclient.errors import HttpError

from sliger import ScalarResult, Sliger
from sliger.slides_utils import batch_update, get_text_elements_from_slide, gslides_element_to_text

LIVE = os.environ.get("SLIGER_LIVE") == "1"
PUBLIC_IMAGE = "https://www.google.com/images/branding/googlelogo/2x/googlelogo_color_92x30dp.png"
_REPO_TOKEN = Path(__file__).resolve().parents[1] / ".secrets" / "user-token.json"


def _creds_file() -> Path | None:
    """Prefer SLIGER_CREDS_FILE when it is a real file; else repo .secrets/user-token.json."""
    raw = os.environ.get("SLIGER_CREDS_FILE")
    if raw and "..." not in Path(raw).parts and Path(raw).is_file():
        return Path(raw)
    if _REPO_TOKEN.is_file():
        return _REPO_TOKEN
    return None


def _client(*, config_path: str | Path | None = None) -> Sliger:
    creds = _creds_file()
    kwargs: dict = {"full_drive": True}
    if config_path is not None:
        kwargs["config_path"] = config_path
    if creds:
        return Sliger(creds, "pending", **kwargs)
    return Sliger(None, "pending", use_adc=True, **kwargs)


def _live_title(kind: str) -> str:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    return f"sliger-live-{kind}-{stamp}-{uuid.uuid4().hex[:6]}"


def _oid() -> str:
    return "s" + uuid.uuid4().hex[:16]


def _delete_presentations(client: Sliger, file_ids: list[str]) -> None:
    for file_id in file_ids:
        with suppress(HttpError):
            client.drive_service.files().delete(fileId=file_id, supportsAllDrives=True).execute()


def _create_presentation(client: Sliger, title: str) -> dict:
    parent = os.environ.get("SLIGER_SHARED_DRIVE_ID")
    if parent:
        created = (
            client.drive_service.files()
            .create(
                body={
                    "name": title,
                    "mimeType": "application/vnd.google-apps.presentation",
                    "parents": [parent],
                },
                supportsAllDrives=True,
                fields="id,name",
            )
            .execute()
        )
        presentation_id = created["id"]
        return client.slides_service.presentations().get(presentationId=presentation_id).execute()
    return client.slides_service.presentations().create(body={"title": title}).execute()


def _seed_template_boxes(client: Sliger, presentation_id: str, page_id: str) -> None:
    batch_update(
        client.slides_service,
        presentation_id,
        [
            {
                "createShape": {
                    "objectId": "sligerTextBox",
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {
                            "height": {"magnitude": 80, "unit": "PT"},
                            "width": {"magnitude": 400, "unit": "PT"},
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 40,
                            "translateY": 40,
                            "unit": "PT",
                        },
                    },
                }
            },
            {
                "insertText": {
                    "objectId": "sligerTextBox",
                    "insertionIndex": 0,
                    "text": "hello {{ probe }}",
                }
            },
            {
                "createShape": {
                    "objectId": "sligerImageBox",
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {
                            "height": {"magnitude": 60, "unit": "PT"},
                            "width": {"magnitude": 200, "unit": "PT"},
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 40,
                            "translateY": 140,
                            "unit": "PT",
                        },
                    },
                }
            },
            {
                "insertText": {
                    "objectId": "sligerImageBox",
                    "insertionIndex": 0,
                    "text": f"![image]({PUBLIC_IMAGE})",
                }
            },
        ],
    )


def _text_on_first_slide(client: Sliger) -> list[str]:
    slides = (
        client.slides_service.presentations()
        .get(presentationId=client.presentation_id)
        .execute()["slides"]
    )
    page_id = slides[0]["objectId"]
    return [
        gslides_element_to_text(el, page_id)["text"]
        for el in get_text_elements_from_slide(slides[0])
    ]


def _fetch_presentation(client: Sliger, presentation_id: str | None = None) -> dict:
    pid = presentation_id or client.presentation_id
    return client.slides_service.presentations().get(presentationId=pid).execute()


def _visible_texts(client: Sliger, presentation_id: str | None = None) -> list[str]:
    presentation = _fetch_presentation(client, presentation_id)
    texts: list[str] = []
    for slide in presentation.get("slides") or []:
        page_id = slide.get("objectId") or ""
        texts.extend(
            gslides_element_to_text(el, page_id)["text"]
            for el in get_text_elements_from_slide(slide)
        )
    return texts


def _seed_boxes(client: Sliger, presentation_id: str, page_id: str, texts: list[str]) -> list[str]:
    requests: list[dict] = []
    object_ids: list[str] = []
    for index, text in enumerate(texts):
        object_id = _oid()
        object_ids.append(object_id)
        requests.append(
            {
                "createShape": {
                    "objectId": object_id,
                    "shapeType": "TEXT_BOX",
                    "elementProperties": {
                        "pageObjectId": page_id,
                        "size": {
                            "height": {"magnitude": 80, "unit": "PT"},
                            "width": {"magnitude": 400, "unit": "PT"},
                        },
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 40,
                            "translateY": 40 + index * 100,
                            "unit": "PT",
                        },
                    },
                }
            }
        )
        requests.append(
            {
                "insertText": {
                    "objectId": object_id,
                    "insertionIndex": 0,
                    "text": text,
                }
            }
        )
    batch_update(client.slides_service, presentation_id, requests)
    return object_ids


def _table_elements(presentation: dict) -> list[dict]:
    found: list[dict] = []
    for slide in presentation.get("slides") or []:
        for element in slide.get("pageElements") or []:
            if element.get("table"):
                found.append(element)
    return found


def _table_cell_texts(table_element: dict) -> list[str]:
    texts: list[str] = []
    table = table_element.get("table") or {}
    for row in table.get("tableRows") or []:
        for cell in row.get("tableCells") or []:
            parts = [
                (te.get("textRun") or {}).get("content") or ""
                for te in (cell.get("text") or {}).get("textElements") or []
            ]
            texts.append("".join(parts))
    return texts


@pytest.mark.live
@pytest.mark.skipif(not LIVE, reason="set SLIGER_LIVE=1 to call Google APIs")
def test_e2e_throwaway_deck() -> None:
    title = f"sliger-live-{datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')}"
    client = _client()
    created_ids = []
    try:
        created = _create_presentation(client, title)
        presentation_id = created["presentationId"]
        created_ids.append(presentation_id)
        client.presentation_id = presentation_id
        page_id = created["slides"][0]["objectId"]
        _seed_template_boxes(client, presentation_id, page_id)

        preview = client.jinjify({"probe": "gcp"}, dry_run=True)
        assert preview.dry_run is True
        assert any(c.rendered == "hello gcp" for c in preview.changes)

        applied = client.jinjify({"probe": "gcp"})
        assert applied.dry_run is False
        assert applied.updates >= 1
        assert "hello gcp" in _text_on_first_slide(client)

        image_preview = client.imagify(dry_run=True)
        assert image_preview.dry_run is True
        assert image_preview.replaced >= 1

        image_applied = client.imagify()
        assert image_applied.dry_run is False
        assert image_applied.replaced >= 1
        assert not any(t.startswith("![image]") for t in _text_on_first_slide(client))

        before = len(
            client.slides_service.presentations()
            .get(presentationId=presentation_id)
            .execute()["slides"]
        )
        client.duplicate_slide(1)
        after_dup = len(
            client.slides_service.presentations()
            .get(presentationId=presentation_id)
            .execute()["slides"]
        )
        assert after_dup == before + 1
        client.delete_slide(after_dup)
        after_del = len(
            client.slides_service.presentations()
            .get(presentationId=presentation_id)
            .execute()["slides"]
        )
        assert after_del == before

        try:
            copy_id = client.duplicate_presentation(f"{title}-copy")
            created_ids.append(copy_id)
        except HttpError as exc:
            # User ADC from gcloud's default client cannot request Drive.
            if exc.status_code != 403:
                raise
    finally:
        _delete_presentations(client, created_ids)


@pytest.mark.live
@pytest.mark.skipif(not LIVE, reason="set SLIGER_LIVE=1 to call Google APIs")
def test_e2e_formula_survives_jinjify() -> None:
    client = _client()
    created_ids: list[str] = []
    try:
        created = _create_presentation(client, _live_title("formula"))
        presentation_id = created["presentationId"]
        created_ids.append(presentation_id)
        client.presentation_id = presentation_id
        page_id = created["slides"][0]["objectId"]
        _seed_boxes(client, presentation_id, page_id, ["v={{ n }}"])

        first = client.jinjify({"n": "1"})
        assert first.dry_run is False
        texts = _text_on_first_slide(client)
        assert any("v=1" in t for t in texts)
        assert not any("{{ n }}" in t for t in texts)

        second = client.jinjify({"n": "2"})
        assert second.dry_run is False
        texts = _text_on_first_slide(client)
        assert any("v=2" in t for t in texts)
        assert not any("v=1" in t for t in texts)
    finally:
        _delete_presentations(client, created_ids)


@pytest.mark.live
@pytest.mark.skipif(not LIVE, reason="set SLIGER_LIVE=1 to call Google APIs")
def test_e2e_table_from_sql(tmp_path: Path) -> None:
    db_path = tmp_path / "accounts.db"
    conn = sqlite3.connect(db_path)
    try:
        conn.execute("CREATE TABLE accounts (id INTEGER, name TEXT)")
        conn.execute("INSERT INTO accounts (id, name) VALUES (1, 'Alice'), (2, 'Bob')")
        conn.commit()
    finally:
        conn.close()

    config_path = tmp_path / "config.toml"
    config_path.write_text(
        f'[connections.db]\ntype = "sqlite"\nurl = "{db_path.resolve()}"\n',
        encoding="utf-8",
    )

    client = _client(config_path=config_path)
    created_ids: list[str] = []
    try:
        created = _create_presentation(client, _live_title("sql-table"))
        presentation_id = created["presentationId"]
        created_ids.append(presentation_id)
        client.presentation_id = presentation_id
        page_id = created["slides"][0]["objectId"]
        _seed_boxes(
            client,
            presentation_id,
            page_id,
            ["{{ sql('select id, name from accounts order by id') }}"],
        )

        applied = client.jinjify()
        assert applied.dry_run is False
        assert any(change.kind == "table" for change in applied.changes)

        presentation = None
        tables: list[dict] = []
        for _ in range(6):
            presentation = _fetch_presentation(client)
            tables = _table_elements(presentation)
            if tables:
                break
            time.sleep(0.5)
        assert tables, "expected a table pageElement after sql() jinjify"
        cells = " ".join(_table_cell_texts(tables[0]))
        assert "Alice" in cells
        assert "Bob" in cells
    finally:
        _delete_presentations(client, created_ids)


@pytest.mark.live
@pytest.mark.skipif(not LIVE, reason="set SLIGER_LIVE=1 to call Google APIs")
def test_e2e_table_from_bigquery(tmp_path: Path) -> None:
    pytest.importorskip("google.cloud.bigquery")
    project = (
        os.environ.get("SLIGER_BQ_PROJECT")
        or os.environ.get("GOOGLE_CLOUD_PROJECT")
        or os.environ.get("GOOGLE_CLOUD_QUOTA_PROJECT")
    )
    if not project:
        pytest.skip("needs SLIGER_BQ_PROJECT or GOOGLE_CLOUD_PROJECT")

    config_path = tmp_path / "config.toml"
    config_path.write_text(
        f'[connections.warehouse]\ntype = "bigquery"\nproject = "{project}"\n',
        encoding="utf-8",
    )

    client = _client(config_path=config_path)
    created_ids: list[str] = []
    try:
        created = _create_presentation(client, _live_title("bq-table"))
        presentation_id = created["presentationId"]
        created_ids.append(presentation_id)
        client.presentation_id = presentation_id
        page_id = created["slides"][0]["objectId"]
        _seed_boxes(
            client,
            presentation_id,
            page_id,
            ["{{ sql('select :left as name, 1 as n union all select :right as name, 2 as n') }}"],
        )

        applied = client.jinjify({"left": "Ada", "right": "Bob"})
        assert applied.dry_run is False
        assert any(change.kind == "table" for change in applied.changes)

        presentation = None
        tables: list[dict] = []
        for _ in range(6):
            presentation = _fetch_presentation(client)
            tables = _table_elements(presentation)
            if tables:
                break
            time.sleep(0.5)
        assert tables, "expected a table pageElement after BigQuery sql() jinjify"
        cells = " ".join(_table_cell_texts(tables[0]))
        assert "Ada" in cells
        assert "Bob" in cells
        assert "name" in cells
    finally:
        _delete_presentations(client, created_ids)


@pytest.mark.live
@pytest.mark.skipif(not LIVE, reason="set SLIGER_LIVE=1 to call Google APIs")
def test_e2e_inspect_and_render() -> None:
    client = _client()
    created_ids: list[str] = []
    try:
        created = _create_presentation(client, _live_title("inspect"))
        presentation_id = created["presentationId"]
        created_ids.append(presentation_id)
        client.presentation_id = presentation_id
        page_id = created["slides"][0]["objectId"]
        _seed_boxes(client, presentation_id, page_id, ["{{ company }}"])

        filled = client.inspect({"company": "Acme"})
        assert "company" not in filled.missing_data

        empty = client.inspect({})
        assert "company" in empty.missing_data

        rendered = client.render(_live_title("render"), {"company": "Acme"})
        created_ids.append(rendered.presentation_id)
        assert rendered.presentation_id != presentation_id
        assert rendered.dry_run is False
        assert rendered.jinjify.dry_run is False
        texts = _visible_texts(client, rendered.presentation_id)
        assert any("Acme" in t for t in texts)
        assert not any("{{ company }}" in t for t in texts)
    finally:
        _delete_presentations(client, created_ids)


@pytest.mark.live
@pytest.mark.skipif(not LIVE, reason="set SLIGER_LIVE=1 to call Google APIs")
def test_e2e_repeat_slides() -> None:
    client = _client()
    created_ids: list[str] = []
    try:
        created = _create_presentation(client, _live_title("repeat"))
        presentation_id = created["presentationId"]
        created_ids.append(presentation_id)
        client.presentation_id = presentation_id
        page_id = created["slides"][0]["objectId"]
        _seed_boxes(
            client,
            presentation_id,
            page_id,
            ["{{ sliger_repeat('items') }}", "{{ item }}"],
        )

        before = len(_fetch_presentation(client)["slides"])
        applied = client.jinjify({"items": ["A", "B"]})
        assert applied.dry_run is False
        after = len(_fetch_presentation(client)["slides"])
        assert after == before + 1

        texts = {t.strip() for t in _visible_texts(client)}
        assert "A" in texts
        assert "B" in texts
    finally:
        _delete_presentations(client, created_ids)


@pytest.mark.live
@pytest.mark.skipif(not LIVE, reason="set SLIGER_LIVE=1 to call Google APIs")
def test_e2e_error_isolation() -> None:
    client = _client()
    created_ids: list[str] = []
    try:
        created = _create_presentation(client, _live_title("errors"))
        presentation_id = created["presentationId"]
        created_ids.append(presentation_id)
        client.presentation_id = presentation_id
        page_id = created["slides"][0]["objectId"]
        _seed_boxes(client, presentation_id, page_id, ["{{ 1/0 }}", "{{ ok }}"])

        applied = client.jinjify({"ok": "yes"})
        assert applied.dry_run is False
        assert applied.errors
        texts = _visible_texts(client)
        blob = "\n".join(texts).lower()
        assert any(t.strip() == "yes" for t in texts)
        assert "sliger error" in blob
    finally:
        _delete_presentations(client, created_ids)


@pytest.mark.live
@pytest.mark.skipif(not LIVE, reason="set SLIGER_LIVE=1 to call Google APIs")
def test_e2e_eval_template() -> None:
    client = _client()
    result = client.eval_template("{{ 2+2 }}")
    assert isinstance(result, ScalarResult)
    assert result.text == "4"
