from __future__ import annotations

from pathlib import Path

import pytest

from sliger.client import IMAGE_PLACEHOLDER, Sliger
from sliger.exceptions import CredentialsError, ImageNotFoundError, SlideNotFoundError


def test_image_placeholder_regex() -> None:
    assert IMAGE_PLACEHOLDER.fullmatch("![image](graph.jpg)")
    assert IMAGE_PLACEHOLDER.fullmatch("![image](https://example.com/a.png)")
    assert IMAGE_PLACEHOLDER.fullmatch("![image]({{ generate_image }})")
    assert IMAGE_PLACEHOLDER.fullmatch("not an image") is None
    assert IMAGE_PLACEHOLDER.fullmatch("![image](a) trailing") is None


def test_credentials_error_on_missing_file(tmp_path: Path) -> None:
    with pytest.raises(CredentialsError):
        Sliger(tmp_path / "nope.json", "pres")


def test_credentials_error_on_invalid_file(tmp_path: Path, monkeypatch) -> None:
    creds = tmp_path / "creds.json"
    creds.write_text("{}")

    def boom(*args, **kwargs):
        raise ValueError("not a service account")

    monkeypatch.setattr(
        "sliger.client.service_account.Credentials.from_service_account_file",
        boom,
    )
    with pytest.raises(CredentialsError, match="Invalid credentials"):
        Sliger(creds, "pres")


def test_delete_slide_missing(sliger_client: Sliger, monkeypatch) -> None:
    monkeypatch.setattr(
        "sliger.slides_utils.get_presentation_slides",
        lambda *args, **kwargs: [{"objectId": "only"}],
    )
    with pytest.raises(SlideNotFoundError) as exc:
        sliger_client.delete_slide(3)
    assert exc.value.slide_number == 3
    assert exc.value.slide_count == 1


def test_duplicate_slide_ok(sliger_client: Sliger, monkeypatch) -> None:
    called = {}
    monkeypatch.setattr(
        "sliger.slides_utils.get_presentation_slides",
        lambda *args, **kwargs: [{"objectId": "s1"}],
    )
    monkeypatch.setattr(
        "sliger.slides_utils.duplicate_slide_by_id",
        lambda service, presentation_id, slide_id: called.update(id=slide_id) or {},
    )
    sliger_client.duplicate_slide(1)
    assert called["id"] == "s1"


def test_jinjify_noop_when_text_unchanged(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("shape-1", "plain text")],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])

    def fail(*args, **kwargs):
        raise AssertionError("should not update")

    monkeypatch.setattr("sliger.slides_utils.batch_update", fail)
    assert sliger_client.jinjify() == 0


def test_delete_slide_ok(sliger_client: Sliger, monkeypatch) -> None:
    called = {}
    monkeypatch.setattr(
        "sliger.slides_utils.get_presentation_slides",
        lambda *args, **kwargs: [{"objectId": "s1"}, {"objectId": "s2"}],
    )

    def fake_delete(service, presentation_id, slide_id):
        called["args"] = (presentation_id, slide_id)
        return {}

    monkeypatch.setattr("sliger.slides_utils.delete_slide_by_id", fake_delete)
    sliger_client.delete_slide(2)
    assert called["args"] == ("pres-123", "s2")


def test_duplicate_presentation_does_not_share_by_default(
    sliger_client: Sliger, monkeypatch
) -> None:
    monkeypatch.setattr("sliger.drive_utils.copy_file", lambda *a, **k: "new-id")
    shared = []
    monkeypatch.setattr(
        "sliger.drive_utils.set_anyone_permission",
        lambda *a, **k: shared.append(a),
    )
    assert sliger_client.duplicate_presentation("Title") == "new-id"
    assert shared == []


def test_lazy_google_services(sliger_client: Sliger, monkeypatch) -> None:
    sliger_client._slides_service = None
    sliger_client._drive_service = None
    monkeypatch.setattr("sliger.slides_utils.get_slides_service", lambda creds: "slides")
    monkeypatch.setattr("sliger.drive_utils.get_drive_service", lambda creds: "drive")
    assert sliger_client.slides_service == "slides"
    assert sliger_client.drive_service == "drive"


def test_duplicate_presentation_opt_in_share(sliger_client: Sliger, monkeypatch) -> None:
    monkeypatch.setattr("sliger.drive_utils.copy_file", lambda *a, **k: "new-id")
    roles = []
    monkeypatch.setattr(
        "sliger.drive_utils.set_anyone_permission",
        lambda service, file_id, role: roles.append(role) or "perm",
    )
    sliger_client.duplicate_presentation("Title", anyone_can_edit=True)
    assert roles == ["writer"]
    roles.clear()
    sliger_client.duplicate_presentation("Title", anyone_can_view=True)
    assert roles == ["reader"]


def test_jinjify_updates_changed_text_only(
    sliger_client: Sliger, monkeypatch, sample_slide: dict
) -> None:
    monkeypatch.setattr(
        "sliger.slides_utils.get_presentation_slides",
        lambda *a, **k: [sample_slide],
    )
    captured = []

    def fake_batch(service, presentation_id, requests):
        captured.extend(requests)
        return {}

    monkeypatch.setattr("sliger.slides_utils.batch_update", fake_batch)
    updates = sliger_client.jinjify({"name": "Ada"})
    assert updates == 2  # delete + insert for the one changed box
    object_ids = [req[next(iter(req))]["objectId"] for req in captured]
    assert object_ids == ["shape-1", "shape-1"]


def test_imagify_uploads_local_and_uses_url(
    sliger_client: Sliger, monkeypatch, text_element_factory, tmp_path: Path
) -> None:
    local = tmp_path / "graph.jpg"
    local.write_bytes(b"fake")
    slide = {
        "objectId": "slide-1",
        "pageElements": [
            text_element_factory("local-shape", f"![image]({local})"),
            text_element_factory("url-shape", "![image](https://cdn.example/a.png)"),
            text_element_factory("plain", "just text"),
        ],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    monkeypatch.setattr("sliger.drive_utils.upload_image_to_drive", lambda *a, **k: "drive-file")
    monkeypatch.setattr(
        "sliger.drive_utils.drive_image_url",
        lambda file_id: f"https://drive.google.com/uc?export=view&id={file_id}",
    )
    batches = []

    def fake_batch(service, presentation_id, requests):
        batches.append(requests)
        return {}

    monkeypatch.setattr("sliger.slides_utils.batch_update", fake_batch)

    replaced = sliger_client.imagify()
    assert replaced == 2
    urls = [batch[0]["createImage"]["url"] for batch in batches]
    assert urls[0].endswith("drive-file")
    assert urls[1] == "https://cdn.example/a.png"
    deleted = [batch[1]["deleteObject"]["objectId"] for batch in batches]
    assert deleted == ["local-shape", "url-shape"]


def test_imagify_missing_local_file(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("s", "![image](nope.jpg)")],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    with pytest.raises(ImageNotFoundError):
        sliger_client.imagify()
