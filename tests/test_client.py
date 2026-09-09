from __future__ import annotations

from pathlib import Path

import pytest

from sliger.client import IMAGE_PLACEHOLDER, Sliger
from sliger.exceptions import CredentialsError, ImageNotFoundError, SlideNotFoundError
from sliger.results import ErrorResult, ImageResult, RepeatDirective, ScalarResult, TableResult


def test_image_placeholder_regex() -> None:
    assert IMAGE_PLACEHOLDER.fullmatch("![image](graph.jpg)")
    assert IMAGE_PLACEHOLDER.fullmatch("![image](https://example.com/a.png)")
    assert IMAGE_PLACEHOLDER.fullmatch("![image]({{ generate_image }})")
    assert IMAGE_PLACEHOLDER.fullmatch("not an image") is None
    assert IMAGE_PLACEHOLDER.fullmatch("![image](a) trailing") is None


def test_credentials_error_on_missing_file(tmp_path: Path) -> None:
    with pytest.raises(CredentialsError, match="not found"):
        Sliger(tmp_path / "nope.json", "pres")


def test_credentials_error_on_invalid_file(tmp_path: Path, monkeypatch) -> None:
    creds = tmp_path / "creds.json"
    creds.write_text('{"type": "service_account"}')

    def boom(*args, **kwargs):
        raise ValueError("not a service account")

    monkeypatch.setattr(
        "sliger.auth.service_account.Credentials.from_service_account_file",
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
    result = sliger_client.jinjify()
    assert result.updates == 0
    assert result.changes == ()
    assert result.dry_run is False


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
    result = sliger_client.jinjify({"name": "Ada"})
    assert result.updates == 2  # delete + insert for the one changed box
    assert result.dry_run is False
    assert len(result.changes) == 1
    assert result.changes[0].original == "Hello {{ name }}"
    assert result.changes[0].rendered == "Hello Ada"
    object_ids = [req[next(iter(req))]["objectId"] for req in captured]
    assert object_ids == ["shape-1", "shape-1", "shape-1"]


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
    deleted_files: list[str] = []
    monkeypatch.setattr(
        "sliger.drive_utils.delete_file", lambda service, file_id: deleted_files.append(file_id)
    )
    monkeypatch.setattr(
        "sliger.drive_utils.drive_image_url",
        lambda file_id: f"https://drive.google.com/uc?export=view&id={file_id}",
    )
    batches = []

    def fake_batch(service, presentation_id, requests):
        batches.append(requests)
        return {}

    monkeypatch.setattr("sliger.slides_utils.batch_update", fake_batch)

    result = sliger_client.imagify()
    assert result.replaced == 2
    assert result.dry_run is False
    urls = [batch[0]["createImage"]["url"] for batch in batches]
    assert urls[0].endswith("drive-file")
    assert urls[1] == "https://cdn.example/a.png"
    deleted = [batch[1]["deleteObject"]["objectId"] for batch in batches]
    assert deleted == ["local-shape", "url-shape"]
    assert deleted_files == ["drive-file"]


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


def test_jinjify_dry_run_does_not_batch_update(
    sliger_client: Sliger, monkeypatch, sample_slide: dict
) -> None:
    monkeypatch.setattr(
        "sliger.slides_utils.get_presentation_slides",
        lambda *a, **k: [sample_slide],
    )

    def fail(*args, **kwargs):
        raise AssertionError("should not update")

    monkeypatch.setattr("sliger.slides_utils.batch_update", fail)
    result = sliger_client.jinjify({"name": "Ada"}, dry_run=True)
    assert result.dry_run is True
    assert result.updates == 2
    assert len(result.changes) == 1
    change = result.changes[0]
    assert change.slide_number == 1
    assert change.object_id == "shape-1"
    assert change.original == "Hello {{ name }}"
    assert change.rendered == "Hello Ada"


def test_imagify_dry_run_does_not_upload(
    sliger_client: Sliger, monkeypatch, text_element_factory, tmp_path: Path
) -> None:
    local = tmp_path / "graph.jpg"
    local.write_bytes(b"fake")
    slide = {
        "objectId": "slide-1",
        "pageElements": [
            text_element_factory("local-shape", f"![image]({local})"),
            text_element_factory("url-shape", "![image](https://cdn.example/a.png)"),
        ],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])

    def fail_upload(*args, **kwargs):
        raise AssertionError("should not upload")

    def fail_batch(*args, **kwargs):
        raise AssertionError("should not update")

    monkeypatch.setattr("sliger.drive_utils.upload_image_to_drive", fail_upload)
    monkeypatch.setattr("sliger.slides_utils.batch_update", fail_batch)

    result = sliger_client.imagify(dry_run=True)
    assert result.dry_run is True
    assert result.replaced == 2
    assert result.changes[0].object_id == "local-shape"
    assert result.changes[0].image_ref == str(local)
    assert result.changes[0].resolved == str(local)
    assert result.changes[1].image_ref == "https://cdn.example/a.png"
    assert result.changes[1].resolved == "https://cdn.example/a.png"


def test_jinjify_persists_formula_as_alt_text(
    sliger_client: Sliger, monkeypatch, sample_slide: dict
) -> None:
    monkeypatch.setattr(
        "sliger.slides_utils.get_presentation_slides",
        lambda *a, **k: [sample_slide],
    )
    captured: list[dict] = []

    def fake_batch(service, presentation_id, requests):
        captured.extend(requests)
        return {}

    monkeypatch.setattr("sliger.slides_utils.batch_update", fake_batch)
    sliger_client.jinjify({"name": "Ada"})
    alt = [req["updatePageElementAltText"] for req in captured if "updatePageElementAltText" in req]
    assert alt == [{"objectId": "shape-1", "description": "sliger:Hello {{ name }}"}]


def test_jinjify_reuses_sliger_description_as_formula(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    element = text_element_factory("shape-1", "Hello Ada")
    element["description"] = "sliger:Hello {{ name }}"
    slide = {"objectId": "slide-1", "pageElements": [element]}
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    captured: list[dict] = []
    monkeypatch.setattr(
        "sliger.slides_utils.batch_update",
        lambda service, presentation_id, requests: captured.extend(requests) or {},
    )
    result = sliger_client.jinjify({"name": "Grace"})
    assert len(result.changes) == 1
    change = result.changes[0]
    assert change.original == "Hello Ada"
    assert change.rendered == "Hello Grace"
    assert change.formula == "Hello {{ name }}"
    assert any(
        req.get("updatePageElementAltText", {}).get("description") == "sliger:Hello {{ name }}"
        for req in captured
    )


def test_jinjify_table_result_creates_table(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("shape-1", "{{ table }}")],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    monkeypatch.setattr(
        "sliger.jinja_utils.render_box",
        lambda *a, **k: TableResult(headers=("a",), rows=(("1",),)),
    )
    captured: list[dict] = []
    monkeypatch.setattr(
        "sliger.slides_utils.batch_update",
        lambda service, presentation_id, requests: captured.extend(requests) or {},
    )
    result = sliger_client.jinjify()
    assert result.changes[0].kind == "table"
    kinds = [next(iter(req)) for req in captured]
    assert "createTable" in kinds
    assert "deleteObject" in kinds
    assert any(req.get("deleteObject", {}).get("objectId") == "shape-1" for req in captured)


def test_jinjify_error_result_does_not_raise(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("shape-1", "{{ boom }}")],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    monkeypatch.setattr(
        "sliger.jinja_utils.render_box",
        lambda *a, **k: ErrorResult("kaboom"),
    )
    monkeypatch.setattr("sliger.slides_utils.batch_update", lambda *a, **k: {})
    result = sliger_client.jinjify()
    assert result.changes[0].kind == "error"
    assert result.changes[0].error == "kaboom"
    assert result.errors == ("kaboom",)


def test_expand_repeaters_dry_run_does_not_duplicate(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("shape-1", "{{ sliger_repeat('deals') }}")],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    monkeypatch.setattr(
        "sliger.jinja_utils.render_box",
        lambda *a, **k: RepeatDirective("deals"),
    )

    def fail(*args, **kwargs):
        raise AssertionError("should not duplicate")

    monkeypatch.setattr("sliger.slides_utils.batch_update", fail)
    changes = sliger_client.expand_repeaters({"deals": [1, 2, 3]}, dry_run=True)
    assert len(changes) == 1
    assert changes[0].kind == "repeat"
    assert changes[0].rendered == "[repeat deals x3]"
    assert sliger_client._item_by_slide == {}


def test_expand_repeaters_duplicates_for_each_item(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("shape-1", "{{ sliger_repeat('deals') }}")],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    monkeypatch.setattr(
        "sliger.jinja_utils.render_box",
        lambda *a, **k: RepeatDirective("deals"),
    )
    batches: list[list] = []

    def fake_batch(service, presentation_id, requests):
        batches.append(list(requests))
        if requests and "duplicateObject" in requests[0]:
            return {
                "replies": [
                    {"duplicateObject": {"objectId": "dup-1"}},
                    {"duplicateObject": {"objectId": "dup-2"}},
                ]
            }
        return {}

    monkeypatch.setattr("sliger.slides_utils.batch_update", fake_batch)
    changes = sliger_client.expand_repeaters({"deals": [1, 2, 3]})
    assert changes[0].kind == "repeat"
    dup_requests = [req for batch in batches for req in batch if "duplicateObject" in req]
    assert dup_requests == [
        {"duplicateObject": {"objectId": "slide-1"}},
        {"duplicateObject": {"objectId": "slide-1"}},
    ]
    assert sliger_client._item_by_slide == {"slide-1": 1, "dup-1": 2, "dup-2": 3}


def test_render_dry_run_does_not_copy(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("shape-1", "Hello {{ name }}")],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    copied: list[str] = []
    monkeypatch.setattr(
        sliger_client,
        "duplicate_presentation",
        lambda *a, **k: copied.append("called") or "new-id",
    )
    result = sliger_client.render("Title", {"name": "Ada"}, dry_run=True)
    assert copied == []
    assert result.dry_run is True
    assert result.presentation_id == "pres-123"
    assert result.url == "https://docs.google.com/presentation/d/pres-123/edit"


def test_render_copies_then_jinjify_imagify(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("shape-1", "Hello {{ name }}")],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    monkeypatch.setattr("sliger.slides_utils.batch_update", lambda *a, **k: {})
    dup_titles: list[str] = []
    monkeypatch.setattr(
        sliger_client,
        "duplicate_presentation",
        lambda title, **k: dup_titles.append(title) or "new-id",
    )
    with_ids: list[str] = []
    real_with = sliger_client.with_presentation

    def fake_with(presentation_id: str) -> Sliger:
        with_ids.append(presentation_id)
        child = real_with(presentation_id)
        child._slides_service = sliger_client._slides_service
        child._drive_service = sliger_client._drive_service
        return child

    monkeypatch.setattr(sliger_client, "with_presentation", fake_with)
    result = sliger_client.render("Title", {"name": "Ada"})
    assert dup_titles == ["Title"]
    assert with_ids == ["new-id"]
    assert result.presentation_id == "new-id"
    assert result.dry_run is False
    assert result.url == "https://docs.google.com/presentation/d/new-id/edit"


def test_inspect_returns_formulas_from_sample_slide(
    sliger_client: Sliger, monkeypatch, sample_slide: dict
) -> None:
    monkeypatch.setattr(
        "sliger.slides_utils.get_presentation_slides",
        lambda *a, **k: [sample_slide],
    )
    report = sliger_client.inspect()
    assert [item.formula for item in report.formulas] == ["Hello {{ name }}"]
    assert report.formulas[0].object_id == "shape-1"
    assert report.formulas[0].slide_number == 1
    assert report.variables == ("name",)


def test_eval_template_adds_one_and_one(sliger_client: Sliger) -> None:
    result = sliger_client.eval_template("{{ 1+1 }}")
    assert isinstance(result, ScalarResult)
    assert result.text == "2"


def test_write_provenance_inserts_notes_when_shape_id_present(
    sliger_client: Sliger, monkeypatch, sample_slide: dict
) -> None:
    monkeypatch.setattr(
        "sliger.slides_utils.get_presentation_slides",
        lambda *a, **k: [sample_slide],
    )
    monkeypatch.setattr("sliger.client.notes_shape_id", lambda slide: "notes-shape")
    captured: list[dict] = []
    monkeypatch.setattr(
        "sliger.slides_utils.batch_update",
        lambda service, presentation_id, requests: captured.extend(requests) or {},
    )
    sliger_client.write_provenance({"changes": 3})
    # sample_slide has no notes text; deleteText on empty notes 400s.
    assert len(captured) == 1
    assert captured[0]["insertText"]["objectId"] == "notes-shape"
    assert "sliger" in captured[0]["insertText"]["text"]
    assert "changes: 3" in captured[0]["insertText"]["text"]


def test_write_provenance_skips_when_notes_shape_id_is_none(
    sliger_client: Sliger, monkeypatch, sample_slide: dict
) -> None:
    monkeypatch.setattr(
        "sliger.slides_utils.get_presentation_slides",
        lambda *a, **k: [sample_slide],
    )
    monkeypatch.setattr("sliger.client.notes_shape_id", lambda slide: None)

    def fail(*args, **kwargs):
        raise AssertionError("should not update")

    monkeypatch.setattr("sliger.slides_utils.batch_update", fail)
    sliger_client.write_provenance()


def test_imagify_image_result_uses_materialize_image(
    sliger_client: Sliger, monkeypatch, text_element_factory
) -> None:
    slide = {
        "objectId": "slide-1",
        "pageElements": [text_element_factory("img-shape", "{{ chart }}")],
    }
    monkeypatch.setattr("sliger.slides_utils.get_presentation_slides", lambda *a, **k: [slide])
    boxed = ImageResult(url="https://cdn.example/chart.png")
    monkeypatch.setattr("sliger.jinja_utils.render_box", lambda *a, **k: boxed)
    seen: list[ImageResult] = []

    def fake_materialize(result: ImageResult) -> str:
        seen.append(result)
        return "https://materialized.example/chart.png"

    monkeypatch.setattr("sliger.client.materialize_image", fake_materialize)
    batches: list[list] = []
    monkeypatch.setattr(
        "sliger.slides_utils.batch_update",
        lambda service, presentation_id, requests: batches.append(requests) or {},
    )
    result = sliger_client.imagify()
    assert seen == [boxed]
    assert result.replaced == 1
    assert result.changes[0].image_ref == "https://materialized.example/chart.png"
    assert result.changes[0].resolved == "https://materialized.example/chart.png"
    assert batches[0][0]["createImage"]["url"] == "https://materialized.example/chart.png"
    assert batches[0][1]["deleteObject"]["objectId"] == "img-shape"
