from __future__ import annotations

from pathlib import Path
from unittest.mock import Mock

import pytest
from googleapiclient.errors import HttpError

from sliger import drive_utils
from sliger.drive_utils import (
    copy_file,
    drive_image_url,
    set_anyone_permission,
    upload_image_to_drive,
)
from sliger.exceptions import GoogleAPIError, ImageNotFoundError


def test_get_drive_service(monkeypatch) -> None:
    monkeypatch.setattr(drive_utils, "build", lambda *a, **k: "service")
    assert drive_utils.get_drive_service("creds") == "service"


def test_drive_image_url() -> None:
    assert drive_image_url("abc") == "https://drive.google.com/uc?export=view&id=abc"


def test_upload_missing_file_raises(tmp_path: Path) -> None:
    with pytest.raises(ImageNotFoundError):
        upload_image_to_drive(object(), tmp_path / "missing.png")


def test_copy_file_returns_id() -> None:
    service = Mock()
    service.files.return_value.copy.return_value.execute.return_value = {"id": "copied"}
    assert copy_file(service, "src", "Title") == "copied"


def test_copy_file_missing_id_raises() -> None:
    service = Mock()
    service.files.return_value.copy.return_value.execute.return_value = {}
    with pytest.raises(GoogleAPIError, match="no file ID"):
        copy_file(service, "src", "Title")


def test_copy_file_wraps_http_error() -> None:
    response = Mock(status=403, reason="Forbidden")
    response.getheaders.return_value = []
    service = Mock()
    service.files.return_value.copy.return_value.execute.side_effect = HttpError(response, b"{}")
    with pytest.raises(GoogleAPIError, match="copy file"):
        copy_file(service, "src", "Title")


def test_set_anyone_permission() -> None:
    service = Mock()
    service.permissions.return_value.create.return_value.execute.return_value = {"id": "perm"}
    assert set_anyone_permission(service, "file", "reader") == "perm"


def test_upload_image_to_drive(tmp_path: Path, monkeypatch) -> None:
    image = tmp_path / "chart.unknown"
    image.write_bytes(b"png")
    monkeypatch.setattr("sliger.drive_utils.MediaFileUpload", lambda *a, **k: "media")
    service = Mock()
    service.files.return_value.create.return_value.execute.return_value = {"id": "img-id"}
    service.permissions.return_value.create.return_value.execute.return_value = {"id": "perm"}
    assert upload_image_to_drive(service, image) == "img-id"


def test_upload_image_missing_id_raises(tmp_path: Path, monkeypatch) -> None:
    image = tmp_path / "chart.jpg"
    image.write_bytes(b"jpg")
    monkeypatch.setattr("sliger.drive_utils.MediaFileUpload", lambda *a, **k: "media")
    service = Mock()
    service.files.return_value.create.return_value.execute.return_value = {}
    with pytest.raises(GoogleAPIError, match="no file ID"):
        upload_image_to_drive(service, image)
