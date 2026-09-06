"""Thin wrappers around the Google Drive API."""

from __future__ import annotations

import logging
import mimetypes
from pathlib import Path
from typing import Any, Literal

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload

from sliger.exceptions import GoogleAPIError, ImageNotFoundError

logger = logging.getLogger(__name__)

PermissionRole = Literal["reader", "writer", "commenter"]


def get_drive_service(creds: Any) -> Any:
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def _execute(request: Any, action: str) -> Any:
    try:
        return request.execute()
    except HttpError as err:
        raise GoogleAPIError(f"{action}: {err}") from err


def copy_file(drive_service: Any, file_id: str, copy_title: str) -> str:
    """Copy a Drive file and return the new file ID."""
    response = _execute(
        drive_service.files().copy(fileId=file_id, body={"name": copy_title}),
        f"copy file {file_id}",
    )
    new_id = response.get("id")
    if not new_id:
        raise GoogleAPIError(f"Drive copy of {file_id} returned no file ID")
    logger.debug("Copied %s -> %s (%s)", file_id, new_id, copy_title)
    return new_id


def set_anyone_permission(drive_service: Any, file_id: str, role: PermissionRole) -> str | None:
    """Grant 'anyone with the link' the given role. Returns the permission id."""
    response = _execute(
        drive_service.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": role},
            fields="id",
        ),
        f"set anyone/{role} permission on {file_id}",
    )
    permission_id = response.get("id")
    logger.debug("Set anyone/%s on %s (permission %s)", role, file_id, permission_id)
    return permission_id


def upload_image_to_drive(drive_service: Any, image_path: str | Path) -> str:
    """Upload a local image, make it link-readable, and return its Drive file ID.

    Google Slides ``createImage`` fetches the file over HTTP, so the upload has
    to be readable by anyone with the link.
    """
    path = Path(image_path)
    if not path.is_file():
        raise ImageNotFoundError(f"Image path does not exist: {path}")

    mime_type, _ = mimetypes.guess_type(path.name)
    if not mime_type or not mime_type.startswith("image/"):
        mime_type = "image/png"

    media = MediaFileUpload(str(path), mimetype=mime_type, resumable=True)
    img_file = _execute(
        drive_service.files().create(
            body={"name": path.name},
            media_body=media,
            fields="id",
        ),
        f"upload image {path}",
    )
    img_file_id = img_file.get("id")
    if not img_file_id:
        raise GoogleAPIError(f"Drive upload of {path} returned no file ID")

    set_anyone_permission(drive_service, img_file_id, "reader")
    logger.debug("Uploaded %s as Drive file %s (%s)", path, img_file_id, mime_type)
    return img_file_id


def drive_image_url(file_id: str) -> str:
    return f"https://drive.google.com/uc?export=view&id={file_id}"
