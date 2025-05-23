import sys
from typing import Any

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload


def get_drive_service(creds: Any):
    """Builds and returns the Google Drive API service."""
    return build("drive", "v3", credentials=creds)


def copy_file(drive_service: Any, file_id: str, copy_title: str) -> str | None:
    """Copies a file on Google Drive and returns the new file ID."""
    try:
        body = {"name": copy_title}
        drive_response = (
            drive_service.files().copy(fileId=file_id, body=body).execute()
        )
        return drive_response.get("id")
    except HttpError as err:
        print(err, file=sys.stderr)
        return None


def set_file_permissions_anyone_writer(drive_service: Any, file_id: str, verbose: bool) -> str | None:
    """Sets file permissions to 'anyone with link can write'."""
    try:
        user_permission = {
            "type": "anyone",
            "role": "writer",
        }
        permission_response = drive_service.permissions().create(
            fileId=file_id,
            body=user_permission,
            fields="id",
        ).execute()
        permission_id = permission_response.get("id")
        if verbose:
            print(f"Permission Id: {permission_id}")
        return permission_id
    except HttpError as err:
        print(err, file=sys.stderr)
        return None


def upload_image_to_drive(drive_service: Any, image_path: str, verbose: bool) -> str | None:
    """Uploads an image to Google Drive and makes it publicly readable.

    Args:
        drive_service: Authorized Google Drive API service instance.
        image_path: Local path to the image file.
        verbose: Boolean for verbose output.

    Returns:
        The ID of the uploaded image file on Google Drive, or None on error.
    """
    try:
        # Upload the image to Google Drive
        # Use image_path for the name on Drive for easier identification
        file_metadata = {"name": image_path, "parents": ["root"]}
        media = MediaFileUpload(image_path, mimetype="image/jpeg", resumable=True)

        img_file = (
            drive_service.files()
            .create(body=file_metadata, media_body=media, fields="id")
            .execute()
        )
        img_file_id = img_file.get("id")

        if verbose:
            print(f"Uploaded image '{image_path}' to Drive with ID: {img_file_id}")

        # Make the image publicly accessible (readable by anyone)
        permission = {"type": "anyone", "role": "reader"}
        drive_service.permissions().create(fileId=img_file_id, body=permission).execute()

        if verbose:
            print(f"Made image ID {img_file_id} publicly readable.")

        return img_file_id
    except HttpError as error:
        print(f"An error occurred during image upload/permission setting: {error}", file=sys.stderr)
        return None 