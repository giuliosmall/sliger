"""Load Google API credentials from a service account or user OAuth."""

from __future__ import annotations

import json
import os
from contextlib import suppress
from pathlib import Path
from typing import Any

import google.auth
from google.auth.exceptions import DefaultCredentialsError, RefreshError
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

from sliger.exceptions import CredentialsError

PRESENTATIONS_SCOPE = "https://www.googleapis.com/auth/presentations"
DRIVE_FILE_SCOPE = "https://www.googleapis.com/auth/drive.file"
DRIVE_SCOPE = "https://www.googleapis.com/auth/drive"


def default_token_path() -> Path:
    """~/.config/sliger/token.json"""
    return Path.home() / ".config" / "sliger" / "token.json"


def scopes(*, service_account: bool, full_drive: bool) -> tuple[str, ...]:
    """Service account always gets presentations+drive (files shared with it).

    User OAuth: presentations+drive.file unless full_drive, then presentations+drive.
    """
    if service_account or full_drive:
        return (PRESENTATIONS_SCOPE, DRIVE_SCOPE)
    return (PRESENTATIONS_SCOPE, DRIVE_FILE_SCOPE)


def credential_kind(path: Path) -> str:
    """Classify a credentials JSON file.

    Returns ``service_account``, ``client_secrets``, ``authorized_user``, or ``unknown``.
    """
    try:
        data = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError):
        return "unknown"
    if not isinstance(data, dict):
        return "unknown"
    if data.get("type") == "service_account":
        return "service_account"
    if "installed" in data or "web" in data:
        return "client_secrets"
    if data.get("type") == "authorized_user":
        return "authorized_user"
    if {"refresh_token", "client_id", "client_secret"} <= data.keys():
        return "authorized_user"
    return "unknown"


def load_credentials(
    *,
    creds_file: Path | None = None,
    client_secrets: Path | None = None,
    token_file: Path | None = None,
    full_drive: bool = False,
    use_adc: bool = False,
) -> Any:
    user_scopes = scopes(service_account=False, full_drive=full_drive)
    if use_adc:
        return _load_adc(user_scopes)

    secrets_path = Path(client_secrets) if client_secrets is not None else None

    if creds_file is not None:
        creds_path = Path(creds_file)
        if not creds_path.is_file():
            raise CredentialsError(
                f"Credentials file not found: {creds_path}. "
                "Use the real path to a token or service-account JSON "
                "(for this repo: .secrets/user-token.json)."
            )
        kind = credential_kind(creds_path)
        if kind == "service_account":
            return _load_service_account(creds_path)
        if kind == "client_secrets":
            secrets_path = creds_path
        elif kind == "authorized_user":
            token_file = creds_path
        else:
            raise CredentialsError(
                f"Unrecognized credentials file {creds_path}. Expected a service-account "
                "JSON, a Desktop OAuth client secrets file, or an authorized-user token."
            )

    token_path = Path(token_file) if token_file is not None else default_token_path()

    creds = _load_stored_token(token_path, user_scopes)
    if creds is not None and creds.valid:
        return creds

    if secrets_path is None:
        return _load_adc(user_scopes)

    return _run_oauth_flow(secrets_path, token_path, user_scopes)


def _load_adc(user_scopes: tuple[str, ...]) -> Any:
    try:
        creds, _project = google.auth.default(scopes=list(user_scopes))
    except DefaultCredentialsError as exc:
        raise CredentialsError(
            "No credentials found. For a gcloud user login (ADC), run:\n"
            "  gcloud auth application-default login --scopes="
            "https://www.googleapis.com/auth/presentations,"
            "https://www.googleapis.com/auth/drive\n"
            "  gcloud auth application-default set-quota-project YOUR_PROJECT_ID\n"
            "Or pass --creds-file / --client-secrets. "
            f"Original error: {exc}"
        ) from exc
    return _with_quota_project(creds)


def _with_quota_project(creds: Any) -> Any:
    quota = os.environ.get("GOOGLE_CLOUD_QUOTA_PROJECT") or os.environ.get("GOOGLE_CLOUD_PROJECT")
    if quota and hasattr(creds, "with_quota_project"):
        return creds.with_quota_project(quota)
    return creds


def _load_service_account(path: Path) -> Any:
    try:
        return _with_quota_project(
            service_account.Credentials.from_service_account_file(
                str(path),
                scopes=scopes(service_account=True, full_drive=False),
            )
        )
    except OSError as exc:
        raise CredentialsError(f"Could not read credentials file {path}: {exc}") from exc
    except ValueError as exc:
        raise CredentialsError(f"Invalid credentials file {path}: {exc}") from exc


def _scopes_satisfied(granted: Any, needed: tuple[str, ...]) -> bool:
    granted_set = set(granted or ())
    needed_set = set(needed)
    if DRIVE_SCOPE in granted_set:
        needed_set.discard(DRIVE_FILE_SCOPE)
    return needed_set <= granted_set


def _load_stored_token(token_path: Path, user_scopes: tuple[str, ...]) -> Any | None:
    if not token_path.is_file():
        return None
    try:
        creds = Credentials.from_authorized_user_file(str(token_path), scopes=list(user_scopes))
    except (OSError, ValueError):
        return None
    if not _scopes_satisfied(creds.scopes, user_scopes):
        return None
    if creds.valid:
        return creds
    if creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except RefreshError:
            return None
        _save_token(token_path, creds)
        return creds
    return None


def _require_installed_client(client_secrets: Path) -> None:
    try:
        data = json.loads(client_secrets.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise CredentialsError(
            f"Could not read OAuth client secrets {client_secrets}: {exc}"
        ) from exc
    if not isinstance(data, dict):
        raise CredentialsError(f"OAuth client secrets {client_secrets} is not a JSON object")
    if "web" in data and "installed" not in data:
        raise CredentialsError(
            f"{client_secrets} is a Web application OAuth client. sliger needs a "
            "Desktop app client (JSON with an 'installed' key) so Google will accept "
            "a random http://localhost redirect. In GCP: APIs & Services → Credentials "
            "→ Create credentials → OAuth client ID → Desktop app."
        )
    if "installed" not in data:
        raise CredentialsError(
            f"{client_secrets} is not a Desktop OAuth client secrets file "
            "(expected a top-level 'installed' object)."
        )


def _run_oauth_flow(client_secrets: Path, token_path: Path, user_scopes: tuple[str, ...]) -> Any:
    _require_installed_client(client_secrets)
    try:
        flow = InstalledAppFlow.from_client_secrets_file(
            str(client_secrets), scopes=list(user_scopes)
        )
        # access_type=offline + prompt=consent is what actually returns a refresh_token
        # on installed apps; without it the next run cannot refresh after ~1 hour.
        creds = flow.run_local_server(
            port=0,
            access_type="offline",
            prompt="consent",
        )
    except CredentialsError:
        raise
    except Exception as exc:
        raise CredentialsError(f"OAuth sign-in failed: {exc}") from exc
    if not creds.refresh_token:
        raise CredentialsError(
            "Google did not return a refresh token. Delete the saved token and retry, "
            "and confirm the OAuth client is a Desktop app with offline access."
        )
    _save_token(token_path, creds)
    return creds


def _save_token(token_path: Path, creds: Any) -> None:
    token_path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.loads(creds.to_json())
    payload["type"] = "authorized_user"
    token_path.write_text(json.dumps(payload), encoding="utf-8")
    with suppress(OSError):
        token_path.chmod(0o600)
