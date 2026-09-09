from __future__ import annotations

import json
from pathlib import Path
from unittest.mock import Mock

import pytest

from sliger.auth import (
    DRIVE_FILE_SCOPE,
    DRIVE_SCOPE,
    PRESENTATIONS_SCOPE,
    credential_kind,
    default_token_path,
    load_credentials,
    scopes,
)
from sliger.exceptions import CredentialsError


def test_scopes_service_account() -> None:
    assert scopes(service_account=True, full_drive=False) == (PRESENTATIONS_SCOPE, DRIVE_SCOPE)
    assert scopes(service_account=True, full_drive=True) == (PRESENTATIONS_SCOPE, DRIVE_SCOPE)


def test_scopes_user_oauth() -> None:
    assert scopes(service_account=False, full_drive=False) == (
        PRESENTATIONS_SCOPE,
        DRIVE_FILE_SCOPE,
    )
    assert scopes(service_account=False, full_drive=True) == (PRESENTATIONS_SCOPE, DRIVE_SCOPE)


def test_default_token_path() -> None:
    assert default_token_path() == Path.home() / ".config" / "sliger" / "token.json"


def test_credential_kind(tmp_path: Path) -> None:
    sa = tmp_path / "sa.json"
    sa.write_text(json.dumps({"type": "service_account", "client_email": "a@b.c"}))
    installed = tmp_path / "installed.json"
    installed.write_text(json.dumps({"installed": {"client_id": "id"}}))
    web = tmp_path / "web.json"
    web.write_text(json.dumps({"web": {"client_id": "id"}}))
    user = tmp_path / "user.json"
    user.write_text(json.dumps({"type": "authorized_user", "refresh_token": "r"}))
    junk = tmp_path / "junk.json"
    junk.write_text(json.dumps({"foo": "bar"}))
    missing = tmp_path / "missing.json"

    assert credential_kind(sa) == "service_account"
    assert credential_kind(installed) == "client_secrets"
    assert credential_kind(web) == "client_secrets"
    assert credential_kind(user) == "authorized_user"
    assert credential_kind(junk) == "unknown"
    assert credential_kind(missing) == "unknown"


def test_credential_kind_token_without_type_field(tmp_path: Path) -> None:
    token = tmp_path / "token.json"
    token.write_text(
        json.dumps(
            {
                "refresh_token": "r",
                "client_id": "id",
                "client_secret": "s",
                "token": "t",
            }
        )
    )
    assert credential_kind(token) == "authorized_user"


def test_load_credentials_service_account(tmp_path: Path, monkeypatch) -> None:
    path = tmp_path / "sa.json"
    path.write_text(json.dumps({"type": "service_account"}))
    ignored = tmp_path / "client.json"
    ignored.write_text(json.dumps({"installed": {}}))

    def fake_from_file(filename: str, scopes=None):
        return {"filename": filename, "scopes": scopes}

    monkeypatch.setattr(
        "sliger.auth.service_account.Credentials.from_service_account_file",
        fake_from_file,
    )
    result = load_credentials(
        creds_file=path, client_secrets=ignored, token_file=tmp_path / "t.json"
    )
    assert result["filename"] == str(path)
    assert result["scopes"] == (PRESENTATIONS_SCOPE, DRIVE_SCOPE)


def test_load_credentials_missing_everything_raises(tmp_path: Path, monkeypatch) -> None:
    from google.auth.exceptions import DefaultCredentialsError

    def boom(**kwargs):
        raise DefaultCredentialsError("no adc")

    monkeypatch.setattr("sliger.auth.google.auth.default", boom)
    with pytest.raises(CredentialsError, match="No credentials"):
        load_credentials(token_file=tmp_path / "no-token.json")


def test_load_credentials_adc(monkeypatch) -> None:
    monkeypatch.setattr(
        "sliger.auth.google.auth.default",
        lambda scopes=None: ("adc-creds", "project"),
    )
    assert load_credentials(use_adc=True, full_drive=True) == "adc-creds"


def test_load_credentials_client_secrets_via_creds_file(tmp_path: Path, monkeypatch) -> None:
    secrets = tmp_path / "client.json"
    secrets.write_text(json.dumps({"installed": {"client_id": "id"}}))
    token = tmp_path / "nested" / "token.json"

    fake_creds = Mock(
        token="t",
        refresh_token="r",
        token_uri="https://oauth2.googleapis.com/token",
        client_id="id",
        client_secret="secret",
        scopes=[PRESENTATIONS_SCOPE, DRIVE_FILE_SCOPE],
    )
    fake_creds.to_json.return_value = json.dumps(
        {
            "token": "t",
            "refresh_token": "r",
            "token_uri": "https://oauth2.googleapis.com/token",
            "client_id": "id",
            "client_secret": "secret",
            "scopes": [PRESENTATIONS_SCOPE, DRIVE_FILE_SCOPE],
            "expiry": "2099-01-01T00:00:00Z",
        }
    )
    flow = Mock()
    flow.run_local_server.return_value = fake_creds
    monkeypatch.setattr(
        "sliger.auth.InstalledAppFlow.from_client_secrets_file",
        lambda *a, **k: flow,
    )

    result = load_credentials(creds_file=secrets, token_file=token)
    assert result is fake_creds
    flow.run_local_server.assert_called_once_with(port=0, access_type="offline", prompt="consent")
    saved = json.loads(token.read_text())
    assert saved["type"] == "authorized_user"
    assert saved["token"] == "t"
    assert saved["refresh_token"] == "r"
    assert saved["expiry"] == "2099-01-01T00:00:00Z"
    assert saved["client_id"] == "id"
    assert saved["client_secret"] == "secret"
    assert saved["scopes"] == [PRESENTATIONS_SCOPE, DRIVE_FILE_SCOPE]


def test_web_client_secrets_are_rejected(tmp_path: Path) -> None:
    secrets = tmp_path / "web.json"
    secrets.write_text(json.dumps({"web": {"client_id": "id", "client_secret": "s"}}))
    with pytest.raises(CredentialsError, match="Desktop app"):
        load_credentials(client_secrets=secrets, token_file=tmp_path / "t.json")


def test_oauth_without_refresh_token_raises(tmp_path: Path, monkeypatch) -> None:
    secrets = tmp_path / "client.json"
    secrets.write_text(json.dumps({"installed": {"client_id": "id"}}))
    fake_creds = Mock(refresh_token=None)
    flow = Mock()
    flow.run_local_server.return_value = fake_creds
    monkeypatch.setattr(
        "sliger.auth.InstalledAppFlow.from_client_secrets_file",
        lambda *a, **k: flow,
    )
    with pytest.raises(CredentialsError, match="refresh token"):
        load_credentials(client_secrets=secrets, token_file=tmp_path / "t.json")
