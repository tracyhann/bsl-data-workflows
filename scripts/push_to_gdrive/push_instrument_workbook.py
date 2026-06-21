#!/usr/bin/env python3
"""Push a local instrument workbook into an existing native Google Sheet.

This replaces the contents of the target Google Sheet with the supplied Excel
workbook while preserving the target Drive file ID and URL. It uses the Google
Drive files.update upload endpoint, which is the same operation used by Drive
when an Excel workbook is imported over an existing Google Sheet.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import uuid
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Protocol


SPREADSHEET_URL_RE = re.compile(
    r"https?://docs\.google\.com/spreadsheets/d/([A-Za-z0-9_-]+)"
)
DRIVE_UPDATE_FIELDS = "id,name,mimeType,webViewLink,modifiedTime"
XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@dataclass(frozen=True)
class DriveUploadResponse:
    status_code: int
    payload: dict[str, Any]
    raw_text: str


@dataclass(frozen=True)
class PushResult:
    spreadsheet_id: str
    workbook: Path
    dry_run: bool
    status_code: int | None = None
    payload: dict[str, Any] | None = None
    upload_url: str | None = None

    @property
    def web_url(self) -> str:
        if self.payload and self.payload.get("webViewLink"):
            return str(self.payload["webViewLink"])
        return f"https://docs.google.com/spreadsheets/d/{self.spreadsheet_id}/edit"


class HttpClient(Protocol):
    def patch(
        self,
        url: str,
        body: bytes,
        headers: Mapping[str, str],
        timeout: float,
    ) -> DriveUploadResponse:
        ...


class UrllibHttpClient:
    def patch(
        self,
        url: str,
        body: bytes,
        headers: Mapping[str, str],
        timeout: float,
    ) -> DriveUploadResponse:
        request = urllib.request.Request(
            url,
            data=body,
            headers=dict(headers),
            method="PATCH",
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                raw = response.read().decode("utf-8")
                payload = json.loads(raw) if raw.strip() else {}
                return DriveUploadResponse(response.status, payload, raw)
        except urllib.error.HTTPError as exc:
            raw = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(
                f"Google Drive upload failed with HTTP {exc.code}: {raw}"
            ) from exc


def resolve_spreadsheet_id(target: str | Path) -> str:
    """Resolve a Google Sheet URL, raw id, or text file containing either."""
    text = str(target).strip()
    path = Path(text).expanduser()
    if path.exists() and path.is_file():
        text = path.read_text(encoding="utf-8").strip()

    url_match = SPREADSHEET_URL_RE.search(text)
    if url_match:
        return url_match.group(1)

    candidate = text.strip()
    if re.fullmatch(r"[A-Za-z0-9_-]{20,}", candidate):
        return candidate

    raise ValueError(
        "Could not resolve a Google Sheet id. Pass a Google Sheets URL, "
        "a raw spreadsheet id, or a text file containing one."
    )


def validate_workbook_path(workbook: str | Path) -> Path:
    path = Path(workbook).expanduser().resolve()
    if not path.exists():
        raise FileNotFoundError(path)
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise ValueError(f"Expected an Excel workbook (.xlsx or .xlsm), got: {path}")
    return path


def build_drive_update_url(spreadsheet_id: str) -> str:
    return (
        f"https://www.googleapis.com/upload/drive/v3/files/{spreadsheet_id}"
        f"?uploadType=multipart&fields={DRIVE_UPDATE_FIELDS}"
    )


def build_multipart_body(
    *,
    workbook_path: Path,
    metadata: Mapping[str, Any] | None = None,
    boundary: str | None = None,
) -> tuple[bytes, str]:
    boundary = boundary or f"codex_{uuid.uuid4().hex}"
    metadata_json = json.dumps(dict(metadata or {}), separators=(",", ":")).encode("utf-8")
    workbook_bytes = workbook_path.read_bytes()

    chunks = [
        f"--{boundary}\r\n".encode("utf-8"),
        b"Content-Type: application/json; charset=UTF-8\r\n\r\n",
        metadata_json,
        b"\r\n",
        f"--{boundary}\r\n".encode("utf-8"),
        f"Content-Type: {XLSX_MIME_TYPE}\r\n\r\n".encode("utf-8"),
        workbook_bytes,
        b"\r\n",
        f"--{boundary}--\r\n".encode("utf-8"),
    ]
    return b"".join(chunks), boundary


def access_token_from_gcloud() -> str | None:
    commands = (
        ("gcloud", "auth", "print-access-token"),
        ("gcloud", "auth", "application-default", "print-access-token"),
    )
    for command in commands:
        try:
            completed = subprocess.run(
                command,
                check=True,
                capture_output=True,
                text=True,
            )
        except (FileNotFoundError, subprocess.CalledProcessError):
            continue
        token = completed.stdout.strip()
        if token:
            return token
    return None


def resolve_access_token(access_token: str | None) -> str:
    if access_token:
        return access_token.strip()

    env_token = os.environ.get("GOOGLE_OAUTH_ACCESS_TOKEN")
    if env_token:
        return env_token.strip()

    gcloud_token = access_token_from_gcloud()
    if gcloud_token:
        return gcloud_token

    raise RuntimeError(
        "No Google OAuth access token found. Pass --access-token, set "
        "GOOGLE_OAUTH_ACCESS_TOKEN, or install/login with gcloud."
    )


def push_instrument_workbook(
    *,
    target: str | Path,
    workbook: str | Path,
    access_token: str | None = None,
    http_client: HttpClient | None = None,
    timeout: float = 120.0,
    dry_run: bool = False,
    title: str | None = None,
) -> PushResult:
    spreadsheet_id = resolve_spreadsheet_id(target)
    workbook_path = validate_workbook_path(workbook)
    upload_url = build_drive_update_url(spreadsheet_id)

    if dry_run:
        return PushResult(
            spreadsheet_id=spreadsheet_id,
            workbook=workbook_path,
            dry_run=True,
            upload_url=upload_url,
        )

    token = resolve_access_token(access_token)
    metadata: dict[str, Any] = {}
    if title:
        metadata["name"] = title

    body, boundary = build_multipart_body(
        workbook_path=workbook_path,
        metadata=metadata,
    )
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": f"multipart/related; boundary={boundary}",
        "Content-Length": str(len(body)),
    }
    client = http_client or UrllibHttpClient()
    response = client.patch(upload_url, body, headers, timeout)
    return PushResult(
        spreadsheet_id=spreadsheet_id,
        workbook=workbook_path,
        dry_run=False,
        status_code=response.status_code,
        payload=response.payload,
        upload_url=upload_url,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--target",
        required=True,
        help="Target Google Sheet URL, raw spreadsheet id, or text file containing either.",
    )
    parser.add_argument(
        "--workbook",
        required=True,
        help="Local instrument workbook (.xlsx/.xlsm) to push into the target Google Sheet.",
    )
    parser.add_argument(
        "--access-token",
        default=None,
        help=(
            "Google OAuth access token with Drive write access. If omitted, the "
            "script tries GOOGLE_OAUTH_ACCESS_TOKEN, then gcloud."
        ),
    )
    parser.add_argument(
        "--title",
        default=None,
        help="Optional Drive file title to set while replacing the sheet contents.",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=120.0,
        help="Upload timeout in seconds.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Resolve and validate inputs without uploading anything.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    result = push_instrument_workbook(
        target=args.target,
        workbook=args.workbook,
        access_token=args.access_token,
        timeout=args.timeout,
        dry_run=args.dry_run,
        title=args.title,
    )
    print(
        json.dumps(
            {
                "dry_run": result.dry_run,
                "spreadsheet_id": result.spreadsheet_id,
                "workbook": str(result.workbook),
                "status_code": result.status_code,
                "url": result.web_url,
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
