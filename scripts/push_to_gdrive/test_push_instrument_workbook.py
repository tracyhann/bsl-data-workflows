import json
import tempfile
import unittest
from pathlib import Path

from scripts.push_to_gdrive.push_instrument_workbook import (
    DriveUploadResponse,
    build_drive_update_url,
    push_instrument_workbook,
    resolve_spreadsheet_id,
)


class FakeHttpClient:
    def __init__(self):
        self.calls = []

    def patch(self, url, body, headers, timeout):
        self.calls.append(
            {
                "url": url,
                "body": body,
                "headers": headers,
                "timeout": timeout,
            }
        )
        return DriveUploadResponse(
            status_code=200,
            payload={
                "id": "1XKA1n3lEdvZs_PGCsNMO7U-QiMVi4-nxmxgo4AYWAg0",
                "mimeType": "application/vnd.google-apps.spreadsheet",
                "webViewLink": "https://docs.google.com/spreadsheets/d/1XKA1n3lEdvZs_PGCsNMO7U-QiMVi4-nxmxgo4AYWAg0/edit",
            },
            raw_text="{}",
        )


class PushInstrumentWorkbookTests(unittest.TestCase):
    def test_resolves_google_sheet_url_to_id(self):
        url = "https://docs.google.com/spreadsheets/d/1XKA1n3lEdvZs_PGCsNMO7U-QiMVi4-nxmxgo4AYWAg0/edit?usp=sharing"

        self.assertEqual(
            resolve_spreadsheet_id(url),
            "1XKA1n3lEdvZs_PGCsNMO7U-QiMVi4-nxmxgo4AYWAg0",
        )

    def test_resolves_google_sheet_url_stored_in_text_file(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            target_file = Path(tmpdir) / "sheet_link.txt"
            target_file.write_text(
                "https://docs.google.com/spreadsheets/d/1XKA1n3lEdvZs_PGCsNMO7U-QiMVi4-nxmxgo4AYWAg0/edit",
                encoding="utf-8",
            )

            self.assertEqual(
                resolve_spreadsheet_id(str(target_file)),
                "1XKA1n3lEdvZs_PGCsNMO7U-QiMVi4-nxmxgo4AYWAg0",
            )

    def test_dry_run_reports_target_and_workbook_without_uploading(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            workbook = Path(tmpdir) / "instrument.xlsx"
            workbook.write_bytes(b"fake xlsx bytes")
            fake_client = FakeHttpClient()

            result = push_instrument_workbook(
                target="1XKA1n3lEdvZs_PGCsNMO7U-QiMVi4-nxmxgo4AYWAg0",
                workbook=workbook,
                access_token="token",
                http_client=fake_client,
                dry_run=True,
            )

            self.assertTrue(result.dry_run)
            self.assertEqual(result.spreadsheet_id, "1XKA1n3lEdvZs_PGCsNMO7U-QiMVi4-nxmxgo4AYWAg0")
            self.assertEqual(result.workbook, workbook.resolve())
            self.assertEqual(fake_client.calls, [])

    def test_push_uploads_xlsx_bytes_to_drive_update_endpoint(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            workbook = Path(tmpdir) / "instrument.xlsx"
            workbook.write_bytes(b"fake xlsx bytes")
            fake_client = FakeHttpClient()

            result = push_instrument_workbook(
                target="https://docs.google.com/spreadsheets/d/1XKA1n3lEdvZs_PGCsNMO7U-QiMVi4-nxmxgo4AYWAg0/edit",
                workbook=workbook,
                access_token="token",
                http_client=fake_client,
                dry_run=False,
            )

            self.assertFalse(result.dry_run)
            self.assertEqual(result.status_code, 200)
            self.assertEqual(len(fake_client.calls), 1)
            call = fake_client.calls[0]
            self.assertEqual(
                call["url"],
                build_drive_update_url("1XKA1n3lEdvZs_PGCsNMO7U-QiMVi4-nxmxgo4AYWAg0"),
            )
            self.assertEqual(call["headers"]["Authorization"], "Bearer token")
            self.assertIn(b"fake xlsx bytes", call["body"])
            self.assertIn(b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", call["body"])

    def test_rejects_non_xlsx_workbooks(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            workbook = Path(tmpdir) / "instrument.csv"
            workbook.write_text("a,b\n1,2\n", encoding="utf-8")

            with self.assertRaisesRegex(ValueError, "Expected an Excel workbook"):
                push_instrument_workbook(
                    target="1XKA1n3lEdvZs_PGCsNMO7U-QiMVi4-nxmxgo4AYWAg0",
                    workbook=workbook,
                    access_token="token",
                    http_client=FakeHttpClient(),
                    dry_run=True,
                )


if __name__ == "__main__":
    unittest.main()
