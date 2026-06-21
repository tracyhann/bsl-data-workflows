import json
import tempfile
import unittest
from pathlib import Path
from urllib.parse import unquote

from openpyxl import Workbook

from scripts.push_to_gdrive.fill_in_overview import SheetsApiResponse
from scripts.push_to_gdrive.push_data_map import (
    data_map_tab_name,
    push_data_map,
)


class FakeSheetsHttpClient:
    def __init__(self, metadata, headers_by_sheet):
        self.metadata = metadata
        self.headers_by_sheet = headers_by_sheet
        self.get_calls = []
        self.post_calls = []

    def get(self, url, headers, timeout):
        self.get_calls.append({"url": url, "headers": headers, "timeout": timeout})
        if "values/" in url:
            encoded_range = url.split("/values/", 1)[1].split("?", 1)[0]
            sheet_name = unquote(encoded_range).split("!", 1)[0].strip("'")
            return SheetsApiResponse(
                status_code=200,
                payload={"values": [self.headers_by_sheet.get(sheet_name, [])]},
                raw_text="{}",
            )
        return SheetsApiResponse(
            status_code=200,
            payload=self.metadata,
            raw_text=json.dumps(self.metadata),
        )

    def post(self, url, body, headers, timeout):
        self.post_calls.append(
            {
                "url": url,
                "payload": json.loads(body.decode("utf-8")),
                "headers": headers,
                "timeout": timeout,
            }
        )
        return SheetsApiResponse(status_code=200, payload={}, raw_text="{}")


def make_data_map(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "data_map"
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)
    workbook.save(path)


class PushDataMapTests(unittest.TestCase):
    def test_data_map_tab_name_strips_suffix(self):
        self.assertEqual(
            data_map_tab_name(Path("biologics_biometrics-data-map.xlsx")),
            "biologics_biometrics",
        )
        self.assertEqual(data_map_tab_name(Path("platforms-data-map.xlsx")), "platforms")

    def test_pushes_each_data_map_workbook_to_matching_google_tab(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            data_map_dir = Path(tmpdir)
            make_data_map(
                data_map_dir / "platforms-data-map.xlsx",
                ["stage", "privacy", "description", "location"],
                [["raw", "phi", "REDCap", "./data/raw_exports/redcap/all"]],
            )
            make_data_map(
                data_map_dir / "assessments-data-map.xlsx",
                ["stage", "description", "location"],
                [
                    ["raw", "", ""],
                    ["cleaned/processed", "MADRS", "./data/cleaned/assessments/madrs.xlsx"],
                ],
            )
            fake_client = FakeSheetsHttpClient(
                metadata={
                    "sheets": [
                        {
                            "properties": {
                                "sheetId": 111,
                                "title": "Platforms",
                                "gridProperties": {"rowCount": 1, "columnCount": 4},
                            }
                        },
                        {
                            "properties": {
                                "sheetId": 222,
                                "title": "Assessments",
                                "gridProperties": {"rowCount": 1, "columnCount": 3},
                            }
                        },
                    ]
                },
                headers_by_sheet={
                    "Platforms": ["stage", "privacy", "description", "location"],
                    "Assessments": ["location", "stage", "description"],
                },
            )

            result = push_data_map(
                target="https://docs.google.com/spreadsheets/d/1MCjkVtR1lOJol8f95sK_W7bYdm8Bz1BkZbwJy410sYQ/edit",
                data_map_dir=data_map_dir,
                access_token="token",
                http_client=fake_client,
            )

            self.assertEqual(result.updated_tabs, ["Assessments", "Platforms"])
            self.assertEqual(result.skipped_tabs, [])
            self.assertEqual(result.updated_cell_count, 10)
            update_calls = [
                call for call in fake_client.post_calls if call["url"].endswith("/values:batchUpdate")
            ]
            self.assertEqual(len(update_calls), 1)
            ranges = {
                update["range"]: update["values"]
                for update in update_calls[0]["payload"]["data"]
            }
            self.assertEqual(
                ranges["'Assessments'!B2:B3"],
                [["raw"], ["cleaned/processed"]],
            )
            self.assertEqual(
                ranges["'Assessments'!C2:C3"],
                [[""], ["MADRS"]],
            )
            self.assertEqual(
                ranges["'Assessments'!A2:A3"],
                [[""], ["./data/cleaned/assessments/madrs.xlsx"]],
            )
            self.assertEqual(
                ranges["'Platforms'!A2:A2"],
                [["raw"]],
            )

    def test_dry_run_accepts_explicit_data_map_files_without_writes(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            data_map_file = Path(tmpdir) / "subjects-data-map.xlsx"
            make_data_map(
                data_map_file,
                ["stage", "description", "location"],
                [["cleaned/processed", "subject_timepoints", "./data/cleaned/subjects/subject_timepoints.xlsx"]],
            )
            fake_client = FakeSheetsHttpClient(
                metadata={
                    "sheets": [
                        {
                            "properties": {
                                "sheetId": 333,
                                "title": "Subjects",
                                "gridProperties": {"rowCount": 10, "columnCount": 3},
                            }
                        }
                    ]
                },
                headers_by_sheet={"Subjects": ["stage", "description", "location"]},
            )

            result = push_data_map(
                target="1MCjkVtR1lOJol8f95sK_W7bYdm8Bz1BkZbwJy410sYQ",
                data_map_files=[data_map_file],
                access_token="token",
                http_client=fake_client,
                dry_run=True,
            )

            self.assertTrue(result.dry_run)
            self.assertEqual(result.updated_tabs, ["Subjects"])
            self.assertEqual(result.updated_cell_count, 3)
            self.assertEqual(fake_client.post_calls, [])


if __name__ == "__main__":
    unittest.main()
