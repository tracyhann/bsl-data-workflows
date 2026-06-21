import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from scripts.workflows.clean_redcap_instruments.steps.drop_stale_instruments import drop_stale_instruments


def write_workbook(path: Path, cleaned_rows: list[list[object]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "cleaned"
    for row in cleaned_rows:
        worksheet.append(row)
    workbook.save(path)


class DropStaleInstrumentsTests(unittest.TestCase):
    def test_deletes_instrument_workbooks_with_no_cleaned_data_rows_and_logs_latest_history(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            study_folder = Path(tmpdir) / "study"
            cleaned_dir = study_folder / "data" / "cleaned" / "redcap"
            older_history = study_folder / "histories" / "2026-06-17"
            latest_history = study_folder / "histories" / "2026-06-18"
            cleaned_dir.mkdir(parents=True)
            older_history.mkdir(parents=True)
            latest_history.mkdir(parents=True)

            stale_path = cleaned_dir / "63771-stale.xlsx"
            kept_path = cleaned_dir / "63771-kept.xlsx"
            dictionary_path = cleaned_dir / "dictionary.xlsx"
            write_workbook(stale_path, [["IRB", "subid", "arm", "visit", "date", "record_id", "redcap_event_name"]])
            write_workbook(
                kept_path,
                [
                    ["IRB", "subid", "arm", "visit", "date", "record_id", "redcap_event_name"],
                    ["63771", "s001", "1", "V1", "2026-01-01", "63771_s001", "baseline_arm_1"],
                ],
            )
            write_workbook(dictionary_path, [["not", "an", "instrument"]])
            (latest_history / "log.md").write_text("# REDCap Workflow Log\n", encoding="utf-8")

            result = drop_stale_instruments(study_folder)
            log_text = (latest_history / "log.md").read_text(encoding="utf-8")

            self.assertEqual(result.deleted_paths, [stale_path])
            self.assertEqual(result.kept_paths, [kept_path])
            self.assertFalse(stale_path.exists())
            self.assertTrue(kept_path.exists())
            self.assertTrue(dictionary_path.exists())
            self.assertIn("<summary><h1>Postprocess</h1></summary>", log_text)
            self.assertIn("## Drop Stale Instruments", log_text)
            self.assertIn("Deleted stale instrument workbooks: **1**", log_text)
            self.assertIn("63771-stale.xlsx", log_text)
            self.assertNotIn("63771-kept.xlsx", log_text)
            self.assertNotIn("Kept instrument workbooks", log_text)

    def test_reuses_existing_postprocess_section_for_stale_instrument_log(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            study_folder = Path(tmpdir) / "study"
            cleaned_dir = study_folder / "data" / "cleaned" / "redcap"
            history_dir = study_folder / "histories" / "2026-06-18"
            cleaned_dir.mkdir(parents=True)
            history_dir.mkdir(parents=True)

            stale_path = cleaned_dir / "63771-stale.xlsx"
            kept_path = cleaned_dir / "63771-kept.xlsx"
            write_workbook(stale_path, [["IRB", "subid", "arm", "visit", "date", "record_id", "redcap_event_name"]])
            write_workbook(
                kept_path,
                [
                    ["IRB", "subid", "arm", "visit", "date", "record_id", "redcap_event_name"],
                    ["63771", "s001", "1", "V1", "2026-01-01", "63771_s001", "baseline_arm_1"],
                ],
            )
            (history_dir / "log.md").write_text(
                "# REDCap Workflow Log\n\n"
                "<details>\n"
                "<summary><h1>Postprocess</h1></summary>\n\n"
                "## Existing Postprocess Step\n\n"
                "already here\n\n"
                "</details>\n",
                encoding="utf-8",
            )

            drop_stale_instruments(study_folder)
            log_text = (history_dir / "log.md").read_text(encoding="utf-8")

            self.assertEqual(log_text.count("<summary><h1>Postprocess</h1></summary>"), 1)
            self.assertIn("## Existing Postprocess Step", log_text)
            self.assertIn("## Drop Stale Instruments", log_text)
            self.assertIn("63771-stale.xlsx", log_text)
            self.assertNotIn("63771-kept.xlsx", log_text)


if __name__ == "__main__":
    unittest.main()
