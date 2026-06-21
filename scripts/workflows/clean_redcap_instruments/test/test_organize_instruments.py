import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from scripts.workflows.clean_redcap_instruments.steps.organize_instruments import organize_instruments


def write_workbook(path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "cleaned"
    worksheet.append(["IRB", "subid"])
    worksheet.append(["62822", "s001"])
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def write_classification(path: Path, rows: list[tuple[str, str, str, float]]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "instrument_classification"
    worksheet.append(["instrument_name", "instrument_label", "class", "confidence"])
    for row in rows:
        worksheet.append(row)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


class OrganizeInstrumentsTests(unittest.TestCase):
    def test_moves_remaining_instruments_to_class_folders_and_logs_latest_history(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            study_folder = Path(tmpdir) / "study"
            cleaned_dir = study_folder / "data" / "cleaned"
            redcap_dir = cleaned_dir / "redcap"
            older_history = study_folder / "histories" / "2026-06-17"
            latest_history = study_folder / "histories" / "2026-06-18"
            older_history.mkdir(parents=True)
            latest_history.mkdir(parents=True)
            (latest_history / "log.md").write_text("# REDCap Workflow Log\n", encoding="utf-8")

            write_workbook(redcap_dir / "62822-madrs.xlsx")
            write_workbook(redcap_dir / "62822-daily_tms_checklist.xlsx")
            write_workbook(redcap_dir / "dictionary.xlsx")
            write_classification(
                latest_history / "instrument_classification.xlsx",
                [
                    ("madrs", "MADRS", "assessments", 0.91),
                    ("daily_tms_checklist", "Daily TMS Checklist", "treatments", 0.9),
                    ("mri_scan", "MRI Scan", "neuroimaging", 0.88),
                    ("survey_trigger", "Survey Trigger", "admin", 0.99),
                ],
            )

            result = organize_instruments(study_folder)
            log_text = (latest_history / "log.md").read_text(encoding="utf-8")

            self.assertEqual(set(result.created_class_dirs), {cleaned_dir / "assessments", cleaned_dir / "treatments", cleaned_dir / "neuroimaging"})
            self.assertFalse((redcap_dir / "62822-madrs.xlsx").exists())
            self.assertFalse((redcap_dir / "62822-daily_tms_checklist.xlsx").exists())
            self.assertTrue((cleaned_dir / "assessments" / "62822-madrs.xlsx").exists())
            self.assertTrue((cleaned_dir / "treatments" / "62822-daily_tms_checklist.xlsx").exists())
            self.assertFalse((cleaned_dir / "admin").exists())
            self.assertFalse((redcap_dir / "dictionary.xlsx").exists())
            self.assertTrue((cleaned_dir / "dictionary.xlsx").exists())
            self.assertFalse(redcap_dir.exists())
            self.assertTrue(result.redcap_dir_removed)
            self.assertIn("<summary><h1>Postprocess</h1></summary>", log_text)
            self.assertIn("## Sort Instuments", log_text)
            self.assertIn("Moved instrument workbooks: **2**", log_text)
            self.assertIn("62822-madrs.xlsx", log_text)
            self.assertIn("62822-daily_tms_checklist.xlsx", log_text)
            self.assertNotIn("62822-survey_trigger.xlsx |", log_text)

    def test_removes_redcap_dir_when_empty_after_sorting(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            study_folder = Path(tmpdir) / "study"
            cleaned_dir = study_folder / "data" / "cleaned"
            redcap_dir = cleaned_dir / "redcap"
            history_dir = study_folder / "histories" / "2026-06-18"

            write_workbook(redcap_dir / "62822-madrs.xlsx")
            write_classification(history_dir / "instrument_classification.xlsx", [("madrs", "MADRS", "assessments", 0.91)])

            result = organize_instruments(study_folder)

            self.assertTrue((cleaned_dir / "assessments" / "62822-madrs.xlsx").exists())
            self.assertFalse(redcap_dir.exists())
            self.assertTrue(result.redcap_dir_removed)

    def test_resorts_existing_category_workbooks_when_redcap_staging_dir_is_absent(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            study_folder = Path(tmpdir) / "study"
            cleaned_dir = study_folder / "data" / "cleaned"
            history_dir = study_folder / "histories" / "2026-06-18"
            history_dir.mkdir(parents=True)
            (history_dir / "log.md").write_text("# REDCap Workflow Log\n", encoding="utf-8")

            write_workbook(cleaned_dir / "dictionary.xlsx")
            write_workbook(cleaned_dir / "unknown" / "54909-psqi.xlsx")
            write_classification(history_dir / "instrument_classification.xlsx", [("psqi", "PSQI", "assessments", 0.91)])

            result = organize_instruments(study_folder)

            self.assertFalse((cleaned_dir / "unknown" / "54909-psqi.xlsx").exists())
            self.assertTrue((cleaned_dir / "assessments" / "54909-psqi.xlsx").exists())
            self.assertEqual(len(result.moved), 1)
            self.assertFalse(result.redcap_dir_removed)


if __name__ == "__main__":
    unittest.main()
