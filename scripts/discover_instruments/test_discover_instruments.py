import csv
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from scripts.discover_instruments.instrument_discovery import (
    discover_instruments,
    is_complete_column,
)
from scripts.discover_instruments.run import process_file


class InstrumentDiscoveryTests(unittest.TestCase):
    def test_detects_complete_stop_signals_precisely(self):
        self.assertTrue(is_complete_column("madrs_complete"))
        self.assertTrue(is_complete_column("Complete?"))
        self.assertFalse(is_complete_column("completion_notes"))
        self.assertFalse(is_complete_column("madrs_complete_notes"))

    def test_discovers_blocks_from_complete_columns_and_skips_redcap_metadata(self):
        headers = [
            "record_id",
            "redcap_event_name",
            "redcap_survey_identifier",
            "madrs_recordid_traineerater",
            "madrs_date_traineerater",
            "madrs_total_traineerater",
            "madrsc_trainee_rater_complete",
            "ymrs_recordid",
            "ymrs_score",
            "ymrs_complete",
        ]

        blocks = discover_instruments(headers)

        self.assertEqual([block.instrument_key for block in blocks], ["madrsc_trainee_rater", "ymrs"])
        self.assertEqual(blocks[0].start_index, 3)
        self.assertEqual(blocks[0].end_index, 6)
        self.assertEqual(blocks[0].dominant_prefix, "madrs")
        self.assertIn("traineerater:3", blocks[0].suffix_evidence)
        self.assertEqual(blocks[1].dominant_prefix, "ymrs")

    def test_handles_label_exports_with_complete_question_mark(self):
        headers = ["Record ID", "Event Name", "Question 1", "Question 2", "Complete?"]

        blocks = discover_instruments(headers)

        self.assertEqual(len(blocks), 1)
        self.assertEqual(blocks[0].instrument_key, "complete_status")
        self.assertEqual(blocks[0].start_index, 2)
        self.assertEqual(blocks[0].end_index, 4)

    def test_runner_writes_instrument_summary_and_columns(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "input.csv"
            out_dir = Path(tmpdir) / "out"
            with input_path.open("w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow(
                    [
                        "record_id",
                        "redcap_event_name",
                        "madrs_recordid_mentorrater",
                        "madrs_total_mentorrater",
                        "madrsc_mentor_rater_complete",
                    ]
                )
                writer.writerow(["s001", "baseline_arm_1", "s001", "20", "Complete"])

            outputs = process_file(input_path, out_dir)

            with outputs["instrument_summary"].open(newline="", encoding="utf-8") as file:
                summary_rows = list(csv.DictReader(file))
            with outputs["instrument_columns"].open(newline="", encoding="utf-8") as file:
                column_rows = list(csv.DictReader(file))

        self.assertEqual(summary_rows[0]["instrument_key"], "madrsc_mentor_rater")
        self.assertEqual(summary_rows[0]["dominant_prefix"], "madrs")
        self.assertEqual(len(column_rows), 3)
        self.assertEqual(column_rows[-1]["is_stop_signal"], "True")

    def test_runner_optionally_verifies_against_codebook_form_names_with_fallback(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            input_path = root / "input.csv"
            out_dir = root / "out"
            codebook_path = root / "codebook.xlsx"

            with input_path.open("w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow(
                    [
                        "record_id",
                        "redcap_event_name",
                        "demo_date",
                        "demographics_complete",
                        "madrs_score",
                        "madrs_complete",
                    ]
                )
                writer.writerow(["s001", "baseline_arm_1", "2026-01-01", "2", "10", "2"])

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "58807-instrument"
            worksheet.append(["Instrument", "Form Name", "Events"])
            worksheet.append(["Demographics", "demographics", "baseline_arm_1"])
            worksheet.append(["madrs", "", "baseline_arm_1"])
            worksheet.append(["Missing Form", "missing_form", "baseline_arm_1"])
            workbook.save(codebook_path)

            outputs = process_file(
                input_path,
                out_dir,
                codebook_path=codebook_path,
                codebook_sheet="58807-instrument",
            )

            with outputs["instrument_codebook_verification"].open(newline="", encoding="utf-8") as file:
                rows = list(csv.DictReader(file))

        self.assertIn("instrument_codebook_verification", outputs)
        self.assertEqual([row["expected_form_name"] for row in rows], ["missing_form"])
        self.assertEqual(rows[0]["status"], "missing_from_discovery")


if __name__ == "__main__":
    unittest.main()
