import csv
import tempfile
import unittest
from pathlib import Path

from scripts.standardize_record_id.dev.experiments.generate_data import generate_rows
from scripts.standardize_record_id.presidio.redcap_subid_presidio import (
    REDCAP_SUBJECT_ID,
    find_redcap_subids,
    standardize_redcap_subid,
)
from scripts.standardize_record_id.run import process_file


class RedcapSubidStandardizationTests(unittest.TestCase):
    def test_standardizes_common_redcap_subid_variants_to_base_subject_id(self):
        examples = {
            "58807_s025": "58807_s025",
            "58807s25": "58807_s025",
            "58807-S025": "58807_s025",
            "IRB 58807 subject 25": "58807_s025",
            "54909_sub002a": "54909_s002",
            "58807_b025": "58807_s025",
            "54909_s002XOVER": "54909_s002",
        }

        for raw, expected in examples.items():
            with self.subTest(raw=raw):
                self.assertEqual(standardize_redcap_subid(raw), expected)

    def test_returns_none_for_non_redcap_subject_ids(self):
        for raw in ["MDD_10598", "General Note to File 1", "24", "2024-06-15", ""]:
            with self.subTest(raw=raw):
                self.assertIsNone(standardize_redcap_subid(raw))

    def test_ignores_long_cells_that_contain_url_or_hash_like_text(self):
        flywheel_url = (
            "https://cni.flywheel.io/#/projects/60a2971370df1eea9567f84f/"
            "sessions/62864b2d66c7c8a0677f57b8?tab=analyses"
        )

        self.assertEqual(standardize_redcap_subid("62864b2"), "62864_s002")
        self.assertIsNone(standardize_redcap_subid(flywheel_url))
        self.assertEqual(find_redcap_subids(flywheel_url), [])

    def test_finds_custom_presidio_entity_span(self):
        matches = find_redcap_subids("ID 58807-s25")

        self.assertEqual(matches[0].entity_type, REDCAP_SUBJECT_ID)
        self.assertEqual(matches[0].text, "58807-s25")
        self.assertEqual(matches[0].canonicalized, "58807_s025")

    def test_runner_processes_named_csv_column_and_adds_detection_columns(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "input.csv"
            output_path = Path(tmpdir) / "output.csv"
            with input_path.open("w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow(["Record ID", "Event Name"])
                writer.writerow(["58807s25", "Screening"])
                writer.writerow(["MDD_10598", "Screening"])

            process_file(input_path, output_path, column_name="Record ID", column_index=0)

            with output_path.open(newline="", encoding="utf-8") as file:
                rows = list(csv.DictReader(file))

        self.assertEqual(rows[0]["redcap_subid_detected"], "True")
        self.assertEqual(rows[0]["redcap_subid_canonicalized"], "58807_s025")
        self.assertEqual(rows[1]["redcap_subid_detected"], "False")
        self.assertEqual(rows[1]["redcap_subid_canonicalized"], "NA")

    def test_runner_writes_naming_pattern_summary_for_selected_column(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "input.csv"
            output_path = Path(tmpdir) / "output.csv"
            with input_path.open("w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow(["Record ID", "Event Name"])
                writer.writerow(["58807s25", "Screening"])
                writer.writerow(["58807_s026_xover", "Baseline"])
                writer.writerow(["MDD_10598", "Screening"])

            outputs = process_file(input_path, output_path, column_name="Record ID", column_index=0)

            with outputs["format_summary"].open(newline="", encoding="utf-8") as file:
                summary_rows = list(csv.DictReader(file))

        summary_by_format = {row["coded_format"]: row for row in summary_rows}
        self.assertIn("DIGITS+[SPACER]+SUBJECT_TOKEN", summary_by_format)
        self.assertIn("DIGITS+[SPACER]+SUBJECT_TOKEN+[SPACER]+SUFFIX", summary_by_format)
        self.assertIn("WORD+[SPACER]+DIGITS", summary_by_format)
        self.assertEqual(summary_by_format["DIGITS+[SPACER]+SUBJECT_TOKEN"]["row_occurrences"], "1")

    def test_experiment_generator_builds_raw_and_canonicalized_rows(self):
        rows = generate_rows(total_rows=100, random_seed=7)

        self.assertEqual(len(rows), 100)
        self.assertEqual(set(rows[0]), {"raw", "canonicalized"})
        self.assertTrue(any(row["canonicalized"] == "NA" for row in rows))
        self.assertTrue(any(row["canonicalized"].endswith("_s025") for row in rows))


if __name__ == "__main__":
    unittest.main()
