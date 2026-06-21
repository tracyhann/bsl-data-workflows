import csv
import tempfile
import unittest
from pathlib import Path

from scripts.standardize_date.date_standardizer import find_dates, standardize_date
from scripts.standardize_date.dev.experiments.generate_data import generate_rows
from scripts.standardize_date.run import process_file


class DateStandardizationTests(unittest.TestCase):
    def test_standardizes_common_date_variants_to_iso_yyyy_mm_dd(self):
        examples = {
            "2026-06-15": "2026-06-15",
            "2026/06/15": "2026-06-15",
            "06/15/2026": "2026-06-15",
            "6/15/26": "2026-06-15",
            "Jun 15, 2026": "2026-06-15",
            "15-Jun-2026": "2026-06-15",
            "2026-06-15 14:30": "2026-06-15",
        }

        for raw, expected in examples.items():
            with self.subTest(raw=raw):
                self.assertEqual(standardize_date(raw), expected)

    def test_rejects_non_dates_and_risky_partials(self):
        for raw in ["", "NA", "MDD_10598", "58807_s025", "March 2026", "6/7", "2026", "02/30/2026"]:
            with self.subTest(raw=raw):
                self.assertIsNone(standardize_date(raw))

    def test_finds_date_match_span_and_status(self):
        matches = find_dates("Assessment completed on 06/15/2026 by coordinator.")

        self.assertEqual(matches[0].text, "06/15/2026")
        self.assertEqual(matches[0].canonicalized, "2026-06-15")
        self.assertEqual(matches[0].parse_status, "parsed")

    def test_runner_processes_named_csv_column_and_adds_date_columns(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "input.csv"
            output_path = Path(tmpdir) / "output.csv"
            with input_path.open("w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow(["Record ID", "visit_date"])
                writer.writerow(["58807_s025", "06/15/2026"])
                writer.writerow(["58807_s026", "March 2026"])

            process_file(input_path, output_path, column_name="visit_date", column_index=0)

            with output_path.open(newline="", encoding="utf-8") as file:
                rows = list(csv.DictReader(file))

        self.assertEqual(rows[0]["date_detected"], "True")
        self.assertEqual(rows[0]["date_standardized"], "2026-06-15")
        self.assertEqual(rows[1]["date_detected"], "False")
        self.assertEqual(rows[1]["date_standardized"], "NA")

    def test_runner_reads_csv_with_quoted_comma_dates(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "input.csv"
            output_path = Path(tmpdir) / "output.csv"
            with input_path.open("w", newline="", encoding="utf-8") as file:
                writer = csv.DictWriter(file, fieldnames=["raw", "canonicalized"])
                writer.writeheader()
                writer.writerows(generate_rows(total_rows=200, random_seed=42))

            process_file(input_path, output_path, column_name="raw", column_index=0)

            with output_path.open(newline="", encoding="utf-8") as file:
                rows = list(csv.DictReader(file))

        mismatches = [
            row
            for row in rows
            if (row["date_standardized"] if row["date_detected"] == "True" else "NA") != row["canonicalized"]
        ]
        self.assertEqual(mismatches, [])

    def test_experiment_generator_builds_raw_and_canonicalized_rows(self):
        rows = generate_rows(total_rows=200, random_seed=7)

        self.assertEqual(len(rows), 200)
        self.assertEqual(set(rows[0]), {"raw", "canonicalized"})
        self.assertTrue(any(row["canonicalized"] == "NA" for row in rows))
        self.assertTrue(any(row["canonicalized"] == "2026-06-15" for row in rows))


if __name__ == "__main__":
    unittest.main()
