import csv
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook

from scripts.discover_events.dev.experiments.generate_data import generate_rows
from scripts.discover_events.event_discovery import (
    discover_event_groups,
    parse_event_value,
    resolve_event_column_index,
)
from scripts.discover_events.run import process_file


class EventDiscoveryTests(unittest.TestCase):
    def test_parses_arm_variants_to_same_arm_number(self):
        examples = [
            "baseline_arm2",
            "baseline arm 2",
            "baseline arm 2 a",
            "baseline arm 2a",
            "baseline_arm_2",
            "baseline_arm_2fxyz",
            "Baseline (Arm 2: Baseline/Active Study)",
            "arm2 baseline",
            "arm second baseline",
        ]

        for raw in examples:
            with self.subTest(raw=raw):
                self.assertEqual(parse_event_value(raw).arm, "2")

    def test_extracts_event_name_without_arm_suffix(self):
        examples = {
            "screening_visit_1_arm_1": ("1", "screening_visit_1"),
            "Screening (Visit 1) (Arm 1: Screening)": ("1", "Screening (Visit 1)"),
            "baseline arm 2a": ("2", "baseline"),
            "arm2 baseline": ("2", "baseline"),
        }

        for raw, expected in examples.items():
            with self.subTest(raw=raw):
                parsed = parse_event_value(raw)
                self.assertEqual((parsed.arm, parsed.event_name), expected)

    def test_groups_events_by_arm_preserving_first_appearance_order(self):
        raw_events = [
            "baseline_arm_2",
            "screening_arm_1",
            "study_day_1_arm_2",
            "baseline arm 2a",
            "follow_up_arm_1",
        ]

        groups = discover_event_groups(raw_events)

        arm_2 = [group.event_name for group in groups if group.arm == "2"]
        arm_1 = [group.event_name for group in groups if group.arm == "1"]
        baseline = next(group for group in groups if group.arm == "2" and group.event_name == "baseline")

        self.assertEqual(arm_2, ["baseline", "study_day_1"])
        self.assertEqual(arm_1, ["screening", "follow_up"])
        self.assertEqual(baseline.count, 2)

    def test_resolves_default_event_column_from_index_one_or_known_name(self):
        self.assertEqual(resolve_event_column_index(["record_id", "redcap_event_name"]), 1)
        self.assertEqual(resolve_event_column_index(["event_name", "unique_event_name"]), 1)
        self.assertEqual(resolve_event_column_index(["event_name", "event_id"]), 0)

    def test_runner_outputs_events_by_arm_summary(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = Path(tmpdir) / "events.csv"
            out_dir = Path(tmpdir) / "out"
            with input_path.open("w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow(["record_id", "redcap_event_name"])
                writer.writerow(["s001", "baseline_arm_2"])
                writer.writerow(["s001", "screening_arm_1"])
                writer.writerow(["s002", "baseline arm 2a"])

            outputs = process_file(input_path, out_dir)

            with outputs["events_by_arm"].open(newline="", encoding="utf-8") as file:
                rows = list(csv.DictReader(file))

        self.assertEqual(rows[0]["arm"], "1")
        self.assertEqual(rows[0]["event_name"], "screening")
        self.assertEqual(rows[1]["arm"], "2")
        self.assertEqual(rows[1]["event_name"], "baseline")
        self.assertEqual(rows[1]["count"], "2")

    def test_runner_optionally_verifies_against_codebook_unique_event_names(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            root = Path(tmpdir)
            input_path = root / "events.csv"
            out_dir = root / "out"
            codebook_path = root / "codebook.xlsx"
            with input_path.open("w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file)
                writer.writerow(["record_id", "redcap_event_name"])
                writer.writerow(["s001", "baseline_arm_1"])
                writer.writerow(["s001", "followup_arm_2"])

            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "58807-event"
            worksheet.append(["Event Name", "Unique event name", "Event ID"])
            worksheet.append(["Baseline", "baseline_arm_1", "1"])
            worksheet.append(["Followup A", "followup_arm_2a", "2"])
            worksheet.append(["Missing", "missing_arm_2", "3"])
            workbook.save(codebook_path)

            outputs = process_file(
                input_path,
                out_dir,
                codebook_path=codebook_path,
                codebook_sheet="58807-event",
            )

            with outputs["event_codebook_verification"].open(newline="", encoding="utf-8") as file:
                rows = list(csv.DictReader(file))

        self.assertIn("event_codebook_verification", outputs)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["status"], "missing_from_discovery")
        self.assertEqual(rows[0]["expected_event_name"], "missing")

    def test_experiment_generator_builds_real_and_synthetic_rows(self):
        rows = generate_rows(total_rows=200, random_seed=7, source_paths=[])

        self.assertEqual(len(rows), 200)
        self.assertEqual(set(rows[0]), {"raw", "arm", "event_name", "source"})
        self.assertTrue(any(row["raw"] == "synthetic_event_arm_2fxyz" for row in rows))
        self.assertTrue(any(row["arm"] == "2" for row in rows))


if __name__ == "__main__":
    unittest.main()
