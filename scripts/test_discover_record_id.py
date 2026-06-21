import csv
import tempfile
import unittest
from pathlib import Path

from scripts.discover_record_id import (
    canonicalize_record_id,
    classify_record_id,
    discover_candidate_columns,
    format_summary_rows,
    profile_csv,
    profile_csv_with_value_rows,
)


class DiscoverRecordIdTests(unittest.TestCase):
    def test_canonicalizes_separator_variants_as_same_family(self):
        variants = [
            "58807_s025",
            "58807s025",
            "58807 s025",
            "58807-s025",
            "58807.S025",
        ]

        results = [canonicalize_record_id(value) for value in variants]

        self.assertEqual({result.canonical for result in results}, {"58807_s025"})
        self.assertEqual({result.format_family for result in results}, {"numeric_prefix_subject_token"})
        self.assertEqual({result.subject_token for result in results}, {"s025"})

    def test_keeps_different_subject_token_prefixes_same_family_not_same_id(self):
        s_result = canonicalize_record_id("58807s025")
        b_result = canonicalize_record_id("58807b025")

        self.assertEqual(s_result.format_family, b_result.format_family)
        self.assertEqual(s_result.canonical, "58807_s025")
        self.assertEqual(b_result.canonical, "58807_b025")
        self.assertNotEqual(s_result.canonical, b_result.canonical)

    def test_extracts_suffixes_without_hardcoding_suffix_values(self):
        result = canonicalize_record_id("58807s025XOVER")

        self.assertEqual(result.canonical, "58807_s025_xover")
        self.assertEqual(result.subject_token, "s025")
        self.assertEqual(result.suffix_tokens, ("xover",))

    def test_coded_format_distinguishes_suffix_patterns(self):
        self.assertEqual(
            canonicalize_record_id("58807_s025").coded_format,
            "DIGITS+[SPACER]+SUBJECT_TOKEN",
        )
        self.assertEqual(
            canonicalize_record_id("58807s025XOVER").coded_format,
            "DIGITS+[SPACER]+SUBJECT_TOKEN+[SPACER]+SUFFIX",
        )
        self.assertEqual(
            canonicalize_record_id("58807_s025XOVER").coded_format,
            "DIGITS+[SPACER]+SUBJECT_TOKEN+[SPACER]+SUFFIX",
        )
        self.assertEqual(canonicalize_record_id("58807_s002a").canonical, "58807_s002_a")
        self.assertEqual(
            canonicalize_record_id("MDD_10598").coded_format,
            "WORD+[SPACER]+DIGITS",
        )
        self.assertEqual(
            canonicalize_record_id("MDD10598new").coded_format,
            "WORD+[SPACER]+DIGITS+[SPACER]+SUFFIX",
        )

    def test_discovers_duplicate_candidate_headers_by_physical_column(self):
        headers = ["Record ID", "Event Name", "Record ID", "Subject ID", "Survey Identifier"]

        candidates = discover_candidate_columns(headers)

        self.assertEqual([c.index for c in candidates], [0, 2, 3])
        self.assertEqual([c.header for c in candidates], ["Record ID", "Record ID", "Subject ID"])

    def test_header_discovery_is_token_based_and_tolerates_subject_typo(self):
        headers = [
            "Did Subject Continue in Study?",
            "Suicidal Ideation Score",
            "Valid Symptoms",
            "Subect ID",
            "Participant Study ID (or screening ID if not enrolled): ",
            "Participant Responding to task accurately: (choice=No)",
            "Subject No",
        ]

        candidates = discover_candidate_columns(headers)

        self.assertEqual([c.index for c in candidates], [3, 4, 6])

    def test_profiles_csv_candidate_columns_and_patterns(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / "sample.csv"
            with csv_path.open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Record ID", "Event Name", "Subject ID"])
                writer.writerow(["58807_s025", "Baseline", "58807s025"])
                writer.writerow(["58807s025XOVER", "Follow-up", "58807_s025_xover"])
                writer.writerow(["TEST", "Admin", ""])

            profile = profile_csv(csv_path)

        record_profile = next(col for col in profile.columns if col.header == "Record ID")
        self.assertEqual(record_profile.total_rows, 3)
        self.assertEqual(record_profile.unique_values, 3)
        self.assertEqual(record_profile.class_counts["numeric_prefix_subject_token"], 2)
        self.assertEqual(record_profile.class_counts["literal_word"], 1)

    def test_profile_defaults_to_index_zero_record_id_only(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / "sample.csv"
            with csv_path.open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Record ID", "Event Name", "Record ID", "Subject ID"])
                writer.writerow(["58807_s001", "Baseline", "MDD_100", "58807_s001"])
                writer.writerow(["58807_s002", "Follow-up", "MDD_101", "58807_s002"])

            profile, value_rows = profile_csv_with_value_rows(csv_path)

        self.assertEqual([column.column_index for column in profile.columns], [0])
        self.assertEqual({row["column_index"] for row in value_rows}, {0})

    def test_profile_accepts_record_id_snake_case_as_primary_column(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / "sample.csv"
            with csv_path.open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["record_id", "redcap_event_name"])
                writer.writerow(["62822_s001", "baseline"])
                writer.writerow(["53879_s002", "baseline"])

            profile, value_rows = profile_csv_with_value_rows(csv_path)

        self.assertEqual([column.column_index for column in profile.columns], [0])
        self.assertEqual(profile.columns[0].header, "record_id")
        self.assertEqual({row["coded_format"] for row in value_rows}, {"DIGITS+[SPACER]+SUBJECT_TOKEN"})

    def test_profile_can_still_scan_all_candidate_columns_when_requested(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / "sample.csv"
            with csv_path.open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Record ID", "Event Name", "Record ID", "Subject ID"])
                writer.writerow(["58807_s001", "Baseline", "MDD_100", "58807_s001"])

            profile, value_rows = profile_csv_with_value_rows(csv_path, all_candidate_columns=True)

        self.assertEqual([column.column_index for column in profile.columns], [0, 2, 3])
        self.assertEqual({row["column_index"] for row in value_rows}, {0, 2, 3})

    def test_builds_coded_format_summary_rows_with_seeded_examples(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            csv_path = Path(tmpdir) / "sample.csv"
            with csv_path.open("w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["Record ID", "Event Name"])
                writer.writerow(["58807_s001", "Baseline"])
                writer.writerow(["58807s002XOVER", "Follow-up"])
                writer.writerow(["MDD_100", "Baseline"])
                writer.writerow(["MDD101new", "Follow-up"])
                writer.writerow(["12345", "Admin"])
                writer.writerow(["TEST", "Admin"])

            _profile, value_rows = profile_csv_with_value_rows(csv_path)
            summary = format_summary_rows(csv_path, value_rows, random_seed=3, examples_per_format=5)

        by_format = {row["coded_format"]: row for row in summary}

        self.assertEqual(
            by_format["DIGITS+[SPACER]+SUBJECT_TOKEN"]["unique_id_values"],
            1,
        )
        self.assertEqual(
            by_format["DIGITS+[SPACER]+SUBJECT_TOKEN+[SPACER]+SUFFIX"]["examples"],
            "58807s002XOVER",
        )
        self.assertEqual(
            by_format["WORD+[SPACER]+DIGITS+[SPACER]+SUFFIX"]["unique_id_values"],
            1,
        )
        self.assertEqual(by_format["DIGITS"]["examples"], "12345")
        self.assertLessEqual(len(by_format["WORD"]["examples"].split("; ")), 5)


if __name__ == "__main__":
    unittest.main()
