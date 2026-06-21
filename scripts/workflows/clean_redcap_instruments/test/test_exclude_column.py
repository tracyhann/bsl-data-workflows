import unittest

from scripts.workflows.clean_redcap_instruments.steps.exclude_column import DEFAULT_EXCLUDE_KEYWORDS, flag_sensitive_columns


class ExcludeColumnTests(unittest.TestCase):
    def test_flags_sensitive_column_names_and_labels(self):
        flagged = flag_sensitive_columns(
            raw_headers=[
                "record_id",
                "patient_email",
                "home_address",
                "emergency_contact",
                "patient_mrn",
                "phone_number",
                "madrs_total",
                "other_field",
            ],
            label_headers=[
                "Record ID",
                "Patient Email",
                "Home Address",
                "Emergency Contact",
                "MRN",
                "Phone Number",
                "MADRS Total",
                "Participant alternate phone",
            ],
        )

        self.assertEqual(DEFAULT_EXCLUDE_KEYWORDS, ("email", "address", "contact", "mrn", "phone"))
        self.assertEqual(set(flagged), {1, 2, 3, 4, 5, 7})
        self.assertEqual(flagged[1].keyword, "email")
        self.assertEqual(flagged[2].keyword, "address")
        self.assertEqual(flagged[3].keyword, "contact")
        self.assertEqual(flagged[4].keyword, "mrn")
        self.assertEqual(flagged[5].keyword, "phone")
        self.assertEqual(flagged[7].keyword, "phone")

    def test_matches_hyphenated_email(self):
        flagged = flag_sensitive_columns(["participant_e_mail"], ["Participant E-mail"])

        self.assertEqual(set(flagged), {0})
        self.assertEqual(flagged[0].keyword, "email")


if __name__ == "__main__":
    unittest.main()
