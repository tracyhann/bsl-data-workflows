import importlib
import unittest


class CleanRedcapInstrumentsPackageLayoutTest(unittest.TestCase):
    def test_workflow_wrapper_and_steps_are_importable_from_new_package(self):
        wrapper = importlib.import_module("scripts.workflows.clean_redcap_instruments.run")
        self.assertTrue(hasattr(wrapper, "run_study_workflow"))

        for module_name in [
            "create_instrument_excels",
            "create_redcap_index_dictionary",
            "discover_and_standardize",
            "drop_stale_instruments",
            "exclude_column",
            "exclude_instruments",
            "final_verify",
            "match_text_labels",
            "organize_instruments",
        ]:
            with self.subTest(module_name=module_name):
                module = importlib.import_module(
                    f"scripts.workflows.clean_redcap_instruments.steps.{module_name}"
                )
                self.assertIsNotNone(module)


if __name__ == "__main__":
    unittest.main()
