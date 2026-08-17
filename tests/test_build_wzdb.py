import unittest
from pathlib import Path

from scripts.build_wzdb import StringTable, record_value


class RecordValueTests(unittest.TestCase):
    def test_string_table_reserves_empty_string_at_zero(self) -> None:
        strings = StringTable()

        self.assertEqual(strings.values[0], "")
        self.assertEqual(strings.intern(""), 0)

    def test_numeric_score_decodes_through_string_table(self) -> None:
        strings = StringTable()

        encoded_score = record_value(7, strings)

        self.assertIsInstance(encoded_score, int)
        self.assertEqual(strings.values[encoded_score], "7")

    def test_result_value_variants_are_preserved(self) -> None:
        strings = StringTable()
        values = [0, 1, 10, 14, 7.0, 3.5, "14+2", "3,2,1", "d", "u", "t"]

        decoded = [strings.values[record_value(value, strings)] for value in values]

        self.assertEqual(
            decoded,
            ["0", "1", "10", "14", "7", "3.5", "14+2", "3,2,1", "d", "u", "t"],
        )

    def test_empty_cells_are_null_references(self) -> None:
        strings = StringTable()

        self.assertIsNone(record_value(None, strings))
        self.assertIsNone(record_value("   ", strings))


class UpdateWorkflowTests(unittest.TestCase):
    def test_drive_download_bypasses_cache_and_checks_frequently(self) -> None:
        workflow = Path(".github/workflows/update-db.yml").read_text(encoding="utf-8")

        self.assertIn('cron: "*/15 * * * *"', workflow)
        self.assertIn("Cache-Control: no-cache", workflow)
        self.assertIn("cachebust=$REQUEST_NONCE", workflow)
        self.assertIn("Google Drive source SHA-256", workflow)


if __name__ == "__main__":
    unittest.main()
