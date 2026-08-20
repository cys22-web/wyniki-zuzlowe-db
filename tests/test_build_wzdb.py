import unittest
from pathlib import Path

from openpyxl import Workbook

from scripts.build_wzdb import StringTable, build_database, record_value


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
        self.assertIn("sha256sum data/event_dates_2026.json", workflow)
        self.assertIn("date_map_sha256", workflow)


class RecordLayoutTests(unittest.TestCase):
    def test_column_n_is_appended_without_moving_existing_fields(self) -> None:
        import tempfile

        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "source.xlsx"
            workbook = Workbook()
            players = workbook.active
            players.title = "Zawodnicy"
            players.append(["Zawodnik", "Narodowość", "Data urodzenia"])
            players.append(["Test Rider", "Polska", None])
            for year in range(2010, 2027):
                sheet = workbook.create_sheet(str(year))
                if year == 2026:
                    values = {
                        2: "Test Rider", 3: "11+2", 4: "3,3,2*,2,1*",
                        7: "Home", 8: "Away", 9: "49-41", 10: "Liga",
                        11: "Krosno", 12: "Rozgrywki", 13: "1 runda",
                        14: "95", 15: "500cc", 16: "Uwagi",
                    }
                    for column, value in values.items():
                        sheet.cell(4, column, value)
            workbook.save(source)
            database = build_database(source, "test", "2026-08-20T00:00:00Z")

        record = database["years"]["2026"][0]
        decoded = [None if value is None else database["strings"][value] for value in record[1:]]
        self.assertEqual(len(record), 15)
        self.assertEqual(decoded[0], "11+2")
        self.assertEqual(decoded[1], "3,3,2*,2,1*")
        self.assertEqual(decoded[11], "500cc")
        self.assertEqual(decoded[12], "Uwagi")
        self.assertEqual(decoded[13], "95")

    def test_missing_date_map_keeps_generation_compatible(self) -> None:
        self.assertEqual(record_value(None, StringTable()), None)


if __name__ == "__main__":
    unittest.main()
