import unittest
from datetime import date, datetime

from pm_plan_auto_schedule.generator import (
    ALL_BACKLINE_DE_DROSS_TEXT,
    DAY_START_COL,
    GenerationError,
    REMOVE_CHEMICAL_TEXT,
    extract_schedule_rules,
    first_occurrence,
    iter_occurrences,
    iter_weekdays,
    parse_template_month,
)


class FakeCell:
    def __init__(self, value=None, text=""):
        self.Value = value
        self.Text = text


class FakeWorksheet:
    def __init__(self, cell):
        self.cell = cell

    def Cells(self, row, col):
        return self.cell


class FakeGridWorksheet:
    def __init__(self, cells):
        self.cells = cells

    def Cells(self, row, col):
        return self.cells.get((row, col), FakeCell())


class GeneratorDateTests(unittest.TestCase):
    def test_iter_occurrences_handles_anchor_after_target_month(self):
        occurrences = list(iter_occurrences(date(2026, 3, 5), 28, 2026, 1))

        self.assertEqual(
            [item.isoformat() for item in occurrences],
            ["2026-01-08"],
        )

    def test_first_occurrence_limits_pm_plan_to_one_per_month(self):
        occurrence = first_occurrence(date(2026, 1, 1), 28, 2026, 1)

        self.assertEqual(occurrence, date(2026, 1, 1))

    def test_parse_template_month_from_date_value(self):
        worksheet = FakeWorksheet(FakeCell(value=datetime(2026, 3, 1)))

        self.assertEqual(parse_template_month(worksheet), 3)

    def test_parse_template_month_from_text(self):
        worksheet = FakeWorksheet(FakeCell(text="1-Feb-2026"))

        self.assertEqual(parse_template_month(worksheet), 2)

    def test_parse_template_month_rejects_unknown_format(self):
        worksheet = FakeWorksheet(FakeCell(text="not a date"))

        with self.assertRaises(GenerationError):
            parse_template_month(worksheet)

    def test_all_backline_without_template_schedule_gets_auto_de_dross(self):
        worksheet = FakeGridWorksheet(
            {
                (9, 2): FakeCell(text="1-Jan-2026"),
                (18, 2): FakeCell(text="ALL BACKLINE"),
            }
        )

        _template_month, rules, _default_source = extract_schedule_rules(worksheet)

        self.assertEqual(len(rules), 1)
        self.assertEqual(rules[0].row, 18)
        self.assertEqual(rules[0].de_dross_start_day, 1)
        self.assertEqual(rules[0].de_dross_source_col, DAY_START_COL)
        self.assertEqual(rules[0].de_dross_text, ALL_BACKLINE_DE_DROSS_TEXT)
        self.assertEqual(rules[0].de_dross_text, "DE-DROSS\n30 MIN")
        self.assertTrue(rules[0].auto_de_dross)

    def test_cleaning_pallet_room_gets_friday_chemical_rule(self):
        worksheet = FakeGridWorksheet(
            {
                (9, 2): FakeCell(text="1-Jan-2026"),
                (18, 2): FakeCell(text="CLEANING PALLET ROOM"),
            }
        )

        _template_month, rules, _default_source = extract_schedule_rules(worksheet)

        self.assertEqual(len(rules), 1)
        self.assertTrue(rules[0].auto_remove_unused_chemical)
        self.assertEqual(rules[0].chemical_source_col, DAY_START_COL)

    def test_iter_weekdays_finds_all_fridays(self):
        fridays = list(iter_weekdays(2026, 1, 4))

        self.assertEqual(
            [item.isoformat() for item in fridays],
            ["2026-01-02", "2026-01-09", "2026-01-16", "2026-01-23", "2026-01-30"],
        )
        self.assertEqual(REMOVE_CHEMICAL_TEXT, "REMOVE\nCHEMICAL")


if __name__ == "__main__":
    unittest.main()
