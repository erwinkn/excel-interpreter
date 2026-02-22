import pytest
from excel_interpreter.errors import ExcelFunctionError
from excel_interpreter.functions import ExcelFunctions


class TestRegexExtract:
    def test_first_match(self):
        # Example inspired by docs: extract name based on capital letters
        assert ExcelFunctions.REGEXEXTRACT("DylanWilliams", "[A-Z][a-z]+") == "Dylan"

    def test_all_matches(self):
        text = (
            "Sonia Rees (378) 555-4195 Angel Brown (878) 555-8622"
            " Blake Martin (437) 555-8987"
        )
        pattern = "[0-9()]+ [0-9-]+"
        result = ExcelFunctions.REGEXEXTRACT(text, pattern, 1)
        assert result == ["(378) 555-4195", "(878) 555-8622", "(437) 555-8987"]

    def test_capturing_groups_first_match(self):
        text = "Report_2025_08_12"
        pattern = r"(.+)_(\d{4})_(\d{2})_(\d{2})"
        result = ExcelFunctions.REGEXEXTRACT(text, pattern, 2)
        assert result == ["Report", "2025", "08", "12"]

    def test_case_insensitive(self):
        assert ExcelFunctions.REGEXEXTRACT("abc ABC", "abc", 1, 1) == [
            "abc",
            "ABC",
        ]

    def test_no_match_returns_na(self):
        assert ExcelFunctions.REGEXEXTRACT("foo", "bar") == "#N/A"

    def test_bad_pattern_returns_value(self):
        # Unbalanced parenthesis
        assert ExcelFunctions.REGEXEXTRACT("abc", "(") == "#VALUE!"

    def test_mode_2_without_groups_returns_na(self):
        assert ExcelFunctions.REGEXEXTRACT("abc", "abc", 2) == "#N/A"

    def test_array_text_rejected(self):
        with pytest.raises(ExcelFunctionError):
            assert ExcelFunctions.REGEXEXTRACT(["abc"], "abc") == "#VALUE!"


class TestAggregateModeSingle:
    def test_mode_single_basic(self):
        assert ExcelFunctions.AGGREGATE(13, 0, [1, 2, 2, 3]) == 2.0

    def test_mode_single_tie_returns_smallest(self):
        # 1 appears twice, 2 appears twice -> smallest is 1
        assert ExcelFunctions.AGGREGATE(13, 0, [1, 1, 2, 2, 3]) == 1.0
