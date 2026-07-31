"""Tests for the Excel/CSV merge pipeline."""
import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from excel_merger import (  # noqa: E402
    COLUMN_ALIASES,
    clean_string,
    dedupe_columns,
    merge_files,
    normalize_column_name,
    normalize_currency,
    normalize_dataframe,
)


class TestNormalizeCurrency:
    @pytest.mark.parametrize("raw,expected", [
        ("$5,400.00", 5400.0),
        ("5400", 5400.0),
        ("  $1,299.50 ", 1299.5),
        ("-250.75", -250.75),
        (3100.0, 3100.0),
    ])
    def test_parses_currency_forms(self, raw, expected):
        assert normalize_currency(raw) == expected

    @pytest.mark.parametrize("raw", ["", "N/A", None, float("nan")])
    def test_unparseable_returns_none(self, raw):
        assert normalize_currency(raw) is None

    def test_rejects_series_with_clear_message(self):
        """Guards the duplicate-column failure mode explicitly."""
        with pytest.raises(TypeError, match="duplicate column names"):
            normalize_currency(pd.Series([1, 2]))


class TestNormalizeColumnName:
    @pytest.mark.parametrize("raw,expected", [
        ("Sales Amount", "revenue"),
        ("  TERRITORY  ", "region"),
        ("Transaction Date", "date"),
        ("Dept", "department"),
        ("Unmapped Column", "unmapped column"),
    ])
    def test_aliases_and_normalization(self, raw, expected):
        assert normalize_column_name(raw) == expected

    def test_alias_map_is_idempotent(self):
        """Canonical names must map to themselves, or merging is unstable."""
        for canonical in set(COLUMN_ALIASES.values()):
            assert normalize_column_name(canonical) == canonical


class TestDedupeColumns:
    def test_disambiguates_aliased_collisions(self):
        # "Amount" and "Total" both alias to "revenue"
        assert dedupe_columns(["name", "revenue", "revenue"], "f.csv") == [
            "name", "revenue", "revenue_2",
        ]

    def test_leaves_unique_columns_untouched(self):
        cols = ["name", "email", "revenue"]
        assert dedupe_columns(cols, "f.csv") == cols

    def test_handles_triple_collision(self):
        assert dedupe_columns(["r", "r", "r"], "f.csv") == ["r", "r_2", "r_3"]


class TestNormalizeDataframe:
    def test_collision_does_not_raise(self):
        """Regression: previously raised 'truth value of a Series is ambiguous'."""
        df = pd.DataFrame({"Name": ["a"], "Amount": ["$10"], "Total": ["$20"]})
        out = normalize_dataframe(df, "collide.csv")
        assert out["revenue"].iloc[0] == 10.0
        assert "revenue_2" in out.columns

    def test_strips_whitespace_and_tags_source(self):
        df = pd.DataFrame({"Name": ["  Sam Turner "], "Amount": ["$5,400.00"]})
        out = normalize_dataframe(df, "leads.csv")
        assert out["name"].iloc[0] == "Sam Turner"
        assert out["revenue"].iloc[0] == 5400.0
        assert out["source_file"].iloc[0] == "leads.csv"

    def test_drops_fully_empty_rows(self):
        df = pd.DataFrame({"Name": ["a", None], "Amount": ["$1", None]})
        assert len(normalize_dataframe(df, "f.csv")) == 1


class TestMergeFiles:
    def test_merges_dedups_and_writes_both_sheets(self, tmp_path):
        src = tmp_path / "in"
        src.mkdir()
        pd.DataFrame({"Name": ["a", "b"], "Sales Amount": ["$1", "$2"]}).to_csv(
            src / "one.csv", index=False)
        pd.DataFrame({"name": ["b", "c"], "amount": ["$2", "$3"]}).to_csv(
            src / "two.csv", index=False)

        out = tmp_path / "master.xlsx"
        merge_files(src, out, deduplicate=True)

        assert out.exists()
        sheets = pd.read_excel(out, sheet_name=None)
        assert set(sheets) == {"Merged Data", "Merge Summary"}
        merged = sheets["Merged Data"]
        assert len(merged) == 3          # b/$2 deduplicated across files
        assert list(merged.columns)[-1] == "source_file"

    def test_no_dedup_keeps_duplicates(self, tmp_path):
        src = tmp_path / "in"
        src.mkdir()
        for n in ("one.csv", "two.csv"):
            pd.DataFrame({"Name": ["a"], "Amount": ["$1"]}).to_csv(src / n, index=False)
        out = tmp_path / "m.xlsx"
        merge_files(src, out, deduplicate=False)
        assert len(pd.read_excel(out, sheet_name="Merged Data")) == 2

    def test_empty_folder_raises(self, tmp_path):
        empty = tmp_path / "empty"
        empty.mkdir()
        with pytest.raises(FileNotFoundError):
            merge_files(empty, tmp_path / "o.xlsx")


def test_clean_string_passes_through_non_strings():
    assert clean_string(42) == 42
    assert clean_string("  x ") == "x"
