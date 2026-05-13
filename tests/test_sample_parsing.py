"""
Tests for sample-name parsing helpers in app.py.

These functions are the heart of the QC pipeline — they decide whether
each row is a Standard, CVS, Blank, or Sample, and extract concentrations
and dilution factors from instrument-export naming conventions.
"""

from __future__ import annotations

import math

import pytest

from spincycle_utils import (
    _sanitize_sheet_name,
    classify_sample,
    clean_name,
    extract_cvs_conc,
    extract_cvs_conc_masshunter,
    extract_dilution,
    extract_std_conc,
    flag_lloq_uloq,
    strip_rep_suffix,
)


# ─────────────────────────────────────────────────────────────────
# classify_sample
# ─────────────────────────────────────────────────────────────────
class TestClassifySample:
    @pytest.mark.parametrize(
        "name",
        ["STD_5ppm", "std-10", "STD1", "Std_calibration_3", "std-mix_2ppm"],
    )
    def test_standards_classified_as_std(self, name: str):
        assert classify_sample(name) == "STD"

    @pytest.mark.parametrize(
        "name",
        ["CVS_5ppm", "cvs-low", "CVS1", "cvs_mid_check"],
    )
    def test_cvs_classified_as_cvs(self, name: str):
        assert classify_sample(name) == "CVS"

    @pytest.mark.parametrize("name", ["Blank", "blank_run", "BLANK_1"])
    def test_blanks_classified_as_blank(self, name: str):
        assert classify_sample(name) == "Blank"

    @pytest.mark.parametrize("name", ["Patient_001", "sample_A12", "extract_42", "QC_unknown"])
    def test_unknown_names_default_to_sample(self, name: str):
        assert classify_sample(name) == "Sample"


# ─────────────────────────────────────────────────────────────────
# extract_std_conc — pulls numeric STD concentration from name
# ─────────────────────────────────────────────────────────────────
class TestExtractStdConc:
    def test_simple_underscore_separated(self):
        # STD_PATTERN looks for digits after _ or -, not followed by xX\d.
        assert extract_std_conc("STD_5") == 5.0

    def test_dash_separated(self):
        assert extract_std_conc("STD-10") == 10.0

    def test_decimal_concentration(self):
        assert extract_std_conc("STD_2.5") == 2.5

    def test_returns_last_match_when_multiple(self):
        # If a name has multiple candidate numbers, the highest-level
        # concentration is the last one (often the actual conc).
        assert extract_std_conc("STD_run3_5") == 5.0

    def test_missing_concentration_returns_nan(self):
        result = extract_std_conc("STD")
        assert math.isnan(result)


# ─────────────────────────────────────────────────────────────────
# extract_cvs_conc — pulls numeric CVS concentration from name
# ─────────────────────────────────────────────────────────────────
class TestExtractCvsConc:
    def test_underscore_separator(self):
        assert extract_cvs_conc("CVS_5") == 5.0

    def test_dash_separator(self):
        assert extract_cvs_conc("CVS-2.5") == 2.5

    def test_no_separator(self):
        # Pattern allows cvs<digits> directly
        assert extract_cvs_conc("cvs5") == 5.0

    def test_case_insensitive(self):
        assert extract_cvs_conc("CVS_3.7") == 3.7
        assert extract_cvs_conc("cvs_3.7") == 3.7

    def test_missing_returns_nan(self):
        assert math.isnan(extract_cvs_conc("Patient_001"))


# ─────────────────────────────────────────────────────────────────
# extract_cvs_conc_masshunter — MassHunter-specific extractor
# ─────────────────────────────────────────────────────────────────
class TestExtractCvsConcMasshunter:
    def test_masshunter_pattern_returns_spike_not_index(self):
        # 'cvs1_5ppm' → spike concentration is 5, not 1
        assert extract_cvs_conc_masshunter("DKmix_cvs1_5ppm") == 5.0

    def test_falls_back_to_chemstation_pattern(self):
        # When no cvs<idx>_<conc> form, generic CVS_PATTERN kicks in
        assert extract_cvs_conc_masshunter("cvs_5ppm") == 5.0
        assert extract_cvs_conc_masshunter("cvs5") == 5.0

    def test_missing_returns_nan(self):
        assert math.isnan(extract_cvs_conc_masshunter("blank_run"))


# ─────────────────────────────────────────────────────────────────
# extract_dilution — pulls dilution factor like 10x, 2X
# ─────────────────────────────────────────────────────────────────
class TestExtractDilution:
    def test_lowercase_x_suffix(self):
        assert extract_dilution("sample_10x") == 10.0

    def test_uppercase_x_suffix(self):
        assert extract_dilution("sample_10X") == 10.0

    def test_decimal_dilution(self):
        assert extract_dilution("sample_2.5x") == 2.5

    def test_no_dilution_defaults_to_one(self):
        # Per the implementation, the default-no-dilution sentinel is 1.0
        assert extract_dilution("plain_sample") == 1.0


# ─────────────────────────────────────────────────────────────────
# clean_name — removes dilution tokens and collapses separators
# ─────────────────────────────────────────────────────────────────
class TestCleanName:
    def test_strips_dilution_token(self):
        assert clean_name("sample_10x_run1") == "sample_run1"

    def test_collapses_double_separators(self):
        # Removing the dilution token can leave "__" which gets collapsed
        assert "__" not in clean_name("sample_10x_extra")

    def test_strips_leading_trailing_separators(self):
        result = clean_name("_sample_")
        assert not result.startswith("_")
        assert not result.endswith("_")

    def test_no_dilution_returns_unchanged(self):
        assert clean_name("Patient_001") == "Patient_001"


# ─────────────────────────────────────────────────────────────────
# strip_rep_suffix — inverse of replicate-suffix tagging
# ─────────────────────────────────────────────────────────────────
class TestStripRepSuffix:
    def test_strips_single_digit_rep(self):
        assert strip_rep_suffix("sample_A_rep1") == "sample_A"

    def test_strips_multi_digit_rep(self):
        assert strip_rep_suffix("sample_A_rep12") == "sample_A"

    def test_leaves_non_rep_names_unchanged(self):
        assert strip_rep_suffix("sample_A") == "sample_A"

    def test_only_strips_trailing_rep_not_middle(self):
        assert strip_rep_suffix("rep1_sample") == "rep1_sample"


# ─────────────────────────────────────────────────────────────────
# flag_lloq_uloq — quantitation-limit flagging
# ─────────────────────────────────────────────────────────────────
class TestFlagLloqUloq:
    def test_below_lloq(self):
        assert flag_lloq_uloq(0.5, lloq=1.0, uloq=100.0) == "< LLOQ"

    def test_above_uloq(self):
        assert flag_lloq_uloq(150.0, lloq=1.0, uloq=100.0) == "> ULOQ"

    def test_within_range_returns_empty_or_passthrough(self):
        # Function returns "" or another in-range token; verify it's NOT a flag
        result = flag_lloq_uloq(50.0, lloq=1.0, uloq=100.0)
        assert result not in ("< LLOQ", "> ULOQ")


# ─────────────────────────────────────────────────────────────────
# _sanitize_sheet_name — Excel-safe sheet names
# ─────────────────────────────────────────────────────────────────
class TestSanitizeSheetName:
    def test_short_name_passes_through(self):
        used: set = set()
        assert _sanitize_sheet_name("MySheet", used) == "MySheet"

    def test_truncated_to_31_chars(self):
        used: set = set()
        long_name = "x" * 50
        assert len(_sanitize_sheet_name(long_name, used)) <= 31

    def test_invalid_chars_replaced(self):
        used: set = set()
        result = _sanitize_sheet_name("a/b\\c?d*e[f]g:h", used)
        for forbidden in r":\/?*[]":
            assert forbidden not in result

    def test_duplicate_gets_suffix(self):
        used: set = set()
        first = _sanitize_sheet_name("Data", used)
        second = _sanitize_sheet_name("Data", used)
        assert first == "Data"
        assert second != first
        assert "Data" in second

    def test_empty_name_gets_default(self):
        used: set = set()
        result = _sanitize_sheet_name("", used)
        assert result  # non-empty
