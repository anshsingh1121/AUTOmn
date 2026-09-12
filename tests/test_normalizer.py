"""
Tests for the Normalizer module.
Tests priority parsing, whitespace normalization, and substring matching.
"""

import pytest
from src.normalizer import (
    normalize_whitespace,
    extract_priority_class,
    is_priority_3,
    normalize_for_matching,
    is_blank,
    contains_substring,
)


class TestNormalizeWhitespace:
    def test_normal_string(self):
        assert normalize_whitespace("hello") == "hello"

    def test_leading_trailing_spaces(self):
        assert normalize_whitespace("  hello  ") == "hello"

    def test_none(self):
        assert normalize_whitespace(None) is None

    def test_empty_string(self):
        assert normalize_whitespace("") is None

    def test_whitespace_only(self):
        assert normalize_whitespace("   ") is None

    def test_nan_string(self):
        assert normalize_whitespace("nan") is None

    def test_numeric(self):
        assert normalize_whitespace(42) == "42"


class TestExtractPriorityClass:
    def test_full_format(self):
        assert extract_priority_class("3 - Moderate") == 3

    def test_critical(self):
        assert extract_priority_class("1 - Critical") == 1

    def test_high(self):
        assert extract_priority_class("2 - High") == 2

    def test_integer(self):
        assert extract_priority_class(3) == 3

    def test_string_integer(self):
        assert extract_priority_class("3") == 3

    def test_none(self):
        assert extract_priority_class(None) is None

    def test_blank(self):
        assert extract_priority_class("") is None

    def test_custom_separator(self):
        config = {"separator": " | "}
        assert extract_priority_class("3 | Moderate", config) == 3


class TestIsPriority3:
    def test_is_priority_3_full(self):
        assert is_priority_3("3 - Moderate") is True

    def test_is_priority_3_int(self):
        assert is_priority_3(3) is True

    def test_not_priority_3(self):
        assert is_priority_3("1 - Critical") is False

    def test_none(self):
        assert is_priority_3(None) is False


class TestContainsSubstring:
    def test_contains_svb(self):
        assert contains_substring("Tier 3 Application Support SVB", "SVB") is True

    def test_contains_svb_case_insensitive(self):
        assert contains_substring("tier 3 application support svb", "SVB") is True

    def test_svb_at_start(self):
        assert contains_substring("SVB Go Support", "SVB") is True

    def test_no_svb(self):
        assert contains_substring("Tier 3 Application Support FCB", "SVB") is False

    def test_blank_value(self):
        assert contains_substring(None, "SVB") is False

    def test_blank_substring(self):
        assert contains_substring("something", None) is False


class TestIsBlank:
    def test_none(self):
        assert is_blank(None) is True

    def test_empty(self):
        assert is_blank("") is True

    def test_spaces(self):
        assert is_blank("   ") is True

    def test_nan(self):
        assert is_blank("nan") is True

    def test_not_blank(self):
        assert is_blank("hello") is False
