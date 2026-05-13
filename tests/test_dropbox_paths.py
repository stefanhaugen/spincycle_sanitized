"""
Tests for Dropbox path normalization in app._normalize_dropbox_path.

The function must accept three input formats and produce a clean
forward-slash path:

    - Already-clean paths:        "/Analytical"  →  "/Analytical"
    - Bare names:                 "Analytical"   →  "/Analytical"
    - Dropbox web URLs:           "https://...home/Apps/X" → "/Apps/X"
    - Empty / None:               ""             →  "/"
"""

from __future__ import annotations

import pytest

from spincycle_utils import _normalize_dropbox_path


class TestBasicPaths:
    """Inputs that don't involve URL parsing."""

    def test_already_normalized_path_unchanged(self):
        assert _normalize_dropbox_path("/Analytical") == "/Analytical"

    def test_bare_name_gets_leading_slash(self):
        assert _normalize_dropbox_path("Analytical") == "/Analytical"

    def test_empty_string_returns_root(self):
        assert _normalize_dropbox_path("") == "/"

    def test_whitespace_only_returns_root(self):
        assert _normalize_dropbox_path("   ") == "/"

    def test_trailing_slash_stripped(self):
        assert _normalize_dropbox_path("/Analytical/") == "/Analytical"

    def test_nested_path_preserved(self):
        assert _normalize_dropbox_path("/Lab/Q1/Reports") == "/Lab/Q1/Reports"


class TestDropboxUrls:
    """Inputs pasted from a Dropbox web browser session."""

    def test_home_url_with_folder(self):
        url = "https://www.dropbox.com/home/Apps/Analytical"
        assert _normalize_dropbox_path(url) == "/Apps/Analytical"

    def test_home_url_root_only(self):
        url = "https://www.dropbox.com/home"
        assert _normalize_dropbox_path(url) == "/"

    def test_home_url_nested_folder(self):
        url = "https://www.dropbox.com/home/Lab/2026/Q1"
        assert _normalize_dropbox_path(url) == "/Lab/2026/Q1"


class TestEdgeCases:
    """Behaviors that should be stable even when the input is weird."""

    @pytest.mark.parametrize("raw", ["  /Analytical  ", " /Analytical"])
    def test_leading_whitespace_trimmed(self, raw: str):
        assert _normalize_dropbox_path(raw) == "/Analytical"

    def test_root_slash_returns_root(self):
        assert _normalize_dropbox_path("/") == "/"
