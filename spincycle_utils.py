"""
Pure utility functions extracted from app.py for testability and reuse.

These functions have no dependency on Streamlit and can be imported by
the test suite without triggering any UI side effects. app.py imports
from this module to avoid duplicating logic.

Functions here must remain pure (deterministic, no I/O, no global state)
so they're trivially unit-testable.
"""

from __future__ import annotations

import re

import numpy as np
import pandas as pd

# ─────────────────────────────────────────────────────────────────
# Regex patterns for instrument-export sample names
#
# Naming conventions in HPLC/MS exports vary by lab, but tokens like
# "STD_5", "cvs1_5ppm", and "sample_10x" are common signals.
# ─────────────────────────────────────────────────────────────────

# STD concentration after _ or - (avoid matching 10x dilution tokens)
STD_PATTERN = re.compile(r"(?<=[_\-])(\d+\.?\d*)(?![xX\d.])")

# CVS concentration (Chemstation style: cvs_5, cvs-5, cvs5)
CVS_PATTERN = re.compile(r"(?:^|[_\-\s])cvs[_\-]?([\d\.]+)", re.IGNORECASE)

# MassHunter-specific: spike concentration is the second numeric token
# in patterns like cvs1_5ppm (index=1, spike=5).
CVS_PATTERN_MASSHUNTER = re.compile(r"(?:^|[_\-\s])cvs\d+[_\-]([\d\.]+)(?:\s*ppm)?", re.IGNORECASE)

# Dilution token: digits followed by xX before _/-/whitespace/end
DIL_PATTERN = re.compile(r"(?<=[_\-])(\d+\.?\d*)[xX](?=[_\-\s]|$)", re.IGNORECASE)

# Replicate suffix appended by add_rep_suffix (e.g. "_rep2")
_REP_SUFFIX_RE = re.compile(r"_rep\d+$")


# ─────────────────────────────────────────────────────────────────
# Sample classification + concentration extraction
# ─────────────────────────────────────────────────────────────────


def classify_sample(name: str) -> str:
    """Classify a sample by name into STD / CVS / Blank / Sample."""
    n = str(name).lower()
    if "std_" in n or "std-" in n or n.startswith("std"):
        return "STD"
    if "cvs_" in n or "cvs-" in n or n.startswith("cvs"):
        return "CVS"
    if "blank" in n:
        return "Blank"
    return "Sample"


def extract_std_conc(name: str):
    """Extract the trailing numeric STD concentration; NaN if absent."""
    matches = STD_PATTERN.findall(str(name))
    return float(matches[-1]) if matches else np.nan


def extract_cvs_conc(name: str):
    """Extract Chemstation-style CVS concentration; NaN if absent."""
    m = CVS_PATTERN.search(str(name))
    return float(m.group(1)) if m else np.nan


def extract_cvs_conc_masshunter(name: str):
    """MassHunter-only CVS concentration extractor.

    Pulls the SPIKE concentration from names like 'DKmix_cvs1_5ppm' → 5.0
    (the cvs index '1' is ignored). Falls back to the Chemstation regex
    if no `cvs<idx>_<conc>` pattern is found, so names like 'cvs_5ppm' or
    'cvs5' still parse correctly.
    """
    s = str(name)
    m = CVS_PATTERN_MASSHUNTER.search(s)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            pass
    m2 = CVS_PATTERN.search(s)
    return float(m2.group(1)) if m2 else np.nan


def extract_dilution(name: str):
    """Extract dilution factor (e.g. '10x' → 10.0); defaults to 1.0."""
    m = DIL_PATTERN.search(str(name))
    return float(m.group(1)) if m else 1.0


def clean_name(name: str) -> str:
    """Strip dilution token + collapse stray separators left behind."""
    s = DIL_PATTERN.sub("", str(name))
    s = re.sub(r"[_\-]{2,}", "_", s)
    return s.strip("_- ").strip()


def strip_rep_suffix(name) -> str:
    """Inverse of add_rep_suffix's '_repN'. Returns base sample name."""
    return _REP_SUFFIX_RE.sub("", str(name))


# ─────────────────────────────────────────────────────────────────
# LLOQ / ULOQ quantitation-limit flagging
# ─────────────────────────────────────────────────────────────────


def flag_lloq_uloq(value, lloq, uloq) -> str:
    """Return flag string for a measured value vs LLOQ/ULOQ bounds."""
    try:
        v = float(value)
        if pd.notna(lloq) and v < float(lloq):
            return "< LLOQ"
        if pd.notna(uloq) and v > float(uloq):
            return "> ULOQ"
        return "In Range"
    except (TypeError, ValueError):
        return ""


# ─────────────────────────────────────────────────────────────────
# Dropbox path normalization
# ─────────────────────────────────────────────────────────────────


def _normalize_dropbox_path(path: str) -> str:
    """Accept various Dropbox path formats and return /folder form.

    Handles:
      'https://www.dropbox.com/home/Apps/Analytical' → '/Apps/Analytical'
      '/Analytical'                                   → '/Analytical'
      'Analytical'                                    → '/Analytical'
      ''                                              → '/'
    """
    if not path:
        return "/"
    p = path.strip()
    if p.startswith("http"):
        m = re.search(r"/home(/.+)?$", p)
        if m:
            p = m.group(1) or "/"
        else:
            m = re.search(r"/scl/fo/[^/]+/[^/]+(/.+)?$", p)
            p = m.group(1) if m else "/"
    if not p.startswith("/"):
        p = "/" + p
    return p.rstrip("/") or "/"


# ─────────────────────────────────────────────────────────────────
# Excel sheet-name sanitization
# ─────────────────────────────────────────────────────────────────


def _sanitize_sheet_name(name: str, used: set | None = None) -> str:
    """Excel sheet names: <=31 chars, no `: \\ / ? * [ ]`, must be unique."""
    used = used if used is not None else set()
    safe = re.sub(r"[:\\/?*\[\]]", "_", str(name))[:31].strip()
    if not safe:
        safe = "sheet"
    candidate = safe
    i = 1
    while candidate in used:
        suffix = f"_{i}"
        candidate = safe[: 31 - len(suffix)] + suffix
        i += 1
    used.add(candidate)
    return candidate
