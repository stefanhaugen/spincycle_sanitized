"""
Analytical Chemistry Data Cleaning Pipeline
Author: Stefan Haugen
Supports: Chemstation, MassHunter
"""

import re

import numpy as np
import pandas as pd
import streamlit as st

# ── Pandas 3.0 compat: disable pyarrow string backend ────────
# Pandas 3.0 defaults to ArrowStringArray which breaks .reshape(),
# JSON roundtrip column names, 'in' checks, and Streamlit Arrow
# serialization. This single flag restores pandas 2.x behavior.
try:
    pd.set_option("future.infer_string", False)
except (pd.errors.OptionError, KeyError, AttributeError):
    pass  # pandas version <2.1 does not include this option, and does not require it
import json
import os
import sqlite3
from datetime import datetime
from io import BytesIO, StringIO

import openpyxl

# ─────────────────────────────────────────────
#  PAGE CONFIGURATION
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Chemstation / MassHunter Data Pipeline",
    layout="wide",
    page_icon="",
    initial_sidebar_state="collapsed",
)
st.title("🧪 SpinCycle Data Pipeline")
st.caption("Chemstation | MassHunter  |  Single & Batch modes")

# ─────────────────────────────────────────────
#  CONSTANTS
# ─────────────────────────────────────────────
_CACHE_VER = "v12"  # bump this to invalidate all @st.cache_data after code changes
# Export Type — the software that produced the raw file, not the physical instrument.
# Chemstation = Agilent ChemStation HPLC export (Data + Labels sheets).
# MassHunter  = Agilent MassHunter export (Sheet1 with *Results headers).
INSTRUMENT_TYPES = ["Chemstation", "MassHunter"]


# ═══════════════════════════════════════════════════════════════
#  DROPBOX  —  group-folder push
# ═══════════════════════════════════════════════════════════════
# Paste your group Dropbox folder path here.
# pasted from a browser (https://www.dropbox.com/home/...) are auto-
# normalized by _normalize_dropbox_path() below.
DROPBOX_DEFAULT_DEST = "/Analytical"


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


def _resolve_dropbox_token() -> str | None:
    """Token resolution priority: st.secrets → session_state → None."""
    try:
        if hasattr(st, "secrets"):
            tok = st.secrets.get("dropbox_token")
            if tok:
                return str(tok)
    except Exception:
        pass
    return st.session_state.get("_dropbox_token")


def _resolve_dropbox_dest() -> str:
    """Destination priority: st.secrets → DROPBOX_DEFAULT_DEST."""
    try:
        if hasattr(st, "secrets"):
            v = st.secrets.get("dropbox_dest")
            if v:
                return _normalize_dropbox_path(str(v))
    except Exception:
        pass
    return _normalize_dropbox_path(DROPBOX_DEFAULT_DEST)


@st.cache_resource(show_spinner=False)
def _dbx_client(token: str):
    """One Dropbox client per token (cached for reuse across reruns)."""
    import dropbox as _dbx

    return _dbx.Dropbox(token)


@st.cache_data(show_spinner=False, ttl=300)
def _dbx_verify_token(token: str) -> tuple[bool, str]:
    """Returns (ok, display_name_or_error). Cached 5 min to avoid per-rerun API call."""
    if not token:
        return False, "No token"
    try:
        acct = _dbx_client(token).users_get_current_account()
        return True, acct.name.display_name
    except Exception as e:
        return False, str(e)[:200]


def _dbx_path_exists(token: str, full_path: str) -> bool:
    try:
        _dbx_client(token).files_get_metadata(full_path)
        return True
    except Exception:
        return False


def _dbx_next_available_name(token: str, folder: str, filename: str) -> str:
    """If folder/filename exists, return filename with (2), (3), … suffix.
    Probes Dropbox sequentially — fine for the typical 1-handful collision case."""
    base, ext = os.path.splitext(filename)
    folder_clean = folder.rstrip("/")
    candidate = filename
    n = 2
    while _dbx_path_exists(token, f"{folder_clean}/{candidate}"):
        candidate = f"{base} ({n}){ext}"
        n += 1
        if n > 99:
            break
    return candidate


def _dbx_upload(token: str, folder: str, filename: str, data: bytes) -> str:
    """Upload bytes to <folder>/<filename>. Returns the full path written.

    Uses WriteMode('add') — Dropbox refuses to overwrite, which is what we
    want since the caller already resolved name collisions via
    _dbx_next_available_name. Auto-creates the destination folder if it
    doesn't exist (idempotent — silently ignores 'already exists' errors).
    """
    import dropbox as _dbx

    client = _dbx_client(token)
    folder_clean = folder.rstrip("/")
    try:
        client.files_create_folder_v2(folder_clean)
    except Exception:
        pass
    full_path = f"{folder_clean}/{filename}"
    client.files_upload(data, full_path, mode=_dbx.files.WriteMode("add"))
    return full_path


def _resolve_dbx_filename(meta: dict, suffix: str, fallback_stem: str) -> str:
    """Compose the Dropbox filename from tracker + optional suffix.

    Default = '{tracking_number}.xlsx'. Falls back to '{fallback_stem}.xlsx'
    if no tracker present. Suffix, when non-empty, is appended with an
    underscore. Illegal filename chars in the suffix are coerced to
    underscores so Dropbox never rejects the upload.
    """
    base = (meta.get("Tracking Number", "") or "").strip() or fallback_stem
    clean_suffix = re.sub(r'[\\/:*?"<>|]+', "_", (suffix or "").strip()).strip("_")
    if clean_suffix:
        return f"{base}_{clean_suffix}.xlsx"
    return f"{base}.xlsx"


def render_dropbox_push_ui(
    key_prefix: str, meta: dict, fallback_stem: str, xlsx_bytes: bytes | None
) -> None:
    """Render the entire Push-to-Dropbox UI block: token, filename, preview, confirm.

    State lives in session_state under:
      {key_prefix}dbx_suffix     — analyst-typed suffix
      {key_prefix}dbx_preview    — populated by Preview Push, cleared on cancel/upload
      _dropbox_token             — global, shared across both modes
    """
    # ── Token row ────────────────────────────────────────────
    tok = _resolve_dropbox_token()
    if tok is None:
        new_tok = st.text_input(
            "Dropbox access token",
            type="password",
            key=f"{key_prefix}dbx_token_input",
            help="Generate at dropbox.com/developers/apps → 'Generate access token'. "
            "Stored in session only — disappears when the app restarts. "
            "Put 'dropbox_token = \"…\"' in .streamlit/secrets.toml to persist.",
        )
        if new_tok:
            ok, msg = _dbx_verify_token(new_tok)
            if ok:
                st.session_state["_dropbox_token"] = new_tok
                tok = new_tok
                st.success(f"Connected as **{msg}**")
            else:
                st.error(f"Invalid token: {msg}")
    else:
        ok, msg = _dbx_verify_token(tok)
        if ok:
            st.caption(f"🔗 Connected to Dropbox as **{msg}**")
        else:
            st.warning(f"Stored token isn't working: {msg}")
            if st.button("Clear token", key=f"{key_prefix}dbx_clear_tok"):
                st.session_state.pop("_dropbox_token", None)
                _dbx_verify_token.clear()
                st.rerun()
            tok = None

    # ── Filename composer ───────────────────────────────────
    suffix = st.text_input(
        "Filename suffix (optional)",
        value="",
        key=f"{key_prefix}dbx_suffix",
        help="Appended to the tracking number with an underscore. "
        "Example: tracker 'T2025-0042' + suffix 'rerun' → T2025-0042_rerun.xlsx",
    )
    resolved_name = _resolve_dbx_filename(meta, suffix, fallback_stem)
    dest = _resolve_dropbox_dest()
    st.caption(f"→ `{dest}/{resolved_name}`")

    # ── Preview button ──────────────────────────────────────
    can_preview = (tok is not None) and (xlsx_bytes is not None)
    if not can_preview:
        st.button(
            "👁 Preview push",
            key=f"{key_prefix}dbx_preview_btn",
            disabled=True,
            help=(
                "Build the Excel first."
                if xlsx_bytes is None
                else "Enter a valid Dropbox token first."
            ),
        )
    elif st.button("👁 Preview push", key=f"{key_prefix}dbx_preview_btn"):
        with st.spinner("Checking destination…"):
            try:
                final_name = _dbx_next_available_name(tok, dest, resolved_name)
                st.session_state[f"{key_prefix}dbx_preview"] = {
                    "dest": dest,
                    "requested_name": resolved_name,
                    "final_name": final_name,
                    "collision": (final_name != resolved_name),
                    "size": len(xlsx_bytes),
                    "data": xlsx_bytes,
                }
            except Exception as e:
                st.error(f"Couldn't reach Dropbox: {e}")

    # ── Preview panel ───────────────────────────────────────
    prev = st.session_state.get(f"{key_prefix}dbx_preview")
    if prev:
        with st.container(border=True):
            st.markdown(f"**Destination:** `{prev['dest']}`")
            badge = "⚠️ exists, will save as new" if prev["collision"] else "✅ new file"
            st.markdown(
                f"**Filename:** `{prev['final_name']}` &nbsp; {badge}", unsafe_allow_html=True
            )
            if prev["collision"]:
                st.caption(f"`{prev['requested_name']}` already exists in `{prev['dest']}`.")
            st.markdown(f"**Size:** {prev['size'] / 1024:.1f} KB")

            cA, cB = st.columns(2)
            with cA:
                if st.button("✅ Confirm & upload", key=f"{key_prefix}dbx_confirm", type="primary"):
                    with st.spinner("Uploading…"):
                        try:
                            full = _dbx_upload(tok, prev["dest"], prev["final_name"], prev["data"])
                            st.success(f"Uploaded → `{full}`")
                            st.session_state.pop(f"{key_prefix}dbx_preview", None)
                        except Exception as e:
                            st.error(f"Upload failed: {e}")
            with cB:
                if st.button("✖ Cancel", key=f"{key_prefix}dbx_cancel"):
                    st.session_state.pop(f"{key_prefix}dbx_preview", None)
                    st.rerun()


INSTRUMENT_NAMES = [
    "",
    "R2P2",
    "Yoshi",
    "Crush",
    "WallE",
    "Spider",
    "BB8",
    "Luigi",
    "Mario",
    "Peach",
    "Black Dragon",
    "Red Dragon",
    "Eve",
    "Toucan",
    "Groot",
    "D-0",
    "Rocket",
]

# FIX: \b fails when std/cvs is preceded by _ (both are word chars, no boundary).
# Use (?:^|[_\-\s]) lookbehind-style alternation to anchor on start, _, -, or space.
STD_PATTERN = re.compile(r"(?<=[_\-])(\d+\.?\d*)(?![xX\d.])")
CVS_PATTERN = re.compile(r"(?:^|[_\-\s])cvs[_\-]?([\d\.]+)", re.IGNORECASE)
# MassHunter-specific: spike concentration regex.
# MassHunter sample names follow `cvs<idx>_<conc>(ppm)?` (e.g. DKmix_cvs1_5ppm
# → CVS injection #1 at 5 ppm). The Chemstation CVS_PATTERN would capture the
# index (1), not the spike concentration. This pattern captures the second
# numeric token. Falls back to CVS_PATTERN behavior if no _<conc> tail exists.
CVS_PATTERN_MASSHUNTER = re.compile(r"(?:^|[_\-\s])cvs\d+[_\-]([\d\.]+)(?:\s*ppm)?", re.IGNORECASE)
# FIX: lookahead replaces \b — \b silently fails when anything follows x (e.g. _10x_rep1)
# Requires dil token be preceded by _/- AND followed by _/-, whitespace, or end-of-string
DIL_PATTERN = re.compile(r"(?<=[_\-])(\d+\.?\d*)[xX](?=[_\-\s]|$)", re.IGNORECASE)

# ─────────────────────────────────────────────
#  SESSION STATE
# ─────────────────────────────────────────────
for _k, _v in {
    "submission_wb": None,
    "submission_meta": {},
    "submission_samples": None,
    # Excel/SQLite bytes stored here so download buttons persist
    "xlsx_bytes": None,
    "xlsx_filename": "",
    "db_bytes": None,
    "batch_master_bytes": None,
    "batch_db_bytes": None,
    "batch_comb": None,
    "batch_all_r": None,
}.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ═══════════════════════════════════════════════════════════════
#  PANDAS 3.0 COMPAT — safe JSON roundtrip
# ═══════════════════════════════════════════════════════════════


def safe_read_json(json_str, **kwargs):
    """pd.read_json wrapper that fixes pandas 3.0 issues after JSON roundtrip:
    - Column names may become float NaN → force to str
    - String columns may become object with mixed types → coerce Sample Name
    """
    df = pd.read_json(StringIO(json_str), **kwargs)
    if isinstance(df, pd.DataFrame):
        df.columns = [str(c) if pd.notna(c) else f"_col_{i}" for i, c in enumerate(df.columns)]
        if "Sample Name" in df.columns:
            df["Sample Name"] = df["Sample Name"].astype(str)
    return df


# ═══════════════════════════════════════════════════════════════
#  PURE UTILITIES
# ═══════════════════════════════════════════════════════════════


def classify_sample(name: str) -> str:
    n = str(name).lower()
    # Simple substring match — any name containing std_ or std- → STD
    if "std_" in n or "std-" in n or n.startswith("std"):
        return "STD"
    if "cvs_" in n or "cvs-" in n or n.startswith("cvs"):
        return "CVS"
    if "blank" in n:
        return "Blank"
    return "Sample"


def extract_std_conc(name: str):
    matches = STD_PATTERN.findall(str(name))
    return float(matches[-1]) if matches else np.nan


def extract_cvs_conc(name: str):
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
    # Fall back to the generic Chemstation pattern
    m2 = CVS_PATTERN.search(s)
    return float(m2.group(1)) if m2 else np.nan


def extract_dilution(name: str):
    m = DIL_PATTERN.search(str(name))
    return float(m.group(1)) if m else 1.0


def clean_name(name: str) -> str:
    s = DIL_PATTERN.sub("", str(name))
    s = re.sub(r"[_\-]{2,}", "_", s)  # collapse double _/- left after removing dil token
    return s.strip("_- ").strip()


def add_rep_suffix(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    counts = df.groupby("Sample Name").cumcount()
    mask = df.groupby("Sample Name")["Sample Name"].transform("count") > 1
    df["Sample Name"] = df["Sample Name"].where(
        ~mask, df["Sample Name"] + "_rep" + (counts + 1).astype(str)
    )
    return df


_REP_SUFFIX_RE = re.compile(r"_rep\d+$")


def strip_rep_suffix(name) -> str:
    """Inverse of add_rep_suffix's appended _repN. Returns the base sample
    name. Used by CVS analyte assignment so a single assignment row applies
    to every rep of the same injection."""
    return _REP_SUFFIX_RE.sub("", str(name))


def inject_meta(df: pd.DataFrame, meta: dict) -> pd.DataFrame:
    for k, v in meta.items():
        df[k] = v
    return df


def detect_cvs_levels(tidy_df: pd.DataFrame) -> list:
    cvs = tidy_df[tidy_df["isStandard"] == "CVS"]["CVS_Known_Conc"].dropna().unique()
    return sorted([float(c) for c in cvs])


def detect_std_levels(tidy_df: pd.DataFrame) -> list:
    """All unique STD concentrations present in the tidy data."""
    stds = tidy_df[tidy_df["isStandard"] == "STD"]["STD_Known_Conc"].dropna().unique()
    return sorted([float(c) for c in stds])


def get_analyte_pairs(df: pd.DataFrame) -> list:
    """
    Returns list of (analyte_name, raw_col, corrected_col) tuples.
    raw_col      = analyte|Amount
    corrected_col= analyte|AmountxDilutionFactor  (may be None)

    Order preserves the first-seen position of each analyte in df.columns,
    which for Chemstation is the Labels-sheet 'Title' order = retention-time
    order. MassHunter preserves whatever column order the export wrote.
    """
    raw_cols = {c.split("|")[0]: c for c in df.columns if "|Amount" in c and "xDilution" not in c}
    corr_cols = {c.split("|")[0]: c for c in df.columns if "AmountxDilutionFactor" in c}
    ordered: list = []
    seen: set = set()
    for c in df.columns:
        if "|Amount" not in c:
            continue
        a = c.split("|")[0]
        if a not in seen:
            seen.add(a)
            ordered.append(a)
    return [(a, raw_cols.get(a), corr_cols.get(a)) for a in ordered]


# ═══════════════════════════════════════════════════════════════
#  LLOQ / ULOQ  HELPER
# ═══════════════════════════════════════════════════════════════


def flag_lloq_uloq(value, lloq, uloq):
    """Return flag string for a measured value vs LLOQ/ULOQ."""
    try:
        v = float(value)
        if pd.notna(lloq) and v < float(lloq):
            return "< LLOQ"
        if pd.notna(uloq) and v > float(uloq):
            return "> ULOQ"
        return "In Range"
    except (TypeError, ValueError):
        return ""


def decide_cell_style(
    rtype: str,
    sample_name: str,
    known,
    ref_val,
    analyte: str,
    std_lo: float,
    std_hi: float,
    cvs_tolerances: dict,
    analyte_assignment: dict,
    lloq_map: dict,
    uloq_map: dict,
) -> str:
    """Pass/fail decision for a single value cell.

    Returns one of: 'pass', 'fail', 'lloq', 'neutral'. Both the Streamlit
    styler (render_workup_table) and the Excel writer (build_excel) call
    this so cell coloring is identical between on-screen review and the
    downloaded workbook.
    """
    cvs_tolerances = cvs_tolerances or {}
    analyte_assignment = analyte_assignment or {}
    lloq_map = lloq_map or {}
    uloq_map = uloq_map or {}

    # Normalize ref_val to float or NaN
    try:
        rv = float(ref_val)
    except (TypeError, ValueError):
        rv = float("nan")
    if pd.isna(rv):
        return "neutral"

    if rtype == "STD":
        if known is None or known <= 0:
            return "neutral"
        pct = rv / known * 100.0
        return "pass" if (std_lo <= pct <= std_hi) else "fail"
    if rtype == "CVS":
        if known is None or known <= 0:
            return "neutral"
        base = strip_rep_suffix(sample_name)
        assigned = analyte_assignment.get(base, analyte_assignment.get(sample_name, {}))
        if assigned and not assigned.get(analyte, True):
            return "neutral"
        tol = float(cvs_tolerances.get(float(known), 10.0))
        pct = rv / known * 100.0
        return "pass" if (100.0 - tol <= pct <= 100.0 + tol) else "fail"
    if rtype == "Sample":
        lloq = lloq_map.get(analyte)
        uloq = uloq_map.get(analyte)
        if lloq is None and uloq is None:
            return "neutral"
        if lloq is not None and rv < float(lloq):
            return "lloq"
        if uloq is not None and rv > float(uloq):
            return "fail"
        return "pass"
    return "neutral"


# Shared color palette — used by Streamlit styler and Excel writer alike
QC_COLORS = {
    "pass": {"bg": "D4EDDA", "fg": "155724"},
    "fail": {"bg": "F8D7DA", "fg": "721C24"},
    "lloq": {"bg": "FFF3CD", "fg": "856404"},
    "ref": {"bg": "FFF2CC", "fg": "7F6000"},
    "sep": {"bg": "E9ECEF", "fg": "000000"},
    "blank": {"bg": "E8F4F8", "fg": "1F4E5F"},
    "neutral": {"bg": None, "fg": None},
}


def merge_run_meta(ui_meta: dict, extracted_meta: dict | None) -> dict:
    """Merge UI-supplied metadata with parser-extracted metadata.

    Rule: UI wins for fields the user filled in; extracted fills in blanks.
    Exception: 'Run Date' — extracted always wins if present, since the
    instrument timestamp is more authoritative than the date_input's
    today-by-default value.
    """
    merged = {**ui_meta}
    if not extracted_meta:
        return merged
    for k, v in extracted_meta.items():
        if v and not merged.get(k):
            merged[k] = v
    # Run Date: extracted always wins when available
    rd = extracted_meta.get("Run Date")
    if rd:
        merged["Run Date"] = rd
    return merged


# ═══════════════════════════════════════════════════════════════
#  CACHED SUBMISSION FORM PARSER
# ═══════════════════════════════════════════════════════════════

SUBMISSION_FIELD_MAP = {
    "tracker #": "Tracking Number",
    "analyst (poc)": "Analyst",
    "date received": "Date Received",
    "notes": "Notes",
    "sample submitter": "Sample Submitter",
    "task pi": "Task PI",
    "task number": "Task Number",
    "task name": "Task Name",
    "analysis requested": "Analysis Requested",
    "research phase": "Research Phase",
    "storage conditions": "Storage Conditions",
    "sample return post analysis": "Sample Return",
    "sample container label": "Sample Container Label",
    "previous analysis": "Previous Analysis",
    "experiment description": "Experiment Description",
}


@st.cache_data(show_spinner=False)
def cached_parse_submission(file_bytes: bytes, _ver=_CACHE_VER) -> tuple:
    # ── Performance: read_only=True switches openpyxl to a streaming SAX
    # parser instead of building the full in-memory cell grid. On a typical
    # submission form (one info sheet + one sample table) this cuts parse
    # time by 5-10x. data_only=True still reads cached values not formulas.
    # keep_links=False skips external-link resolution. close() is required
    # in read_only mode to release the underlying ZIP file handle.
    wb = openpyxl.load_workbook(
        BytesIO(file_bytes),
        data_only=True,
        read_only=True,
        keep_links=False,
    )
    try:
        meta = {}
        if "Submission Information" in wb.sheetnames:
            for row in wb["Submission Information"].iter_rows(values_only=True):
                if row[0] is None:
                    continue
                label = str(row[0]).strip().lower().rstrip(":")
                value = row[2] if len(row) > 2 else None
                if isinstance(value, datetime):
                    value = value.strftime("%Y-%m-%d")
                for raw_key, clean_key in SUBMISSION_FIELD_MAP.items():
                    if label.startswith(raw_key):
                        meta[clean_key] = str(value) if value is not None else ""
                        break
        samples_df = pd.DataFrame()
        if "Submission Sample Table" in wb.sheetnames:
            rows = list(wb["Submission Sample Table"].iter_rows(values_only=True))
            header_idx = next(
                (
                    i
                    for i, r in enumerate(rows)
                    if any(
                        str(c).lower() in ("sample number", "sample name on vial")
                        for c in r
                        if c is not None
                    )
                ),
                None,
            )
            if header_idx is not None:
                headers = [
                    str(c).strip() if c else f"col_{j}" for j, c in enumerate(rows[header_idx])
                ]
                data_rows = [r for r in rows[header_idx + 1 :] if any(v is not None for v in r)]
                samples_df = pd.DataFrame(data_rows, columns=headers)
                samples_df.dropna(how="all", axis=1, inplace=True)
    finally:
        wb.close()
    return meta, samples_df.to_json(), file_bytes


# ═══════════════════════════════════════════════════════════════
#  CACHED INSTRUMENT PARSERS
# ═══════════════════════════════════════════════════════════════


@st.cache_data(show_spinner=False)
def cached_parse_chemstation(file_bytes: bytes, _ver=_CACHE_VER) -> tuple:
    """Parse an Agilent ChemStation HPLC export.

    Canonical structure: a 'Data' sheet (sample × analyte values) and a
    'Labels' sheet (column metadata, including the batch path at E2).
    """
    data_df = pd.read_excel(BytesIO(file_bytes), sheet_name="Data")
    labels_df = pd.read_excel(BytesIO(file_bytes), sheet_name="Labels")

    # ── Extract embedded metadata BEFORE any column drops ────
    # Default pd.read_excel(header=0) → Excel row 1 = header,
    # Excel row 2 = df.iloc[0]. Column E = df.iloc[:, 4].
    # So Excel cell E2  ==  labels_df.iloc[0, 4]  on the freshly-read sheet.
    extracted_meta = {}
    try:
        if labels_df.shape[0] > 0 and labels_df.shape[1] > 4:
            v = labels_df.iloc[0, 4]
            if pd.notna(v) and str(v).strip():
                extracted_meta["Batch Path"] = str(v).strip()
    except Exception:
        pass

    # ── Run Date — ChemStation auto-batches with the date embedded
    #    in the batch path as YYYY-MM-DD, e.g.
    #       C:\…\2026-04-15_Run42\…\batch.dxd
    #    so a regex on the Batch Path string is the most reliable extraction.
    _bp = extracted_meta.get("Batch Path", "")
    if _bp:
        m = re.search(r"(\d{4}-\d{2}-\d{2})", _bp)
        if m:
            try:
                # Validate it's a real date — rejects e.g. 2026-13-99
                extracted_meta["Run Date"] = pd.Timestamp(m.group(1)).strftime("%Y-%m-%d")
            except Exception:
                pass

    data_df.drop(data_df.columns[:2], axis=1, inplace=True)
    labels_df.drop(labels_df.columns[:2], axis=1, inplace=True)
    data_df.columns = np.array(labels_df.loc[1:, "Title"].tolist()).reshape(-1)
    raw_df = (
        data_df.copy()
    )  # FIX: was data_df.copy() — ensure mutations below don't corrupt raw sheet
    data_df["Sample"] = data_df["Sample"].astype(str)
    # NOTE: blanks are NOT filtered here. classify_sample tags them 'Blank' so
    # they flow through to Workup. Summary filters them out via checkbox.
    data_df["isStandard"] = data_df["Sample"].apply(classify_sample)
    data_df["Dilution Factor"] = data_df["Sample"].apply(extract_dilution)
    data_df["STD_Known_Conc"] = data_df["Sample"].apply(extract_std_conc)
    data_df["CVS_Known_Conc"] = data_df["Sample"].apply(extract_cvs_conc)
    for col in [c for c in data_df.columns if "Amount" in c and "xDilution" not in c]:
        data_df[f"{col}xDilutionFactor"] = data_df[col] * data_df["Dilution Factor"]
    keep = [
        "Sample",
        "Location",
        "Dilution Factor",
        "isStandard",
        "STD_Known_Conc",
        "CVS_Known_Conc",
        "Run",
    ] + [c for c in data_df.columns if "Amount" in c]
    data_df = data_df[[c for c in keep if c in data_df.columns]]
    data_df.rename(columns={"Sample": "Sample Name"}, inplace=True)
    return raw_df.to_json(), data_df.to_json(), json.dumps(extracted_meta)


@st.cache_data(show_spinner=False)
def _find_masshunter_sheet(file_bytes: bytes) -> str:
    """Locate the MassHunter results sheet.

    The canonical name is 'Sheet1'. If a file has been renamed (it happens),
    fall back to scanning every sheet for the row-0 ``*Results`` signature
    that uniquely identifies a MassHunter analyte header row. Returns
    'Sheet1' as a last-ditch attempt so the caller surfaces a real read error.
    """
    try:
        xl = pd.ExcelFile(BytesIO(file_bytes))
    except Exception:
        return "Sheet1"
    # Prefer Sheet1 (canonical) if present, then any other sheet.
    candidates = (["Sheet1"] if "Sheet1" in xl.sheet_names else []) + [
        s for s in xl.sheet_names if s != "Sheet1"
    ]
    for sn in candidates:
        try:
            row0 = pd.read_excel(BytesIO(file_bytes), sheet_name=sn, header=None, nrows=1).iloc[0]
            if any("Results" in str(v) for v in row0 if pd.notna(v)):
                return sn
        except Exception:
            continue
    return "Sheet1"


@st.cache_data(show_spinner=False)
def cached_parse_masshunter(file_bytes: bytes, _ver=_CACHE_VER) -> tuple:
    """Parse an Agilent MassHunter export.

    Canonical layout (Sheet1):
      Row 0: analyte 'Results' headers at cols 10, 15, 20, ... (every 5 cols)
      Row 1: identity sub-headers at cols 2-9
             (Name, Data File, Type, Level, Dil., Acq. Date-Time, Pos.,
              Acq. Method File), then RT / Area / Calc. Conc. / Final Conc. /
              Accuracy repeating per analyte starting at col 10
      Row 2+: data, including instrument-classified `Type` (Sample / Cal /
              QC) and `Level` (known standard concentration).

    Rules (parity with Chemstation, plus MassHunter-specific behavior):
      • `Type` is authoritative for isStandard when populated — overrides
        the regex-based classifier from sample name. Falls back to
        ``classify_sample`` otherwise.
      • `Level` is authoritative for STD_Known_Conc when populated —
        overrides regex-based extraction. Falls back to ``extract_std_conc``.
      • CVS spike concentration uses the MassHunter-specific regex (captures
        the `_<conc>` after `cvs<idx>`, e.g. ``cvs1_5ppm`` → 5.0).
      • Analytical Method auto-populated from the modal (most-common)
        ``Acq. Method File`` value across data rows.
      • ``|Calc. Conc.`` → ``|Amount`` (raw, pre-dilution).
        ``|Final Conc.`` → ``|AmountxDilutionFactor`` (corrected; recomputed
        from raw × dilution factor so sample-name dil tokens always win).
      • Workup output keeps `Type` and `Level` columns alongside the cleaned
        amounts; drops `Data File` and `Acq. Method File`. The downstream
        Summary builder further drops `Type` and `Level` for the polished
        deliverable.
    """
    sheet_name = _find_masshunter_sheet(file_bytes)
    df = pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, header=None)

    # ── Extract analyte names from row 0 ─────────────────────
    # Format: 'JUL25_hexamethylene diamine_HD Results'
    # Strip the trailing ' Results' and any batch-prefix like 'JUL25_'.
    # Preserve the trailing short-code (e.g. '_HD', '_AA') alongside the
    # long form so both naming conventions remain visible downstream.
    analyte_map = {}
    for i, v in enumerate(df.iloc[0]):
        if pd.notna(v) and "Results" in str(v):
            name = re.sub(r"\s*Results$", "", str(v)).strip()
            name = re.sub(r"^[A-Z]{2,5}\d{2,6}_", "", name).strip()  # strip batch prefix
            analyte_map[i] = name

    # ── Build flat column names from row 1 sub-headers ───────
    row1 = df.iloc[1].tolist()
    cols = []
    current_analyte = None
    for i, h in enumerate(row1):
        if i in analyte_map:
            current_analyte = analyte_map[i]
        if i < 2:
            cols.append(f"_drop_{i}")
        elif i < 10:
            cols.append(str(h) if pd.notna(h) else f"_drop_{i}")
        else:
            sub = str(h) if pd.notna(h) else "unknown"
            cols.append(f"{current_analyte}|{sub}" if current_analyte else sub)

    data = df.iloc[2:].copy()
    data.columns = cols
    data = data[[c for c in cols if not c.startswith("_drop_")]].reset_index(drop=True)

    # Preserve the raw (pre-cleaning) frame for the 'Raw Data' export sheet.
    raw_original = data.copy()

    # ── Drop superfluous identity columns per spec ───────────
    # 'Data File' and 'Acq. Method File' are not useful in the workup view
    # (the method file is hoisted to extracted_meta instead, see below).
    for drop_col in ("Data File", "Acq. Method File"):
        if drop_col in data.columns:
            # Hoist Acq. Method File modal value to extracted_meta before dropping
            if drop_col == "Acq. Method File":
                pass  # handled below after extracted_meta is built
            data = data.drop(columns=drop_col)

    # ── Auto-extract Analytical Method (modal Acq. Method File) ──
    extracted_meta = {}
    try:
        amf = raw_original.get("Acq. Method File")
        if amf is not None:
            non_blank = amf.dropna().astype(str)
            non_blank = non_blank[non_blank.str.strip() != ""]
            if len(non_blank) > 0:
                modal = non_blank.mode()
                if len(modal) > 0:
                    extracted_meta["Analytical Method"] = str(modal.iloc[0]).strip()
    except Exception:
        pass

    # ── Auto-extract Run Date (earliest Acq. Date-Time across the run) ──
    # MassHunter writes one timestamp per injection; the earliest is the
    # sequence start, which is what the analyst calls "the run date".
    # Wins over the UI default of today() via merge_run_meta's Run Date
    # special case — same behavior Chemstation already gets from the
    # batch path's YYYY-MM-DD.
    try:
        adt = raw_original.get("Acq. Date-Time")
        if adt is not None:
            times = pd.to_datetime(adt, errors="coerce").dropna()
            if len(times) > 0:
                extracted_meta["Run Date"] = times.min().strftime("%Y-%m-%d")
    except Exception:
        pass

    # ── Rename identity columns to standard names ────────────
    data.rename(
        columns={
            "Name": "Sample Name",
            "Dil.": "Dilution Factor",
            "Pos.": "Location",
            "Acq. Date-Time": "Acq DateTime",
        },
        inplace=True,
    )

    # NOTE: blanks are NOT filtered — they flow through tagged as 'Blank'
    # by classify_sample so they show up in Workup. Summary filters via checkbox.

    data["Sample Name"] = data["Sample Name"].astype(str)

    # ── isStandard: prefer instrument `Type`, fall back to name parsing ──
    # MassHunter Type values: 'Sample', 'Cal' (calibration STD), occasionally
    # 'QC' or 'Blank'. Map them to the app's canonical labels.
    # NOTE: MassHunter writes Type='Sample' for CVS injections too — the
    # instrument has no concept of "CVS". When Type doesn't disambiguate,
    # consult the sample name, including the MassHunter-specific `cvs<digit>`
    # pattern (e.g. cvs1, cvs2) which Chemstation's classify_sample misses
    # because it only recognizes cvs_ / cvs- / leading cvs.
    _type_to_label = {
        "cal": "STD",
        "std": "STD",
        "qc": "CVS",
        "cvs": "CVS",
        "blank": "Blank",
    }
    _cvs_name_pattern = re.compile(r"(?:^|[_\-\s])cvs\d", re.IGNORECASE)

    def _classify_masshunter(row):
        t = row.get("Type")
        name = str(row.get("Sample Name", ""))
        if pd.notna(t):
            mapped = _type_to_label.get(str(t).strip().lower())
            if mapped is not None:
                return mapped
            # Type == 'Sample' (or anything unrecognized) — fall through
        # MassHunter-specific name pattern: cvs<digit> with no separator.
        if _cvs_name_pattern.search(name):
            return "CVS"
        return classify_sample(name)

    if "Type" in data.columns:
        data["isStandard"] = data.apply(_classify_masshunter, axis=1)
    else:
        # No Type column — still apply the MassHunter cvs<digit> rule.
        def _classify_no_type(name):
            if _cvs_name_pattern.search(str(name)):
                return "CVS"
            return classify_sample(name)

        data["isStandard"] = data["Sample Name"].apply(_classify_no_type)

    # ── STD_Known_Conc: prefer instrument `Level`, fall back to name parsing ──
    if "Level" in data.columns:
        level_numeric = pd.to_numeric(data["Level"], errors="coerce")
        name_std = data["Sample Name"].apply(extract_std_conc)
        # Level wins where populated; otherwise fall back to name regex
        data["STD_Known_Conc"] = level_numeric.where(level_numeric.notna(), name_std)
    else:
        data["STD_Known_Conc"] = data["Sample Name"].apply(extract_std_conc)

    # ── CVS_Known_Conc: MassHunter-specific spike-concentration regex ──
    data["CVS_Known_Conc"] = data["Sample Name"].apply(extract_cvs_conc_masshunter)

    # ── Dilution factor: instrument column first, sample name as override ──
    data["Dilution Factor"] = pd.to_numeric(data.get("Dilution Factor"), errors="coerce").fillna(
        1.0
    )
    name_dil = data["Sample Name"].apply(extract_dilution)
    data["Dilution Factor"] = data.apply(
        lambda r: name_dil[r.name] if name_dil[r.name] != 1.0 else r["Dilution Factor"],
        axis=1,
    )

    # ── Rename concentration columns to canonical app suffixes ──
    #     |Calc. Conc.  → |Amount                  (raw, pre-dilution)
    #     |Final Conc.  → |AmountxDilutionFactor   (corrected)
    rename_map = {}
    for c in data.columns:
        if "|Calc. Conc." in c:
            rename_map[c] = c.replace("|Calc. Conc.", "|Amount")
        elif "|Final Conc." in c:
            rename_map[c] = c.replace("|Final Conc.", "|AmountxDilutionFactor")
    data.rename(columns=rename_map, inplace=True)

    # ── Recompute corrected columns from raw × dilution factor ──
    # This ensures any sample-name dilution token (e.g. _10x) is always
    # applied, overriding whatever MassHunter stored in Final Conc.
    for col in [c for c in data.columns if "|Amount" in c and "xDilution" not in c]:
        analyte = col.split("|")[0]
        data[f"{analyte}|AmountxDilutionFactor"] = (
            pd.to_numeric(data[col], errors="coerce") * data["Dilution Factor"]
        )

    # ── Final column selection for the Workup tidy frame ─────
    # Keep Type and Level so they flow into the rendered Workup table; the
    # Summary builder drops them downstream for the polished output.
    base = [
        c
        for c in [
            "Sample Name",
            "Location",
            "Dilution Factor",
            "Type",
            "Level",
            "isStandard",
            "STD_Known_Conc",
            "CVS_Known_Conc",
            "Acq DateTime",
        ]
        if c in data.columns
    ]
    amts = [c for c in data.columns if "|Amount" in c]
    return raw_original.to_json(), data[base + amts].to_json(), json.dumps(extracted_meta)


CACHED_PARSERS = {
    "Chemstation": cached_parse_chemstation,
    "MassHunter": cached_parse_masshunter,
}


# ═══════════════════════════════════════════════════════════════
#  CACHED QC BUILDERS
# ═══════════════════════════════════════════════════════════════


@st.cache_data(show_spinner=False)
def cached_std_workup(tidy_json: str, unit: str, lo: float, hi: float, _ver=_CACHE_VER) -> str:
    df = safe_read_json(tidy_json)
    std = df[df["isStandard"] == "STD"].copy()
    if std.empty:
        return pd.DataFrame({"Note": ["No STD rows found"]}).to_json()
    a_cols = [c for c in std.columns if "|Amount" in c and "xDilution" not in c]
    rows = []
    for _, row in std.iterrows():
        known = pd.to_numeric(row.get("STD_Known_Conc"), errors="coerce")
        for col in a_cols:
            analyte = col.split("|")[0]
            measured = pd.to_numeric(row[col], errors="coerce")
            pct = round(measured / known * 100, 2) if (pd.notna(known) and known > 0) else np.nan
            rows.append(
                {
                    "Sample Name": row["Sample Name"],
                    "Analyte": analyte,
                    f"Known ({unit})": known,
                    f"Measured ({unit})": measured,
                    "% Recovery": pct,
                    f"Pass ({lo}-{hi}%)": "✅" if (pd.notna(pct) and lo <= pct <= hi) else "❌",
                }
            )
    return pd.DataFrame(rows).to_json()


@st.cache_data(show_spinner=False)
def cached_cvs_workup(
    tidy_json: str,
    unit: str,
    cvs_tol_json: str,
    analyte_assignment_json: str = "{}",
    _ver=_CACHE_VER,
) -> str:
    """
    Build the CVS recovery table.

    analyte_assignment_json: optional JSON of {sample_name: {analyte: bool}}.
    Combinations marked False are emitted with NaN %Recovery and Pass="—",
    so the heatmap renders them as gray and they don't pollute pass-rate stats.
    Missing entries default to True (evaluate every analyte) — backward-compatible
    with the prior behavior.
    """
    df = safe_read_json(tidy_json)
    tol = pd.read_json(StringIO(cvs_tol_json), typ="series").to_dict()
    try:
        assignment = json.loads(analyte_assignment_json) if analyte_assignment_json else {}
    except Exception:
        assignment = {}
    cvs = df[df["isStandard"] == "CVS"].copy()
    if cvs.empty:
        return pd.DataFrame({"Note": ["No CVS rows found"]}).to_json()
    a_cols = [c for c in cvs.columns if "|Amount" in c and "xDilution" not in c]
    rows = []
    for _, row in cvs.iterrows():
        known = pd.to_numeric(row.get("CVS_Known_Conc"), errors="coerce")
        tolerance = 10.0
        if pd.notna(known):
            for k, v in tol.items():
                try:
                    if abs(float(k) - known) < 0.001:
                        tolerance = float(v)
                        break
                except (ValueError, TypeError):
                    pass
        lo, hi = 100 - tolerance, 100 + tolerance
        sample_name = row["Sample Name"]
        base_name = strip_rep_suffix(sample_name)
        # Assignment table keys are base names (rep suffix stripped).
        # Fall back to exact match for backward compat.
        sample_assign = assignment.get(base_name, assignment.get(sample_name, {}))
        for col in a_cols:
            analyte = col.split("|")[0]
            measured = pd.to_numeric(row[col], errors="coerce")
            active = bool(sample_assign.get(analyte, True))
            if not active:
                rows.append(
                    {
                        "Sample Name": sample_name,
                        "Analyte": analyte,
                        "CVS Level": known,
                        f"Known ({unit})": known,
                        f"Measured ({unit})": measured,
                        "Tolerance (±%)": tolerance,
                        "% Recovery": np.nan,
                        "Pass": "—",
                        "Active": False,
                    }
                )
                continue
            pct = round(measured / known * 100, 2) if (pd.notna(known) and known > 0) else np.nan
            rows.append(
                {
                    "Sample Name": sample_name,
                    "Analyte": analyte,
                    "CVS Level": known,
                    f"Known ({unit})": known,
                    f"Measured ({unit})": measured,
                    "Tolerance (±%)": tolerance,
                    "% Recovery": pct,
                    "Pass": "✅" if (pd.notna(pct) and lo <= pct <= hi) else "❌",
                    "Active": True,
                }
            )
    return pd.DataFrame(rows).to_json()


@st.cache_data(show_spinner=False)
def cached_merge_submission(
    tidy_json: str, sub_meta_json: str, sub_samples_json: str, _ver=_CACHE_VER
) -> str:
    df = safe_read_json(tidy_json)
    sub_meta = pd.read_json(StringIO(sub_meta_json), typ="series").to_dict()
    for k, v in sub_meta.items():
        if k not in df.columns:
            df[k] = v
    if sub_samples_json:
        try:
            sub_df = safe_read_json(sub_samples_json)
        except Exception:
            sub_df = pd.DataFrame()
        if not sub_df.empty:
            vial_col = next(
                (
                    c
                    for c in sub_df.columns
                    if "sample name" in str(c).lower() or "vial" in str(c).lower()
                ),
                None,
            )
            if vial_col:
                sub_df["_mk"] = sub_df[vial_col].astype(str).str.lower().str.strip()
                df["_mk"] = df["Sample Name"].astype(str).str.lower().str.strip()
                extra = [c for c in sub_df.columns if c not in (vial_col, "_mk", "Sample number")]
                for col in extra:
                    df[f"sub_{col}"] = None
                for i, row in df.iterrows():
                    for _, sr in sub_df.iterrows():
                        if sr["_mk"] in row["_mk"] or row["_mk"] in sr["_mk"]:
                            for col in extra:
                                df.at[i, f"sub_{col}"] = sr.get(col)
                            break
                df.drop(columns=["_mk"], inplace=True)
    return df.to_json()


# ═══════════════════════════════════════════════════════════════
#  WORKUP BUILDER
#  Single integrated review surface that replaces the old
#  Summary + STD Workup + CVS Workup sheets.
#
#  Row order: [LLOQ/ULOQ ref] + [STDs] + [blank] + [CVSs] + [blank] + [Samples]
#  Columns:   Sample Name, [Dilution Factor], Known Conc,
#             then per analyte: value [unit], corrected [unit] (opt), QC
#
#  Returns (workup_df, row_types) where row_types is parallel to the
#  dataframe rows: 'ref' | 'std' | 'cvs' | 'sample' | 'blank'.
#  This parallel list lets the Streamlit styler decide cell coloring
#  without needing a 'Type' column in the visible output.
# ═══════════════════════════════════════════════════════════════


def build_workup(
    tidy_df: pd.DataFrame,
    lloq_map: dict,
    uloq_map: dict,
    unit: str,
    std_lo: float = 90,
    std_hi: float = 110,
    cvs_tolerances: dict | None = None,
    analyte_assignment: dict | None = None,
) -> tuple:
    """
    Build the integrated workup dataframe — the full review surface.
    Always includes every row type and both raw + corrected value columns.
    Does NOT emit per-analyte QC text columns; pass/fail is decided at
    render time by the styler from row_meta (known/lloq/uloq) and the
    same std_lo/std_hi/cvs_tolerances/analyte_assignment used by the heatmap.

    Returns:
        (workup_df, row_types, row_meta)

        row_types entries:
          'ref'    — LLOQ/ULOQ reference label row
          'std'    — STD injection (coloring vs std_lo/std_hi)
          'cvs'    — CVS injection (coloring vs per-level tol + assignment)
          'blank'  — Blank injection (no coloring, just values)
          'sample' — actual sample (LLOQ/ULOQ flag coloring)
          'sep'    — visual separator between sections

        row_meta is a parallel list of dicts, one per row, with keys:
          'sample_name', 'known', 'std_lo', 'std_hi', 'tol',
          'lloq_map', 'uloq_map' — whatever the styler needs.
    """
    cvs_tolerances = cvs_tolerances or {}
    analyte_assignment = analyte_assignment or {}
    pairs = get_analyte_pairs(tidy_df)

    # MassHunter parser emits instrument-classified `Type` and `Level` columns
    # alongside the cleaned amounts. They flow through to Workup as audit
    # columns and are stripped by build_summary for the polished output.
    include_type = "Type" in tidy_df.columns
    include_level = "Level" in tidy_df.columns

    # ── Column structure — both raw and corrected always shown ────
    cols: list = ["Sample Name", "Dilution Factor"]
    if include_type:
        cols.append("Type")
    if include_level:
        cols.append("Level")
    cols.append("Known Conc")
    for analyte, _raw_col, _corr_col in pairs:
        cols.append(f"{analyte} [{unit}]")
        cols.append(f"{analyte} corrected [{unit}]")

    def empty_row() -> dict:
        return dict.fromkeys(cols, "")

    def make_row(src: pd.Series, type_label: str) -> tuple:
        row = empty_row()
        row["Sample Name"] = src.get("Sample Name", "")
        row["Dilution Factor"] = src.get("Dilution Factor", "")
        if include_type:
            t = src.get("Type", "")
            row["Type"] = "" if (t is None or (isinstance(t, float) and pd.isna(t))) else str(t)
        if include_level:
            lv = src.get("Level", "")
            if lv is None or (isinstance(lv, float) and pd.isna(lv)):
                row["Level"] = ""
            else:
                # Keep numeric levels as floats so they sort & format nicely;
                # leave non-numeric strings (e.g. blank-but-non-NaN) as-is.
                lv_num = pd.to_numeric(lv, errors="coerce")
                row["Level"] = float(lv_num) if pd.notna(lv_num) else str(lv)
        if type_label == "STD":
            known = pd.to_numeric(src.get("STD_Known_Conc"), errors="coerce")
        elif type_label == "CVS":
            known = pd.to_numeric(src.get("CVS_Known_Conc"), errors="coerce")
        else:
            known = np.nan
        if pd.notna(known):
            row["Known Conc"] = float(known)
        for analyte, raw_col, corr_col in pairs:
            raw_val = pd.to_numeric(src.get(raw_col), errors="coerce") if raw_col else np.nan
            corr_val = pd.to_numeric(src.get(corr_col), errors="coerce") if corr_col else np.nan
            if pd.notna(raw_val):
                row[f"{analyte} [{unit}]"] = float(raw_val)
            if pd.notna(corr_val):
                row[f"{analyte} corrected [{unit}]"] = float(corr_val)
        # Per-row meta — what the styler will use to decide pass/fail.
        sample_name = str(src.get("Sample Name", ""))
        meta = {
            "type": type_label,
            "name": sample_name,
            "known": float(known) if pd.notna(known) else None,
        }
        return row, meta

    # ── Reference row ────────────────────────────────────────
    ref = empty_row()
    ref["Sample Name"] = "⬇ LLOQ / ULOQ reference"
    for analyte, _raw_col, _corr_col in pairs:
        lloq = lloq_map.get(analyte)
        uloq = uloq_map.get(analyte)
        lbl = (
            f"LLOQ={lloq if lloq is not None else 'n/a'} | "
            f"ULOQ={uloq if uloq is not None else 'n/a'}"
        )
        ref[f"{analyte} [{unit}]"] = lbl
        ref[f"{analyte} corrected [{unit}]"] = lbl

    all_rows: list = [ref]
    types: list = ["ref"]
    row_meta: list = [{"type": "ref", "name": "", "known": None}]

    _type_map = {"STD": "std", "CVS": "cvs", "Blank": "blank", "Sample": "sample"}
    for _, src in tidy_df.iterrows():
        label = src.get("isStandard", "Sample")
        if label not in _type_map:
            label = "Sample"
        row, meta = make_row(src, label)
        all_rows.append(row)
        types.append(_type_map[label])
        row_meta.append(meta)

    workup_df = pd.DataFrame(all_rows, columns=cols)
    return workup_df, types, row_meta


def build_summary(
    workup_df: pd.DataFrame,
    row_types: list,
    *,
    use_dilution: bool,
    include_blanks: bool,
    include_std: bool,
    include_cvs: bool,
    unit: str,
) -> tuple:
    """
    Project Workup into the polished Summary based on user checkbox state.

    Always kept:  reference row, Sample rows.
    Optionally kept:  Blank / STD / CVS rows (by checkbox).
    Visual section separators ('sep') are always dropped from Summary.
    If use_dilution is False, the 'corrected [unit]' columns are dropped.

    Returns (summary_df, summary_row_types).
    """
    if workup_df.empty or not row_types:
        return workup_df.copy(), list(row_types)

    keep_mask = []
    for rt in row_types:
        if rt == "ref" or rt == "sample":
            keep_mask.append(True)
        elif rt == "blank":
            keep_mask.append(include_blanks)
        elif rt == "std":
            keep_mask.append(include_std)
        elif rt == "cvs":
            keep_mask.append(include_cvs)
        elif rt == "sep":
            keep_mask.append(False)
        else:
            keep_mask.append(False)

    kept_types = [rt for rt, k in zip(row_types, keep_mask) if k]
    summary_df = workup_df.loc[keep_mask].reset_index(drop=True)

    # Polished Summary drops the instrument-audit columns surfaced in Workup.
    # No-op when the parser didn't emit them (e.g. Chemstation).
    for _audit_col in ("Type", "Level"):
        if _audit_col in summary_df.columns:
            summary_df = summary_df.drop(columns=_audit_col)

    if not use_dilution:
        drop_cols = [c for c in summary_df.columns if c.endswith(f" corrected [{unit}]")]
        if drop_cols:
            summary_df = summary_df.drop(columns=drop_cols)

    return summary_df, kept_types


# ═══════════════════════════════════════════════════════════════
#  UI RENDER HELPERS  (deduplicate single-file vs batch mode)
# ═══════════════════════════════════════════════════════════════


def render_cvs_tolerance_ui(cvs_levels: list, unit: str, key_prefix: str) -> dict:
    """Render the CVS-tolerance number_inputs. Returns {level: tolerance_pct}."""
    cvs_tolerances: dict = {}
    if not cvs_levels:
        st.caption("No CVS samples detected.")
        return cvs_tolerances
    st.subheader("CVS Tolerance per Level", anchor=False)
    st.caption(f"{len(cvs_levels)} CVS level(s) detected.")
    tol_cols = st.columns(min(len(cvs_levels), 6))
    for i, lvl in enumerate(cvs_levels):
        with tol_cols[i % 6]:
            cvs_tolerances[lvl] = st.number_input(
                f"CVS {lvl} {unit}  ±%",
                min_value=1,
                max_value=50,
                value=10,
                step=1,
                key=f"{key_prefix}cvs_tol_{lvl}",
            )
    return cvs_tolerances


def render_std_qc_bounds(key_prefix: str) -> tuple:
    """STD recovery pass-window inputs. Returns (std_lo, std_hi)."""
    st.subheader("STD QC Bounds", anchor=False)
    st.caption("Set the acceptable recovery range for STD samples.")
    c1, c2 = st.columns(2)
    with c1:
        std_lo = st.number_input(
            "STD Recovery Lower (%)", value=90, step=1, key=f"{key_prefix}std_lo"
        )
    with c2:
        std_hi = st.number_input(
            "STD Recovery Upper (%)", value=110, step=1, key=f"{key_prefix}std_hi"
        )
    return std_lo, std_hi


def render_inclusion_checkboxes(key_prefix: str) -> tuple:
    """The four Summary-inclusion toggles. Returns (use_dil, inc_blanks, inc_std, inc_cvs)."""
    st.subheader("Summary inclusion options", anchor=False)
    st.caption(
        "These toggles shape the **Summary** dataset (a polished projection "
        "of Workup). They don't affect the Workup view — Workup always shows "
        "everything."
    )
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        use_dil = st.checkbox(
            "Include dilution workup in Summary", value=True, key=f"{key_prefix}use_dil"
        )
    with c2:
        inc_blanks = st.checkbox(
            "Include blank injections in Summary", value=False, key=f"{key_prefix}inc_blanks"
        )
    with c3:
        inc_std = st.checkbox(
            "Include STD assessment in Summary", value=False, key=f"{key_prefix}inc_std"
        )
    with c4:
        inc_cvs = st.checkbox(
            "Include CVS assessment in Summary", value=False, key=f"{key_prefix}inc_cvs"
        )
    return use_dil, inc_blanks, inc_std, inc_cvs


def render_cvs_analyte_assignment_editor(
    cvs_names: list,
    analytes: list,
    key_prefix: str,
    *,
    name_to_conc: dict | None = None,
    unit: str = "mg/L",
) -> None:
    """Render ONLY the analyte/CVS checkbox table.

    The caller is expected to wrap this in `st.form(...)` together with a
    `st.form_submit_button(...)`. That submit button is what commits the
    user's edits via `commit_cvs_assignment(key_prefix)`. Without the form
    wrapper every checkbox click would trigger a rerun and the heatmap +
    Workup styling would flicker (and in practice Streamlit's reactive
    cycle racing the data_editor's pending widget state caused boxes to
    re-check themselves) — the form is the fix.

    Two display modes:

      ▸ **Per-injection** (default): one column per unique CVS base name.
        Used by Chemstation runs where each CVS name encodes a different
        analyte/level combination (e.g. tpa_cvs_50ppm vs hdo_aro_cvs_2ppm).

      ▸ **Per-concentration**: one column per unique CVS concentration.
        Activates automatically when `name_to_conc` is provided AND the
        number of unique concentrations is strictly less than the number
        of unique CVS base names (i.e. multiple injections share a level —
        the MassHunter pattern, where DKmix_cvs1_5ppm, DKmix_cvs2_5ppm,
        … all belong together at 5 mg/L). One checkbox per analyte covers
        every injection at that concentration; the fan-out happens in
        `get_applied_cvs_assignment`.

    Default state: each analyte is checked for a CVS column iff the
    analyte's name (case-insensitive substring) appears in any underlying
    CVS base name attached to that column.

    State lives in three session-state slots:
      `{key_prefix}cvs_assign_draft_df`   — live editor contents
      `{key_prefix}cvs_assign_applied_df` — committed version (drives QC)
      `{key_prefix}cvs_assign_grouping`   — {column_label: [base_names]}
                                            used to fan grouped choices
                                            back out to individual CVS
                                            injections at commit time
    All three get reset whenever the CVS list, analyte list, or grouping
    mode changes (i.e. a new file is uploaded).
    """
    # Strip rep suffixes, dedup preserving first-seen order
    base_names: list = []
    seen = set()
    for n in cvs_names:
        b = strip_rep_suffix(n)
        if b not in seen:
            seen.add(b)
            base_names.append(b)
    if not base_names or not analytes:
        return

    # ── Decide between per-injection and per-concentration display ──
    # Build base_name → conc map (drop NaN / None entries).
    base_to_conc: dict = {}
    if name_to_conc:
        for b in base_names:
            # name_to_conc may be keyed by full sample name OR base name;
            # try both, preferring the base lookup.
            v = name_to_conc.get(b)
            if v is None:
                # Fall back to any rep variant whose strip matches this base
                for raw_name, c in name_to_conc.items():
                    if strip_rep_suffix(raw_name) == b:
                        v = c
                        break
            if v is not None and pd.notna(v):
                base_to_conc[b] = float(v)

    unique_concs = sorted(set(base_to_conc.values())) if base_to_conc else []
    use_grouping = len(base_to_conc) == len(base_names) and 0 < len(  # every base has a conc
        unique_concs
    ) < len(
        base_names
    )  # grouping is reductive

    if use_grouping:
        # column_label = "5 mg/L"; grouping_map = {label: [base_names]}
        column_labels = [f"{c:g} {unit}" for c in unique_concs]
        grouping_map: dict = {f"{c:g} {unit}": [] for c in unique_concs}
        for b in base_names:
            grouping_map[f"{base_to_conc[b]:g} {unit}"].append(b)
    else:
        column_labels = base_names
        grouping_map = {b: [b] for b in base_names}

    draft_key = f"{key_prefix}cvs_assign_draft_df"
    applied_key = f"{key_prefix}cvs_assign_applied_df"
    grouping_key = f"{key_prefix}cvs_assign_grouping"
    version_key = f"{key_prefix}cvs_editor_version"

    def _build_default() -> pd.DataFrame:
        rows = []
        for a in analytes:
            a_lc = a.lower()
            row = {"Analyte": a}
            for label in column_labels:
                # Checked if analyte name appears in ANY base name grouped here
                row[label] = any(a_lc in base.lower() for base in grouping_map[label])
            rows.append(row)
        return pd.DataFrame(rows)

    # Reset all three slots whenever the table layout changes.
    existing = st.session_state.get(draft_key)
    existing_grouping = st.session_state.get(grouping_key)
    needs_reset = (
        existing is None
        or list(existing["Analyte"]) != analytes
        or [c for c in existing.columns if c != "Analyte"] != column_labels
        or existing_grouping != grouping_map
    )
    if needs_reset:
        default_df = _build_default()
        st.session_state[draft_key] = default_df
        st.session_state[applied_key] = default_df.copy()
        st.session_state[grouping_key] = grouping_map
        # Bump editor version so the data_editor widget reinitializes from
        # the freshly-built default instead of clinging to internal state
        # from a previous file's edits.
        st.session_state[version_key] = st.session_state.get(version_key, 0) + 1

    # ── Column config — per-column help text in grouped mode shows the
    # underlying injections so the user knows what each column actually
    # covers.
    column_config = {"Analyte": st.column_config.TextColumn(disabled=True)}
    for label in column_labels:
        underlying = grouping_map[label]
        if use_grouping and len(underlying) > 1:
            help_txt = (
                "Applies to " + str(len(underlying)) + " injection(s): " + ", ".join(underlying)
            )
        else:
            help_txt = None
        column_config[label] = st.column_config.CheckboxColumn(help=help_txt)

    # Version-bumped key forces the widget to reinitialize on bulk actions
    # (Check All / Clear All / file change), bypassing the data_editor's
    # internal state cache that would otherwise stick to the prior value.
    version = st.session_state.get(version_key, 0)
    edited = st.data_editor(
        st.session_state[draft_key],
        key=f"{key_prefix}cvs_assign_editor_v{version}",
        column_config=column_config,
        hide_index=True,
        use_container_width=True,
    )
    # Track live edits; downstream consumers read from the applied slot, not this.
    st.session_state[draft_key] = edited


def set_all_cvs_assignment(key_prefix: str, value: bool) -> None:
    """Bulk-set every checkbox in the draft to ``value``.

    Bumps the editor version so the data_editor widget reinitializes from
    the new draft on next render — without the bump, Streamlit's internal
    widget cache keyed on the editor's `key=` would shadow our changes and
    show stale checkboxes.
    """
    draft_key = f"{key_prefix}cvs_assign_draft_df"
    version_key = f"{key_prefix}cvs_editor_version"
    df = st.session_state.get(draft_key)
    if df is None:
        return
    df = df.copy()
    for c in df.columns:
        if c != "Analyte":
            df[c] = value
    st.session_state[draft_key] = df
    st.session_state[version_key] = st.session_state.get(version_key, 0) + 1


def render_cvs_bulk_action_buttons(key_prefix: str) -> None:
    """Render Check-All / Clear-All buttons. Call OUTSIDE the form wrapper.

    Updates the DRAFT only — the user still clicks "🔄 Refresh QC" inside
    the form to commit. This matches the manual-edit workflow: edits live
    in the draft until explicitly committed.
    """
    cA, cB, _ = st.columns([1, 1, 3])
    with cA:
        if st.button(
            "☑️ Check all analytes",
            key=f"{key_prefix}cvs_check_all",
            help="Tick every checkbox in the table below. "
            "Click 'Refresh QC' afterwards to commit.",
        ):
            set_all_cvs_assignment(key_prefix, True)
            st.rerun()
    with cB:
        if st.button(
            "☐ Clear all",
            key=f"{key_prefix}cvs_clear_all",
            help="Untick every checkbox. Click 'Refresh QC' to commit.",
        ):
            set_all_cvs_assignment(key_prefix, False)
            st.rerun()


def commit_cvs_assignment(key_prefix: str) -> None:
    """Copy draft → applied for the assignment table at this key_prefix.
    Called by the form-wrapping caller when its submit button fires."""
    draft_key = f"{key_prefix}cvs_assign_draft_df"
    applied_key = f"{key_prefix}cvs_assign_applied_df"
    draft = st.session_state.get(draft_key)
    if draft is not None:
        st.session_state[applied_key] = draft.copy()


def get_applied_cvs_assignment(key_prefix: str) -> dict:
    """Return the applied {cvs_base_name: {analyte: bool}} dict for downstream QC.

    When the editor ran in per-concentration grouped mode, each user-facing
    column covers multiple underlying CVS base names. We fan the per-column
    choices out so the returned dict ALWAYS keys on individual base names —
    that's what the QC code below us expects.
    """
    applied_key = f"{key_prefix}cvs_assign_applied_df"
    grouping_key = f"{key_prefix}cvs_assign_grouping"
    df = st.session_state.get(applied_key)
    if df is None or len(df) == 0 or "Analyte" not in df.columns:
        return {}
    grouping_map = st.session_state.get(grouping_key) or {}
    column_labels = [c for c in df.columns if c != "Analyte"]
    # Resolve every column_label to its list of base names. If the grouping
    # map is missing (older session), assume identity (label IS the base).
    out: dict = {}
    for label in column_labels:
        for base in grouping_map.get(label, [label]):
            out.setdefault(base, {})
    for _, row in df.iterrows():
        analyte = str(row["Analyte"])
        for label in column_labels:
            val = bool(row[label])
            for base in grouping_map.get(label, [label]):
                out[base][analyte] = val
    return out


def render_refresh_button(key_prefix: str):
    """Standalone refresh affordance. Largely redundant now that the
    assignment editor lives inside a form (its submit button drives the
    refresh), but kept here for callers that want an extra explicit
    button outside the form. The form's submit button is preferred."""
    if st.button(
        "🔄 Refresh QC",
        key=f"{key_prefix}refresh_qc",
        help="Re-render the heatmap with current tolerances and assignments.",
    ):
        st.rerun()


def memoized_build_workup(
    cache_key: str,
    session_key: str,
    tidy_df,
    lloq_map,
    uloq_map,
    unit,
    std_lo,
    std_hi,
    cvs_tolerances,
    analyte_assignment=None,
) -> tuple:
    """Wrap build_workup with session-state caching keyed on inputs.

    Workup is independent of the four Summary checkboxes, so we don't want
    to rebuild it every time the user toggles one. cache_key should hash
    everything Workup actually depends on; session_key is the storage slot.
    """
    if st.session_state.get(f"{session_key}_key") != cache_key:
        wk_df, wk_types, wk_meta = build_workup(
            tidy_df,
            lloq_map,
            uloq_map,
            unit,
            std_lo=std_lo,
            std_hi=std_hi,
            cvs_tolerances=cvs_tolerances,
            analyte_assignment=analyte_assignment,
        )
        st.session_state[f"{session_key}_cache"] = (wk_df, wk_types, wk_meta)
        st.session_state[f"{session_key}_key"] = cache_key
    return st.session_state[f"{session_key}_cache"]


def render_workup_table(
    workup_df: pd.DataFrame,
    row_types: list,
    unit: str,
    row_meta: list | None = None,
    std_lo: float = 90,
    std_hi: float = 110,
    cvs_tolerances: dict | None = None,
    analyte_assignment: dict | None = None,
    lloq_map: dict | None = None,
    uloq_map: dict | None = None,
):
    """Render workup_df with tolerance-based cell coloring computed inline.

    row_meta is the parallel list from build_workup giving each row's
    type, name, and known concentration. The styler combines that with
    the same std_lo/std_hi/cvs_tolerances/analyte_assignment/LLOQ-ULOQ
    the heatmap uses, so a value cell's color always tracks the QC inputs
    above. No '<analyte> QC' columns are needed in the dataframe itself.
    """
    if workup_df.empty or not row_types:
        st.caption("No data to render.")
        return

    cvs_tolerances = cvs_tolerances or {}
    analyte_assignment = analyte_assignment or {}
    lloq_map = lloq_map or {}
    uloq_map = uloq_map or {}
    if row_meta is None or len(row_meta) != len(workup_df):
        row_meta = [{"type": "sep", "name": "", "known": None}] * len(workup_df)

    # Map each value column to its analyte (and whether it's the corrected variant).
    # Used to (a) find the cell's "ref value" for QC math, and (b) decide whether
    # to look up corrected or raw values.
    value_cols: list = []  # (col_name, analyte, is_corrected)
    for col in workup_df.columns:
        if not col.endswith(f" [{unit}]"):
            continue
        stripped = col[: -(len(unit) + 3)]
        is_corrected = stripped.endswith(" corrected")
        if is_corrected:
            stripped = stripped[: -len(" corrected")]
        value_cols.append((col, stripped.strip(), is_corrected))

    # For each row, decide the reference value per analyte (corrected if present,
    # else raw). We do this once up-front to keep style_row cheap.
    ref_value_by_row_analyte: dict = {}  # (row_idx, analyte) -> float
    raw_vals_by_analyte: dict = {}  # analyte -> col
    corr_vals_by_analyte: dict = {}
    for col, analyte, is_corrected in value_cols:
        if is_corrected:
            corr_vals_by_analyte[analyte] = col
        else:
            raw_vals_by_analyte[analyte] = col
    for analyte in set(list(raw_vals_by_analyte) + list(corr_vals_by_analyte)):
        corr_col = corr_vals_by_analyte.get(analyte)
        raw_col = raw_vals_by_analyte.get(analyte)
        for idx in range(len(workup_df)):
            v = np.nan
            if corr_col is not None:
                v_c = workup_df.iat[idx, workup_df.columns.get_loc(corr_col)]
                if isinstance(v_c, (int, float)) and pd.notna(v_c):
                    v = float(v_c)
            if (not isinstance(v, float) or pd.isna(v)) and raw_col is not None:
                v_r = workup_df.iat[idx, workup_df.columns.get_loc(raw_col)]
                if isinstance(v_r, (int, float)) and pd.notna(v_r):
                    v = float(v_r)
            ref_value_by_row_analyte[(idx, analyte)] = v

    REF_STYLE = "background-color: #fff2cc; color: #7f6000; font-weight: bold; font-style: italic"
    SEP_STYLE = (
        "background-color: #e9ecef; border-top: 1px solid #adb5bd; border-bottom: 1px solid #adb5bd"
    )
    BLANK_INJ_STYLE = "background-color: #e8f4f8; color: #1f4e5f"
    _CSS_BY_KEY = {
        "pass": "background-color: #d4edda; color: #155724",
        "fail": "background-color: #f8d7da; color: #721c24",
        "lloq": "background-color: #fff3cd; color: #856404",
        "neutral": "",
    }

    def _decide_value_style(row_idx: int, col: str, analyte: str) -> str:
        meta = row_meta[row_idx]
        decision = decide_cell_style(
            rtype=meta.get("type", ""),
            sample_name=meta.get("name", ""),
            known=meta.get("known"),
            ref_val=ref_value_by_row_analyte.get((row_idx, analyte), np.nan),
            analyte=analyte,
            std_lo=std_lo,
            std_hi=std_hi,
            cvs_tolerances=cvs_tolerances,
            analyte_assignment=analyte_assignment,
            lloq_map=lloq_map,
            uloq_map=uloq_map,
        )
        return _CSS_BY_KEY.get(decision, "")

    value_col_to_analyte = {c: a for c, a, _ in value_cols}

    def style_row(idx: int) -> list:
        rtype = row_types[idx]
        n = len(workup_df.columns)
        if rtype == "ref":
            return [REF_STYLE] * n
        if rtype == "sep":
            return [SEP_STYLE] * n
        if rtype == "blank":
            return [BLANK_INJ_STYLE] * n
        styles = [""] * n
        for col_idx, col in enumerate(workup_df.columns):
            analyte = value_col_to_analyte.get(col)
            if analyte:
                styles[col_idx] = _decide_value_style(idx, col, analyte)
        return styles

    def _styler(_df):
        return pd.DataFrame(
            [style_row(i) for i in range(len(workup_df))],
            index=workup_df.index,
            columns=workup_df.columns,
        )

    def _fmt_value(v):
        if isinstance(v, bool):
            return v
        if isinstance(v, (int, float)) and pd.notna(v):
            return f"{v:.3f}"
        return v if v is not None else ""

    fmt_targets = [c for c, _, _ in value_cols]
    if "Known Conc" in workup_df.columns:
        fmt_targets.append("Known Conc")
    styled = workup_df.style.apply(_styler, axis=None).format(
        _fmt_value, subset=fmt_targets, na_rep=""
    )

    row_height = min(len(workup_df) * 35 + 60, 800)
    st.dataframe(styled, use_container_width=True, hide_index=True, height=row_height)


# ═══════════════════════════════════════════════════════════════
#  HEATMAP VISUALIZATION HELPERS
# ═══════════════════════════════════════════════════════════════


def render_qc_heatmap(
    std_df: pd.DataFrame, cvs_df: pd.DataFrame, std_lo: float, std_hi: float, cvs_tolerances: dict
):
    """Render a single stacked QC heatmap: STD rows on top, CVS rows below."""
    has_std = not std_df.empty and "% Recovery" in std_df.columns and "Analyte" in std_df.columns
    has_cvs = not cvs_df.empty and "% Recovery" in cvs_df.columns and "Analyte" in cvs_df.columns

    if not has_std and not has_cvs:
        st.caption("No QC data to visualize.")
        return

    # ── Pivot each type ──────────────────────────────────────
    std_pivot = pd.DataFrame()
    if has_std:
        std_pivot = std_df.pivot_table(
            index="Sample Name", columns="Analyte", values="% Recovery", aggfunc="first"
        )

    cvs_pivot = pd.DataFrame()
    cvs_tol_lookup = {}
    if has_cvs:
        cvs_pivot = cvs_df.pivot_table(
            index="Sample Name", columns="Analyte", values="% Recovery", aggfunc="first"
        )
        for _, row in cvs_df.iterrows():
            sn = row.get("Sample Name")
            lvl = row.get("CVS Level")
            if pd.notna(lvl):
                cvs_tol_lookup[sn] = cvs_tolerances.get(float(lvl), 10.0)

    if std_pivot.empty and cvs_pivot.empty:
        st.caption("No QC data to visualize.")
        return

    # ── Unify columns and stack with a divider row ───────────
    all_analytes = sorted(set(std_pivot.columns.tolist()) | set(cvs_pivot.columns.tolist()))
    std_pivot = std_pivot.reindex(columns=all_analytes)
    cvs_pivot = cvs_pivot.reindex(columns=all_analytes)

    # Build a divider row (NaN values, styled differently)
    divider_idx = "── CVS ──" if not std_pivot.empty else ""
    parts = []
    std_rows = set()
    cvs_rows = set()

    if not std_pivot.empty:
        std_pivot.index = ["[STD] " + str(n) for n in std_pivot.index]
        std_rows = set(std_pivot.index)
        parts.append(std_pivot)

    if not cvs_pivot.empty:
        cvs_pivot.index = ["[CVS] " + str(n) for n in cvs_pivot.index]
        cvs_rows = set(cvs_pivot.index)
        if not std_pivot.empty:
            # Insert a visual divider row between STD and CVS sections
            divider = pd.DataFrame(
                [[np.nan] * len(all_analytes)], columns=all_analytes, index=["─── ─── ───"]
            )
            parts.append(divider)
        parts.append(cvs_pivot)

    combined = pd.concat(parts)
    divider_rows = {"─── ─── ───"}

    # ── Styling function ─────────────────────────────────────
    def _style_cell(val, row_name, col_name):
        if row_name in divider_rows:
            return "background-color: #e9ecef; color: #e9ecef; font-size:1px"
        if pd.isna(val):
            return "background-color: #f0f0f0; color: #999"
        if row_name in std_rows:
            passed = std_lo <= val <= std_hi
        elif row_name in cvs_rows:
            # Strip "[CVS] " prefix to look up tolerance
            orig_name = row_name.replace("[CVS] ", "", 1)
            tol = cvs_tol_lookup.get(orig_name, 10.0)
            lo, hi = 100 - tol, 100 + tol
            passed = lo <= val <= hi
        else:
            passed = True
        if passed:
            return "background-color: #d4edda; color: #155724"
        return "background-color: #f8d7da; color: #721c24"

    def _apply_styles(df):
        return pd.DataFrame(
            [
                [
                    _style_cell(df.iloc[r, c], df.index[r], df.columns[c])
                    for c in range(len(df.columns))
                ]
                for r in range(len(df))
            ],
            index=df.index,
            columns=df.columns,
        )

    styled = combined.style.apply(lambda _: _apply_styles(combined), axis=None).format(
        "{:.1f}%", na_rep="—"
    )
    total_rows = len(combined)
    row_height = min(total_rows * 40 + 50, 500)

    legend = f"STD pass: {std_lo}–{std_hi}%"
    if has_cvs:
        legend += " &nbsp;|&nbsp; CVS pass: per-level tolerance"
    st.markdown(f"**QC Recovery Heatmap** &nbsp; _({legend})_")
    st.dataframe(styled, use_container_width=True, height=row_height)


# ═══════════════════════════════════════════════════════════════
#  EXCEL BUILD  (lazy — bytes stored in session_state)
# ═══════════════════════════════════════════════════════════════


def build_excel(
    tidy_df,
    raw_df,
    workup_df,
    workup_row_types,
    summary_df,
    summary_row_types,
    meta: dict,
    submission_wb_bytes=None,
    qc_context: dict | None = None,
) -> bytes:
    """
    Emit the deliverable workbook.

    Sheets (in order, after any pre-existing submission sheets):
        Workup            — full review table: STD + CVS + Blank + Sample with
                            section separators AND per-cell pass/fail coloring
                            that matches the Streamlit Workup view.
        Summary           — polished output filtered by the user's checkboxes
                            (dilution / blanks / STD / CVS). LLOQ/ULOQ ref row
                            preserved at top. Same per-cell coloring as Workup.
        Tidy Data         — full merged tidy frame (every parsed row)
        Raw Data          — original instrument export
        Method Information — flat key/value metadata block (every entry in meta)

    qc_context (optional): dict with keys
        unit, std_lo, std_hi, cvs_tolerances, analyte_assignment,
        lloq_map, uloq_map, workup_row_meta, summary_row_meta
      When provided, value cells in Workup and Summary get colored exactly
      as the Streamlit styler would (via shared decide_cell_style helper).
      Without it, only row-level tints (ref / sep / blank) are applied.
    """
    output = BytesIO()

    REF_FG, SEP_FG, BLANK_FG = "FFF2CC", "E9ECEF", "E8F4F8"
    PASS_FG, FAIL_FG, LLOQ_FG = "D4EDDA", "F8D7DA", "FFF3CD"
    PASS_FONT, FAIL_FONT, LLOQ_FONT = "155724", "721C24", "856404"

    def _df_to_ws(wb, df, name):
        ws = wb.create_sheet(title=name)
        ws.append([str(c) for c in df.columns.tolist()])
        for row in df.itertuples(index=False):
            ws.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
        return ws

    def _build_value_col_index(df: pd.DataFrame, unit: str):
        """Return list of (col_idx, analyte, is_corrected) for value cols."""
        out = []
        suffix = f" [{unit}]"
        for c_idx, col in enumerate(df.columns):
            if not isinstance(col, str) or not col.endswith(suffix):
                continue
            stripped = col[: -len(suffix)]
            is_corr = stripped.endswith(" corrected")
            if is_corr:
                stripped = stripped[: -len(" corrected")]
            out.append((c_idx, stripped.strip(), is_corr))
        return out

    def _row_ref_value(df, row_idx, value_cols_for_analyte):
        """Pick corrected value if present, else raw, for a given analyte's
        column-index pair. value_cols_for_analyte = (raw_idx, corr_idx)."""
        raw_idx, corr_idx = value_cols_for_analyte
        for idx in (corr_idx, raw_idx):
            if idx is None:
                continue
            v = df.iat[row_idx, idx]
            try:
                f = float(v)
                if not np.isnan(f):
                    return f
            except (TypeError, ValueError):
                pass
        return float("nan")

    def _color_value_cells(
        ws,
        df,
        row_types,
        row_meta,
        ctx_unit,
        ctx_std_lo,
        ctx_std_hi,
        ctx_cvs_tol,
        ctx_assignment,
        ctx_lloq,
        ctx_uloq,
    ):
        """Paint per-cell coloring on value columns based on decide_cell_style."""
        if row_meta is None or len(row_meta) != len(df):
            return
        value_col_index = _build_value_col_index(df, ctx_unit)
        # Group by analyte → (raw_col_idx, corr_col_idx)
        by_analyte: dict = {}
        for ci, analyte, is_corr in value_col_index:
            slot = by_analyte.setdefault(analyte, [None, None])
            slot[1 if is_corr else 0] = ci

        fill_pass = openpyxl.styles.PatternFill("solid", fgColor=PASS_FG)
        fill_fail = openpyxl.styles.PatternFill("solid", fgColor=FAIL_FG)
        fill_lloq = openpyxl.styles.PatternFill("solid", fgColor=LLOQ_FG)
        font_pass = openpyxl.styles.Font(color=PASS_FONT)
        font_fail = openpyxl.styles.Font(color=FAIL_FONT)
        font_lloq = openpyxl.styles.Font(color=LLOQ_FONT)

        rtype_to_label = {
            "std": "STD",
            "cvs": "CVS",
            "sample": "Sample",
            "blank": "Blank",
            "ref": "ref",
            "sep": "sep",
        }

        for i, rtype in enumerate(row_types):
            if rtype in ("ref", "sep", "blank"):
                continue  # row-level fills handled separately
            xl_row = i + 2  # +1 header, +1 1-based
            meta_i = row_meta[i] if i < len(row_meta) else {}
            normalized_type = rtype_to_label.get(rtype, "Sample")
            name = meta_i.get("name", "")
            known = meta_i.get("known")
            for analyte, (raw_idx, corr_idx) in by_analyte.items():
                ref_val = _row_ref_value(df, i, (raw_idx, corr_idx))
                decision = decide_cell_style(
                    rtype=normalized_type,
                    sample_name=name,
                    known=known,
                    ref_val=ref_val,
                    analyte=analyte,
                    std_lo=ctx_std_lo,
                    std_hi=ctx_std_hi,
                    cvs_tolerances=ctx_cvs_tol,
                    analyte_assignment=ctx_assignment,
                    lloq_map=ctx_lloq,
                    uloq_map=ctx_uloq,
                )
                if decision == "neutral":
                    continue
                fill, font = {
                    "pass": (fill_pass, font_pass),
                    "fail": (fill_fail, font_fail),
                    "lloq": (fill_lloq, font_lloq),
                }[decision]
                # Apply to both raw and corrected columns for this analyte
                for ci in (raw_idx, corr_idx):
                    if ci is None:
                        continue
                    cell = ws.cell(row=xl_row, column=ci + 1)
                    cell.fill = fill
                    # Don't squash bold/italic on header/ref already applied — these rows are skipped above
                    cell.font = font

    def _row_tint_workup(ws, row_types):
        ref_fill = openpyxl.styles.PatternFill("solid", fgColor=REF_FG)
        ref_font = openpyxl.styles.Font(bold=True, italic=True, color="7F6000")
        sep_fill = openpyxl.styles.PatternFill("solid", fgColor=SEP_FG)
        blank_fill = openpyxl.styles.PatternFill("solid", fgColor=BLANK_FG)
        for i, rtype in enumerate(row_types):
            xl_row = i + 2
            if rtype == "ref":
                for cell in ws[xl_row]:
                    cell.fill = ref_fill
                    cell.font = ref_font
            elif rtype == "sep":
                for cell in ws[xl_row]:
                    cell.fill = sep_fill
            elif rtype == "blank":
                for cell in ws[xl_row]:
                    cell.fill = blank_fill

    def _row_tint_summary(ws, row_types):
        if not row_types:
            return
        ref_fill = openpyxl.styles.PatternFill("solid", fgColor=REF_FG)
        ref_font = openpyxl.styles.Font(bold=True, italic=True, color="7F6000")
        blank_fill = openpyxl.styles.PatternFill("solid", fgColor=BLANK_FG)
        for i, rtype in enumerate(row_types):
            xl_row = i + 2
            if rtype == "ref":
                for cell in ws[xl_row]:
                    cell.fill = ref_fill
                    cell.font = ref_font
            elif rtype == "blank":
                for cell in ws[xl_row]:
                    cell.fill = blank_fill

    # ── Assemble workbook ────────────────────────────────────
    if submission_wb_bytes:
        wb_out = openpyxl.load_workbook(BytesIO(submission_wb_bytes))
        for sh in [
            "Summary",
            "Workup",
            "Tidy Data",
            "STD Workup",
            "CVS Workup",
            "Raw Data",
            "Method Information",
        ]:
            if sh in wb_out.sheetnames:
                del wb_out[sh]
    else:
        wb_out = openpyxl.Workbook()
        if "Sheet" in wb_out.sheetnames:
            del wb_out["Sheet"]

    ws_wk = _df_to_ws(wb_out, workup_df, "Workup")
    ws_sum = _df_to_ws(wb_out, summary_df, "Summary")
    _df_to_ws(wb_out, tidy_df, "Tidy Data")
    _df_to_ws(wb_out, raw_df, "Raw Data")

    # ── Per-cell coloring first, then row-level tints (which include
    #    the value cells in their rows — row tints "win" because they're
    #    applied AFTER per-cell coloring). The order keeps blank/sep/ref
    #    rows uniformly tinted while non-special rows get value-cell color.
    if qc_context:
        _color_value_cells(
            ws_wk,
            workup_df,
            workup_row_types,
            qc_context.get("workup_row_meta"),
            qc_context.get("unit", ""),
            qc_context.get("std_lo", 90),
            qc_context.get("std_hi", 110),
            qc_context.get("cvs_tolerances"),
            qc_context.get("analyte_assignment"),
            qc_context.get("lloq_map"),
            qc_context.get("uloq_map"),
        )
        _color_value_cells(
            ws_sum,
            summary_df,
            summary_row_types,
            qc_context.get("summary_row_meta"),
            qc_context.get("unit", ""),
            qc_context.get("std_lo", 90),
            qc_context.get("std_hi", 110),
            qc_context.get("cvs_tolerances"),
            qc_context.get("analyte_assignment"),
            qc_context.get("lloq_map"),
            qc_context.get("uloq_map"),
        )
    _row_tint_workup(ws_wk, workup_row_types)
    _row_tint_summary(ws_sum, summary_row_types)

    # Column widths — apply uniformly. openpyxl uses column_dimensions.
    for ws in (ws_wk, ws_sum):
        for col_letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            ws.column_dimensions[col_letter].width = 22
    for sh_name in ("Tidy Data", "Raw Data"):
        if sh_name in wb_out.sheetnames:
            ws_x = wb_out[sh_name]
            for col_letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
                ws_x.column_dimensions[col_letter].width = 20

    # ── Method Information sheet — every metadata field, never trimmed ──
    ws_meta = wb_out.create_sheet(title="Method Information")
    ws_meta.append(["Field", "Value"])
    for k, v in (meta or {}).items():
        ws_meta.append([str(k), "" if v is None else str(v)])
    ws_meta.column_dimensions["A"].width = 30
    ws_meta.column_dimensions["B"].width = 50

    # Sheet order
    order = [
        s for s in ["Submission Information", "Submission Sample Table"] if s in wb_out.sheetnames
    ]
    order += [
        s
        for s in ["Workup", "Summary", "Tidy Data", "Raw Data", "Method Information"]
        if s in wb_out.sheetnames
    ]
    order += [s for s in wb_out.sheetnames if s not in order]
    wb_out._sheets = [wb_out[s] for s in order]
    wb_out.save(output)
    return output.getvalue()


def _sanitize_sheet_name(name: str, used: set | None = None) -> str:
    """Excel sheet names: ≤31 chars, no `: \\ / ? * [ ]`, must be unique."""
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


def build_batch_excel(
    per_file: list,
    combined_tidy: pd.DataFrame,
    per_file_meta: dict,
    submission_wb_bytes=None,
    qc_context_by_file: dict | None = None,
) -> bytes:
    """
    Batch master Excel — each source file becomes its own trio of sheets
    (Workup + Summary + Raw Data), then the Tidy / Method Information
    sheets are shared (one row per source file in the methods grid;
    one combined sheet with a Source File column for tidy).

    per_file:    list of dicts, each with keys:
                   "name"           — source filename
                   "workup_df"      — that file's Workup DataFrame
                   "workup_types"   — row_types parallel list
                   "summary_df"     — that file's Summary DataFrame
                   "summary_types"  — row_types parallel list
                   "raw_df"         — that file's raw parsed DataFrame
                                       (pre-cleaning, exactly as the
                                       parser emitted it)
    per_file_meta: {filename: {field: value, ...}} — drives the
                   Method Information sheet (one column per file).
    qc_context_by_file (optional): {filename: qc_context_dict} — when given,
                   each file's Workup/Summary value cells get per-cell QC
                   coloring matching the Streamlit view.
    """
    output = BytesIO()
    REF_FG, SEP_FG, BLANK_FG = "FFF2CC", "E9ECEF", "E8F4F8"
    PASS_FG, FAIL_FG, LLOQ_FG = "D4EDDA", "F8D7DA", "FFF3CD"
    PASS_FONT, FAIL_FONT, LLOQ_FONT = "155724", "721C24", "856404"
    used_names: set = set()
    qc_context_by_file = qc_context_by_file or {}

    def _tint_ws(ws, types):
        ref_fill = openpyxl.styles.PatternFill("solid", fgColor=REF_FG)
        ref_font = openpyxl.styles.Font(bold=True, italic=True, color="7F6000")
        sep_fill = openpyxl.styles.PatternFill("solid", fgColor=SEP_FG)
        blank_fill = openpyxl.styles.PatternFill("solid", fgColor=BLANK_FG)
        for i, rtype in enumerate(types or []):
            xl_row = i + 2
            if rtype == "ref":
                for cell in ws[xl_row]:
                    cell.fill = ref_fill
                    cell.font = ref_font
            elif rtype == "sep":
                for cell in ws[xl_row]:
                    cell.fill = sep_fill
            elif rtype == "blank":
                for cell in ws[xl_row]:
                    cell.fill = blank_fill

    def _df_to_ws(wb, df, name):
        ws = wb.create_sheet(title=name)
        ws.append([str(c) for c in df.columns.tolist()])
        for row in df.itertuples(index=False):
            ws.append([None if (isinstance(v, float) and np.isnan(v)) else v for v in row])
        return ws

    def _color_value_cells_local(ws, df, row_types, row_meta, ctx):
        """Per-cell coloring on a file's sheet using its qc_context."""
        if row_meta is None or len(row_meta) != len(df) or not ctx:
            return
        unit = ctx.get("unit", "")
        suffix = f" [{unit}]"
        # Group columns by analyte → (raw_col_idx, corr_col_idx)
        by_analyte: dict = {}
        for ci, col in enumerate(df.columns):
            if not isinstance(col, str) or not col.endswith(suffix):
                continue
            stripped = col[: -len(suffix)]
            is_corr = stripped.endswith(" corrected")
            if is_corr:
                stripped = stripped[: -len(" corrected")]
            slot = by_analyte.setdefault(stripped.strip(), [None, None])
            slot[1 if is_corr else 0] = ci

        fill_pass = openpyxl.styles.PatternFill("solid", fgColor=PASS_FG)
        fill_fail = openpyxl.styles.PatternFill("solid", fgColor=FAIL_FG)
        fill_lloq = openpyxl.styles.PatternFill("solid", fgColor=LLOQ_FG)
        font_pass = openpyxl.styles.Font(color=PASS_FONT)
        font_fail = openpyxl.styles.Font(color=FAIL_FONT)
        font_lloq = openpyxl.styles.Font(color=LLOQ_FONT)
        rtype_to_label = {
            "std": "STD",
            "cvs": "CVS",
            "sample": "Sample",
            "blank": "Blank",
            "ref": "ref",
            "sep": "sep",
        }

        for i, rtype in enumerate(row_types):
            if rtype in ("ref", "sep", "blank"):
                continue
            xl_row = i + 2
            meta_i = row_meta[i] if i < len(row_meta) else {}
            normalized_type = rtype_to_label.get(rtype, "Sample")
            name = meta_i.get("name", "")
            known = meta_i.get("known")
            for analyte, (raw_idx, corr_idx) in by_analyte.items():
                # Pick corrected first, fall back to raw
                ref_val = float("nan")
                for idx in (corr_idx, raw_idx):
                    if idx is None:
                        continue
                    try:
                        f = float(df.iat[i, idx])
                        if not np.isnan(f):
                            ref_val = f
                            break
                    except (TypeError, ValueError):
                        pass
                decision = decide_cell_style(
                    rtype=normalized_type,
                    sample_name=name,
                    known=known,
                    ref_val=ref_val,
                    analyte=analyte,
                    std_lo=ctx.get("std_lo", 90),
                    std_hi=ctx.get("std_hi", 110),
                    cvs_tolerances=ctx.get("cvs_tolerances"),
                    analyte_assignment=ctx.get("analyte_assignment"),
                    lloq_map=ctx.get("lloq_map"),
                    uloq_map=ctx.get("uloq_map"),
                )
                if decision == "neutral":
                    continue
                fill, font = {
                    "pass": (fill_pass, font_pass),
                    "fail": (fill_fail, font_fail),
                    "lloq": (fill_lloq, font_lloq),
                }[decision]
                for ci in (raw_idx, corr_idx):
                    if ci is None:
                        continue
                    cell = ws.cell(row=xl_row, column=ci + 1)
                    cell.fill = fill
                    cell.font = font

    # ── Build the workbook ────────────────────────────────────
    if submission_wb_bytes:
        wb = openpyxl.load_workbook(BytesIO(submission_wb_bytes))
        for sh in list(wb.sheetnames):
            if sh not in ("Submission Information", "Submission Sample Table"):
                del wb[sh]
    else:
        wb = openpyxl.Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

    for n in wb.sheetnames:
        used_names.add(n)

    per_file_sheet_triples: list = []
    for entry in per_file:
        stem = os.path.splitext(entry["name"])[0]
        wk_name = _sanitize_sheet_name(f"{stem} Workup", used_names)
        sm_name = _sanitize_sheet_name(f"{stem} Summary", used_names)
        raw_name = _sanitize_sheet_name(f"{stem} Raw Data", used_names)
        ws_wk = _df_to_ws(wb, entry["workup_df"], wk_name)
        ws_sm = _df_to_ws(wb, entry["summary_df"], sm_name)
        # Per-file raw — falls back gracefully to an empty sheet if absent.
        if entry.get("raw_df") is not None:
            ws_raw = _df_to_ws(wb, entry["raw_df"], raw_name)
            for col_letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
                ws_raw.column_dimensions[col_letter].width = 20
        else:
            ws_raw = wb.create_sheet(title=raw_name)
        ctx = qc_context_by_file.get(entry["name"])
        # Per-cell color first, then row tints overlay
        if ctx:
            _color_value_cells_local(
                ws_wk,
                entry["workup_df"],
                entry.get("workup_types"),
                ctx.get("workup_row_meta"),
                ctx,
            )
            _color_value_cells_local(
                ws_sm,
                entry["summary_df"],
                entry.get("summary_types"),
                ctx.get("summary_row_meta"),
                ctx,
            )
        _tint_ws(ws_wk, entry.get("workup_types"))
        _tint_ws(ws_sm, entry.get("summary_types"))
        # Column widths
        for col_letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
            ws_wk.column_dimensions[col_letter].width = 22
            ws_sm.column_dimensions[col_letter].width = 22
        per_file_sheet_triples.append((wk_name, sm_name, raw_name))

    tidy_name = _sanitize_sheet_name("Tidy Data", used_names)
    _df_to_ws(wb, combined_tidy, tidy_name)
    ws_tidy = wb[tidy_name]
    for col_letter in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
        ws_tidy.column_dimensions[col_letter].width = 20

    # ── Method Information: rows = field names, columns = files ──
    method_name = _sanitize_sheet_name("Method Information", used_names)
    ws_meta = wb.create_sheet(title=method_name)
    all_fields: list = []
    seen = set()
    file_order = [e["name"] for e in per_file]
    for fname in file_order:
        for k in (per_file_meta.get(fname) or {}).keys():
            if k not in seen:
                seen.add(k)
                all_fields.append(k)
    header = ["Field"] + [os.path.splitext(n)[0] for n in file_order]
    ws_meta.append(header)
    for field in all_fields:
        row = [field]
        for fname in file_order:
            v = (per_file_meta.get(fname) or {}).get(field, "")
            row.append("" if v is None else str(v))
        ws_meta.append(row)
    ws_meta.column_dimensions["A"].width = 30
    for col_letter in "BCDEFGHIJKLMNOPQRSTUVWXYZ":
        ws_meta.column_dimensions[col_letter].width = 35

    # ── Sheet order ────────────────────────────────────────
    order = [s for s in ("Submission Information", "Submission Sample Table") if s in wb.sheetnames]
    for wk_name, sm_name, raw_name in per_file_sheet_triples:
        order += [wk_name, sm_name, raw_name]
    order += [tidy_name, method_name]
    order += [s for s in wb.sheetnames if s not in order]
    wb._sheets = [wb[s] for s in order]

    wb.save(output)
    return output.getvalue()


# ═══════════════════════════════════════════════════════════════
#  SQLITE EXPORT
# ═══════════════════════════════════════════════════════════════


def export_sqlite(runs_meta, tidy_df: pd.DataFrame, path: str = "/tmp/analytical_data.db") -> bytes:
    """Two-table archive: ``runs`` (one row per source file) + ``tidy_data``
    (every parsed row, with a foreign key to runs).

    runs_meta : dict | list[dict]
        Single-mode call passes ONE meta dict (the META built from the form,
        ideally with 'Source File' set to the raw filename and a 'Run Date').
        Batch-mode call passes a list of META dicts — one per source file in
        the export. The dict's 'Source File' value is what we match against
        the tidy_df['Source File'] column to assign run_id values.
    tidy_df : pd.DataFrame
        Single-mode: the per-run tidy frame (no 'Source File' column needed —
        every row gets run_id=1).
        Batch-mode: the combined tidy frame with a 'Source File' column
        whose values match the 'Source File' field on each entry in
        runs_meta.

    Tables written:
        runs(run_id PK, tracker, analyst, run_date, task_*, analysis,
             batch, analytical_method, processing_method, unit,
             column_serial, instrument_name, export_type, source_file,
             ingested_at)
        tidy_data(run_id FK → runs, <every column from tidy_df>)

    Returns the raw bytes of the resulting .db file.
    """
    # Normalize to a list
    if isinstance(runs_meta, dict):
        runs_meta = [runs_meta]

    # Always write to a fresh file — no accumulation across button clicks.
    if os.path.exists(path):
        try:
            os.remove(path)
        except OSError:
            pass
    conn = sqlite3.connect(path)

    # Build the runs frame + source-file → run_id mapping
    runs_rows = []
    src_to_run_id: dict = {}
    now_iso = datetime.now().isoformat(timespec="seconds")
    for i, m in enumerate(runs_meta, start=1):
        src = (m.get("Source File") or "").strip()
        src_to_run_id[src] = i
        runs_rows.append(
            {
                "run_id": i,
                "tracker": m.get("Tracking Number", ""),
                "analyst": m.get("Analyst", ""),
                "run_date": m.get("Run Date", ""),
                "task_number": m.get("Task Number", ""),
                "task_name": m.get("Task Name", ""),
                "task_pi": m.get("Task PI", ""),
                "analysis": m.get("Analysis", ""),
                "batch": m.get("Batch", ""),
                "analytical_method": m.get("Analytical Method", ""),
                "processing_method": m.get("Processing Method", ""),
                "unit": m.get("Unit", ""),
                "column_serial": m.get("Column Serial", ""),
                "instrument_name": m.get("Instrument Name", ""),
                "export_type": m.get("Export Type", ""),
                "source_file": src,
                "ingested_at": now_iso,
            }
        )
    runs_df = pd.DataFrame(runs_rows)

    # Build tidy_data with run_id prepended
    tidy_clean = tidy_df.copy()
    if "Source File" in tidy_clean.columns and len(runs_meta) > 1:
        # Batch: map each row's Source File to its run_id (unknowns → NaN)
        tidy_clean.insert(0, "run_id", tidy_clean["Source File"].map(src_to_run_id))
    else:
        # Single: every row gets run_id=1
        tidy_clean.insert(0, "run_id", 1)

    runs_df.to_sql("runs", conn, if_exists="append", index=False)
    tidy_clean.to_sql("tidy_data", conn, if_exists="append", index=False)
    conn.close()

    with open(path, "rb") as f:
        return f.read()


# ═══════════════════════════════════════════════════════════════
#  CORE PIPELINE
# ═══════════════════════════════════════════════════════════════


def run_processing(
    file_bytes,
    inst,
    meta,
    cvs_tolerances,
    sub_meta,
    sub_samples_json,
    analyte_assignment: dict | None = None,
    std_lo: float = 90,
    std_hi: float = 110,
):
    """
    Parse + transform + merge submission. Also runs std/cvs QC workups.

    The QC workups call internally-cached helpers (cached_std_workup,
    cached_cvs_workup), so when the wrapping session-state cache in the
    caller is keyed only on file bytes (not on tolerances or assignments),
    tolerance/assignment changes still trigger reactive QC recomputation
    via @st.cache_data on those helpers — provided the caller invokes
    run_processing on every rerun. The current pattern in single-file mode
    short-circuits via _run_hash, so callers should also expose recompute_qc()
    (below) which skips parsing and just re-runs the QC piece.
    """
    raw_json, tidy_json, extracted_meta_json = CACHED_PARSERS[inst](file_bytes)
    extracted_meta = {}
    try:
        extracted_meta = json.loads(extracted_meta_json) if extracted_meta_json else {}
    except Exception:
        extracted_meta = {}
    effective_meta = {**meta}
    for k, v in extracted_meta.items():
        if v and not effective_meta.get(k):
            effective_meta[k] = v
    tidy_df = add_rep_suffix(inject_meta(safe_read_json(tidy_json), effective_meta))
    tidy_json = tidy_df.to_json()
    if sub_meta:
        tidy_json = cached_merge_submission(
            tidy_json,
            pd.Series(sub_meta).to_json(),
            sub_samples_json or "",
        )
    unit = effective_meta.get("Unit", "")
    std_json = cached_std_workup(tidy_json, unit, std_lo, std_hi)
    cvs_tol_j = pd.Series({str(k): v for k, v in cvs_tolerances.items()}).to_json()
    assign_j = json.dumps(analyte_assignment or {})
    cvs_json = cached_cvs_workup(tidy_json, unit, cvs_tol_j, assign_j)
    return {
        "raw": safe_read_json(raw_json),
        "tidy": safe_read_json(tidy_json),
        "std": safe_read_json(std_json),
        "cvs": safe_read_json(cvs_json),
        "extracted_meta": extracted_meta,
        "_tidy_json": tidy_json,  # for recompute_qc()
    }


def recompute_qc(
    tidy_json: str,
    unit: str,
    cvs_tolerances: dict,
    analyte_assignment: dict | None,
    std_lo: float,
    std_hi: float,
) -> tuple:
    """Re-run std/cvs workups without re-parsing. Used when tolerances or
    analyte-assignment change but the underlying tidy data hasn't."""
    std_json = cached_std_workup(tidy_json, unit, std_lo, std_hi)
    cvs_tol_j = pd.Series({str(k): v for k, v in cvs_tolerances.items()}).to_json()
    assign_j = json.dumps(analyte_assignment or {})
    cvs_json = cached_cvs_workup(tidy_json, unit, cvs_tol_j, assign_j)
    return safe_read_json(std_json), safe_read_json(cvs_json)


# (STD QC Bounds moved inline — rendered after data upload in both modes)


# ═══════════════════════════════════════════════════════════════
#  ① SUBMISSION FORM
# ═══════════════════════════════════════════════════════════════

st.header("1. Submission Form", divider="gray")
sub_file = st.file_uploader(
    "Upload submission form (.xlsx) — optional but recommended",
    type=["xlsx"],
    key="sub_uploader",
)

if sub_file is not None:
    # ── Performance: detect new uploads via Streamlit's stable `file_id`
    # attribute, not by re-reading and re-hashing the file bytes on every
    # rerun. On a 5 MB form the previous read+MD5 cost ~30 ms per widget
    # interaction (every checkbox click, every text input change); file_id
    # is a string comparison.
    _sub_file_id = getattr(sub_file, "file_id", None)
    # Fall back to size+name for older Streamlit (file_id added in 1.18+).
    if _sub_file_id is None:
        _sub_file_id = f"{sub_file.name}::{sub_file.size}"

    if st.session_state.get("_sub_file_id") != _sub_file_id:
        # ── Genuine new upload — do the expensive work once. ──
        sub_bytes = sub_file.getvalue()
        with st.spinner("Reading submission form…"):
            sub_meta_parsed, sub_samples_json, _wb_bytes = cached_parse_submission(sub_bytes)
        # Pre-materialize the samples DataFrame so the per-rerun render
        # never has to parse the samples JSON twice (was happening once
        # for the row-count message and again for st.dataframe).
        _samples_df = safe_read_json(sub_samples_json) if sub_samples_json else pd.DataFrame()
        st.session_state["submission_meta"] = sub_meta_parsed
        st.session_state["submission_samples"] = (
            sub_samples_json  # legacy contract for run_processing
        )
        st.session_state["submission_samples_df"] = _samples_df  # render-fast handle
        st.session_state["submission_wb"] = sub_bytes
        st.session_state["_sub_file_id"] = _sub_file_id

    # Always read from session_state for display — variables only exist inside the if block above
    _display_meta = st.session_state.get("submission_meta", {})
    _display_samples_df = st.session_state.get("submission_samples_df", pd.DataFrame())
    c1, c2 = st.columns([1, 2])
    with c1:
        st.success(
            f"Loaded — Tracker **{_display_meta.get('Tracking Number','?')}** | "
            f"{len(_display_samples_df)} samples"
        )
        st.dataframe(
            pd.DataFrame(_display_meta.items(), columns=["Field", "Value"]),
            use_container_width=True,
            hide_index=True,
            height=340,
        )
    with c2:
        st.caption("Sample Manifest")
        st.dataframe(_display_samples_df, use_container_width=True, height=340)

_sm = st.session_state.get("submission_meta", {})


# ═══════════════════════════════════════════════════════════════
#   Metadata extraction and user input
# ═══════════════════════════════════════════════════════════════


st.header("2. RawData Upload + Processing", divider="gray")
st.caption(
    "Pick the export type and drop in the raw file. The parser runs once on upload; "
    "any values it can pull from the file (Run Date, Analytical Method, …) get pushed "
    "up into the metadata form above on the next rerun."
)

_upload_c1, _upload_c2 = st.columns([1, 2])
with _upload_c1:
    instrument = st.selectbox(
        "Export Type",
        INSTRUMENT_TYPES,
        key="single_export_type",
        help="The raw-file format being uploaded — i.e., the software that produced it. "
        "Chemstation = Agilent ChemStation HPLC export (Data + Labels sheets). "
        "MassHunter = Agilent MassHunter export (Sheet1 with *Results headers).",
    )
with _upload_c2:
    _early_upload = st.file_uploader(
        "Upload raw data file (.xlsx / .xls)",
        type=["xlsx", "xls"],
        key="single_upload",
        help="Uploaded once here — used by the metadata pre-fill below and by the "
        "processing pipeline further down. No need to re-upload.",
    )

# Pre-parse the moment a file is present so Run Date / Analytical Method
# land in extracted_meta BEFORE the metadata form renders. The processing
# section below reads file_bytes / tidy_preview straight out of session_state.
if _early_upload is not None:
    _eu_bytes = _early_upload.getvalue()
    _eu_id = getattr(_early_upload, "file_id", f"{_early_upload.name}::{_early_upload.size}")
    # Re-parse only when the file OR the Export Type changes — switching
    # Chemstation ↔ MassHunter on the same upload should still re-parse.
    _cache_key = f"{_eu_id}::{instrument}"
    if st.session_state.get("_single_parse_key") != _cache_key:
        with st.spinner("Scanning file…"):
            try:
                _raw_j, _tidy_j, _em_j = CACHED_PARSERS[instrument](_eu_bytes)
            except Exception as e:
                st.error(f"Parse error — verify Export Type. Detail: {e}")
                st.stop()
        try:
            _new_extracted = json.loads(_em_j) if _em_j else {}
        except Exception:
            _new_extracted = {}
        st.session_state["_single_extracted_meta"] = _new_extracted
        st.session_state["_single_file_bytes"] = _eu_bytes
        st.session_state["_single_file_stem"] = os.path.splitext(_early_upload.name)[0]
        st.session_state["_single_tidy_preview"] = _tidy_j
        st.session_state["_single_parse_key"] = _cache_key

        # ── Push extracted values into the metadata form's widget keys.
        # Streamlit forbids writing to a widget's key AFTER the widget is
        # instantiated. §② Metadata widgets are already instantiated by the
        # time we get here, so write to *pending* keys instead. The top of
        # the script (above §②) copies pending → widget key BEFORE §②'s
        # widgets render on the next rerun. We then call st.rerun() to make
        # that next rerun happen immediately — otherwise the user sees the
        # form remain empty until they touch something else.
        _did_pending = False
        if _new_extracted.get("Run Date"):
            try:
                st.session_state["_pending_meta_run_date"] = pd.to_datetime(
                    _new_extracted["Run Date"]
                ).date()
                _did_pending = True
            except Exception:
                pass
        if _new_extracted.get("Analytical Method"):
            st.session_state["_pending_meta_analytical_method"] = str(
                _new_extracted["Analytical Method"]
            )
            _did_pending = True
        if _did_pending:
            st.rerun()

    if st.session_state.get("_single_extracted_meta"):
        with st.container():
            st.subheader("Auto-extracted from raw file", anchor=False)
            for k, v in st.session_state["_single_extracted_meta"].items():
                st.code(f"{k}: {v}", language=None)
            st.caption(
                "These values pre-fill the metadata form below and get written into "
                "the Method Information sheet. Anything you type in the form wins."
            )
else:
    # Clear stashed bytes when the user removes the upload, so a stale
    # file doesn't silently flow into processing later.
    for _k in (
        "_single_extracted_meta",
        "_single_file_bytes",
        "_single_file_stem",
        "_single_tidy_preview",
        "_single_parse_key",
    ):
        st.session_state.pop(_k, None)


# Mode selector & per-mode flow continue the same §③ section started above.
mode = st.radio("Processing Mode", ["Single File", "Batch / Multiple Files"], horizontal=True)

# ═══════════════════════════════════════════════════════════════
#  RawData UPload and processing
# ═══════════════════════════════════════════════════════════════

st.header("3. MetaData collection", divider="gray")

# ─── Consume pending auto-extract values from §③ (set on the previous
# rerun by the upload section). Streamlit forbids writing to a widget's
# session_state key AFTER the widget has been instantiated, so §③
# stashes auto-extracted values in '_pending_*' keys + calls st.rerun().
# We copy them to the widget keys here, BEFORE any §② widget renders.

st.caption("Pre-filled from submission form where available — edit any field before processing.")

r1c1, r1c2, r1c3, r1c4 = st.columns(4)
with r1c1:
    tracking = st.text_input("Tracking Number", value=_sm.get("Tracking Number", ""))
    analyst = st.text_input("Analyst (POC)", value=_sm.get("Analyst", ""))
with r1c2:
    task_number = st.text_input("Task Number", value=_sm.get("Task Number", ""))
    task_name = st.text_input("Task Name", value=_sm.get("Task Name", ""))
with r1c3:
    task_pi = st.text_input("Task PI", value=_sm.get("Task PI", ""))
    analysis = st.text_input("Analysis Requested", value=_sm.get("Analysis Requested", ""))
with r1c4:
    batch = st.text_input("Batch", value="")

r2c1, r2c2, r2c3, r2c4 = st.columns(4)
with r2c1:
    unit = st.text_input("Units", value="mg/L")
    col_serial = st.text_input("Column Serial (last 4 digits)", value="")
with r2c2:
    # Analytical Method auto-populates from MassHunter's modal `Acq. Method File`
    # value when available (parity with Chemstation auto-extracting Run Date from
    # the batch path). UI value still wins once the user has typed something —
    # the auto-populated value is pushed into session_state["meta_analytical_method"]
    # ONLY when a brand-new file is parsed (see the upload section), so manual
    # edits between uploads are preserved.
    st.session_state.setdefault("meta_analytical_method", "")
    analytical_method = st.text_input(
        "Analytical Method",
        key="meta_analytical_method",
        help="The instrument/chromatography method (e.g., HPLC method file or chromatographic conditions). "
        "Auto-populated from the modal 'Acq. Method File' on MassHunter exports.",
    )
    processing_method = st.text_input(
        "Processing Method",
        value="",
        help="The data-processing / integration method (distinct from the analytical method).",
    )
with r2c3:
    # Default to today, but if a file was already parsed this session and we
    # extracted a Run Date from it, prefer that. The upload section pushes the
    # extracted date into session_state["meta_run_date"] on a new parse.
    st.session_state.setdefault("meta_run_date", datetime.today().date())
    run_date = st.date_input(
        "Run Date",
        key="meta_run_date",
        help="Auto-populated from the raw file when a date " "is found there; otherwise today.",
    )
with r2c4:
    instrument_name = st.text_input(
        "Instrument Name",
        value="",
        placeholder="e.g. R2P2 or R2P2; Mario",
        help=(
            "Free-text entry. Use semicolons to separate multiple instruments.\n\n"
            "Suggested names: " + ", ".join(n for n in INSTRUMENT_NAMES if n)
        ),
    )

    output_dir = st.text_input(
        "Output folder (optional)",
        value="",
        key="output_dir",
        help="Paste a folder path to enable direct save, e.g. C:\\\\Data\\\\Results",
    )

# Export Type selectbox lives in §③ which renders BELOW this section. Its
# session_state key ("single_export_type") is set when the selectbox renders;
# we read it here with INSTRUMENT_TYPES[0] as the first-paint fallback so the
# META dict below always has a value to use.
instrument = st.session_state.get("single_export_type", INSTRUMENT_TYPES[0])

META = {
    "Tracking Number": tracking,
    "Analyst": analyst,
    "Task Number": task_number,
    "Task Name": task_name,
    "Task PI": task_pi,
    "Analysis": analysis,
    "Batch": batch,
    "Analytical Method": analytical_method,
    "Processing Method": processing_method,
    "Unit": unit,
    "Column Serial": col_serial,
    "Instrument Name": instrument_name,
    "Run Date": str(run_date),
    "Export Type": instrument,
}
st.divider()
for _src, _dst in [
    ("_pending_meta_run_date", "meta_run_date"),
    ("_pending_meta_analytical_method", "meta_analytical_method"),
]:
    if _src in st.session_state:
        try:
            _val = st.session_state[_src]
            st.session_state[_dst] = _val
            del st.session_state[_src]
        except Exception:
            pass
        st.session_state[_dst] = st.session_state.pop(_src)

# ───────────────────────────────────────────────────────────────
#  LLOQ / ULOQ  CHECKBOX UI
#  Renders a compact table: rows = STD levels, cols = analytes
#  Two checkbox groups per analyte: LLOQ | ULOQ (radio-like)
# ───────────────────────────────────────────────────────────────


def render_lloq_uloq_ui(tidy_df: pd.DataFrame, unit: str) -> tuple:
    """
    Renders the LLOQ/ULOQ selector.
    Returns (lloq_map, uloq_map) dicts keyed by analyte name.
    """
    std_levels = detect_std_levels(tidy_df)
    pairs = get_analyte_pairs(tidy_df)
    analytes = [a for a, _, _ in pairs]

    if not std_levels or not analytes:
        st.caption("No STD levels detected — LLOQ/ULOQ cannot be auto-populated.")
        return dict.fromkeys(analytes), dict.fromkeys(analytes)

    def _check_row(level, analyte_list, kind: str):
        """kind ◘ {'lloq','uloq'} - tick every analyte's box in this row."""
        for a in analyte_list:
            st.session_state[f"{kind}_cb_{a}_{level}"] = True

    st.subheader("LLOQ / ULOQ Definition", anchor=False)
    st.caption(
        f"Detected **{len(std_levels)}** STD levels across **{len(analytes)}** analyte(s). "
        "Check one level as **LLOQ** and one as **ULOQ** per analyte, "
        "or leave blank and enter values manually below."
    )

    # ── Checkbox grid ──────────────────────────────────────────
    # header row
    header_cols = st.columns([1.4] + [1] * len(analytes))
    header_cols[0].markdown("**STD Level**")
    for i, a in enumerate(analytes):
        header_cols[i + 1].markdown(f"**{a}**")

    # track which box is ticked per analyte — store in session state
    # keys: lloq_cb_{analyte}_{level}, uloq_cb_{analyte}_{level}
    for lvl in std_levels:
        row_cols = st.columns([1.4] + [1] * len(analytes))
        with row_cols[0]:
            st.markdown(f"`{lvl} {unit}`")
            bc1, bc2 = st.columns(2)
            bc1.button(
                "All LLOQ", key=f"all_lloq_{lvl}", on_click=_check_row, args=(lvl, analytes, "lloq")
            )
            bc2.button(
                "All ULOQ", key=f"all_uloq_{lvl}", on_click=_check_row, args=(lvl, analytes, "uloq")
            )
        for i, a in enumerate(analytes):
            with row_cols[i + 1]:
                lkey = f"lloq_cb_{a}_{lvl}"
                ukey = f"uloq_cb_{a}_{lvl}"
                st.checkbox("LLOQ", key=lkey)
                st.checkbox("ULOQ", key=ukey)
        st.markdown("<hr style='margin:4px 0; border-top:1px solid #eee'>", unsafe_allow_html=True)
    # Collect selections — last ticked box wins
    lloq_map: dict = {}
    uloq_map: dict = {}
    for a in analytes:
        lloq_val = None
        uloq_val = None
        for lvl in std_levels:
            if st.session_state.get(f"lloq_cb_{a}_{lvl}"):
                lloq_val = lvl
            if st.session_state.get(f"uloq_cb_{a}_{lvl}"):
                uloq_val = lvl
        lloq_map[a] = lloq_val
        uloq_map[a] = uloq_val

    # ── Manual override row ────────────────────────────────────
    st.caption("Manual override — leave 0 to use checkbox selection above.")
    ov_cols = st.columns([1.4] + [1] * len(analytes))
    ov_cols[0].markdown("**LLOQ override**")
    for i, a in enumerate(analytes):
        with ov_cols[i + 1]:
            ov = st.number_input(
                "",
                min_value=0.0,
                value=0.0,
                step=0.001,
                format="%.4f",
                key=f"lloq_ov_{a}",
                label_visibility="collapsed",
            )
            if ov > 0:
                lloq_map[a] = ov

    ov_cols2 = st.columns([1.4] + [1] * len(analytes))
    ov_cols2[0].markdown("**ULOQ override**")
    for i, a in enumerate(analytes):
        with ov_cols2[i + 1]:
            ov = st.number_input(
                "",
                min_value=0.0,
                value=0.0,
                step=0.001,
                format="%.4f",
                key=f"uloq_ov_{a}",
                label_visibility="collapsed",
            )
            if ov > 0:
                uloq_map[a] = ov

    return lloq_map, uloq_map


# ═══════════════════════════════════════════════════════════════
#  SINGLE FILE MODE
# ═══════════════════════════════════════════════════════════════

if mode == "Single File":
    # The upload + pre-parse happened up in §② so the metadata form had
    # the auto-extracted values to pre-fill from. Pick those results up
    # from session_state. If nothing has been uploaded yet, point the
    # user back up the page.
    if not st.session_state.get("_single_file_bytes"):
        st.info("👆 Upload a raw data file in section ② to continue.")
        st.stop()

    file_bytes = st.session_state.get("_single_file_bytes")
    file_stem = st.session_state.get("_single_file_stem", "")
    _tidy_j = st.session_state.get("_single_tidy_preview")
    if file_bytes is None or _tidy_j is None:
        st.info("👆 Upload a raw data file in section ③ to continue.")
        st.stop()
    tidy_preview = safe_read_json(_tidy_j)

    # Defensive: if a stale parse_key from a prior Export Type is in
    # session_state but the user has since switched the selectbox in §②,
    # the bytes/tidy/extracted_meta got rebuilt up there already — we
    # just read whatever's now current.

    # ── Sample Classification Diagnostic ─────────────────────
    if "isStandard" in tidy_preview.columns:
        counts = tidy_preview["isStandard"].value_counts()
        with st.expander(f"Sample classification: {len(tidy_preview)} rows total"):
            st.dataframe(
                counts.reset_index().rename(
                    columns={"index": "Type", "isStandard": "Type", "count": "Count"}
                ),
                use_container_width=True,
                hide_index=True,
            )
            for stype in ["STD", "CVS", "Blank", "Sample"]:
                subset = tidy_preview[tidy_preview["isStandard"] == stype]
                if not subset.empty:
                    names = subset["Sample Name"].head(5).tolist()
                    st.caption(f"**{stype}** examples: {', '.join(str(n) for n in names)}")

    # ── STD QC Bounds ────────────────────────────────────────
    std_lo, std_hi = render_std_qc_bounds(key_prefix="single_")
    st.divider()

    # ── CVS Tolerance UI ──────────────────────────────────────
    cvs_levels = detect_cvs_levels(tidy_preview)
    cvs_tolerances = render_cvs_tolerance_ui(cvs_levels, unit, key_prefix="single_")

    # ── CVS → Analyte Assignment (form-wrapped) ───────────────
    # Wrapping in st.form means checkbox edits don't trigger reruns. The user
    # makes all their edits, then clicks the submit button (the refresh) to
    # commit and have the heatmap + Workup re-render with the new assignment.
    _cvs_names = (
        tidy_preview[tidy_preview["isStandard"] == "CVS"]["Sample Name"]
        .astype(str)
        .unique()
        .tolist()
        if "isStandard" in tidy_preview.columns
        else []
    )
    _analyte_list = [a for a, _, _ in get_analyte_pairs(tidy_preview)]

    # Build {CVS sample name → concentration} map. Drives the editor's
    # decision to group columns by concentration when multiple injections
    # share a level (the MassHunter pattern).
    _cvs_name_to_conc: dict = {}
    if (
        _cvs_names
        and "CVS_Known_Conc" in tidy_preview.columns
        and "Sample Name" in tidy_preview.columns
    ):
        cvs_rows = tidy_preview[tidy_preview["isStandard"] == "CVS"]
        for nm, grp in cvs_rows.groupby("Sample Name"):
            uc = grp["CVS_Known_Conc"].dropna().unique()
            if len(uc) == 1:
                _cvs_name_to_conc[str(nm)] = float(uc[0])

    if _cvs_names and _analyte_list:
        st.subheader("CVS → Analyte Assignment", anchor=False)
        st.caption(
            "Rows are analytes, columns are CVS injection groups (rep suffixes "
            "collapsed — one row applies to every rep). On MassHunter runs "
            "columns auto-collapse to one per concentration level when multiple "
            "injections share a level — hover the header to see which injections "
            "each column covers. Defaults to checked where the analyte name "
            "appears in the CVS name. Edit checkboxes freely — they won't take "
            "effect until you click **Refresh QC**."
        )
        # Bulk-action buttons go OUTSIDE the form (regular st.button can't
        # live inside an st.form). They mutate the draft via session_state
        # and rerun; the user then clicks Refresh QC to commit.
        render_cvs_bulk_action_buttons("single_")
        with st.form("single_cvs_assign_form", clear_on_submit=False):
            render_cvs_analyte_assignment_editor(
                _cvs_names,
                _analyte_list,
                key_prefix="single_",
                name_to_conc=_cvs_name_to_conc,
                unit=unit,
            )
            _refresh_clicked = st.form_submit_button(
                "🔄 Refresh QC",
                type="primary",
                help="Commit the checkbox edits and re-render the heatmap + Workup.",
            )
        if _refresh_clicked:
            commit_cvs_assignment("single_")
    cvs_analyte_assignment = get_applied_cvs_assignment("single_")

    # ── Full parse pipeline (cached on file bytes; tolerances NOT in key) ──
    import hashlib as _hl

    _run_hash = _hl.md5(file_bytes).hexdigest() + instrument + _CACHE_VER
    if st.session_state.get("_single_run_hash") != _run_hash:
        with st.spinner("Running pipeline…"):
            try:
                res = run_processing(
                    file_bytes,
                    instrument,
                    META,
                    cvs_tolerances,
                    st.session_state.get("submission_meta", {}),
                    st.session_state.get("submission_samples"),
                    analyte_assignment=cvs_analyte_assignment,
                    std_lo=std_lo,
                    std_hi=std_hi,
                )
            except Exception as e:
                st.error(f"Processing error: {e}")
                st.stop()
            st.session_state["_single_res"] = res
            st.session_state["_single_run_hash"] = _run_hash
    res = st.session_state["_single_res"]
    tidy_df = res["tidy"]
    raw_df = res["raw"]
    _tidy_json = res["_tidy_json"]

    # ── QC workups: recompute every rerun via cached helpers ─
    # This is what makes tolerance/assignment changes reactive — the heavy
    # parse step is gated by file bytes, but std_df/cvs_df rebuild whenever
    # any QC input changes (and the @st.cache_data layer makes repeat
    # invocations with identical inputs free).
    std_df, cvs_df = recompute_qc(
        _tidy_json,
        unit,
        cvs_tolerances,
        cvs_analyte_assignment,
        std_lo,
        std_hi,
    )

    # ── QC Heatmap (sits directly under tolerance/assignment area) ──
    st.subheader("QC Recovery Heatmap", anchor=False)
    render_qc_heatmap(std_df, cvs_df, std_lo, std_hi, cvs_tolerances)

    # ── Per-level CVS pass rates — rendered HERE (right after heatmap)
    #    so the numbers update as soon as Refresh QC commits new tolerances
    #    or assignment edits, without scrolling to the bottom of the page.
    if "Pass" in cvs_df.columns and "CVS Level" in cvs_df.columns and len(cvs_df):
        cvs_lvl_vals = sorted(cvs_df["CVS Level"].dropna().unique())
        if cvs_lvl_vals:
            lc = st.columns(min(len(cvs_lvl_vals), 6))
            for i, lvl in enumerate(cvs_lvl_vals):
                sub = cvs_df[(cvs_df["CVS Level"] == lvl) & (cvs_df["Pass"].isin(["✅", "❌"]))]
                lr = (sub["Pass"] == "✅").mean() * 100 if len(sub) else float("nan")
                tol = cvs_tolerances.get(float(lvl), 10)
                with lc[i % len(lc)]:
                    st.metric(f"CVS {lvl} {unit}  ±{tol}%", f"{lr:.1f}%" if pd.notna(lr) else "—")
    st.divider()

    # ── LLOQ / ULOQ UI ───────────────────────────────────────
    lloq_map, uloq_map = render_lloq_uloq_ui(tidy_preview, unit)
    st.divider()

    # ── Summary inclusion toggles ────────────────────────────
    use_dil, inc_blanks, inc_std, inc_cvs = render_inclusion_checkboxes(key_prefix="single_")

    # ── Build (or reuse cached) Workup ───────────────────────
    _wk_key = (
        _run_hash
        + "|"
        + json.dumps(lloq_map, sort_keys=True, default=str)
        + "|"
        + json.dumps(uloq_map, sort_keys=True, default=str)
        + "|"
        + f"{std_lo}-{std_hi}|"
        + json.dumps({str(k): v for k, v in cvs_tolerances.items()}, sort_keys=True)
        + "|"
        + json.dumps(cvs_analyte_assignment, sort_keys=True)
    )
    workup_df, workup_row_types, workup_row_meta = memoized_build_workup(
        _wk_key,
        "_single_workup",
        tidy_df,
        lloq_map,
        uloq_map,
        unit,
        std_lo,
        std_hi,
        cvs_tolerances,
        analyte_assignment=cvs_analyte_assignment,
    )
    # ── Project Workup → Summary based on checkboxes (cheap) ─
    summary_df, summary_row_types = build_summary(
        workup_df,
        workup_row_types,
        use_dilution=use_dil,
        include_blanks=inc_blanks,
        include_std=inc_std,
        include_cvs=inc_cvs,
        unit=unit,
    )

    # ── QC pass-rate summary (above the table) ───────────────
    mc1, mc2, mc3, mc4 = st.columns(4)
    with mc1:
        pass_col = f"Pass ({std_lo}-{std_hi}%)"
        if pass_col in std_df.columns and len(std_df):
            rate = (std_df[pass_col] == "✅").mean() * 100
            st.metric(f"STD Pass Rate ({std_lo}-{std_hi}%)", f"{rate:.1f}%")
        else:
            st.metric("STD Pass Rate", "—")
    with mc2:
        if "Pass" in cvs_df.columns and len(cvs_df):
            active = cvs_df[cvs_df["Pass"].isin(["✅", "❌"])]
            if len(active):
                overall = (active["Pass"] == "✅").mean() * 100
                st.metric("CVS Overall Pass Rate", f"{overall:.1f}%")
            else:
                st.metric("CVS Overall Pass Rate", "—")
        else:
            st.metric("CVS Overall Pass Rate", "—")
    with mc3:
        n_blanks = (
            int((tidy_df["isStandard"] == "Blank").sum()) if "isStandard" in tidy_df.columns else 0
        )
        st.metric("Blank Injections", f"{n_blanks}")
    with mc4:
        n_samples = (
            int((tidy_df["isStandard"] == "Sample").sum()) if "isStandard" in tidy_df.columns else 0
        )
        st.metric("Sample Rows", f"{n_samples}")

    # (Per-level CVS pass rates were moved up to the heatmap area —
    #  rendered right after Refresh QC so they update without scrolling.)

    # ── Tabs: Workup (review) → Summary (polished) → Tidy → Raw ──
    t1, t2, t3, t4 = st.tabs(["📊 Workup", "📋 Summary", "📋 Tidy Data", "🗂️ Raw Data"])

    with t1:
        st.caption(
            "Full review surface — everything is here. STD (recovery vs "
            f"{std_lo}-{std_hi}% window) · CVS (per-level tolerance + assignment) · "
            "Blank injections (pale blue) · Samples (LLOQ/ULOQ flags). "
            "Use the toggles above to shape the Summary."
        )
        render_workup_table(
            workup_df,
            workup_row_types,
            unit,
            row_meta=workup_row_meta,
            std_lo=std_lo,
            std_hi=std_hi,
            cvs_tolerances=cvs_tolerances,
            analyte_assignment=cvs_analyte_assignment,
            lloq_map=lloq_map,
            uloq_map=uloq_map,
        )

    with t2:
        kept_bits = []
        if use_dil:
            kept_bits.append("dilution-corrected columns")
        if inc_blanks:
            kept_bits.append("blank injections")
        if inc_std:
            kept_bits.append("STD rows")
        if inc_cvs:
            kept_bits.append("CVS rows")
        kept_str = ", ".join(kept_bits) if kept_bits else "Sample rows only"
        st.caption(
            f"Polished Summary — what will be written to the Summary sheet. "
            f"Currently including: {kept_str}."
        )
        # Build the parallel row_meta for Summary by filtering workup_row_meta
        # with the same mask build_summary used. The mask is reproducible from
        # the row types and the toggles, so do it here in lockstep.
        _kept_mask = []
        for rt in workup_row_types:
            if rt == "ref" or rt == "sample":
                _kept_mask.append(True)
            elif rt == "blank":
                _kept_mask.append(inc_blanks)
            elif rt == "std":
                _kept_mask.append(inc_std)
            elif rt == "cvs":
                _kept_mask.append(inc_cvs)
            else:
                _kept_mask.append(False)
        summary_row_meta = [m for m, k in zip(workup_row_meta, _kept_mask) if k]
        render_workup_table(
            summary_df,
            summary_row_types,
            unit,
            row_meta=summary_row_meta,
            std_lo=std_lo,
            std_hi=std_hi,
            cvs_tolerances=cvs_tolerances,
            analyte_assignment=cvs_analyte_assignment,
            lloq_map=lloq_map,
            uloq_map=uloq_map,
        )

    with t3:
        st.caption(f"{len(tidy_df)} rows · {len(tidy_df.columns)} columns")
        st.dataframe(tidy_df, use_container_width=True, hide_index=True)

    with t4:
        st.dataframe(raw_df, use_container_width=True, hide_index=True)

    # (QC Heatmap is rendered above, under the tolerance area.)
    color_summary = st.checkbox(
        "Format Summary sheet with 'Workup'colors",
        value=False,
        key="color_summary_xlsx",
        help="Apply the same formatting including pass/fail/ref/blank coloring to the Summary sheet that's used for the Workup sheet. Uncheck for uncolored Summary",
    )
    # ── Export ────────────────────────────────────────────────
    st.divider()
    ec1, ec2 = st.columns(2)

    with ec1:
        if st.button("📥 Build Excel", key="xl_build"):
            with st.spinner("Building workbook…"):
                _effective_meta = merge_run_meta(
                    META, res.get("extracted_meta") if isinstance(res, dict) else None
                )
                _qc_ctx = {
                    "unit": unit,
                    "std_lo": std_lo,
                    "std_hi": std_hi,
                    "cvs_tolerances": cvs_tolerances,
                    "analyte_assignment": cvs_analyte_assignment,
                    "lloq_map": lloq_map,
                    "uloq_map": uloq_map,
                    "workup_row_meta": workup_row_meta,
                    "summary_row_meta": summary_row_meta,
                }
                st.session_state["xlsx_bytes"] = build_excel(
                    tidy_df,
                    raw_df,
                    workup_df,
                    workup_row_types,
                    summary_df,
                    summary_row_types if color_summary else [],
                    _effective_meta,
                    submission_wb_bytes=st.session_state.get("submission_wb"),
                    qc_context={
                        **_qc_ctx,
                        "summary_row_meta": _qc_ctx["summary_row_meta"] if color_summary else None,
                    },
                )
                # Derive filename from Tracking Number (YYYY-###) if available
                _track = META.get("Tracking Number", "").strip()
                if _track:
                    st.session_state["xlsx_filename"] = f"{_track}.xlsx"
                else:
                    st.session_state["xlsx_filename"] = f"cleaned_{file_stem}.xlsx"

        if st.session_state["xlsx_bytes"] is not None:
            label = (
                "⬇️ Download Excel (+ submission form)"
                if st.session_state.get("submission_wb")
                else "⬇️ Download Excel"
            )
            st.download_button(
                label,
                data=st.session_state["xlsx_bytes"],
                file_name=st.session_state["xlsx_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_xl",
            )
            # ── Direct save to output folder (with _v1, _v2 … dedup) ──
            _out_dir = st.session_state.get("output_dir", "").strip()
            if _out_dir and os.path.isdir(_out_dir):
                _base, _ext = os.path.splitext(st.session_state["xlsx_filename"])
                _candidate = os.path.join(_out_dir, st.session_state["xlsx_filename"])
                if os.path.exists(_candidate):
                    _ver = 1
                    while os.path.exists(os.path.join(_out_dir, f"{_base}_v{_ver}{_ext}")):
                        _ver += 1
                    _candidate = os.path.join(_out_dir, f"{_base}_v{_ver}{_ext}")
                if st.button(f"💾 Save to `{_out_dir}`", key="save_to_dir"):
                    with open(_candidate, "wb") as _fout:
                        _fout.write(st.session_state["xlsx_bytes"])
                    st.success(f"Saved → `{_candidate}`")

    with ec2:
        st.markdown("**Push to Dropbox**")
        # The pushed bytes ARE st.session_state["xlsx_bytes"] — the same
        # bytes the Download button serves. No second generation path.
        render_dropbox_push_ui(
            key_prefix="single_",
            meta=META,
            fallback_stem=f"cleaned_{file_stem}",
            xlsx_bytes=st.session_state.get("xlsx_bytes"),
        )

    # ── Advanced: SQLite export (out-of-the-way for the lab user) ────
    with st.expander("Advanced: SQLite export", expanded=False):
        st.caption(
            "Programmatic archive. Two tables: `runs` (one row of metadata) "
            "and `tidy_data` (every parsed row, linked by run_id). "
            "Use a SQLite browser or `pd.read_sql(...)` to query."
        )
        if st.button("💾 Build SQLite DB", key="db_build"):
            with st.spinner("Writing SQLite…"):
                _sql_meta = merge_run_meta(
                    META, res.get("extracted_meta") if isinstance(res, dict) else None
                )
                _sql_meta = {**_sql_meta, "Source File": f"{file_stem}.xlsx"}
                st.session_state["db_bytes"] = export_sqlite(_sql_meta, tidy_df)

        if st.session_state["db_bytes"] is not None:
            _db_fn = f"{META.get('Tracking Number', '').strip() or 'analytical_data'}.db"
            st.download_button(
                "⬇️ Download SQLite DB",
                data=st.session_state["db_bytes"],
                file_name=_db_fn,
                mime="application/octet-stream",
                key="dl_db",
            )
            st.caption("Tables: `runs` · `tidy_data`")


# ═══════════════════════════════════════════════════════════════
#  BATCH MODE
# ═══════════════════════════════════════════════════════════════

else:
    uploaded_files = st.file_uploader(
        "Upload multiple raw data files (.xlsx / .xls)",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="batch_upload",
    )
    if not uploaded_files:
        st.info("👆 Upload one or more files to begin.")
        st.stop()

    # Per-file instrument assignment
    st.write("**Assign export type per file:**")
    fi_cols = st.columns(min(len(uploaded_files), 3))
    file_instruments: dict = {}
    for i, f in enumerate(uploaded_files):
        with fi_cols[i % 3]:
            file_instruments[f.name] = st.selectbox(
                f.name,
                INSTRUMENT_TYPES,
                index=INSTRUMENT_TYPES.index(instrument),
                key=f"inst_{i}",
            )

    # Detect CVS + STD levels across all files
    all_cvs_levels: set = set()
    combined_preview = pd.DataFrame()
    batch_extracted_meta: dict = {}  # filename -> {Batch Path: ...}
    per_file_preview: dict = {}  # filename -> tidy preview df
    for f in uploaded_files:
        fb = f.read()
        f.seek(0)
        try:
            _, tj, em_json = CACHED_PARSERS[file_instruments[f.name]](fb)
            prev = safe_read_json(tj)
            all_cvs_levels.update(detect_cvs_levels(prev))
            combined_preview = pd.concat([combined_preview, prev], ignore_index=True)
            per_file_preview[f.name] = prev
            try:
                em = json.loads(em_json) if em_json else {}
                if em:
                    batch_extracted_meta[f.name] = em
            except Exception:
                pass
        except Exception:
            pass

    if batch_extracted_meta:
        with st.expander(f"Auto-extracted metadata from {len(batch_extracted_meta)} file(s)"):
            for fname, em in batch_extracted_meta.items():
                st.markdown(f"**{fname}**")
                for k, v in em.items():
                    st.code(f"{k}: {v}", language=None)
    st.session_state["_batch_extracted_meta"] = batch_extracted_meta

    # ── STD QC Bounds (global to batch) ──────────────────────
    std_lo, std_hi = render_std_qc_bounds(key_prefix="batch_")
    st.divider()

    # ── CVS Tolerance UI (global to batch) ───────────────────
    cvs_tolerances = render_cvs_tolerance_ui(sorted(all_cvs_levels), unit, key_prefix="batch_")

    # ── CVS → Analyte Assignment (PER FILE, form-wrapped) ────
    # Each file gets its own checkbox table, since different runs may have
    # different intended QC scope. ALL of them live in a single st.form so
    # checkbox edits across all files are batched — the user can edit
    # multiple file's assignments and click ONE Refresh QC button to apply.
    st.subheader("CVS → Analyte Assignment (per file)", anchor=False)
    st.caption(
        "Each uploaded file has its own table inside an expander. Rows are "
        "analytes, columns are CVS injection groups. Defaults to checked where "
        "the analyte name appears in the CVS name. Edit freely across files — "
        "one **Refresh QC** click commits everything."
    )

    # Compute per-file key prefixes up front so we can use them both inside
    # the form (to render the editors) and afterwards (to commit + fetch).
    import hashlib as _hl_a

    per_file_key_prefix: dict = {
        fname: f"batch_{_hl_a.md5(fname.encode()).hexdigest()[:8]}_" for fname in per_file_preview
    }
    per_file_cvs_analytes: dict = {}  # fname -> (cvs_names, analyte_list, name_to_conc)
    for fname, prev in per_file_preview.items():
        cvs_names_f = (
            prev[prev["isStandard"] == "CVS"]["Sample Name"].astype(str).unique().tolist()
            if "isStandard" in prev.columns
            else []
        )
        analyte_list_f = [a for a, _, _ in get_analyte_pairs(prev)]
        # Per-file concentration map for grouped MassHunter display
        name_to_conc_f: dict = {}
        if cvs_names_f and "CVS_Known_Conc" in prev.columns and "Sample Name" in prev.columns:
            cvs_rows = prev[prev["isStandard"] == "CVS"]
            for nm, grp in cvs_rows.groupby("Sample Name"):
                uc = grp["CVS_Known_Conc"].dropna().unique()
                if len(uc) == 1:
                    name_to_conc_f[str(nm)] = float(uc[0])
        per_file_cvs_analytes[fname] = (cvs_names_f, analyte_list_f, name_to_conc_f)

    # Global bulk-action buttons (above the form). One click toggles every
    # checkbox across every file's table — handy when the experiment has
    # all analytes valid for all CVS levels in every run.
    if per_file_preview:
        st.caption("Bulk actions apply to **every file's** assignment table below.")
        bA, bB, _ = st.columns([1, 1, 3])
        with bA:
            if st.button("☑️ Check all (every file)", key="batch_cvs_check_all_global"):
                for fname in per_file_preview:
                    set_all_cvs_assignment(per_file_key_prefix[fname], True)
                st.rerun()
        with bB:
            if st.button("☐ Clear all (every file)", key="batch_cvs_clear_all_global"):
                for fname in per_file_preview:
                    set_all_cvs_assignment(per_file_key_prefix[fname], False)
                st.rerun()

    with st.form("batch_cvs_assign_form", clear_on_submit=False):
        for fname, prev in per_file_preview.items():
            cvs_names_f, analyte_list_f, name_to_conc_f = per_file_cvs_analytes[fname]
            with st.expander(f"📁 {fname}", expanded=(len(per_file_preview) <= 2)):
                if not cvs_names_f:
                    st.caption("No CVS injections found in this file.")
                    continue
                if not analyte_list_f:
                    st.caption("No analytes detected.")
                    continue
                render_cvs_analyte_assignment_editor(
                    cvs_names_f,
                    analyte_list_f,
                    key_prefix=per_file_key_prefix[fname],
                    name_to_conc=name_to_conc_f,
                    unit=unit,
                )
        _batch_refresh_clicked = st.form_submit_button(
            "🔄 Refresh QC",
            type="primary",
            help="Commit checkbox edits across ALL file tables and re-render the heatmap.",
        )

    if _batch_refresh_clicked:
        for fname in per_file_preview:
            commit_cvs_assignment(per_file_key_prefix[fname])

    # Pull the applied per-file assignments for downstream code
    per_file_assignment: dict = {
        fname: get_applied_cvs_assignment(per_file_key_prefix[fname]) for fname in per_file_preview
    }

    # ── Pre-Run QC Heatmap: combined, but each file's assignment honored ──
    # Built by running recompute_qc PER FILE with each file's assignment,
    # then concatenating results with a [filename] prefix on Sample Name so
    # the heatmap rows are scannable by file.
    if per_file_preview:
        try:
            _std_parts, _cvs_parts = [], []
            for fname, prev in per_file_preview.items():
                _std_f, _cvs_f = recompute_qc(
                    prev.to_json(),
                    unit,
                    cvs_tolerances,
                    per_file_assignment.get(fname, {}),
                    std_lo,
                    std_hi,
                )
                file_tag = f"[{os.path.splitext(fname)[0]}]"
                if not _std_f.empty and "Sample Name" in _std_f.columns:
                    _std_f = _std_f.copy()
                    _std_f["Sample Name"] = file_tag + " " + _std_f["Sample Name"].astype(str)
                    _std_parts.append(_std_f)
                if not _cvs_f.empty and "Sample Name" in _cvs_f.columns:
                    _cvs_f = _cvs_f.copy()
                    _cvs_f["Sample Name"] = file_tag + " " + _cvs_f["Sample Name"].astype(str)
                    _cvs_parts.append(_cvs_f)
            _pre_std = pd.concat(_std_parts, ignore_index=True) if _std_parts else pd.DataFrame()
            _pre_cvs = pd.concat(_cvs_parts, ignore_index=True) if _cvs_parts else pd.DataFrame()
            st.subheader("QC Recovery Heatmap (pre-run preview)", anchor=False)
            st.caption(
                "Row labels prefixed with `[file_stem]` so each file's CVS / STD injections are identifiable."
            )
            render_qc_heatmap(_pre_std, _pre_cvs, std_lo, std_hi, cvs_tolerances)
        except Exception as _e:
            st.caption(f"_Pre-run heatmap unavailable: {_e}_")
    st.divider()

    if not combined_preview.empty:
        lloq_map, uloq_map = render_lloq_uloq_ui(combined_preview, unit)
        st.divider()
    else:
        lloq_map, uloq_map = {}, {}

    # ── Summary inclusion toggles (batch) ────────────────────
    batch_use_dil, batch_inc_blanks, batch_inc_std, batch_inc_cvs = render_inclusion_checkboxes(
        key_prefix="batch_"
    )

    if st.button("▶️ Run Workup", type="primary"):
        all_r: dict = {k: [] for k in ["tidy", "raw", "std", "cvs"]}
        per_file_workup: list = []
        per_file_workup_types: list = []
        per_file_workup_meta: list = []
        per_file_summary: list = []
        per_file_summary_types: list = []
        per_file_summary_meta: list = []
        per_file_names: list = []  # parallel to the per_file_* lists
        errors = []
        prog = st.progress(0)
        status = st.empty()

        for i, f in enumerate(uploaded_files):
            inst = file_instruments[f.name]
            fmeta = {**META, "Source File": f.name, "Export Type": inst}
            status.text(f"Processing {f.name}…")
            try:
                fb = f.read()
                # Per-file assignment — falls back to empty dict (all-assigned)
                # if the file had no CVS injections.
                file_assignment = per_file_assignment.get(f.name, {})
                res = run_processing(
                    fb,
                    inst,
                    fmeta,
                    cvs_tolerances,
                    st.session_state.get("submission_meta", {}),
                    st.session_state.get("submission_samples"),
                    analyte_assignment=file_assignment,
                    std_lo=std_lo,
                    std_hi=std_hi,
                )
                wk_df, wk_types, wk_meta = build_workup(
                    res["tidy"],
                    lloq_map,
                    uloq_map,
                    unit,
                    std_lo=std_lo,
                    std_hi=std_hi,
                    cvs_tolerances=cvs_tolerances,
                    analyte_assignment=file_assignment,
                )
                sm_df, sm_types = build_summary(
                    wk_df,
                    wk_types,
                    use_dilution=batch_use_dil,
                    include_blanks=batch_inc_blanks,
                    include_std=batch_inc_std,
                    include_cvs=batch_inc_cvs,
                    unit=unit,
                )
                _kept_mask = []
                for rt in wk_types:
                    if rt == "ref" or rt == "sample":
                        _kept_mask.append(True)
                    elif rt == "blank":
                        _kept_mask.append(batch_inc_blanks)
                    elif rt == "std":
                        _kept_mask.append(batch_inc_std)
                    elif rt == "cvs":
                        _kept_mask.append(batch_inc_cvs)
                    else:
                        _kept_mask.append(False)
                sm_meta = [m for m, k in zip(wk_meta, _kept_mask) if k]
                for key in ["tidy", "raw", "std", "cvs"]:
                    res[key]["Source File"] = f.name
                    all_r[key].append(res[key])
                wk_df["Source File"] = f.name
                sm_df["Source File"] = f.name
                per_file_names.append(f.name)
                per_file_workup.append(wk_df)
                per_file_workup_types.append(wk_types)
                per_file_workup_meta.append(wk_meta)
                per_file_summary.append(sm_df)
                per_file_summary_types.append(sm_types)
                per_file_summary_meta.append(sm_meta)
                st.success(f"✅ {f.name}")
            except Exception as e:
                errors.append((f.name, str(e)))
                st.error(f"❌ {f.name}: {e}")
            prog.progress((i + 1) / len(uploaded_files))

        status.empty()
        if not all_r["tidy"]:
            st.error("No files processed successfully.")
            st.stop()

        comb = {k: pd.concat(v, ignore_index=True) for k, v in all_r.items()}

        # ── Combined workup (still built — used for QC metrics and as a
        #    fallback aggregate view; primary display is per-file tabs).
        def _stitch(per_file_dfs, per_file_types, per_file_metas):
            rows = []
            types = []
            metas = []
            sep_template = None
            sep_meta = {"type": "sep", "name": "", "known": None}
            for fi, (df_, t_, m_) in enumerate(zip(per_file_dfs, per_file_types, per_file_metas)):
                if fi > 0:
                    if sep_template is None:
                        sep_template = dict.fromkeys(df_.columns, "")
                    rows.append(sep_template)
                    types.append("sep")
                    metas.append(sep_meta)
                for _, r in df_.iterrows():
                    rows.append(r.to_dict())
                types.extend(t_)
                metas.extend(m_)
            return (pd.DataFrame(rows) if rows else pd.DataFrame()), types, metas

        combined_workup_df, combined_workup_types, combined_workup_meta = _stitch(
            per_file_workup, per_file_workup_types, per_file_workup_meta
        )
        combined_summary_df, combined_summary_types, combined_summary_meta = _stitch(
            per_file_summary, per_file_summary_types, per_file_summary_meta
        )

        comb["workup"] = combined_workup_df
        comb["workup_types"] = combined_workup_types
        comb["workup_meta"] = combined_workup_meta
        comb["summary"] = combined_summary_df
        comb["summary_types"] = combined_summary_types
        comb["summary_meta"] = combined_summary_meta
        st.session_state["batch_comb"] = comb
        st.session_state["batch_all_r"] = all_r
        st.session_state["batch_per_file_names"] = per_file_names
        st.session_state["batch_per_file_workup"] = per_file_workup
        st.session_state["batch_per_file_workup_types"] = per_file_workup_types
        st.session_state["batch_per_file_workup_meta"] = per_file_workup_meta
        st.session_state["batch_per_file_summary"] = per_file_summary
        st.session_state["batch_per_file_summary_types"] = per_file_summary_types
        st.session_state["batch_per_file_summary_meta"] = per_file_summary_meta
        st.session_state["batch_per_file_assignment"] = per_file_assignment
        # Per-file metadata: UI fields + Source File + Instrument, then merge
        # extracted fields (Batch Path, Run Date) so every field shows up in
        # Method Information. Run Date from the instrument always wins.
        st.session_state["batch_per_file_meta"] = {
            n: merge_run_meta(
                {**META, "Source File": n, "Export Type": file_instruments.get(n, "")},
                batch_extracted_meta.get(n, {}),
            )
            for n in per_file_names
        }
        # QC inputs — needed by export builders for per-cell Excel coloring
        st.session_state["batch_qc_inputs"] = {
            "unit": unit,
            "std_lo": std_lo,
            "std_hi": std_hi,
            "cvs_tolerances": cvs_tolerances,
            "lloq_map": lloq_map,
            "uloq_map": uloq_map,
        }

        # ── QC pass-rate metrics (combined across all files) ──
        mc1, mc2, mc3, mc4 = st.columns(4)
        with mc1:
            pass_col = f"Pass ({std_lo}-{std_hi}%)"
            if pass_col in comb["std"].columns and len(comb["std"]):
                rate = (comb["std"][pass_col] == "✅").mean() * 100
                st.metric(f"STD Pass Rate ({std_lo}-{std_hi}%)", f"{rate:.1f}%")
            else:
                st.metric("STD Pass Rate", "—")
        with mc2:
            if "Pass" in comb["cvs"].columns and len(comb["cvs"]):
                active = comb["cvs"][comb["cvs"]["Pass"].isin(["✅", "❌"])]
                if len(active):
                    overall = (active["Pass"] == "✅").mean() * 100
                    st.metric("CVS Overall Pass Rate", f"{overall:.1f}%")
                else:
                    st.metric("CVS Overall Pass Rate", "—")
            else:
                st.metric("CVS Overall Pass Rate", "—")
        with mc3:
            n_blanks = (
                int((comb["tidy"]["isStandard"] == "Blank").sum())
                if "isStandard" in comb["tidy"].columns
                else 0
            )
            st.metric("Blank Injections (all files)", f"{n_blanks}")
        with mc4:
            n_samples = (
                int((comb["tidy"]["isStandard"] == "Sample").sum())
                if "isStandard" in comb["tidy"].columns
                else 0
            )
            st.metric("Sample Rows (all files)", f"{n_samples}")

        # ── Per-file tabs + combined Tidy/Raw at the end ─────
        tab_labels = [f"📁 {os.path.splitext(n)[0]}" for n in per_file_names]
        tab_labels += ["📋 Combined Tidy", "🗂️ Combined Raw"]
        tabs = st.tabs(tab_labels)
        for idx, fname in enumerate(per_file_names):
            with tabs[idx]:
                st.caption(f"**File:** `{fname}` — instrument: {file_instruments.get(fname,'?')}")
                file_assignment = per_file_assignment.get(fname, {})
                st.subheader("Workup", anchor=False)
                st.caption("Full review surface for this file.")
                render_workup_table(
                    per_file_workup[idx],
                    per_file_workup_types[idx],
                    unit,
                    row_meta=per_file_workup_meta[idx],
                    std_lo=std_lo,
                    std_hi=std_hi,
                    cvs_tolerances=cvs_tolerances,
                    analyte_assignment=file_assignment,
                    lloq_map=lloq_map,
                    uloq_map=uloq_map,
                )
                st.subheader("Summary", anchor=False)
                st.caption("Polished projection per inclusion toggles.")
                render_workup_table(
                    per_file_summary[idx],
                    per_file_summary_types[idx],
                    unit,
                    row_meta=per_file_summary_meta[idx],
                    std_lo=std_lo,
                    std_hi=std_hi,
                    cvs_tolerances=cvs_tolerances,
                    analyte_assignment=file_assignment,
                    lloq_map=lloq_map,
                    uloq_map=uloq_map,
                )
        with tabs[-2]:
            st.dataframe(comb["tidy"], use_container_width=True, hide_index=True)
        with tabs[-1]:
            st.dataframe(comb["raw"], use_container_width=True, hide_index=True)

        # (QC Heatmap is rendered above the Run Workup button — see pre-run preview.)

# ── Batch export buttons — live outside Run Batch block so they survive reruns ──
if st.session_state.get("batch_comb") is not None:
    _comb = st.session_state["batch_comb"]
    _all_r = st.session_state["batch_all_r"]
    _pf_wk = st.session_state.get("batch_per_file_workup", [])
    _pf_wk_tp = st.session_state.get("batch_per_file_workup_types", [])
    _pf_sm = st.session_state.get("batch_per_file_summary", [])
    _pf_sm_tp = st.session_state.get("batch_per_file_summary_types", [])

    st.divider()
    bc1, bc2 = st.columns(2)

    with bc1:
        if st.button("📥 Build Master Excel", key="master_b"):
            with st.spinner("Building master Excel…"):
                _pf_names = st.session_state.get("batch_per_file_names", [])
                _pf_wk_meta = st.session_state.get("batch_per_file_workup_meta", [])
                _pf_sm_meta = st.session_state.get("batch_per_file_summary_meta", [])
                _pf_meta = st.session_state.get("batch_per_file_meta", {})
                _pf_assign = st.session_state.get("batch_per_file_assignment", {})
                _qc_in = st.session_state.get("batch_qc_inputs", {})
                # Each per-file entry now carries its own raw_df so the
                # master workbook can write per-file Raw Data sheets.
                per_file_payload = [
                    {
                        "name": _pf_names[i],
                        "workup_df": _pf_wk[i],
                        "workup_types": _pf_wk_tp[i],
                        "summary_df": _pf_sm[i],
                        "summary_types": _pf_sm_tp[i],
                        "raw_df": _all_r["raw"][i] if i < len(_all_r["raw"]) else None,
                    }
                    for i in range(len(_pf_names))
                ]
                qc_ctx_by_file = {
                    _pf_names[i]: {
                        **_qc_in,
                        "analyte_assignment": _pf_assign.get(_pf_names[i], {}),
                        "workup_row_meta": _pf_wk_meta[i] if i < len(_pf_wk_meta) else None,
                        "summary_row_meta": _pf_sm_meta[i] if i < len(_pf_sm_meta) else None,
                    }
                    for i in range(len(_pf_names))
                }
                st.session_state["batch_master_bytes"] = build_batch_excel(
                    per_file_payload,
                    combined_tidy=_comb["tidy"],
                    per_file_meta=_pf_meta,
                    submission_wb_bytes=st.session_state.get("submission_wb"),
                    qc_context_by_file=qc_ctx_by_file,
                )
                # Filename for the download button — same composer the Dropbox
                # push uses, so the local download and the cloud copy match.
                st.session_state["batch_master_filename"] = (
                    f"{META.get('Tracking Number', '').strip() or 'batch_master'}.xlsx"
                )
        if st.session_state["batch_master_bytes"]:
            st.download_button(
                "⬇️ Download Master Excel",
                data=st.session_state["batch_master_bytes"],
                file_name=st.session_state.get("batch_master_filename", "batch_master.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_master",
            )

    with bc2:
        st.markdown("**Push to Dropbox**")
        # Same bytes that the Download button serves.
        render_dropbox_push_ui(
            key_prefix="batch_",
            meta=META,
            fallback_stem="batch_master",
            xlsx_bytes=st.session_state.get("batch_master_bytes"),
        )

    # ── Advanced: SQLite export ─────────────────────────────────
    with st.expander("Advanced: SQLite export", expanded=False):
        st.caption(
            "Programmatic archive. Two tables: `runs` (one row per source "
            "file) and `tidy_data` (every parsed row, linked by run_id)."
        )
        if st.button("💾 Build SQLite DB", key="db_b"):
            with st.spinner("Writing SQLite…"):
                # One META dict per source file. The combined tidy carries a
                # Source File column that the exporter matches against
                # meta['Source File'] to assign run_id values.
                _pf_meta_b = st.session_state.get("batch_per_file_meta", {})
                _pf_names_b = st.session_state.get("batch_per_file_names", [])
                runs_meta_list = [
                    {**(_pf_meta_b.get(n) or {}), "Source File": n} for n in _pf_names_b
                ]
                st.session_state["batch_db_bytes"] = export_sqlite(runs_meta_list, _comb["tidy"])
        if st.session_state["batch_db_bytes"]:
            _b_db_fn = f"{META.get('Tracking Number', '').strip() or 'analytical_data_batch'}.db"
            st.download_button(
                "⬇️ Download SQLite DB",
                data=st.session_state["batch_db_bytes"],
                file_name=_b_db_fn,
                mime="application/octet-stream",
                key="dl_db_b",
            )
            st.caption("Tables: `runs` · `tidy_data`")
