"""
Shared pytest configuration and fixtures for the SpinCycle test suite.

Currently empty — all tests target pure functions in `spincycle_utils.py`
which has no Streamlit dependency. If we later add tests that need to
import `app.py` itself, we'll need to mock streamlit here to prevent
Streamlit's runtime calls (st.set_page_config, st.file_uploader, etc.)
from blowing up at module-import time.
"""

from __future__ import annotations
