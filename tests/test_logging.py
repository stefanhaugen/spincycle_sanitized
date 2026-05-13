"""
Tests for the spincycle_logging module.

These tests verify the setup_logging() side effects: that the root logger
gets the expected handler and level, that env vars are honored, and that
the JSON formatter produces parseable output.
"""

from __future__ import annotations

import json
import logging
import os

import pytest

from spincycle_logging import _JsonFormatter, setup_logging


@pytest.fixture(autouse=True)
def _reset_logging():
    """Reset the setup_logging idempotency flag and the env between tests."""
    # Reset the flag so each test gets a fresh setup
    if hasattr(setup_logging, "_done"):
        delattr(setup_logging, "_done")

    # Snapshot env vars we touch
    saved = {k: os.environ.get(k) for k in ("SPINCYCLE_LOG_LEVEL", "SPINCYCLE_LOG_JSON")}
    yield
    # Restore env
    for k, v in saved.items():
        if v is None:
            os.environ.pop(k, None)
        else:
            os.environ[k] = v


class TestSetupLogging:
    def test_default_level_is_info(self, monkeypatch):
        monkeypatch.delenv("SPINCYCLE_LOG_LEVEL", raising=False)
        setup_logging(force=True)
        assert logging.getLogger().level == logging.INFO

    def test_env_var_overrides_level(self, monkeypatch):
        monkeypatch.setenv("SPINCYCLE_LOG_LEVEL", "DEBUG")
        setup_logging(force=True)
        assert logging.getLogger().level == logging.DEBUG

    def test_invalid_level_falls_back_to_info(self, monkeypatch):
        monkeypatch.setenv("SPINCYCLE_LOG_LEVEL", "NOT_A_LEVEL")
        setup_logging(force=True)
        assert logging.getLogger().level == logging.INFO

    def test_handler_installed_on_root(self, monkeypatch):
        monkeypatch.delenv("SPINCYCLE_LOG_JSON", raising=False)
        setup_logging(force=True)
        assert len(logging.getLogger().handlers) == 1

    def test_idempotent_without_force(self, monkeypatch):
        monkeypatch.delenv("SPINCYCLE_LOG_JSON", raising=False)
        setup_logging()
        first = logging.getLogger().handlers[0]
        setup_logging()  # second call should be a no-op
        assert logging.getLogger().handlers[0] is first

    def test_force_replaces_handler(self, monkeypatch):
        monkeypatch.delenv("SPINCYCLE_LOG_JSON", raising=False)
        setup_logging()
        first = logging.getLogger().handlers[0]
        setup_logging(force=True)
        assert logging.getLogger().handlers[0] is not first


class TestJsonFormatter:
    def test_basic_record_produces_valid_json(self):
        record = logging.LogRecord(
            name="test_logger",
            level=logging.INFO,
            pathname=__file__,
            lineno=1,
            msg="hello %s",
            args=("world",),
            exc_info=None,
        )
        formatted = _JsonFormatter().format(record)
        parsed = json.loads(formatted)
        assert parsed["message"] == "hello world"
        assert parsed["level"] == "INFO"
        assert parsed["logger"] == "test_logger"
        assert "timestamp" in parsed

    def test_exception_info_included(self):
        try:
            raise ValueError("boom")
        except ValueError:
            import sys

            record = logging.LogRecord(
                name="test_logger",
                level=logging.ERROR,
                pathname=__file__,
                lineno=1,
                msg="problem",
                args=(),
                exc_info=sys.exc_info(),
            )
        formatted = _JsonFormatter().format(record)
        parsed = json.loads(formatted)
        assert "exception" in parsed
        assert "ValueError" in parsed["exception"]
        assert "boom" in parsed["exception"]


class TestThirdPartyLoggersTamed:
    def test_urllib3_set_to_warning(self):
        setup_logging(force=True)
        assert logging.getLogger("urllib3").level == logging.WARNING

    def test_dropbox_set_to_warning(self):
        setup_logging(force=True)
        assert logging.getLogger("dropbox").level == logging.WARNING
