"""
Centralized logging configuration for SpinCycle.

Call ``setup_logging()`` once at application startup. After that, any module
can do::

    import logging
    logger = logging.getLogger(__name__)
    logger.warning("Dropbox token resolution failed", exc_info=True)

…and the message is routed through the configured handler.

Environment variables
---------------------
SPINCYCLE_LOG_LEVEL
    One of DEBUG / INFO / WARNING / ERROR / CRITICAL. Default: ``INFO``.
SPINCYCLE_LOG_JSON
    Set to ``1`` to emit logs as one JSON object per line (for log
    aggregation systems like Datadog/Splunk/ELK). Default: plain text.

Both flags can be overridden at runtime by re-running ``setup_logging``
with ``force=True``.
"""

from __future__ import annotations

import json
import logging
import os
import sys


def setup_logging(*, force: bool = False) -> None:
    """Configure the root logger from environment variables.

    Idempotent by default — calling it multiple times is a no-op after
    the first call. Pass ``force=True`` to re-apply the configuration
    (useful in tests or when env vars change at runtime).
    """
    if getattr(setup_logging, "_done", False) and not force:
        return

    level_name = os.getenv("SPINCYCLE_LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_name, logging.INFO)

    handler = logging.StreamHandler(sys.stderr)

    if os.getenv("SPINCYCLE_LOG_JSON") == "1":
        handler.setFormatter(_JsonFormatter())
    else:
        handler.setFormatter(
            logging.Formatter(
                fmt="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
                datefmt="%Y-%m-%dT%H:%M:%S",
            )
        )

    root = logging.getLogger()
    root.handlers = [handler]
    root.setLevel(level)

    # Tame noisy third-party loggers — these emit DEBUG/INFO spam by default
    # that drowns out our own messages.
    for noisy in ("urllib3", "dropbox", "requests"):
        logging.getLogger(noisy).setLevel(logging.WARNING)

    setup_logging._done = True  # type: ignore[attr-defined]


class _JsonFormatter(logging.Formatter):
    """Emit each log record as a single JSON object on one line.

    Compatible with most log aggregation systems out of the box. Includes
    the standard fields plus exception info when present.
    """

    def format(self, record: logging.LogRecord) -> str:
        payload: dict[str, object] = {
            "timestamp": self.formatTime(record, "%Y-%m-%dT%H:%M:%S"),
            "level": record.levelname,
            "logger": record.name,
            "message": record.getMessage(),
        }
        if record.exc_info:
            payload["exception"] = self.formatException(record.exc_info)
        if record.stack_info:
            payload["stack"] = record.stack_info
        return json.dumps(payload, default=str)
