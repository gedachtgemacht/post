"""Audit log: what the AI read or wrote, and when.

Metadata only (tool, account, counts, outcome), never mail contents.
"""

from __future__ import annotations

import json
import os
import threading
from datetime import datetime
from pathlib import Path
from typing import Any

from ggpost.config import config_dir

_LOCK = threading.Lock()
MAX_BYTES = 5 * 1024 * 1024


def audit_path() -> Path:
    return config_dir() / "audit.log"


def record(tool: str, ok: bool, **fields: Any) -> None:
    entry = {"ts": datetime.now().astimezone().isoformat(timespec="seconds"), "tool": tool, "ok": ok}
    entry.update({k: v for k, v in fields.items() if v is not None})
    line = json.dumps(entry, ensure_ascii=False)
    path = audit_path()
    try:
        with _LOCK:
            path.parent.mkdir(parents=True, exist_ok=True)
            if path.exists() and path.stat().st_size > MAX_BYTES:
                os.replace(path, path.with_suffix(".log.1"))
            with path.open("a", encoding="utf-8") as fh:
                fh.write(line + "\n")
    except OSError:
        pass  # logging must never block the actual work
