"""Opaque message ids.

An id carries everything the backend needs to find the message again
(account, folder, UID or Outlook EntryID). To the AI it is just a string
it hands back.
"""

from __future__ import annotations

import base64
import json
from typing import Any


class InvalidId(ValueError):
    pass


def encode_id(data: dict[str, Any]) -> str:
    raw = json.dumps(data, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    return base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")


def decode_id(value: str) -> dict[str, Any]:
    if not isinstance(value, str) or not value:
        raise InvalidId("Empty message id.")
    padded = value + "=" * (-len(value) % 4)
    try:
        data = json.loads(base64.urlsafe_b64decode(padded.encode("ascii")).decode("utf-8"))
    except Exception as exc:
        raise InvalidId(
            "Unknown message id. Please use an id returned by mail_search or mail_read."
        ) from exc
    if not isinstance(data, dict) or "a" not in data:
        raise InvalidId("Unknown message id.")
    return data
