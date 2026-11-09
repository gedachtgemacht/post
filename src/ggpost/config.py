"""post. configuration.

The file lives in the user profile (Windows: %APPDATA%\\post\\config.json).
Passwords are never in this file; they live in the operating system's
keychain (see ``vault.py``).
"""

from __future__ import annotations

import json
import logging
import os
import sys
import tempfile
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

CONFIG_VERSION = 1


def config_dir() -> Path:
    override = os.environ.get("POST_CONFIG_DIR")
    if override:
        return Path(override)
    if sys.platform == "win32":
        base = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        return Path(base) / "post"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "post"
    base = os.environ.get("XDG_CONFIG_HOME") or str(Path.home() / ".config")
    return Path(base) / "post"


def config_path() -> Path:
    return config_dir() / "config.json"


@dataclass
class ServerSettings:
    host: str
    port: int
    security: str = "ssl"  # ssl | starttls | plain
    username: str = ""

    @classmethod
    def from_dict(cls, d: dict[str, Any] | None) -> ServerSettings | None:
        if not d:
            return None
        return cls(
            host=str(d["host"]),
            port=int(d["port"]),
            security=str(d.get("security", "ssl")),
            username=str(d.get("username", "")),
        )


@dataclass
class Source:
    """One mail source: either classic Outlook (all or selected accounts) or a
    single IMAP mailbox."""

    kind: str  # "outlook" | "imap"
    email: str = ""  # imap only
    label: str = ""
    imap: ServerSettings | None = None
    smtp: ServerSettings | None = None
    outlook_accounts: list[str] | None = None  # None = every account Outlook has
    signature: str = ""  # IMAP only: Outlook adds its own signature

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> Source:
        accounts = d.get("outlook_accounts")
        return cls(
            kind=str(d["kind"]),
            email=str(d.get("email", "")),
            label=str(d.get("label", "")),
            imap=ServerSettings.from_dict(d.get("imap")),
            smtp=ServerSettings.from_dict(d.get("smtp")),
            outlook_accounts=[str(a) for a in accounts] if accounts is not None else None,
            signature=str(d.get("signature", "")),
        )

    def to_dict(self) -> dict[str, Any]:
        out = asdict(self)
        return {k: v for k, v in out.items() if v not in (None, "")}


@dataclass
class Permissions:
    modify: bool = False  # move (including to the trash), delete, mark as read, flag
    send: bool = False  # actually send a draft

    @classmethod
    def from_dict(cls, d: dict[str, Any] | None) -> Permissions:
        d = d or {}
        return cls(modify=bool(d.get("modify", False)), send=bool(d.get("send", False)))


@dataclass
class Config:
    sources: list[Source] = field(default_factory=list)
    permissions: Permissions = field(default_factory=Permissions)
    max_body_chars: int = 20000
    max_results: int = 50
    audit: bool = True
    version: int = CONFIG_VERSION

    @classmethod
    def from_dict(cls, d: dict[str, Any]) -> Config:
        return cls(
            sources=[Source.from_dict(s) for s in d.get("sources", [])],
            permissions=Permissions.from_dict(d.get("permissions")),
            max_body_chars=int(d.get("max_body_chars", 20000)),
            max_results=int(d.get("max_results", 50)),
            audit=bool(d.get("audit", True)),
            version=int(d.get("version", CONFIG_VERSION)),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "version": self.version,
            "sources": [s.to_dict() for s in self.sources],
            "permissions": asdict(self.permissions),
            "max_body_chars": self.max_body_chars,
            "max_results": self.max_results,
            "audit": self.audit,
        }


def load_config(path: Path | None = None) -> Config | None:
    """Loads the settings. A damaged file is set aside
    (config.json.broken-…) so setup can start over."""
    path = path or config_path()
    if not path.exists():
        return None
    try:
        with path.open("r", encoding="utf-8") as fh:
            return Config.from_dict(json.load(fh))
    except (ValueError, KeyError, TypeError) as exc:
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        broken = path.with_name(f"{path.name}.broken-{stamp}")
        try:
            os.replace(path, broken)
        except OSError:
            pass
        logging.getLogger(__name__).warning("Settings damaged (%s), kept as %s", exc, broken)
        return None


def save_config(cfg: Config, path: Path | None = None) -> Path:
    path = path or config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    data = json.dumps(cfg.to_dict(), indent=2, ensure_ascii=False)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix=".config-", suffix=".json")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(data)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    return path


def env_flag(name: str) -> bool | None:
    """Reads a yes/no value from the environment (e.g. from the Claude Desktop
    extension). An empty value means "not set"."""
    raw = os.environ.get(name, "").strip().lower()
    if not raw or raw.startswith("${"):
        return None
    return raw in {"1", "true", "yes", "ja", "on"}


def env_str(name: str) -> str:
    raw = os.environ.get(name, "").strip()
    if raw.startswith("${"):  # placeholder from manifest.json that was never filled in
        return ""
    return raw
