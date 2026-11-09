"""The interface every backend implements (Outlook COM, IMAP)."""

from __future__ import annotations

import re
from abc import ABC, abstractmethod
from dataclasses import asdict, dataclass, field
from datetime import datetime
from email.utils import getaddresses
from typing import Any

# Roles the AI can use instead of real folder names.
FOLDER_ROLES = ("inbox", "drafts", "sent", "trash", "junk", "archive")

# Usual folder names per role, lowercased. The German names stay on purpose:
# they are how a German mailbox actually labels its folders.
ROLE_NAMES = {
    "inbox": ["inbox", "posteingang"],
    "drafts": ["drafts", "draft", "entwürfe", "entwurf", "brouillons", "concepten"],
    "sent": ["sent", "sent items", "sent messages", "sent mail", "gesendet", "gesendete objekte",
             "gesendete elemente", "gesendete nachrichten"],
    "trash": ["trash", "deleted", "deleted items", "deleted messages", "papierkorb",
              "gelöscht", "gelöschte elemente", "gelöschte objekte", "bin"],
    "junk": ["junk", "spam", "junk e-mail", "junk-e-mail", "spamverdacht"],
    "archive": ["archive", "archiv", "all mail", "alle nachrichten"],
}

# German short forms people like to type instead of the role name. Kept so a
# German-speaking user can say "Papierkorb" and be understood.
ROLE_ALIASES = {
    "posteingang": "inbox", "entwürfe": "drafts", "entwurf": "drafts", "gesendet": "sent",
    "papierkorb": "trash", "spam": "junk", "archiv": "archive",
}


ALL_FOLDERS = {"*", "all", "alle", "überall"}
WEEKDAYS = ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")


def fmt_date(dt: datetime | None) -> str:
    """'2026-10-20 08:00 Mon': sortable and carries the weekday (helps with "by Friday")."""
    if not dt:
        return ""
    if dt.tzinfo is not None:
        dt = dt.astimezone()
    return f"{dt:%Y-%m-%d %H:%M} {WEEKDAYS[dt.weekday()]}"


def is_all_folders(folder: str | None) -> bool:
    return (folder or "").strip().lower() in ALL_FOLDERS


class BackendError(RuntimeError):
    """An error whose message can be shown to the user as it is."""


class NotFound(BackendError):
    pass


@dataclass
class AccountInfo:
    email: str
    name: str
    backend: str  # outlook | imap
    detail: str = ""  # z. B. "IMAP", "Exchange", "POP3"

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class FolderInfo:
    path: str
    role: str = ""
    total: int | None = None
    unread: int | None = None

    def to_dict(self) -> dict[str, Any]:
        return {k: v for k, v in asdict(self).items() if v not in (None, "")}


@dataclass
class MessageSummary:
    id: str
    account: str
    folder: str
    date: str
    sender: str
    to: str
    subject: str
    snippet: str = ""
    unread: bool = False
    flagged: bool = False
    has_attachments: bool = False

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["from"] = d.pop("sender")
        return d


@dataclass
class Attachment:
    name: str
    size: int | None = None
    content_type: str = ""


@dataclass
class Message:
    summary: MessageSummary
    cc: str = ""
    body: str = ""
    attachments: list[Attachment] = field(default_factory=list)
    message_id: str = ""
    in_reply_to: str = ""
    references: list[str] = field(default_factory=list)
    reply_to: str = ""
    sort_key: datetime | None = None


@dataclass
class DraftRequest:
    account: str | None
    to: list[str]
    cc: list[str]
    bcc: list[str]
    subject: str | None
    body: str
    reply_to_id: str | None = None
    reply_all: bool = False
    open_window: bool = False
    forward_of: str | None = None


@dataclass
class DraftResult:
    id: str
    account: str
    folder: str
    subject: str
    to: str
    cc: str = ""
    opened_window: bool = False
    note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {k: v for k, v in asdict(self).items() if v not in (None, "")}


@dataclass
class SearchQuery:
    folder: str = "inbox"
    text: str | None = None
    sender: str | None = None
    since: datetime | None = None
    until: datetime | None = None
    unread_only: bool = False
    limit: int = 20


class MailBackend(ABC):
    kind: str = ""

    @abstractmethod
    def accounts(self) -> list[AccountInfo]: ...

    @abstractmethod
    def owns(self, account: str) -> bool: ...

    @abstractmethod
    def folders(self, account: str) -> list[FolderInfo]: ...

    @abstractmethod
    def search(self, account: str, query: SearchQuery) -> list[MessageSummary]: ...

    @abstractmethod
    def read(self, ref: dict[str, Any]) -> Message: ...

    @abstractmethod
    def thread(self, ref: dict[str, Any], limit: int = 15) -> list[Message]: ...

    @abstractmethod
    def create_draft(self, account: str, req: DraftRequest) -> DraftResult: ...

    @abstractmethod
    def move(self, ref: dict[str, Any], folder: str) -> str: ...

    @abstractmethod
    def mark(self, ref: dict[str, Any], read: bool | None, flagged: bool | None) -> None: ...

    @abstractmethod
    def send_draft(self, ref: dict[str, Any]) -> str: ...

    @abstractmethod
    def discard_draft(self, ref: dict[str, Any]) -> None:
        """Deletes a draft (drafts only, never received mail)."""

    @abstractmethod
    def attachment(self, ref: dict[str, Any], name: str) -> tuple[str, bytes]:
        """Returns the file name and contents of an attachment."""

    def delete(self, ref: dict[str, Any], permanent: bool = False) -> str:
        """Deletes a message. By default it goes to the trash and can be restored.

        ``permanent=True`` means gone for good; backends that cannot do that
        must say so instead of quietly moving the mail to the trash.
        """
        if permanent:
            raise BackendError("This mailbox cannot delete mail permanently; it can only move it to the trash.")
        return self.move(ref, "trash")

    def check(self) -> str:
        """Short connection test; returns a success message or raises BackendError."""
        accs = self.accounts()
        return f"{len(accs)} account(s) reachable"

    def close(self) -> None:  # pragma: no cover - optional
        pass


def _display_addr(name: str, addr: str) -> str:
    """Like formataddr, but without encoding non-ASCII (Outlook and the mail library do that)."""
    if not name:
        return addr
    if re.search(r'[,;:<>@"()\[\]\\.]', name):
        name = '"' + name.replace("\\", "\\\\").replace('"', '\\"') + '"'
    return f"{name} <{addr}>"


def split_addresses(value: str | list[str] | None) -> list[str]:
    """Splits recipients robustly, including "Müller, Hans <hans@x.de>; a@b.de"."""
    if not value:
        return []
    items = value if isinstance(value, list) else [value]
    out: list[str] = []
    for item in items:
        if not item or not str(item).strip():
            continue
        text = str(item).replace(";", ",")
        # "Müller, Hans <hans@x.de>": quote a name that holds a comma before <...>
        text = re.sub(r'(^|,)\s*([^"<>,@]+,[^"<>,@]+?)\s*<', lambda m: f'{m.group(1)}"{m.group(2).strip()}" <', text)
        for name, addr in getaddresses([text]):
            if "@" not in addr:
                continue
            out.append(_display_addr(name, addr))
    seen: set[str] = set()
    unique = []
    for a in out:
        key = a.lower()
        if key not in seen:
            seen.add(key)
            unique.append(a)
    return unique
