"""Outlook backend: drives classic Outlook for Windows from the outside.

This is not an add-in. post. talks to the running Outlook as a program of
its own, through its object model (COM). That makes every account set up in
Outlook work: Exchange, IMAP, POP3, PST archives. No Microsoft 365
subscription, no app registration, no password inside post.

Every COM call runs in exactly one dedicated thread, because Outlook objects
are only valid in the thread that created them.
"""

from __future__ import annotations

import logging
import os
import re
import shutil
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor
from concurrent.futures import TimeoutError as FuturesTimeout
from datetime import datetime, timezone
from email.utils import formataddr
from typing import Any, Callable, TypeVar

from ggpost.backends.base import (
    ROLE_ALIASES,
    ROLE_NAMES,
    AccountInfo,
    Attachment,
    BackendError,
    DraftRequest,
    DraftResult,
    FolderInfo,
    MailBackend,
    Message,
    MessageSummary,
    NotFound,
    SearchQuery,
    fmt_date,
    is_all_folders,
)
from ggpost.config import Source
from ggpost.ids import decode_id, encode_id
from ggpost.text import html_to_text, normalize_whitespace, text_to_html

log = logging.getLogger(__name__)
T = TypeVar("T")

OL_FOLDER = {"inbox": 6, "sent": 5, "drafts": 16, "trash": 3, "junk": 23}
OL_STORE_PUBLIC = 2  # ExchangeStoreType: public folders
OL_MAIL_ITEM = 0
OL_CLASS_MAIL = 43
OL_FORMAT_PLAIN = 1
OL_FLAG_MARKED = 2
OL_TO, OL_CC, OL_BCC = 1, 2, 3
ACCOUNT_TYPES = {0: "Exchange", 1: "IMAP", 2: "POP3", 3: "HTTP", 4: "Exchange ActiveSync", 5: "Sonstiges"}
DISPID_SEND_USING_ACCOUNT = 64209

PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

# HRESULTs that mean the connection to Outlook has to be rebuilt.
_RECONNECT_HRESULTS = {-2147023174, -2147417848, -2147023170, -2147418113}
_SECURITY_DENIED = {-2147467260, -2147024891}
SECURITY_HELP = (
    "Outlook denied access. Usually the security prompt was declined, or Windows reports no "
    "up-to-date antivirus. The setting lives at Outlook > File > Options > Trust Center > "
    "Trust Center Settings > Programmatic Access (this needs administrator rights)."
)

DASL_SUBJECT = "urn:schemas:httpmail:subject"
DASL_BODY = "urn:schemas:httpmail:textdescription"
DASL_FROMNAME = "urn:schemas:httpmail:fromname"
DASL_FROMEMAIL = "urn:schemas:httpmail:fromemail"
DASL_READ = "urn:schemas:httpmail:read"
DASL_TOPIC = "urn:schemas:httpmail:thread-topic"


# ── Thread-Steuerung ─────────────────────────────────────────────────────


class ComRunner:
    """Runs every call in a dedicated, COM-initialised thread."""

    def __init__(self) -> None:
        def init() -> None:
            import pythoncom  # noqa: PLC0415

            pythoncom.CoInitialize()

        self._pool = ThreadPoolExecutor(max_workers=1, thread_name_prefix="post-outlook", initializer=init)

    def call(self, fn: Callable[[], T], timeout: float = 120.0) -> T:
        future = self._pool.submit(fn)
        try:
            return future.result(timeout=timeout)
        except FuturesTimeout as exc:
            raise BackendError(
                "Outlook is not responding. Often a security prompt from Outlook is waiting in the "
                "background (\"A program is trying to access e-mail address information\"). Please look "
                "for the Outlook window in the taskbar and allow the access."
            ) from exc


class DirectRunner:
    """For tests: runs directly in the calling thread."""

    def __init__(self) -> None:
        self._lock = threading.RLock()

    def call(self, fn: Callable[[], T], timeout: float = 0) -> T:
        with self._lock:
            return fn()


def default_app_factory() -> Any:
    import win32com.client.dynamic  # noqa: PLC0415

    return win32com.client.dynamic.Dispatch("Outlook.Application")


DISP_E_EXCEPTION = -2147352567


def _hresult(exc: Exception) -> int | None:
    """pywin32 reports errors as (hresult, text, excepinfo, argerr). With
    DISP_E_EXCEPTION the real code sits in excepinfo[5]."""
    args = getattr(exc, "args", ())
    code = args[0] if args and isinstance(args[0], int) else getattr(exc, "hresult", None)
    if code == DISP_E_EXCEPTION:
        info = getattr(exc, "excepinfo", None) or (args[2] if len(args) > 2 else None)
        if isinstance(info, tuple) and len(info) > 5 and isinstance(info[5], int):
            return info[5]
    return code


# ── Hilfsfunktionen ──────────────────────────────────────────────────────


def com_datetime(value: Any) -> datetime | None:
    """Outlook liefert Ortszeit, pywin32 etikettiert sie teils als UTC.
    We take the clock time literally as local time. Year 4501 means "no date"."""
    if value is None:
        return None
    try:
        if value.year >= 4500:
            return None
        naive = datetime(value.year, value.month, value.day, value.hour, value.minute, value.second)
        return naive.astimezone()
    except Exception:
        return None


_fmt = fmt_date


def _dasl_escape(value: str) -> str:
    return value.replace("'", "''").replace("%", "").replace('"', "")


def build_filter(query: SearchQuery) -> str:
    """Builds a DASL filter for Items.Restrict. With several search words, all of
    them have to appear (in subject, body or sender)."""
    parts: list[str] = []
    if query.unread_only:
        parts.append(f'"{DASL_READ}" = 0')
    if query.text:
        for word in query.text.split():
            w = _dasl_escape(word)
            if not w:
                continue
            ors = " OR ".join(f"\"{p}\" LIKE '%{w}%'" for p in (DASL_SUBJECT, DASL_BODY, DASL_FROMNAME))
            parts.append(f"({ors})")
    if query.sender:
        s = _dasl_escape(query.sender.strip())
        if s:
            parts.append(f"(\"{DASL_FROMNAME}\" LIKE '%{s}%' OR \"{DASL_FROMEMAIL}\" LIKE '%{s}%')")
    return " AND ".join(parts)


def _safe(fn: Callable[[], T], default: T) -> T:
    try:
        value = fn()
        return default if value is None else value
    except Exception:
        return default


def _sender(item: Any) -> str:
    name = _safe(lambda: item.SenderName, "")
    addr = _safe(lambda: item.SenderEmailAddress, "")
    if _safe(lambda: item.SenderEmailType, "") == "EX":
        smtp = _safe(lambda: item.Sender.GetExchangeUser().PrimarySmtpAddress, "")
        addr = smtp or addr
    if addr and "@" in addr:
        return formataddr((name, addr)) if name and name != addr else addr
    return name or addr


def _recipients(item: Any) -> tuple[str, str]:
    to: list[str] = []
    cc: list[str] = []
    recips = _safe(lambda: item.Recipients, None)
    count = _safe(lambda: recips.Count, 0) if recips is not None else 0
    for i in range(1, count + 1):
        r = recips.Item(i)
        name = _safe(lambda: r.Name, "")
        addr = _safe(lambda: r.Address, "")
        if "@" not in (addr or ""):
            addr = _safe(lambda: r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS), addr)
        label = formataddr((name, addr)) if addr and "@" in addr and name != addr else (name or addr)
        rtype = _safe(lambda: r.Type, OL_TO)
        if rtype == OL_TO:
            to.append(label)
        elif rtype == OL_CC:
            cc.append(label)
    return ", ".join(to), ", ".join(cc)


def _set_account(item: Any, account: Any) -> None:
    try:
        item.SendUsingAccount = account
        return
    except Exception:
        pass
    try:  # the known detour for late-bound objects in pywin32
        item._oleobj_.Invoke(DISPID_SEND_USING_ACCOUNT, 0, 8, 0, account)
    except Exception as exc:
        log.warning("Could not set the sending account: %s", exc)


def insert_body(item: Any, text: str) -> None:
    """Puts the text written by the AI at the top, above signature and quote."""
    fmt = _safe(lambda: item.BodyFormat, 2)
    if fmt == OL_FORMAT_PLAIN:
        current = _safe(lambda: item.Body, "")
        item.Body = text.rstrip() + ("\r\n\r\n" + current if current else "")
        return
    html_block = text_to_html(text)
    current = _safe(lambda: item.HTMLBody, "") or ""
    if "MsoNormal" in current:
        html_block = html_block.replace("<p>", '<p class="MsoNormal">')
    else:
        html_block = (
            '<div style="font-family:Calibri,Arial,Helvetica,sans-serif;font-size:11pt">'
            + html_block
            + "</div>"
        )
    match = re.search(r"<body[^>]*>", current, re.I)
    if match:
        new = current[: match.end()] + html_block + current[match.end() :]
    elif current.strip():
        new = html_block + current
    else:
        new = f"<html><body>{html_block}</body></html>"
    item.HTMLBody = new


# ── Backend ──────────────────────────────────────────────────────────────


class Entry:
    """One mailbox in Outlook: an account or an extra store (a shared mailbox,
    a PST archive)."""

    def __init__(self, key: str, name: str, detail: str, store: Any, account: Any = None) -> None:
        self.key = key
        self.name = name
        self.detail = detail
        self.store = store
        self.account = account


class OutlookBackend(MailBackend):
    kind = "outlook"

    def __init__(
        self,
        source: Source,
        app_factory: Callable[[], Any] | None = None,
        runner: Any = None,
    ) -> None:
        self.source = source
        self._factory = app_factory or default_app_factory
        self._runner = runner or ComRunner()
        self._app: Any = None
        self._ns: Any = None

    # -- Connection --

    def _namespace(self) -> Any:
        if self._ns is None:
            try:
                self._app = self._factory()
                self._ns = self._app.GetNamespace("MAPI")
            except Exception as exc:
                self._app = self._ns = None
                raise BackendError(
                    "Classic Outlook could not be started. Is Outlook installed and is there "
                    f"a profile set up? ({exc})"
                ) from exc
        return self._ns

    def _call(self, fn: Callable[[], T]) -> T:
        def guarded() -> T:
            for attempt in (1, 2):
                try:
                    return fn()
                except BackendError:
                    raise
                except Exception as exc:
                    code = _hresult(exc)
                    if code in _RECONNECT_HRESULTS and attempt == 1:
                        log.info("Outlook connection lost, reconnecting")
                        self._app = self._ns = None
                        continue
                    if code in _SECURITY_DENIED:
                        raise BackendError(SECURITY_HELP) from exc
                    raise BackendError(f"Outlook reports an error: {exc}") from exc
            raise AssertionError("unreachable")

        return self._runner.call(guarded)

    # -- Accounts, stores, folders --

    def _entries(self) -> list[Entry]:
        ns = self._namespace()
        wanted = {a.lower() for a in (self.source.outlook_accounts or [])}
        out: list[Entry] = []
        seen_stores: set[str] = set()
        accounts = ns.Accounts
        for i in range(1, accounts.Count + 1):
            acc = accounts.Item(i)
            email = (_safe(lambda: acc.SmtpAddress, "") or "").strip()
            if not email:
                name = _safe(lambda: acc.DisplayName, "")
                email = name if "@" in name else ""
            store = _safe(lambda: acc.DeliveryStore, None)
            if store is not None:
                seen_stores.add(_safe(lambda: store.StoreID, ""))
            if not email:
                continue
            kind = ACCOUNT_TYPES.get(_safe(lambda: acc.AccountType, 5), "Sonstiges")
            out.append(Entry(email, _safe(lambda: acc.DisplayName, email), f"Outlook, {kind}", store or ns.DefaultStore, acc))
        stores = _safe(lambda: ns.Stores, None)
        count = _safe(lambda: stores.Count, 0) if stores is not None else 0
        for i in range(1, count + 1):
            st = stores.Item(i)
            sid = _safe(lambda: st.StoreID, "")
            if not sid or sid in seen_stores:
                continue
            if _safe(lambda: st.ExchangeStoreType, 3) == OL_STORE_PUBLIC:
                continue
            name = (_safe(lambda: st.DisplayName, "") or "").strip()
            if not name:
                continue
            is_file = bool(_safe(lambda: st.IsDataFileStore, False)) and bool(_safe(lambda: st.FilePath, ""))
            detail = "Outlook, archive file (PST)" if is_file else "Outlook, shared mailbox"
            seen_stores.add(sid)
            out.append(Entry(name, name, detail, st, None))
        if wanted:
            out = [e for e in out if e.key.lower() in wanted]
        return out

    def _entry(self, key: str | None) -> Entry:
        entries = self._entries()
        if not entries:
            raise BackendError("Outlook has no matching account set up.")
        if not key:
            return entries[0]
        for e in entries:
            if e.key.lower() == key.lower():
                return e
        raise NotFound(f"Mailbox {key} is not shared in Outlook.")

    def _default_folder(self, store: Any, role: str) -> Any:
        const = OL_FOLDER.get(role)
        if const is not None:
            try:
                folder = store.GetDefaultFolder(const)
                if folder is not None:
                    return folder
            except Exception:
                pass
        root = store.GetRootFolder()
        names = ROLE_NAMES.get(role, [])
        for folder in self._iter_children(root):
            if (_safe(lambda: folder.Name, "") or "").strip().lower() in names:
                return folder
        raise NotFound(f"No folder found for '{role}'.")

    @staticmethod
    def _iter_children(folder: Any) -> list[Any]:
        children = _safe(lambda: folder.Folders, None)
        if children is None:
            return []
        return [children.Item(i) for i in range(1, _safe(lambda: children.Count, 0) + 1)]

    def _child_named(self, folder: Any, name: str) -> Any:
        for child in self._iter_children(folder):
            if (_safe(lambda: child.Name, "") or "").strip().lower() == name.strip().lower():
                return child
        return None

    def _resolve_folder(self, store: Any, name: str | None) -> Any:
        key = (name or "inbox").strip()
        low = key.lower()
        role = low if low in OL_FOLDER or low == "archive" else ROLE_ALIASES.get(low)
        if role:
            return self._default_folder(store, role)
        current = store.GetRootFolder()
        for index, part in enumerate(p for p in re.split(r"[\\/]", key) if p):
            match = self._child_named(current, part)
            if match is None and index == 0:
                # "Customers" may also mean "Inbox/Customers"
                match = self._child_named(self._default_folder(store, "inbox"), part)
            if match is None:
                raise NotFound(f"Folder '{key}' not found. mail_folders lists the available ones.")
            current = match
        return current

    def _mail_folders(self, store: Any, skip_roles: tuple[str, ...] = ()) -> list[Any]:
        skip = set()
        for role in skip_roles:
            skip.add(_safe(lambda: self._default_folder(store, role).EntryID, "-"))
        out: list[Any] = []

        def walk(folder: Any, depth: int) -> None:
            if depth > 4 or len(out) >= 300:
                return
            for child in self._iter_children(folder):
                if _safe(lambda: child.EntryID, "") in skip:
                    continue
                if _safe(lambda: child.DefaultItemType, OL_MAIL_ITEM) == OL_MAIL_ITEM:
                    out.append(child)
                walk(child, depth + 1)

        walk(store.GetRootFolder(), 0)
        return out

    def _folder_path(self, folder: Any) -> str:
        raw = _safe(lambda: folder.FolderPath, "") or _safe(lambda: folder.Name, "")
        parts = [p for p in raw.split("\\") if p]
        return "/".join(parts[1:]) if len(parts) > 1 else (parts[0] if parts else "")

    def accounts(self) -> list[AccountInfo]:
        return self._call(lambda: [AccountInfo(email=e.key, name=e.name, backend="outlook", detail=e.detail) for e in self._entries()])

    def owns(self, account: str) -> bool:
        try:
            return any(a.email.lower() == account.lower() for a in self.accounts())
        except BackendError:
            return False

    def check(self) -> str:
        def op() -> str:
            entries = self._entries()
            if not entries:
                raise BackendError("Outlook is running, but no account is set up.")
            first = entries[0]
            inbox = self._default_folder(first.store, "inbox")
            total = _safe(lambda: inbox.Items.Count, 0)
            # read a sender once, so a security prompt comes now and not later
            item = _safe(lambda: inbox.Items.GetFirst(), None)
            if item is not None:
                _sender(item)
            n = len(entries)
            return f"{n} {'mailbox' if n == 1 else 'mailboxes'}, inbox {first.key}: {total} mails"

        return self._call(op)

    def folders(self, account: str) -> list[FolderInfo]:
        def op() -> list[FolderInfo]:
            store = self._entry(account).store
            roles: dict[str, str] = {}
            for role in OL_FOLDER:
                try:
                    roles[_safe(lambda: self._default_folder(store, role).EntryID, "")] = role
                except Exception:
                    pass
            return [
                FolderInfo(
                    path=self._folder_path(f),
                    role=roles.get(_safe(lambda: f.EntryID, "-"), ""),
                    total=_safe(lambda: f.Items.Count, None),
                    unread=_safe(lambda: f.UnReadItemCount, None),
                )
                for f in self._mail_folders(store)
            ]

        return self._call(op)

    # -- Reading --

    def _summary(self, key: str, item: Any, folder_path: str, with_snippet: bool = True) -> MessageSummary:
        dt = com_datetime(_safe(lambda: item.ReceivedTime, None)) or com_datetime(_safe(lambda: item.SentOn, None))
        snippet = ""
        if with_snippet:
            snippet = re.sub(r"\s+", " ", (_safe(lambda: item.Body, "") or "")[:800]).strip()[:240]
        return MessageSummary(
            id=encode_id({"a": key, "e": item.EntryID, "s": _safe(lambda: item.Parent.StoreID, "")}),
            account=key,
            folder=folder_path,
            date=_fmt(dt),
            sender=_sender(item),
            to=(_safe(lambda: item.To, "") or "")[:300],
            subject=_safe(lambda: item.Subject, ""),
            snippet=snippet,
            unread=bool(_safe(lambda: item.UnRead, False)),
            flagged=_safe(lambda: item.FlagStatus, 0) == OL_FLAG_MARKED,
            has_attachments=_safe(lambda: item.Attachments.Count, 0) > 0,
        )

    def _search_folder(self, key: str, folder: Any, query: SearchQuery) -> list[MessageSummary]:
        path = self._folder_path(folder)
        items = folder.Items
        flt = build_filter(query)
        if flt:
            items = items.Restrict("@SQL=" + flt)
        items.Sort("[ReceivedTime]", True)
        out: list[MessageSummary] = []
        item = items.GetFirst()
        scanned = 0
        while item is not None and len(out) < query.limit and scanned < 5000:
            scanned += 1
            if _safe(lambda: item.Class, 0) == OL_CLASS_MAIL:
                dt = com_datetime(_safe(lambda: item.ReceivedTime, None))
                if query.since and dt and dt < query.since:
                    break
                if not (query.until and dt and dt >= query.until):
                    out.append(self._summary(key, item, path))
            item = items.GetNext()
        return out

    def search(self, account: str, query: SearchQuery) -> list[MessageSummary]:
        def op() -> list[MessageSummary]:
            entry = self._entry(account)
            if is_all_folders(query.folder):
                results: list[MessageSummary] = []
                for folder in self._mail_folders(entry.store, skip_roles=("trash", "junk", "drafts"))[:60]:
                    results.extend(self._search_folder(entry.key, folder, query))
                results.sort(key=lambda m: m.date, reverse=True)
                return results[: query.limit]
            folder = self._resolve_folder(entry.store, query.folder)
            return self._search_folder(entry.key, folder, query)

        return self._call(op)

    def _get_item(self, ref: dict[str, Any]) -> Any:
        ns = self._namespace()
        try:
            if ref.get("s"):
                return ns.GetItemFromID(ref["e"], ref["s"])
            return ns.GetItemFromID(ref["e"])
        except Exception as exc:
            raise NotFound("The message was not found (it may have been moved or deleted).") from exc

    def _to_message(self, key: str, item: Any) -> Message:
        summary = self._summary(key, item, self._folder_path(_safe(lambda: item.Parent, None)), with_snippet=False)
        to, cc = _recipients(item)
        if to:
            summary.to = to
        body = _safe(lambda: item.Body, "") or ""
        if not body.strip():
            body = html_to_text(_safe(lambda: item.HTMLBody, "") or "")
        attachments: list[Attachment] = []
        for _i, a in self._visible_attachments(item):
            attachments.append(
                Attachment(name=_safe(lambda: a.FileName, "") or _safe(lambda: a.DisplayName, ""), size=_safe(lambda: a.Size, None))
            )
        summary.has_attachments = bool(attachments)
        dt = com_datetime(_safe(lambda: item.ReceivedTime, None)) or com_datetime(_safe(lambda: item.SentOn, None))
        return Message(
            summary=summary,
            cc=cc,
            body=normalize_whitespace(body),
            attachments=attachments,
            message_id=_safe(lambda: item.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID), ""),
            sort_key=dt,
        )

    @staticmethod
    def _visible_attachments(item: Any) -> list[tuple[int, Any]]:
        atts = _safe(lambda: item.Attachments, None)
        out = []
        for i in range(1, (_safe(lambda: atts.Count, 0) if atts is not None else 0) + 1):
            a = atts.Item(i)
            if _safe(lambda: a.PropertyAccessor.GetProperty(PR_ATTACHMENT_HIDDEN), False):
                continue
            out.append((i, a))
        return out

    def read(self, ref: dict[str, Any]) -> Message:
        return self._call(lambda: self._to_message(ref["a"], self._get_item(ref)))

    def thread(self, ref: dict[str, Any], limit: int = 15) -> list[Message]:
        def op() -> list[Message]:
            item = self._get_item(ref)
            key = ref["a"]
            ns = self._namespace()
            store_id = ref.get("s") or None
            msgs: dict[str, Message] = {}
            conv = _safe(lambda: item.GetConversation(), None)
            if conv is not None:
                table = conv.GetTable()
                while not table.EndOfTable and len(msgs) < limit * 2:
                    row = table.GetNextRow()
                    entry = _safe(lambda: row.Item("EntryID"), "") or _safe(lambda: row("EntryID"), "")
                    if not entry or entry in msgs:
                        continue
                    it = _safe(lambda: ns.GetItemFromID(entry, store_id) if store_id else ns.GetItemFromID(entry), None)
                    if it is not None and _safe(lambda: it.Class, 0) == OL_CLASS_MAIL:
                        msgs[entry] = self._to_message(key, it)
            if len(msgs) <= 1:
                topic = _dasl_escape(_safe(lambda: item.ConversationTopic, "") or "")
                store = self._entry(key).store
                for role in ("inbox", "sent"):
                    try:
                        folder = self._default_folder(store, role)
                    except NotFound:
                        continue
                    found = folder.Items.Restrict(f"@SQL=\"{DASL_TOPIC}\" = '{topic}'") if topic else None
                    it = found.GetFirst() if found is not None else None
                    while it is not None and len(msgs) < limit * 2:
                        if _safe(lambda: it.Class, 0) == OL_CLASS_MAIL:
                            msgs.setdefault(it.EntryID, self._to_message(key, it))
                        it = found.GetNext()
            msgs.setdefault(item.EntryID, self._to_message(key, item))
            # drafts have no date: sort them last
            ordered = sorted(msgs.values(), key=lambda m: m.sort_key or datetime.max.replace(tzinfo=timezone.utc))
            return ordered[-limit:]

        return self._call(op)

    def attachment(self, ref: dict[str, Any], name: str) -> tuple[str, bytes]:
        def op() -> tuple[str, bytes]:
            item = self._get_item(ref)
            visible = self._visible_attachments(item)
            wanted = name.strip().lower()
            names = []
            for pos, (_i, a) in enumerate(visible, start=1):
                label = _safe(lambda: a.FileName, "") or _safe(lambda: a.DisplayName, "") or f"attachment-{pos}"
                names.append(label)
                if label.lower() == wanted or (wanted.isdigit() and int(wanted) == pos):
                    tmpdir = tempfile.mkdtemp(prefix="post-")
                    safe_name = re.sub(r'[\\/:*?"<>|]', "_", label)[:120] or "attachment"
                    path = os.path.join(tmpdir, safe_name)
                    try:
                        a.SaveAsFile(path)
                        with open(path, "rb") as fh:
                            return label, fh.read()
                    finally:
                        shutil.rmtree(tmpdir, ignore_errors=True)
            raise NotFound(f"Attachment '{name}' not found. Available: {', '.join(names) or 'no attachments'}.")

        return self._call(op)

    # -- Writing --

    def create_draft(self, account: str, req: DraftRequest) -> DraftResult:
        def op() -> DraftResult:
            entry = self._entry(account)
            store = entry.store
            if req.reply_to_id:
                orig = self._get_item(decode_id(req.reply_to_id))
                draft = orig.ReplyAll() if req.reply_all else orig.Reply()
            elif req.forward_of:
                if not req.to:
                    raise BackendError("Please give at least one recipient for the forward.")
                draft = self._get_item(decode_id(req.forward_of)).Forward()
            else:
                if not req.to:
                    raise BackendError("Please give at least one recipient.")
                draft = self._default_folder(store, "drafts").Items.Add(OL_MAIL_ITEM)
            existing = {a.lower() for a in _recipient_addresses(draft)}
            for addr, rtype in [(a, OL_TO) for a in req.to] + [(a, OL_CC) for a in req.cc] + [(a, OL_BCC) for a in req.bcc]:
                bare = _bare(addr)
                if bare.lower() in existing:
                    continue
                existing.add(bare.lower())
                recip = draft.Recipients.Add(addr)
                recip.Type = rtype
            if req.to or req.cc or req.bcc:
                _safe(lambda: draft.Recipients.ResolveAll(), None)
            if req.subject:
                draft.Subject = req.subject
            if entry.account is not None:
                _set_account(draft, entry.account)
            opened = False
            if req.open_window:
                draft.Display(False)  # Outlook inserts the signature now
                opened = True
            else:
                # touching the inspector makes Outlook insert the default signature, no window
                _safe(lambda: draft.GetInspector, None)
            insert_body(draft, req.body)
            draft.Save()
            if not opened:
                target = self._default_folder(store, "drafts")
                if _safe(lambda: draft.Parent.StoreID, "") != _safe(lambda: store.StoreID, ""):
                    draft = draft.Move(target)
            to, cc = _recipients(draft)
            note = (
                "NOT sent. The window is open in Outlook; the user checks it and clicks Send."
                if opened
                else f"NOT sent. It is in Outlook in mailbox {entry.key}, drafts folder. Open it, check it, send it."
            )
            return DraftResult(
                id=encode_id({"a": entry.key, "e": draft.EntryID, "s": _safe(lambda: draft.Parent.StoreID, "")}),
                account=entry.key,
                folder=self._folder_path(_safe(lambda: draft.Parent, None)) or "Drafts",
                subject=_safe(lambda: draft.Subject, ""),
                to=to,
                cc=cc,
                opened_window=opened,
                note=note,
            )

        return self._call(op)

    def move(self, ref: dict[str, Any], folder: str) -> str:
        def op() -> str:
            item = self._get_item(ref)
            target = self._resolve_folder(self._entry(ref["a"]).store, folder)
            moved = item.Move(target)
            return encode_id({"a": ref["a"], "e": moved.EntryID, "s": _safe(lambda: moved.Parent.StoreID, "")})

        return self._call(op)

    def delete(self, ref: dict[str, Any], permanent: bool = False) -> str:
        """Moves a mail to "Deleted Items", or erases it for good.

        Outlook's Delete() only moves an item to "Deleted Items"; it erases
        only what already sits there, so for a permanent delete we move first.
        """
        def op() -> str:
            item = self._get_item(ref)
            trash = self._resolve_folder(self._entry(ref["a"]).store, "trash")
            in_trash = _safe(lambda: item.Parent.EntryID, None) == _safe(lambda: trash.EntryID, "")
            if not permanent:
                if in_trash:
                    return encode_id(ref)
                moved = item.Move(trash)
                return encode_id({"a": ref["a"], "e": moved.EntryID, "s": _safe(lambda: moved.Parent.StoreID, "")})
            if not in_trash:
                item = item.Move(trash)
            item.Delete()
            return ""

        return self._call(op)

    def mark(self, ref: dict[str, Any], read: bool | None, flagged: bool | None) -> None:
        def op() -> None:
            item = self._get_item(ref)
            if read is not None:
                item.UnRead = not read
            if flagged is True:
                item.FlagStatus = OL_FLAG_MARKED
            elif flagged is False:
                try:
                    item.ClearTaskFlag()
                except Exception:
                    item.FlagStatus = 0
            item.Save()

        self._call(op)

    def _require_draft(self, item: Any) -> None:
        if _safe(lambda: item.Sent, False):
            raise BackendError("That is not a draft but a received or sent mail.")

    def send_draft(self, ref: dict[str, Any]) -> str:
        def op() -> str:
            item = self._get_item(ref)
            self._require_draft(item)
            subject = _safe(lambda: item.Subject, "")
            item.Send()
            return subject

        return self._call(op)

    def discard_draft(self, ref: dict[str, Any]) -> None:
        def op() -> None:
            item = self._get_item(ref)
            self._require_draft(item)
            item.Delete()

        self._call(op)


def _bare(addr: str) -> str:
    m = re.search(r"<([^>]+)>", addr)
    return (m.group(1) if m else addr).strip()


def _recipient_addresses(item: Any) -> list[str]:
    recips = _safe(lambda: item.Recipients, None)
    out = []
    for i in range(1, (_safe(lambda: recips.Count, 0) if recips is not None else 0) + 1):
        r = recips.Item(i)
        addr = _safe(lambda: r.Address, "") or ""
        if "@" not in addr:
            addr = _safe(lambda: r.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS), addr) or addr
        if addr:
            out.append(addr)
    return out
