"""IMAP backend: works with any mailbox that speaks IMAP.

For the new Outlook, Thunderbird, Apple Mail, Linux and servers. Drafts are
stored with IMAP APPEND in the drafts folder, so they show up in every mail
program using the same mailbox.
"""

from __future__ import annotations

import logging
import re
import smtplib
import socket
import ssl
import threading
from datetime import datetime, timezone
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
from email.utils import formataddr, format_datetime, getaddresses, make_msgid, parsedate_to_datetime
from typing import Any, Callable, TypeVar

from imapclient import IMAPClient
from imapclient.exceptions import IMAPClientError, LoginError

from ggpost.backends.base import (
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
)
from ggpost.backends.base import ROLE_ALIASES, ROLE_NAMES, fmt_date, is_all_folders
from ggpost.config import ServerSettings, Source
from ggpost.ids import decode_id, encode_id
from ggpost.text import (
    decode_mime_words,
    html_to_text,
    normalize_subject,
    normalize_whitespace,
)

log = logging.getLogger(__name__)
T = TypeVar("T")

SEEN = b"\\Seen"
FLAGGED = b"\\Flagged"
DRAFT = b"\\Draft"
DELETED = b"\\Deleted"

_ROLE_FLAGS = {
    b"\\Drafts": "drafts",
    b"\\Sent": "sent",
    b"\\Trash": "trash",
    b"\\Junk": "junk",
    b"\\Archive": "archive",
}

PREVIEW_BYTES = 24 * 1024
MAX_FULL_BYTES = 30 * 1024 * 1024


def _sort_key(dt: datetime | None) -> datetime:
    if dt is None:
        return datetime.min.replace(tzinfo=timezone.utc)
    if dt.tzinfo is None:
        return dt.astimezone()
    return dt


def _header_date(msg) -> datetime | None:
    raw = msg.get("Date")
    if not raw:
        return None
    try:
        return parsedate_to_datetime(str(raw))
    except Exception:
        return None


def _addr_header(msg, name: str) -> str:
    raw = msg.get(name)
    return decode_mime_words(str(raw)) if raw else ""


def _decode_lenient(payload: bytes, charset: str | None) -> str:
    """Decodes even with unknown or invented charsets ("unknown-8bit")."""
    for cs in (charset, "utf-8", "cp1252"):
        if not cs:
            continue
        try:
            return payload.decode(cs)
        except (LookupError, UnicodeDecodeError):
            continue
    return payload.decode("utf-8", errors="replace")


def extract_body(msg) -> str:
    """Prefers text/plain, otherwise HTML converted to text."""
    plain: str | None = None
    html_part: str | None = None
    parts = msg.walk() if msg.is_multipart() else [msg]
    for part in parts:
        if part.is_multipart():
            continue
        disp = (part.get_content_disposition() or "").lower()
        if disp == "attachment":
            continue
        ctype = part.get_content_type()
        if ctype not in ("text/plain", "text/html"):
            continue
        try:
            content = part.get_content()
        except Exception:
            payload = part.get_payload(decode=True) or b""
            content = _decode_lenient(payload, part.get_content_charset())
        if not isinstance(content, str):
            continue
        if ctype == "text/plain" and plain is None:
            plain = content
        elif ctype == "text/html" and html_part is None:
            html_part = content
    if plain and plain.strip():
        return normalize_whitespace(plain)
    if html_part:
        return html_to_text(html_part)
    return ""


def list_attachments(msg) -> list[Attachment]:
    out: list[Attachment] = []
    if not msg.is_multipart():
        return out
    for part in msg.walk():
        if part.is_multipart():
            continue
        filename = part.get_filename()
        disp = (part.get_content_disposition() or "").lower()
        if not filename and disp != "attachment":
            continue
        if disp == "inline" and part.get("Content-ID") and part.get_content_maintype() == "image":
            continue  # eingebettete Bilder (Logos, Signaturen)
        payload = part.get_payload(decode=True) or b""
        out.append(
            Attachment(
                name=decode_mime_words(filename) if filename else "(unnamed)",
                size=len(payload),
                content_type=part.get_content_type(),
            )
        )
    return out


def _refs(msg) -> list[str]:
    raw = str(msg.get("References", "") or "")
    return re.findall(r"<[^>]+>", raw)


class ImapBackend(MailBackend):
    kind = "imap"

    def __init__(
        self,
        source: Source,
        password_provider: Callable[[str], str | None],
        timeout: float = 30.0,
    ) -> None:
        if not source.imap:
            raise BackendError(f"The IMAP server settings for {source.email} are missing.")
        self.source = source
        self.email = source.email
        self._password_provider = password_provider
        self._timeout = timeout
        self._client: IMAPClient | None = None
        self._lock = threading.RLock()
        self._roles: dict[str, str] | None = None
        self._delimiter = "/"

    # -- Connection -------------------------------------------------------

    def _password(self) -> str:
        pw = self._password_provider(self.email)
        if not pw:
            raise BackendError(
                f"No password is stored for {self.email}. Please run setup again "
                "(command: post setup)."
            )
        return pw

    def _connect(self) -> IMAPClient:
        s = self.source.imap
        assert s is not None
        ctx = ssl.create_default_context()
        try:
            client = IMAPClient(
                s.host,
                port=s.port,
                ssl=(s.security == "ssl"),
                ssl_context=ctx if s.security == "ssl" else None,
                timeout=self._timeout,
            )
            if s.security == "starttls":
                client.starttls(ctx)
            user, password = s.username or self.email, self._password()
            if (user + password).isascii():
                client.login(user, password)
            else:  # LOGIN only knows ASCII, SASL PLAIN can do UTF-8
                client.plain_login(user, password)
        except LoginError as exc:
            raise BackendError(
                f"Login at {s.host} failed. Check the password; some providers need an "
                "app password or IMAP switched on first."
            ) from exc
        except (socket.gaierror, ConnectionRefusedError) as exc:
            raise BackendError(f"Server {s.host}:{s.port} not reachable ({exc}).") from exc
        except ssl.SSLError as exc:
            raise BackendError(f"Encrypted connection to {s.host} failed ({exc}).") from exc
        except (OSError, IMAPClientError) as exc:
            raise BackendError(f"Connection to {s.host} failed ({exc}).") from exc
        return client

    def _run(self, fn: Callable[[IMAPClient], T]) -> T:
        with self._lock:
            for attempt in (1, 2):
                if self._client is None:
                    self._client = self._connect()
                try:
                    return fn(self._client)
                except (IMAPClient.AbortError, OSError, ssl.SSLError) as exc:
                    log.info("IMAP connection lost (%s), reconnecting", exc)
                    self._drop()
                    if attempt == 2:
                        raise BackendError(f"Connection to {self.source.imap.host} was dropped.") from exc
            raise AssertionError("unreachable")

    def _drop(self) -> None:
        if self._client is not None:
            try:
                self._client.shutdown()
            except Exception:
                pass
        self._client = None

    def close(self) -> None:
        with self._lock:
            if self._client is not None:
                try:
                    self._client.logout()
                except Exception:
                    pass
            self._client = None

    # -- Folders ----------------------------------------------------------

    def _load_roles(self, c: IMAPClient) -> dict[str, str]:
        if self._roles is not None:
            return self._roles
        roles: dict[str, str] = {"inbox": "INBOX"}
        listing = c.list_folders()
        by_name: dict[str, str] = {}
        for flags, delim, name in listing:
            if delim:
                self._delimiter = delim.decode() if isinstance(delim, bytes) else delim
            for flag in flags:
                role = _ROLE_FLAGS.get(flag)
                if role and role not in roles:
                    roles[role] = name
            leaf = name.split(self._delimiter)[-1].strip().lower()
            by_name.setdefault(leaf, name)
        for role, names in ROLE_NAMES.items():
            if role in roles or role == "inbox":
                continue
            for candidate in names:
                if candidate in by_name:
                    roles[role] = by_name[candidate]
                    break
        self._roles = roles
        return roles

    def _resolve(self, c: IMAPClient, folder: str | None) -> str:
        folder = (folder or "inbox").strip()
        roles = self._load_roles(c)
        key = folder.lower()
        if key in roles:
            return roles[key]
        if key in ROLE_ALIASES and ROLE_ALIASES[key] in roles:
            return roles[ROLE_ALIASES[key]]
        names = [n for _, _, n in c.list_folders()]
        for n in names:
            if n == folder:
                return n
        for n in names:
            if n.lower() == key or n.replace(self._delimiter, "/").lower() == key:
                return n
        raise NotFound(
            f"There is no folder '{folder}' in {self.email}. mail_folders lists the available ones."
        )

    def _role_of(self, c: IMAPClient, name: str) -> str:
        for role, n in self._load_roles(c).items():
            if n == name:
                return role
        return ""

    # ── Schnittstelle ─────────────────────────────────────────────────────

    def accounts(self) -> list[AccountInfo]:
        return [AccountInfo(email=self.email, name=self.source.label or self.email, backend="imap", detail="IMAP")]

    def owns(self, account: str) -> bool:
        return account.lower() == self.email.lower()

    def check(self) -> str:
        def op(c: IMAPClient) -> str:
            info = c.folder_status("INBOX", [b"MESSAGES", b"UNSEEN"])
            return f"Inbox: {info.get(b'MESSAGES', 0)} mails, {info.get(b'UNSEEN', 0)} unread"

        return self._run(op)

    def folders(self, account: str) -> list[FolderInfo]:
        def op(c: IMAPClient) -> list[FolderInfo]:
            roles = self._load_roles(c)
            inverse = {v: k for k, v in roles.items()}
            out = []
            for flags, _delim, name in c.list_folders():
                if b"\\Noselect" in flags or b"\\NonExistent" in flags:
                    continue
                total = unread = None
                try:
                    st = c.folder_status(name, [b"MESSAGES", b"UNSEEN"])
                    total, unread = st.get(b"MESSAGES"), st.get(b"UNSEEN")
                except Exception:
                    pass
                out.append(FolderInfo(path=name, role=inverse.get(name, ""), total=total, unread=unread))
            return out

        return self._run(op)

    def _summaries(self, c: IMAPClient, folder: str, uidvalidity: int, uids: list[int]) -> list[MessageSummary]:
        if not uids:
            return []
        data = c.fetch(uids, [b"FLAGS", b"INTERNALDATE", f"BODY.PEEK[]<0.{PREVIEW_BYTES}>".encode()])
        out: list[tuple[datetime, MessageSummary]] = []
        for uid in uids:
            item = data.get(uid)
            if not item:
                continue
            raw = next((v for k, v in item.items() if k.startswith(b"BODY[")), b"") or b""
            msg = BytesParser(policy=policy.default).parsebytes(raw)
            flags = item.get(b"FLAGS", ())
            dt = _header_date(msg) or item.get(b"INTERNALDATE")
            try:
                snippet = extract_body(msg)
            except Exception:
                snippet = ""
            snippet = re.sub(r"\s+", " ", snippet)[:240]
            ctype = (msg.get_content_type() or "").lower()
            summary = MessageSummary(
                id=encode_id({"a": self.email, "f": folder, "v": uidvalidity, "u": uid}),
                account=self.email,
                folder=folder,
                date=fmt_date(dt),
                sender=_addr_header(msg, "From"),
                to=_addr_header(msg, "To")[:300],
                subject=_addr_header(msg, "Subject"),
                snippet=snippet,
                unread=SEEN not in flags,
                flagged=FLAGGED in flags,
                has_attachments=ctype == "multipart/mixed",
            )
            out.append((_sort_key(dt), summary))
        out.sort(key=lambda x: x[0], reverse=True)
        return [s for _, s in out]

    def _criteria(self, query: SearchQuery) -> tuple[list[Any], str | None]:
        criteria: list[Any] = []
        if query.unread_only:
            criteria.append("UNSEEN")
        if query.since:
            criteria += ["SINCE", query.since.date()]
        if query.until:
            criteria += ["BEFORE", query.until.date()]
        if query.sender:
            criteria += ["FROM", query.sender]
        for word in (query.text or "").split():
            criteria += ["TEXT", word]  # every word has to appear, same as in Outlook
        if not criteria:
            criteria = ["ALL"]
        needs_utf8 = any(isinstance(x, str) and not x.isascii() for x in criteria)
        return criteria, "UTF-8" if needs_utf8 else None

    def _search_folder(self, c: IMAPClient, folder: str, query: SearchQuery) -> list[MessageSummary]:
        sel = c.select_folder(folder, readonly=True)
        uidvalidity = int(sel.get(b"UIDVALIDITY", 0))
        criteria, charset = self._criteria(query)
        uids = c.search(criteria, charset=charset)
        uids = sorted(uids, reverse=True)[: max(1, query.limit) * 2]
        return self._summaries(c, folder, uidvalidity, uids)

    def _searchable_folders(self, c: IMAPClient) -> list[str]:
        roles = self._load_roles(c)
        skip = {roles.get("trash"), roles.get("junk"), roles.get("drafts")}
        out = []
        for flags, _d, name in c.list_folders():
            if b"\\Noselect" in flags or b"\\NonExistent" in flags or name in skip:
                continue
            if b"\\All" in flags:  # Gmail "All Mail" would return everything twice
                continue
            out.append(name)
        return out[:40]

    def search(self, account: str, query: SearchQuery) -> list[MessageSummary]:
        def op(c: IMAPClient) -> list[MessageSummary]:
            if is_all_folders(query.folder):
                results: list[MessageSummary] = []
                for folder in self._searchable_folders(c):
                    try:
                        results.extend(self._search_folder(c, folder, query))
                    except IMAPClientError as exc:
                        log.info("Search in %s skipped: %s", folder, exc)
                results.sort(key=lambda m: m.date, reverse=True)
                return results[: query.limit]
            folder = self._resolve(c, query.folder)
            return self._search_folder(c, folder, query)[: query.limit]

        return self._run(op)

    def _select_ref(self, c: IMAPClient, ref: dict[str, Any], readonly: bool = True) -> int:
        folder = ref.get("f", "INBOX")
        try:
            sel = c.select_folder(folder, readonly=readonly)
        except IMAPClientError as exc:
            raise NotFound(f"Folder '{folder}' not found.") from exc
        if int(sel.get(b"UIDVALIDITY", 0)) != int(ref.get("v", 0)):
            raise NotFound("The message is no longer in that place. Please search again.")
        return int(ref["u"])

    def _fetch_full(self, c: IMAPClient, uid: int) -> tuple[Any, tuple]:
        data = c.fetch([uid], [b"FLAGS", b"RFC822.SIZE"])
        if uid not in data:
            raise NotFound("The message was not found (it may have been moved or deleted).")
        size = int(data[uid].get(b"RFC822.SIZE", 0))
        if size > MAX_FULL_BYTES:
            spec = f"BODY.PEEK[]<0.{MAX_FULL_BYTES}>".encode()
        else:
            spec = b"BODY.PEEK[]"
        raw_data = c.fetch([uid], [spec])
        raw = next((v for k, v in raw_data[uid].items() if k.startswith(b"BODY[")), b"") or b""
        return BytesParser(policy=policy.default).parsebytes(raw), data[uid].get(b"FLAGS", ())

    def _to_message(self, msg, flags, folder: str, uidvalidity: int, uid: int) -> Message:
        dt = _header_date(msg)
        body = extract_body(msg)
        summary = MessageSummary(
            id=encode_id({"a": self.email, "f": folder, "v": uidvalidity, "u": uid}),
            account=self.email,
            folder=folder,
            date=fmt_date(dt),
            sender=_addr_header(msg, "From"),
            to=_addr_header(msg, "To"),
            subject=_addr_header(msg, "Subject"),
            unread=SEEN not in flags,
            flagged=FLAGGED in flags,
        )
        atts = list_attachments(msg)
        summary.has_attachments = bool(atts)
        return Message(
            summary=summary,
            cc=_addr_header(msg, "Cc"),
            body=body,
            attachments=atts,
            message_id=str(msg.get("Message-ID", "") or "").strip(),
            in_reply_to=str(msg.get("In-Reply-To", "") or "").strip(),
            references=_refs(msg),
            reply_to=_addr_header(msg, "Reply-To"),
            sort_key=_sort_key(dt),
        )

    def read(self, ref: dict[str, Any]) -> Message:
        def op(c: IMAPClient) -> Message:
            uid = self._select_ref(c, ref)
            msg, flags = self._fetch_full(c, uid)
            return self._to_message(msg, flags, ref["f"], int(ref["v"]), uid)

        return self._run(op)

    _THREAD_HEADERS = b"BODY.PEEK[HEADER.FIELDS (MESSAGE-ID IN-REPLY-TO REFERENCES FROM TO CC)]"

    def thread(self, ref: dict[str, Any], limit: int = 15) -> list[Message]:
        def op(c: IMAPClient) -> list[Message]:
            uid = self._select_ref(c, ref)
            msg, flags = self._fetch_full(c, uid)
            target = self._to_message(msg, flags, ref["f"], int(ref["v"]), uid)
            ids = {target.message_id, target.in_reply_to, *target.references} - {""}
            base = normalize_subject(target.summary.subject)
            folders = [ref["f"]]
            roles = self._load_roles(c)
            for role in ("inbox", "sent"):
                name = roles.get(role)
                if name and name not in folders:
                    folders.append(name)
            me = {a.lower() for _, a in getaddresses([target.summary.sender]) if a}
            # Fetch headers first and check the links, load bodies only for hits
            linked: list[tuple[str, int]] = []
            loose: list[tuple[str, int]] = []
            seen: set[str] = {target.message_id} - {""}
            for folder in folders:
                if not base:
                    break
                c.select_folder(folder, readonly=True)
                charset = None if base.isascii() else "UTF-8"
                uids = sorted(c.search(["SUBJECT", base], charset=charset), reverse=True)[: limit * 3]
                if not uids:
                    continue
                heads = c.fetch(uids, [self._THREAD_HEADERS])
                for cand_uid in uids:
                    item = heads.get(cand_uid) or {}
                    raw = next((v for k, v in item.items() if k.startswith(b"BODY[")), b"") or b""
                    h = BytesParser(policy=policy.default).parsebytes(raw, headersonly=True)
                    mid = str(h.get("Message-ID", "") or "").strip()
                    if (folder == ref["f"] and cand_uid == uid) or (mid and mid in seen):
                        continue
                    if mid:
                        seen.add(mid)
                    refs = set(_refs(h)) | {str(h.get("In-Reply-To", "") or "").strip()}
                    if mid in ids or refs & (ids | {target.message_id} - {""}):
                        linked.append((folder, cand_uid))
                    else:
                        people = " ".join(str(h.get(x, "")) for x in ("From", "To", "Cc")).lower()
                        if any(a in people for a in me):
                            loose.append((folder, cand_uid))
            chosen = linked if linked else loose
            found = [target]
            by_folder: dict[str, list[int]] = {}
            for folder, cand_uid in chosen[: limit * 2]:
                by_folder.setdefault(folder, []).append(cand_uid)
            for folder, uids in by_folder.items():
                sel = c.select_folder(folder, readonly=True)
                validity = int(sel.get(b"UIDVALIDITY", 0))
                for cand_uid in uids:
                    try:
                        m, fl = self._fetch_full(c, cand_uid)
                    except NotFound:
                        continue
                    found.append(self._to_message(m, fl, folder, validity, cand_uid))
            msgs = sorted(found, key=lambda m: m.sort_key or datetime.min.replace(tzinfo=timezone.utc))
            return msgs[-limit:]

        return self._run(op)

    # -- Writing ----------------------------------------------------------

    def _german(self) -> bool:
        return self.email.rsplit(".", 1)[-1].lower() in ("de", "at", "ch")

    def _prefixed(self, subject: str, kind: str) -> str:
        """Re/AW and Fwd/WG: keep an existing prefix, otherwise match the domain."""
        if kind == "reply":
            if re.match(r"^\s*(re|aw)\s*:", subject, re.I):
                return subject
            return ("AW: " if self._german() else "Re: ") + subject
        if re.match(r"^\s*(fwd?|wg)\s*:", subject, re.I):
            return subject
        return ("WG: " if self._german() else "Fwd: ") + subject

    def _load_original(self, c: IMAPClient, message_id: str) -> tuple[Any, Message]:
        from ggpost.ids import decode_id  # noqa: PLC0415

        ref = decode_id(message_id)
        uid = self._select_ref(c, ref)
        raw, flags = self._fetch_full(c, uid)
        return raw, self._to_message(raw, flags, ref["f"], int(ref["v"]), uid)

    def _build_draft(self, c: IMAPClient, req: DraftRequest) -> EmailMessage:
        msg = EmailMessage(policy=policy.SMTP)
        domain = self.email.split("@")[-1] or None
        label = self.source.label
        msg["From"] = formataddr((label, self.email)) if label and label != self.email else self.email
        to = list(req.to)
        cc = list(req.cc)
        subject = req.subject or ""
        body = req.body.rstrip() + "\n"
        signature = (self.source.signature or "").strip()
        if signature:
            body += "\n" + signature + "\n"
        forwarded_parts: list[Any] = []
        if req.reply_to_id:
            _raw, orig = self._load_original(c, req.reply_to_id)
            reply_target = orig.reply_to or orig.summary.sender
            own = self.email.lower()
            auto_to = [formataddr(a) for a in getaddresses([reply_target]) if a[1]]
            auto_cc: list[str] = []
            if req.reply_all:
                for name, addr in getaddresses([orig.summary.to, orig.cc]):
                    if addr and addr.lower() != own and addr.lower() not in reply_target.lower():
                        auto_cc.append(formataddr((name, addr)))
            to = auto_to + [a for a in to if a not in auto_to]
            cc = auto_cc + [a for a in cc if a not in auto_cc]
            if not subject:
                subject = self._prefixed(orig.summary.subject, "reply")
            if orig.message_id:
                msg["In-Reply-To"] = orig.message_id
                msg["References"] = " ".join([*orig.references, orig.message_id])
            quoted = "\n".join(f"> {line}" if line else ">" for line in orig.body.splitlines())
            body += f"\nAm {orig.summary.date} schrieb {orig.summary.sender}:\n{quoted}\n"
        elif req.forward_of:
            raw, orig = self._load_original(c, req.forward_of)
            if not subject:
                subject = self._prefixed(orig.summary.subject, "forward")
            body += (
                "\n---------- Forwarded message ----------\n"
                f"From: {orig.summary.sender}\nDate: {orig.summary.date}\n"
                f"Subject: {orig.summary.subject}\nTo: {orig.summary.to}\n"
                + (f"Cc: {orig.cc}\n" if orig.cc else "")
                + f"\n{orig.body}\n"
            )
            if raw.is_multipart():
                for part in raw.walk():
                    if part.is_multipart() or not (part.get_filename() or part.get_content_disposition() == "attachment"):
                        continue
                    forwarded_parts.append(part)
        if not to and not req.reply_to_id:
            raise BackendError("Please give at least one recipient.")
        if to:
            msg["To"] = ", ".join(to)
        if cc:
            msg["Cc"] = ", ".join(cc)
        if req.bcc:
            msg["Bcc"] = ", ".join(req.bcc)
        msg["Subject"] = subject
        msg["Date"] = format_datetime(datetime.now().astimezone())
        msg["Message-ID"] = make_msgid(domain=domain)
        msg["X-Mailer"] = "post. (gedacht.gemacht.)"
        msg.set_content(body)
        for part in forwarded_parts:
            data = part.get_payload(decode=True) or b""
            maintype, _, subtype = part.get_content_type().partition("/")
            msg.add_attachment(data, maintype=maintype or "application", subtype=subtype or "octet-stream",
                               filename=decode_mime_words(part.get_filename() or "attachment"))
        return msg

    def _draft_note(self) -> str:
        host = (self.source.imap.host if self.source.imap else "").lower()
        where = f"It is in mailbox {self.email} in the drafts folder"
        if "gmail" in host or "google" in host:
            where += " (Gmail: https://mail.google.com/mail/#drafts)"
        return f"NOT sent. {where} and shows up in every mail program using this mailbox. Open it, check it, send it."

    def create_draft(self, account: str, req: DraftRequest) -> DraftResult:
        def op(c: IMAPClient) -> DraftResult:
            msg = self._build_draft(c, req)
            drafts = self._load_roles(c).get("drafts")
            if not drafts:
                raise BackendError(f"No drafts folder found in {self.email}.")
            resp = c.append(drafts, msg.as_bytes(), flags=(DRAFT, SEEN), msg_time=datetime.now().astimezone())
            sel = c.select_folder(drafts, readonly=True)
            validity = int(sel.get(b"UIDVALIDITY", 0))
            uid = None
            m = re.search(rb"APPENDUID (\d+) (\d+)", resp if isinstance(resp, bytes) else str(resp).encode())
            if m:
                uid = int(m.group(2))
                validity = int(m.group(1))
            else:
                hits = c.search(["HEADER", "Message-ID", str(msg["Message-ID"])])
                uid = max(hits) if hits else None
            draft_id = encode_id({"a": self.email, "f": drafts, "v": validity, "u": uid}) if uid else ""
            note = self._draft_note()
            if req.open_window:
                note += " Opening a window only works with classic Outlook."
            return DraftResult(
                id=draft_id,
                account=self.email,
                folder=drafts,
                subject=str(msg["Subject"]),
                to=str(msg.get("To", "")),
                cc=str(msg.get("Cc", "")),
                opened_window=False,
                note=note,
            )

        return self._run(op)

    def move(self, ref: dict[str, Any], folder: str) -> str:
        def op(c: IMAPClient) -> str:
            target = self._resolve(c, folder)
            if target == ref["f"]:
                return encode_id(ref)
            # Ask about the capabilities before we select the folder for writing:
            # a CAPABILITY in between can put the mailbox back into read-only,
            # and the server then refuses the MOVE.
            can_move = c.has_capability("MOVE")
            can_uidplus = c.has_capability("UIDPLUS")
            uid = self._select_ref(c, ref, readonly=False)
            head = c.fetch([uid], [b"BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)]"]).get(uid, {})
            raw_head = next((v for k, v in head.items() if k.startswith(b"BODY[")), b"") or b""
            message_id = str(BytesParser(policy=policy.default).parsebytes(raw_head, headersonly=True).get("Message-ID", "") or "").strip()
            c._imap.untagged_responses.pop("COPYUID", None)  # noqa: SLF001
            if can_move:
                resp = c.move([uid], target)
            else:
                resp = c.copy([uid], target)
                c.add_flags([uid], [DELETED])
                if can_uidplus:
                    c.uid_expunge([uid])
                else:
                    c.expunge()
            texts = [resp if isinstance(resp, bytes) else str(resp or "").encode()]
            texts += [b"COPYUID " + x for x in c._imap.untagged_responses.pop("COPYUID", [])]  # noqa: SLF001
            for text in texts:
                m = re.search(rb"COPYUID (\d+) \S+ (\d+)", text)
                if m:
                    return encode_id({"a": self.email, "f": target, "v": int(m.group(1)), "u": int(m.group(2))})
            # Server without UIDPLUS: find it again by Message-ID in the target folder
            if message_id:
                sel = c.select_folder(target, readonly=True)
                hits = c.search(["HEADER", "Message-ID", message_id])
                if hits:
                    return encode_id({"a": self.email, "f": target, "v": int(sel.get(b"UIDVALIDITY", 0)), "u": max(hits)})
            return ""

        return self._run(op)

    def delete(self, ref: dict[str, Any], permanent: bool = False) -> str:
        """Moves a mail to the trash, or erases it for good.

        Erasing goes through the trash on purpose: Gmail and others keep the
        message under another label if we only flag it as deleted where it is,
        so the mail would still be there afterwards.
        """
        if not permanent:
            return self.move(ref, "trash")

        def message_id_of(c: IMAPClient) -> str:
            uid = self._select_ref(c, ref, readonly=True)
            head = c.fetch([uid], [b"BODY.PEEK[HEADER.FIELDS (MESSAGE-ID)]"]).get(uid, {})
            raw = next((v for k, v in head.items() if k.startswith(b"BODY[")), b"") or b""
            parsed = BytesParser(policy=policy.default).parsebytes(raw, headersonly=True)
            return str(parsed.get("Message-ID", "") or "").strip()

        already_in_trash = self._run(lambda c: self._role_of(c, ref.get("f", "")) == "trash")
        if already_in_trash:
            def purge_here(c: IMAPClient) -> None:
                uid = self._select_ref(c, ref, readonly=False)
                self._expunge(c, [uid])

            self._run(purge_here)
            return ""

        message_id = self._run(message_id_of)
        moved_id = self.move(ref, "trash")
        moved = decode_id(moved_id) if moved_id else None

        def purge_in_trash(c: IMAPClient) -> None:
            # The UID that move() reported is the reliable handle. Only when the
            # server gave us none do we fall back to the Message-ID, and a mail
            # without that header leaves us nothing to match on.
            if moved and moved.get("u"):
                self._expunge(c, [self._select_ref(c, moved, readonly=False)])
                return
            target = self._resolve(c, "trash")
            c.select_folder(target, readonly=False)
            hits = c.search(["HEADER", "Message-ID", message_id]) if message_id else []
            if not hits:
                raise BackendError(
                    "The mail was moved to the trash, but could not be erased there. "
                    "Emptying the trash in your mail program will remove it."
                )
            self._expunge(c, list(hits))

        self._run(purge_in_trash)
        return ""

    def _expunge(self, c: IMAPClient, uids: list[int]) -> None:
        can_uidplus = c.has_capability("UIDPLUS")
        c.add_flags(uids, [DELETED])
        if can_uidplus:
            c.uid_expunge(uids)
        else:
            c.expunge()

    def mark(self, ref: dict[str, Any], read: bool | None, flagged: bool | None) -> None:
        def op(c: IMAPClient) -> None:
            uid = self._select_ref(c, ref, readonly=False)
            if read is True:
                c.add_flags([uid], [SEEN])
            elif read is False:
                c.remove_flags([uid], [SEEN])
            if flagged is True:
                c.add_flags([uid], [FLAGGED])
            elif flagged is False:
                c.remove_flags([uid], [FLAGGED])

        self._run(op)

    def send_draft(self, ref: dict[str, Any]) -> str:
        if not self.source.smtp:
            raise BackendError(f"No outgoing server (SMTP) is set up for {self.email}.")

        def op(c: IMAPClient) -> tuple[EmailMessage, int]:
            if self._role_of(c, ref.get("f", "")) != "drafts":
                raise BackendError("Only drafts from the drafts folder can be sent (id from mail_draft).")
            uid = self._select_ref(c, ref)
            msg, _ = self._fetch_full(c, uid)
            return msg, uid

        msg, uid = self._run(op)
        recipients = [a for _, a in getaddresses([str(msg.get(h, "")) for h in ("To", "Cc", "Bcc")]) if a]
        if not recipients:
            raise BackendError("The draft has no recipients.")
        del msg["Bcc"]
        del msg["Date"]
        msg["Date"] = format_datetime(datetime.now().astimezone())
        self._smtp_send(msg, recipients)

        def after(c: IMAPClient) -> None:
            sent = self._load_roles(c).get("sent")
            if sent:
                try:
                    c.append(sent, msg.as_bytes(policy=policy.SMTP), flags=(SEEN,))
                except Exception as exc:  # a copy in "sent" is nice to have
                    log.warning("Copy to sent failed: %s", exc)
            self._select_ref(c, ref, readonly=False)
            c.add_flags([uid], [DELETED])
            if c.has_capability("UIDPLUS"):
                c.uid_expunge([uid])
            else:
                c.expunge()

        self._run(after)
        return str(msg.get("Message-ID", ""))

    def discard_draft(self, ref: dict[str, Any]) -> None:
        def op(c: IMAPClient) -> None:
            if self._role_of(c, ref.get("f", "")) != "drafts":
                raise BackendError("Replacing or deleting only works for drafts.")
            uid = self._select_ref(c, ref, readonly=False)
            c.add_flags([uid], [DELETED])
            if c.has_capability("UIDPLUS"):
                c.uid_expunge([uid])
            else:
                c.expunge()

        self._run(op)

    def attachment(self, ref: dict[str, Any], name: str) -> tuple[str, bytes]:
        def op(c: IMAPClient) -> tuple[str, bytes]:
            uid = self._select_ref(c, ref)
            msg, _flags = self._fetch_full(c, uid)
            wanted = name.strip().lower()
            candidates = []
            for part in msg.walk():
                if part.is_multipart():
                    continue
                filename = part.get_filename()
                if not filename and part.get_content_disposition() != "attachment":
                    continue
                label = decode_mime_words(filename) if filename else "(unnamed)"
                candidates.append(label)
                if label.lower() == wanted or (wanted.isdigit() and len(candidates) == int(wanted)):
                    return label, part.get_payload(decode=True) or b""
            raise NotFound(
                f"Attachment '{name}' not found. Available: {', '.join(candidates) or 'no attachments'}."
            )

        return self._run(op)

    def _smtp_send(self, msg: EmailMessage, recipients: list[str]) -> None:
        s: ServerSettings = self.source.smtp  # type: ignore[assignment]
        ctx = ssl.create_default_context()
        try:
            if s.security == "ssl":
                server: smtplib.SMTP = smtplib.SMTP_SSL(s.host, s.port, context=ctx, timeout=self._timeout)
            else:
                server = smtplib.SMTP(s.host, s.port, timeout=self._timeout)
                if s.security == "starttls":
                    server.starttls(context=ctx)
            with server:
                if s.security != "plain" or s.username:
                    _smtp_login(server, s.username or self.email, self._password())
                server.send_message(msg, from_addr=self.email, to_addrs=recipients)
        except smtplib.SMTPAuthenticationError as exc:
            raise BackendError(f"Login at the outgoing server {s.host} failed.") from exc
        except (smtplib.SMTPException, OSError) as exc:
            raise BackendError(f"Sending through {s.host} failed ({exc}).") from exc


def _smtp_login(server: smtplib.SMTP, user: str, password: str) -> None:
    if (user + password).isascii():
        server.login(user, password)
        return
    # smtplib encodes credentials as ASCII only; do AUTH PLAIN with UTF-8 by hand
    import base64  # noqa: PLC0415

    server.ehlo_or_helo_if_needed()
    token = base64.b64encode(b"\0" + user.encode("utf-8") + b"\0" + password.encode("utf-8")).decode("ascii")
    code, resp = server.docmd("AUTH", "PLAIN " + token)
    if code != 235:
        raise smtplib.SMTPAuthenticationError(code, resp)


def verify_login(settings: ServerSettings, email: str, password: str, timeout: float = 20.0) -> str:
    """Used by setup: logs in and reads the inbox."""
    src = Source(kind="imap", email=email, imap=settings)
    backend = ImapBackend(src, lambda _e: password, timeout=timeout)
    try:
        return backend.check()
    finally:
        backend.close()
