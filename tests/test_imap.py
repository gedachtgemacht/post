from __future__ import annotations

from email import policy
from email.parser import BytesParser

import pytest
from imapclient import IMAPClient

from ggpost.backends.base import BackendError, DraftRequest, NotFound, SearchQuery
from ggpost.backends.imap import ImapBackend, verify_login
from ggpost.ids import decode_id


def make(source, pw="demopass"):
    return ImapBackend(source, lambda _e: pw, timeout=10)


def raw_client(port):
    c = IMAPClient("127.0.0.1", port=port, ssl=False)
    c.login("demouser", "demopass")
    return c


def test_login_and_check(imap_source):
    assert "Posteingang" in verify_login(imap_source.imap, imap_source.email, "demopass")


def test_wrong_password(imap_source):
    with pytest.raises(BackendError, match="Login"):
        make(imap_source, "falsch").check()


def test_folders_roles(imap_source):
    b = make(imap_source)
    folders = {f.path: f for f in b.folders(imap_source.email)}
    assert folders["INBOX"].role == "inbox"
    assert folders["Sent"].role == "sent"
    assert folders["Entwürfe"].role == "drafts"
    assert folders["Trash"].role == "trash"
    assert folders["Archiv"].role == "archive"
    assert folders["INBOX"].total >= 1


def test_search_lists_newest_first(imap_source):
    b = make(imap_source)
    res = b.search(imap_source.email, SearchQuery(limit=10))
    assert len(res) >= 3
    assert res[0].date >= res[-1].date
    assert all(m.account == imap_source.email for m in res)
    assert any(m.unread for m in res)


def test_search_filters(imap_source):
    b = make(imap_source)
    unread = b.search(imap_source.email, SearchQuery(unread_only=True))
    assert unread and all(m.unread for m in unread)
    first = b.search(imap_source.email, SearchQuery(limit=1))[0]
    word = first.subject.split()[0]
    hits = b.search(imap_source.email, SearchQuery(text=word))
    assert any(h.id == first.id for h in hits)


def test_search_unknown_folder(imap_source):
    with pytest.raises(NotFound):
        make(imap_source).search(imap_source.email, SearchQuery(folder="gibtsnicht"))


def test_read_does_not_mark_seen(imap_source, imap_server):
    b = make(imap_source)
    unread = b.search(imap_source.email, SearchQuery(unread_only=True))[0]
    msg = b.read(decode_id(unread.id))
    assert msg.body
    assert msg.summary.subject == unread.subject
    again = b.search(imap_source.email, SearchQuery(unread_only=True))
    assert any(m.id == unread.id for m in again)


def test_draft_new_and_reply(imap_source, imap_server):
    b = make(imap_source)
    res = b.create_draft(
        imap_source.email,
        DraftRequest(account=None, to=["kunde@example.com"], cc=[], bcc=[], subject="Angebot Paletten", body="Moin,\n\nanbei das Angebot.\n\nGruß"),
    )
    assert res.folder == "Entwürfe"
    assert res.id
    ref = decode_id(res.id)
    drafted = b.read(ref)
    assert "anbei das Angebot" in drafted.body
    assert drafted.summary.to == "kunde@example.com"

    # Reply to a mail that has a Message-ID (the draft just created has one)
    reply = b.create_draft(
        imap_source.email,
        DraftRequest(account=None, to=[], cc=[], bcc=[], subject=None, body="Danke, passt.", reply_to_id=res.id),
    )
    assert reply.subject.lower().startswith(("re:", "aw:"))
    c = raw_client(imap_server)
    c.select_folder("Entwürfe", readonly=True)
    raw = c.fetch([decode_id(reply.id)["u"]], [b"BODY.PEEK[]", b"FLAGS"])
    item = next(iter(raw.values()))
    msg = BytesParser(policy=policy.default).parsebytes(item[b"BODY[]"])
    assert b"\\Draft" in item[b"FLAGS"]
    assert msg["In-Reply-To"]
    assert "> " in msg.get_content()
    assert "Danke, passt." in msg.get_content()
    c.logout()


def test_draft_needs_recipient(imap_source):
    with pytest.raises(BackendError, match="recipient"):
        make(imap_source).create_draft(imap_source.email, DraftRequest(account=None, to=[], cc=[], bcc=[], subject="x", body="y"))


def test_mark_and_move(imap_source):
    b = make(imap_source)
    first = b.search(imap_source.email, SearchQuery(limit=1))[0]
    ref = decode_id(first.id)
    b.mark(ref, read=False, flagged=True)
    after = b.read(ref)
    assert after.summary.unread and after.summary.flagged
    b.mark(ref, read=True, flagged=False)
    after = b.read(ref)
    assert not after.summary.unread and not after.summary.flagged

    new_id = b.move(ref, "archiv")
    assert new_id
    moved = b.read(decode_id(new_id))
    assert moved.summary.subject == first.subject
    back = b.move(decode_id(new_id), "inbox")
    assert back


def test_stale_id(imap_source):
    b = make(imap_source)
    first = b.search(imap_source.email, SearchQuery(limit=1))[0]
    ref = decode_id(first.id)
    ref["v"] = 1
    with pytest.raises(NotFound):
        b.read(ref)


def test_thread(imap_source):
    b = make(imap_source)
    first = b.search(imap_source.email, SearchQuery(limit=10))
    thread = b.thread(decode_id(first[0].id))
    assert thread and thread[-1].summary.subject


def test_reconnect_after_drop(imap_source):
    b = make(imap_source)
    b.check()
    b._client.shutdown()  # cut the connection hard
    assert "Posteingang" in b.check()


def _append_with_attachment(port, subject, filename, data, maintype, subtype):
    from email.message import EmailMessage

    m = EmailMessage()
    m["From"] = "Lieferant <info@lieferant.de>"
    m["To"] = "demouser@example.org"
    m["Subject"] = subject
    m["Message-ID"] = f"<{subject.replace(' ', '')}@lieferant.de>"
    m.set_content("Siehe Anhang.")
    m.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
    c = raw_client(port)
    c.append("INBOX", m.as_bytes())
    c.logout()


def test_search_words_independent_and_all_folders(imap_source, imap_server):
    _append_with_attachment(imap_server, "Paletten Angebot Nord", "liste.csv", b"a;b", "text", "csv")
    b = make(imap_source)
    hits = b.search(imap_source.email, SearchQuery(text="Angebot Paletten"))
    assert any(h.subject == "Paletten Angebot Nord" for h in hits)
    c = raw_client(imap_server)
    c.select_folder("INBOX")
    uid = max(c.search(["SUBJECT", "Paletten Angebot Nord"]))
    c.move([uid], "Archiv")
    c.logout()
    assert not any(h.subject == "Paletten Angebot Nord" for h in b.search(imap_source.email, SearchQuery(text="Nord")))
    everywhere = b.search(imap_source.email, SearchQuery(folder="all", text="Nord"))
    assert [h.folder for h in everywhere if h.subject == "Paletten Angebot Nord"] == ["Archiv"]


def test_attachment_and_forward(imap_source, imap_server):
    _append_with_attachment(imap_server, "Rechnung 99", "rechnung.txt", "Betrag: 120 €".encode(), "text", "plain")
    b = make(imap_source)
    hit = b.search(imap_source.email, SearchQuery(text="Rechnung 99"))[0]
    name, data = b.attachment(decode_id(hit.id), "rechnung.txt")
    assert name == "rechnung.txt" and "120 €" in data.decode()
    with pytest.raises(NotFound, match="rechnung.txt"):
        b.attachment(decode_id(hit.id), "nix.pdf")
    fwd = b.create_draft(imap_source.email, DraftRequest(account=None, to=["chef@example.org"], cc=[], bcc=[], subject=None, body="Bitte prüfen.", forward_of=hit.id))
    assert fwd.subject == "Fwd: Rechnung 99"
    assert b.attachment(decode_id(fwd.id), "rechnung.txt")[1] == data
    draft = b.read(decode_id(fwd.id))
    assert "Forwarded message" in draft.body and "Bitte prüfen." in draft.body


def test_signature_prefix_and_discard(imap_source):
    from ggpost.config import Source

    src = Source(kind="imap", email="demouser@example.de", label="Demo", imap=imap_source.imap, signature="Viele Grüße\nDemo GmbH")
    b = make(src)
    first = b.create_draft(src.email, DraftRequest(account=None, to=["x@y.de"], cc=[], bcc=[], subject="Frage", body="Hallo"))
    msg = b.read(decode_id(first.id))
    assert msg.body.endswith("Viele Grüße\nDemo GmbH")
    reply = b.create_draft(src.email, DraftRequest(account=None, to=[], cc=[], bcc=[], subject=None, body="Antwort", reply_to_id=first.id))
    assert reply.subject == "AW: Frage"
    assert "NICHT gesendet" in reply.note
    b.discard_draft(decode_id(first.id))
    with pytest.raises(NotFound):
        b.read(decode_id(first.id))
    inbox_mail = b.search(src.email, SearchQuery(limit=1))[0]
    with pytest.raises(BackendError, match="drafts"):
        b.discard_draft(decode_id(inbox_mail.id))


def test_utf8_password_login(imap_server_umlaut):
    from ggpost.config import ServerSettings, Source

    src = Source(kind="imap", email="demouser@example.org", imap=ServerSettings("127.0.0.1", imap_server_umlaut, "plain", "demouser"))
    assert "Posteingang" in ImapBackend(src, lambda _e: "Grüße€2026", timeout=10).check()
    with pytest.raises(BackendError, match="Login"):
        ImapBackend(src, lambda _e: "Grüße€falsch", timeout=10).check()


def test_unknown_charset_is_readable(imap_server, imap_source):
    raw = (
        b"From: spam@example.com\r\nTo: demouser@example.org\r\nSubject: Komisch\r\n"
        b"Content-Type: text/plain; charset=unknown-8bit\r\nContent-Transfer-Encoding: 8bit\r\n\r\n"
        b"Gr\xfc\xdfe aus dem Nirgendwo\r\n"
    )
    c = raw_client(imap_server)
    c.append("INBOX", raw)
    c.logout()
    b = make(imap_source)
    hit = b.search(imap_source.email, SearchQuery(text="Komisch"))[0]
    assert "aus dem Nirgendwo" in b.read(decode_id(hit.id)).body


def test_thread_links_by_headers(imap_server, imap_source):
    def mail(mid, subject, frm, to, irt=None, refs=None):
        h = f"From: {frm}\r\nTo: {to}\r\nSubject: {subject}\r\nMessage-ID: {mid}\r\nDate: Mon, 19 Oct 2026 0{len(refs or '') % 9}:00:00 +0200\r\n"
        if irt:
            h += f"In-Reply-To: {irt}\r\nReferences: {refs or irt}\r\n"
        return (h + "\r\nText\r\n").encode()

    c = raw_client(imap_server)
    c.append("INBOX", mail("<t1@k.de>", "Lieferung KW 44", "kunde@k.de", "demouser@example.org"))
    c.append("Sent", mail("<t2@me.de>", "AW: Lieferung KW 44", "demouser@example.org", "kunde@k.de", "<t1@k.de>"))
    c.append("INBOX", mail("<t3@k.de>", "AW: Lieferung KW 44", "kunde@k.de", "demouser@example.org", "<t2@me.de>", "<t1@k.de> <t2@me.de>"))
    c.append("INBOX", mail("<x9@other.de>", "Lieferung KW 44", "fremd@other.de", "x@other.de"))
    c.logout()
    b = make(imap_source)
    hit = [h for h in b.search(imap_source.email, SearchQuery(text="Lieferung")) if h.sender == "kunde@k.de"]
    thread = b.thread(decode_id(hit[0].id))
    ids = {m.message_id for m in thread}
    assert ids == {"<t1@k.de>", "<t2@me.de>", "<t3@k.de>"}


def test_send_draft_via_smtp(imap_server, imap_source, smtp_server):
    from ggpost.config import ServerSettings, Source

    port, sink = smtp_server
    src = Source(kind="imap", email=imap_source.email, label="Demo", imap=imap_source.imap,
                 smtp=ServerSettings("127.0.0.1", port, "plain", ""))
    b = make(src)
    draft = b.create_draft(src.email, DraftRequest(account=None, to=["kunde@k.de"], cc=["chef@k.de"], bcc=["geheim@k.de"], subject="Angebot", body="Hallo"))
    b.send_draft(decode_id(draft.id))
    mail_from, rcpts, content = sink.messages[0]
    assert mail_from == src.email
    assert set(rcpts) == {"kunde@k.de", "chef@k.de", "geheim@k.de"}
    assert b"geheim@k.de" not in content  # Bcc is not part of the mail itself
    with pytest.raises(NotFound):
        b.read(decode_id(draft.id))  # the draft is gone
    sent = b.search(src.email, SearchQuery(folder="sent", text="Angebot"))
    assert sent and sent[0].subject == "Angebot"


def test_delete_moves_to_trash_and_permanent_erases(imap_source):
    b = ImapBackend(imap_source, lambda e: "demopass")
    hit = b.search(imap_source.email, SearchQuery())[0]
    ref = decode_id(hit.id)
    b.delete(ref)
    trash = b.search(imap_source.email, SearchQuery(folder="trash"))
    assert any(m.subject == hit.subject for m in trash)

    second = b.search(imap_source.email, SearchQuery())[0]
    subject = second.subject
    b.delete(decode_id(second.id), permanent=True)
    # Neither where it was nor in the trash: really gone
    assert not any(m.subject == subject for m in b.search(imap_source.email, SearchQuery(folder="all")))
