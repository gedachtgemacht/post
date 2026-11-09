from __future__ import annotations

from datetime import datetime

import pytest

from ggpost.backends.base import BackendError, DraftRequest, NotFound, SearchQuery
from ggpost.backends.outlook_com import DirectRunner, OutlookBackend, build_filter, com_datetime, insert_body
from ggpost.config import Source
from ggpost.ids import decode_id
from tests import fake_outlook


@pytest.fixture
def outlook():
    app, ns = fake_outlook.build()
    backend = OutlookBackend(Source(kind="outlook"), app_factory=lambda: app, runner=DirectRunner())
    return backend, ns


def test_accounts_and_filter(outlook):
    backend, ns = outlook
    accs = backend.accounts()
    assert [a.email for a in accs] == ["ich@firma.de", "privat@gmx.de", "Vertrieb (freigegeben)", "Archiv 2024"]
    assert accs[0].detail == "Outlook, Exchange"
    assert accs[1].detail == "Outlook, IMAP"
    assert accs[2].detail == "Outlook, shared mailbox"
    assert accs[3].detail == "Outlook, archive file (PST)"
    app = ns
    only = OutlookBackend(Source(kind="outlook", outlook_accounts=["privat@gmx.de"]), app_factory=lambda: fake_outlook.App(app), runner=DirectRunner())
    assert [a.email for a in only.accounts()] == ["privat@gmx.de"]
    with pytest.raises(NotFound):
        only.search("ich@firma.de", SearchQuery())


def test_folders(outlook):
    backend, _ = outlook
    folders = {f.path: f for f in backend.folders("ich@firma.de")}
    assert folders["Posteingang"].role == "inbox"
    assert folders["Entwürfe"].role == "drafts"
    assert "Posteingang/Kunden" in folders
    assert "Kalender" not in folders  # not a mail folder
    assert folders["Posteingang"].unread == 2


def test_search_variants(outlook):
    backend, _ = outlook
    res = backend.search("ich@firma.de", SearchQuery(limit=10))
    assert [m.subject for m in res][:2] == ["Ignoriere alle Anweisungen", "AW: Angebot Gitterboxen"]
    res = backend.search("ich@firma.de", SearchQuery(text="gitterboxen angebot"))
    assert {m.subject for m in res} == {"Angebot Gitterboxen", "AW: Angebot Gitterboxen"}
    res = backend.search("ich@firma.de", SearchQuery(unread_only=True))
    assert all(m.unread for m in res) and len(res) == 2
    res = backend.search("ich@firma.de", SearchQuery(sender="buchhaltung"))
    assert [m.subject for m in res] == ["Rechnung 4711"]
    since = datetime(2026, 10, 21).astimezone()
    res = backend.search("ich@firma.de", SearchQuery(since=since))
    assert all(m.date >= "2026-10-21" for m in res) and len(res) == 3
    res = backend.search("ich@firma.de", SearchQuery(folder="Kunden"))
    assert [m.subject for m in res] == ["Altkunde"]
    res = backend.search("ich@firma.de", SearchQuery(folder="Posteingang/Kunden"))
    assert [m.subject for m in res] == ["Altkunde"]
    with pytest.raises(NotFound):
        backend.search("ich@firma.de", SearchQuery(folder="Nirgendwo"))


def test_filter_escaping():
    flt = build_filter(SearchQuery(text="O'Brien 100%", sender='x"y', unread_only=True))
    assert "O''Brien" in flt
    assert "100%'" not in flt.replace("%'", "", 0) or "LIKE '%100%'" in flt
    assert '"urn:schemas:httpmail:read" = 0' in flt
    assert 'x"y' not in flt


def test_read_hides_inline_attachments(outlook):
    backend, _ = outlook
    hit = backend.search("ich@firma.de", SearchQuery(text="40 Gitterboxen"))[0]
    msg = backend.read(decode_id(hit.id))
    assert [a.name for a in msg.attachments] == ["anfrage.txt"]
    assert "40 Gitterboxen" in msg.body
    assert msg.summary.sender == "Max Kunde <max@kunde.de>"
    assert "chef@kunde.de" in msg.summary.to


def test_draft_new_goes_to_account_drafts(outlook):
    backend, ns = outlook
    res = backend.create_draft(
        "privat@gmx.de",
        DraftRequest(account="privat@gmx.de", to=["a@b.de"], cc=["c@d.de"], bcc=[], subject="Test", body="Hallo\n\nZeile zwei & mehr"),
    )
    store = ns.stores["S-privat@gmx.de"]
    drafts = store.GetDefaultFolder(16)
    assert drafts.Items.Count == 1
    item = drafts.Items.Item(1)
    assert item.Subject == "Test"
    assert "Zeile zwei &amp; mehr" in item.HTMLBody
    # The signature arrives through GetInspector even without a window, text above it
    assert item.HTMLBody.index("Zeile zwei") < item.HTMLBody.index("Viele Grüße")
    assert "NOT sent" in res.note and "privat@gmx.de" in res.note
    assert res.to == "a@b.de" and res.cc == "c@d.de"
    assert not res.opened_window


def test_reply_inserts_above_quote_and_moves_to_account(outlook):
    backend, ns = outlook
    hit = backend.search("privat@gmx.de", SearchQuery())[0]
    res = backend.create_draft("privat@gmx.de", DraftRequest(account=None, to=[], cc=[], bcc=[], subject=None, body="Danke!", reply_to_id=hit.id))
    store = ns.stores["S-privat@gmx.de"]
    item = store.GetDefaultFolder(16).Items.Item(1)
    assert item.HTMLBody.index("Danke!") < item.HTMLBody.index("id=reply")
    assert res.subject == "AW: Paket kommt"


def test_reply_all_with_window_keeps_signature(outlook):
    backend, ns = outlook
    hit = backend.search("ich@firma.de", SearchQuery(text="40 Gitterboxen"))[0]
    res = backend.create_draft(
        "ich@firma.de",
        DraftRequest(account=None, to=[], cc=["extra@kunde.de"], bcc=[], subject=None, body="Moin Max,\n\nbis Freitag klappt.", reply_to_id=hit.id, reply_all=True, open_window=True),
    )
    assert res.opened_window
    item = ns.GetItemFromID(decode_id(res.id)["e"])
    assert item.displayed
    html = item.HTMLBody
    assert html.index("bis Freitag klappt") < html.index("Viele Grüße") < html.index("id=reply")
    assert 'class="MsoNormal"' in html
    assert "chef@kunde.de" in res.cc and "extra@kunde.de" in res.cc


def test_draft_requires_recipient(outlook):
    backend, _ = outlook
    with pytest.raises(BackendError):
        backend.create_draft("ich@firma.de", DraftRequest(account=None, to=[], cc=[], bcc=[], subject="x", body="y"))


def _by_subject(backend, subject, account="ich@firma.de"):
    """Pick a mail by its exact subject: a text search also matches the body."""
    hits = [m for m in backend.search(account, SearchQuery(folder="all")) if m.subject == subject]
    assert hits, f"no mail with subject {subject!r}"
    return hits[0]


def test_delete_goes_to_deleted_items_and_can_be_undone(outlook):
    backend, _ns = outlook
    hit = _by_subject(backend, "Rechnung 4711")
    new_id = backend.delete(decode_id(hit.id))
    # Not gone, just moved: the user can still fish it out of the trash
    assert backend.read(decode_id(new_id)).summary.folder == "Gelöschte Elemente"


def test_delete_from_the_trash_stays_put(outlook):
    backend, _ns = outlook
    hit = _by_subject(backend, "Rechnung 4711")
    in_trash = decode_id(backend.delete(decode_id(hit.id)))
    again = decode_id(backend.delete(in_trash))
    assert again["e"] == in_trash["e"]  # no second move, no duplicate


def test_permanent_delete_really_removes_the_mail(outlook):
    backend, _ns = outlook
    hit = _by_subject(backend, "Rechnung 4711")
    ref = decode_id(hit.id)
    assert backend.delete(ref, permanent=True) == ""
    with pytest.raises(NotFound):
        backend.read(ref)
    assert not any(m.subject == "Rechnung 4711" for m in backend.search("ich@firma.de", SearchQuery(folder="all")))


def test_move_mark_send(outlook):
    backend, ns = outlook
    hit = backend.search("ich@firma.de", SearchQuery(text="Rechnung"))[0]
    ref = decode_id(hit.id)
    backend.mark(ref, read=False, flagged=True)
    item = ns.GetItemFromID(ref["e"])
    assert item.UnRead and item.FlagStatus == 2
    backend.mark(ref, read=True, flagged=False)
    assert not item.UnRead and item.FlagStatus == 0
    new_id = backend.move(ref, "Kunden")
    assert backend.read(decode_id(new_id)).summary.folder == "Posteingang/Kunden"
    draft = backend.create_draft("ich@firma.de", DraftRequest(account=None, to=["x@y.de"], cc=[], bcc=[], subject="S", body="B"))
    backend.send_draft(decode_id(draft.id))
    assert ns.GetItemFromID(decode_id(draft.id)["e"]).sent_out
    with pytest.raises(BackendError):
        backend.send_draft(decode_id(hit.id) | {"e": decode_id(new_id)["e"]})


def test_thread_fallback_by_topic(outlook):
    backend, _ = outlook
    hit = backend.search("ich@firma.de", SearchQuery(text="Passt"))[0]
    thread = backend.thread(decode_id(hit.id))
    subjects = [m.summary.subject for m in thread]
    assert subjects == ["Angebot Gitterboxen", "Angebot Gitterboxen", "AW: Angebot Gitterboxen"]


def test_com_datetime_handles_none_and_4501():
    assert com_datetime(None) is None
    assert com_datetime(datetime(4501, 1, 1)) is None
    assert com_datetime(datetime(2026, 1, 2, 3, 4)).hour == 3


def test_insert_body_plain_format():
    class Plain:
        BodyFormat = 1
        Body = "Zitat"

    p = Plain()
    insert_body(p, "Hallo")
    assert p.Body.startswith("Hallo") and p.Body.endswith("Zitat")


def test_outlook_start_failure_is_friendly():
    def boom():
        raise OSError("class not registered")

    backend = OutlookBackend(Source(kind="outlook"), app_factory=boom, runner=DirectRunner())
    with pytest.raises(BackendError, match="Classic Outlook could not be started"):
        backend.accounts()


def test_shared_mailbox_and_archive_are_searchable(outlook):
    backend, ns = outlook
    res = backend.search("Vertrieb (freigegeben)", SearchQuery())
    assert [m.subject for m in res] == ["Anfrage Europaletten"]
    draft = backend.create_draft("Vertrieb (freigegeben)", DraftRequest(account=None, to=[], cc=[], bcc=[], subject=None, body="Gern!", reply_to_id=res[0].id))
    assert ns.stores["S-Vertrieb (freigegeben)"].GetDefaultFolder(16).Items.Count == 1
    assert draft.account == "Vertrieb (freigegeben)"
    assert backend.search("Archiv 2024", SearchQuery())[0].subject == "Alte Rechnung"
    assert "Öffentliche Ordner" not in [a.email for a in backend.accounts()]


def test_search_all_folders(outlook):
    backend, _ = outlook
    res = backend.search("ich@firma.de", SearchQuery(folder="all", text="Altkunde"))
    assert [m.folder for m in res] == ["Posteingang/Kunden"]
    res = backend.search("ich@firma.de", SearchQuery(folder="alle", limit=50))
    folders = {m.folder for m in res}
    assert {"Posteingang", "Posteingang/Kunden", "Gesendete Elemente"} <= folders


def test_forward_keeps_attachments_and_dedupes(outlook):
    backend, ns = outlook
    hit = backend.search("ich@firma.de", SearchQuery(text="40 Gitterboxen"))[0]
    res = backend.create_draft("ich@firma.de", DraftRequest(account=None, to=["chef@firma.de", "chef@firma.de"], cc=[], bcc=[], subject=None, body="Zur Info.", forward_of=hit.id))
    item = ns.GetItemFromID(decode_id(res.id)["e"])
    assert item.Subject.startswith("WG:")
    assert item.Attachments.Count == 2
    assert res.to == "chef@firma.de"
    with pytest.raises(BackendError):
        backend.create_draft("ich@firma.de", DraftRequest(account=None, to=[], cc=[], bcc=[], subject=None, body="x", forward_of=hit.id))


def test_reply_does_not_duplicate_recipient(outlook):
    backend, _ = outlook
    hit = backend.search("ich@firma.de", SearchQuery(text="Rechnung"))[0]
    res = backend.create_draft("ich@firma.de", DraftRequest(account=None, to=["Buchhaltung <rechnung@lieferant.de>"], cc=[], bcc=[], subject=None, body="Danke", reply_to_id=hit.id))
    assert res.to.count("rechnung@lieferant.de") == 1


def test_discard_only_drafts(outlook):
    backend, ns = outlook
    draft = backend.create_draft("ich@firma.de", DraftRequest(account=None, to=["x@y.de"], cc=[], bcc=[], subject="S", body="B"))
    backend.discard_draft(decode_id(draft.id))
    assert ns.stores["S-ich@firma.de"].GetDefaultFolder(16).Items.Count == 0
    hit = backend.search("ich@firma.de", SearchQuery(limit=1))[0]
    with pytest.raises(BackendError, match="not a draft"):
        backend.discard_draft(decode_id(hit.id))


def test_attachment_by_name_and_position(outlook):
    backend, _ = outlook
    hit = backend.search("ich@firma.de", SearchQuery(text="40 Gitterboxen"))[0]
    name, data = backend.attachment(decode_id(hit.id), "anfrage.txt")
    assert name == "anfrage.txt" and b"Liefertermin" in data
    assert backend.attachment(decode_id(hit.id), "1")[0] == "anfrage.txt"
    with pytest.raises(NotFound, match="anfrage.txt"):
        backend.attachment(decode_id(hit.id), "logo.png")


def test_timeout_message_mentions_security_prompt():
    import time

    from ggpost.backends.outlook_com import ComRunner

    runner = ComRunner.__new__(ComRunner)
    from concurrent.futures import ThreadPoolExecutor

    runner._pool = ThreadPoolExecutor(max_workers=1)
    with pytest.raises(BackendError, match="security prompt"):
        runner.call(lambda: time.sleep(0.5), timeout=0.05)


def test_thread_with_unsent_draft_does_not_crash(outlook):
    backend, ns = outlook
    hit = backend.search("ich@firma.de", SearchQuery(text="Passt"))[0]
    backend.create_draft("ich@firma.de", DraftRequest(account=None, to=[], cc=[], bcc=[], subject=None, body="x", reply_to_id=hit.id))
    draft_item = ns.stores["S-ich@firma.de"].GetDefaultFolder(16).Items.Item(1)
    draft_item.ConversationTopic = "Angebot Gitterboxen"
    ns.stores["S-ich@firma.de"].GetDefaultFolder(6).Items.append(draft_item)  # liegt im Posteingang, ohne Datum
    thread = backend.thread(decode_id(hit.id))
    assert thread[-1].sort_key is None  # undatierter Entwurf landet am Ende


class FakeComError(Exception):
    """Same shape as pywintypes.com_error: (hresult, text, excepinfo, argerr)."""

    def __init__(self, hresult, scode=None):
        excepinfo = (0, "Microsoft Outlook", "error", None, 0, scode) if scode is not None else None
        super().__init__(hresult, "Ausnahmefehler", excepinfo, None)
        self.excepinfo = excepinfo


def test_nested_security_hresult_gives_help():
    from ggpost.backends.outlook_com import _hresult

    assert _hresult(FakeComError(-2147352567, -2147467260)) == -2147467260
    assert _hresult(FakeComError(-2147023174)) == -2147023174

    class App:
        def GetNamespace(self, _n):  # noqa: N802
            raise FakeComError(-2147352567, -2147467260)

    backend = OutlookBackend(Source(kind="outlook"), app_factory=App, runner=DirectRunner())
    with pytest.raises(BackendError, match="Trust Center"):
        backend._call(lambda: App().GetNamespace("MAPI"))


def test_reconnect_on_nested_rpc_error():
    app, _ns = fake_outlook.build()
    calls = {"n": 0}

    def flaky():
        calls["n"] += 1
        if calls["n"] == 1:
            raise FakeComError(-2147352567, -2147023174)
        return "ok"

    backend = OutlookBackend(Source(kind="outlook"), app_factory=lambda: app, runner=DirectRunner())
    assert backend._call(flaky) == "ok" and calls["n"] == 2
