"""Does post. pick the workable route by itself when there is no settings file?"""

from __future__ import annotations

from ggpost.config import ServerSettings
from ggpost.detect.clients import DetectedAccount, DetectedClient, Detection
from ggpost.service import choose_sources


def _outlook(accounts: list[str]) -> DetectedClient:
    return DetectedClient(
        id="outlook_classic",
        name="Outlook (klassisch)",
        backend="outlook",
        accounts=[DetectedAccount(email=a, via="outlook_classic") for a in accounts],
    )


def _thunderbird(*accounts: DetectedAccount) -> DetectedClient:
    return DetectedClient(id="thunderbird", name="Thunderbird", backend="imap", accounts=list(accounts))


def _tb_account(email: str, host: str = "imap.example.org") -> DetectedAccount:
    return DetectedAccount(
        email=email,
        via="thunderbird",
        kind="IMAP",
        imap=ServerSettings(host=host, port=993, security="ssl", username=email),
    )


def _det(*clients: DetectedClient, default: str = "") -> Detection:
    return Detection(platform="win32", clients=list(clients), default_mail_app=default)


def test_outlook_with_accounts_wins_because_it_needs_no_password():
    choice = choose_sources(_det(_outlook(["chef@firma.de"]), default="outlook_classic"), env={})
    assert choice.via == "outlook_classic"
    assert [s.kind for s in choice.sources] == ["outlook"]
    assert not choice.hint


def test_thunderbird_is_used_when_outlook_is_empty():
    det = _det(_outlook([]), _thunderbird(_tb_account("ich@gmx.de")))
    choice = choose_sources(det, env={})
    assert choice.via == "thunderbird"
    assert [(s.kind, s.email) for s in choice.sources] == [("imap", "ich@gmx.de")]
    assert choice.sources[0].imap.host == "imap.example.org"


def test_the_default_mail_program_decides_the_order():
    det = _det(_outlook(["chef@firma.de"]), _thunderbird(_tb_account("ich@gmx.de")), default="thunderbird")
    choice = choose_sources(det, env={})
    # Thunderbird is the default program, so it comes first; Outlook stays reachable as a second route
    assert [s.kind for s in choice.sources] == ["imap", "outlook"]
    det_outlook = _det(_outlook(["chef@firma.de"]), _thunderbird(_tb_account("ich@gmx.de")), default="outlook_classic")
    assert [s.kind for s in choose_sources(det_outlook, env={}).sources] == ["outlook", "imap"]


def test_post_email_comes_first_and_only_once():
    det = _det(_thunderbird(_tb_account("ich@gmx.de"), _tb_account("zweit@web.de")))
    choice = choose_sources(det, env={"POST_EMAIL": "ich@gmx.de"})
    assert [s.email for s in choice.sources] == ["ich@gmx.de", "zweit@web.de"]
    assert choice.via.startswith("POST_EMAIL")


def test_a_pop3_account_gets_the_automatic_server_lookup():
    pop = DetectedAccount(email="alt@web.de", via="thunderbird", kind="POP3")
    choice = choose_sources(_det(_thunderbird(pop)), env={})
    assert choice.sources[0].imap is None  # looked up on first access


def test_outlook_without_an_account_says_what_is_wrong():
    choice = choose_sources(_det(_outlook([])), env={})
    assert not choice.sources
    assert "no account" in choice.hint
    assert "post_setup" in choice.hint


def test_only_the_new_outlook_points_at_imap():
    neu = DetectedClient(id="outlook_new", name="Outlook (neu)", backend="imap")
    choice = choose_sources(_det(neu), env={})
    assert not choice.sources
    assert "new Outlook" in choice.hint and "post_setup" in choice.hint


def test_finding_nothing_still_leaves_a_clear_hint():
    choice = choose_sources(_det(), env={})
    assert not choice.sources
    assert "No mail program found" in choice.hint


def test_detection_may_crash_without_taking_the_server_down(monkeypatch):
    import ggpost.detect.clients as clients

    monkeypatch.setattr(clients, "detect", lambda **kw: (_ for _ in ()).throw(RuntimeError("registry broken")))
    choice = choose_sources(env={})
    assert not choice.sources and choice.hint
