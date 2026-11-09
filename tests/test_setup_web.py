from __future__ import annotations

import http.client
import json

import pytest

from ggpost.config import ServerSettings, load_config
from ggpost.detect import aiapps, autoconfig, clients
from ggpost.setup.web import SetupApp


def fake_detection():
    return clients.Detection(
        platform="win32",
        clients=[
            clients.DetectedClient(
                id="outlook_classic",
                name="Outlook (classic), profile Outlook",
                in_use=True,
                backend="outlook",
                accounts=[
                    clients.DetectedAccount("info@palettenfirma.de", kind="Exchange", via="outlook_classic"),
                    clients.DetectedAccount("vertrieb@palettenfirma.de", kind="IMAP", via="outlook_classic"),
                ],
            ),
            clients.DetectedClient(
                id="thunderbird",
                name="Thunderbird",
                backend="imap",
                accounts=[
                    clients.DetectedAccount(
                        "privat@gmx.de",
                        kind="IMAP",
                        via="thunderbird",
                        imap=ServerSettings("imap.gmx.net", 993, "ssl", "privat@gmx.de"),
                        smtp=ServerSettings("mail.gmx.net", 587, "starttls", "privat@gmx.de"),
                    )
                ],
            ),
        ],
        default_mail_app="outlook_classic",
    )


def fake_apps():
    return [
        aiapps.AiApp("claude_desktop", "Claude Desktop", True),
        aiapps.AiApp("claude_code", "Claude Code", True, registered=True),
        aiapps.AiApp("codex", "Codex", False),
    ]


@pytest.fixture
def app():
    registered = []
    a = SetupApp(
        detect_fn=fake_detection,
        apps_fn=fake_apps,
        discover_fn=lambda e: autoconfig.Discovery(email=e, provider="GMX", imap=ServerSettings("imap.gmx.net", 993, "ssl", e), hint="IMAP freischalten"),
        test_imap_fn=lambda s, e, p: "Posteingang: 12 Mails, 3 ungelesen" if p == "richtig" else (_ for _ in ()).throw(ValueError("Login failed")),
        test_outlook_fn=lambda: {"message": "2 Konten", "accounts": [{"email": "info@palettenfirma.de", "detail": "Outlook, Exchange"}]},
        register_fn=lambda ids: registered.extend(ids) or [aiapps.RegisterResult(i, True, "ok") for i in ids],
    )
    a.registered = registered
    a.start(open_browser=False)
    yield a
    a.stop()


def req(app, method, path, body=None, token=True, host=None, origin=None):
    conn = http.client.HTTPConnection("127.0.0.1", app.port, timeout=5)
    headers = {"Host": host or f"127.0.0.1:{app.port}"}
    if token:
        headers["X-Post-Token"] = app.token
    if origin:
        headers["Origin"] = origin
    payload = None
    if body is not None:
        payload = json.dumps(body).encode()
        headers["Content-Type"] = "application/json"
    conn.request(method, path, body=payload, headers=headers)
    resp = conn.getresponse()
    raw = resp.read()
    conn.close()
    try:
        return resp.status, json.loads(raw), resp
    except ValueError:
        return resp.status, raw, resp


def test_index_requires_token(app):
    status, body, resp = req(app, "GET", f"/?t={app.token}", token=False)
    assert status == 200 and b"Set up post." in body
    assert "script-src 'self'" in resp.getheader("Content-Security-Policy")
    status, body, _ = req(app, "GET", "/?t=falsch", token=False)
    assert b"no longer valid" in body


def test_dns_rebinding_and_csrf_blocked(app):
    assert req(app, "GET", "/api/state", host="evil.example:80")[0] == 403
    assert req(app, "GET", "/api/state", token=False)[0] == 403
    assert req(app, "POST", "/api/finish", {}, origin="https://evil.example")[0] == 403
    assert req(app, "POST", "/api/finish", {}, token=False)[0] == 403


def test_static_files(app):
    status, body, resp = req(app, "GET", "/static/app.js", token=False)
    assert status == 200 and resp.getheader("Content-Type").startswith("text/javascript")
    assert req(app, "GET", "/static/fonts/dm-sans-latin-400-normal.woff2", token=False)[0] == 200
    assert req(app, "GET", "/static/../web.py", token=False)[0] == 404
    assert req(app, "GET", "/static/index.html", token=False)[0] == 404


def test_full_flow(app, isolated):
    status, state, _ = req(app, "GET", "/api/state")
    assert status == 200
    assert state["detection"]["clients"][0]["accounts"][0]["email"] == "info@palettenfirma.de"
    assert not state["config"]["exists"]

    status, d, _ = req(app, "POST", "/api/discover", {"email": "neu@gmx.de"})
    assert status == 200 and d["imap"]["host"] == "imap.gmx.net" and d["has_password"] is False
    assert req(app, "POST", "/api/discover", {"email": "kaputt"})[0] == 422

    imap = {"host": "imap.gmx.net", "port": 993, "security": "ssl", "username": "privat@gmx.de"}
    status, r, _ = req(app, "POST", "/api/test", {"email": "privat@gmx.de", "imap": imap, "password": "falsch"})
    assert status == 422 and "Login" in r["error"]
    status, r, _ = req(app, "POST", "/api/test", {"email": "privat@gmx.de", "imap": imap, "password": "richtig"})
    assert r["ok"]

    status, r, _ = req(app, "POST", "/api/save", {"sources": [{"kind": "imap", "email": "privat@gmx.de", "imap": imap}]})
    assert status == 422 and "password" in r["error"]
    assert load_config() is None

    status, r, _ = req(app, "POST", "/api/save", {
        "sources": [
            {"kind": "outlook", "outlook_accounts": ["info@palettenfirma.de"]},
            {"kind": "imap", "email": "privat@gmx.de", "imap": imap},
        ],
        "passwords": {"privat@gmx.de": "richtig"},
        "permissions": {"modify": False, "send": False},
    })
    assert status == 200 and r["ok"]
    cfg = load_config()
    assert [s.kind for s in cfg.sources] == ["outlook", "imap"]
    assert cfg.sources[0].outlook_accounts == ["info@palettenfirma.de"]
    assert cfg.permissions.modify is False
    assert isolated.get_password("post-mail-connector", "privat@gmx.de") == "richtig"
    assert "richtig" not in (load_config.__globals__["config_path"]()).read_text(encoding="utf-8")

    status, state, _ = req(app, "GET", "/api/state")
    assert state["config"]["sources"][1]["has_password"] is True

    status, r, _ = req(app, "POST", "/api/register", {"apps": ["claude_desktop"]})
    assert r["results"][0]["ok"] and app.registered == ["claude_desktop"]

    status, r, _ = req(app, "POST", "/api/finish", {})
    assert r["ok"] and app.wait(2)


def test_rejects_bad_input(app):
    assert req(app, "POST", "/api/save", {"sources": []})[0] == 422
    bad = {"host": "x", "port": 99999, "security": "ssl"}
    assert req(app, "POST", "/api/test", {"email": "a@b.de", "imap": bad, "password": "x"})[0] == 422
    assert req(app, "POST", "/api/nope", {})[0] == 404


def test_stored_password_never_goes_to_a_new_host(app, isolated):
    from ggpost.config import Config, ServerSettings, Source, save_config

    seen = []
    app.test_imap_fn = lambda s, e, p: seen.append((s.host, p)) or "ok"
    save_config(Config(sources=[Source(kind="imap", email="privat@gmx.de", imap=ServerSettings("imap.gmx.net", 993, "ssl", "privat@gmx.de"))]))
    isolated.set_password("post-mail-connector", "privat@gmx.de", "geheim")
    evil = {"host": "attacker.example", "port": 993, "security": "ssl", "username": "privat@gmx.de"}
    status, r, _ = req(app, "POST", "/api/test", {"email": "privat@gmx.de", "imap": evil})
    assert status == 422 and "password" in r["error"] and seen == []
    same = {"host": "imap.gmx.net", "port": 993, "security": "ssl", "username": "privat@gmx.de"}
    status, r, _ = req(app, "POST", "/api/test", {"email": "privat@gmx.de", "imap": same})
    assert status == 200 and seen == [("imap.gmx.net", "geheim")]


def test_state_survives_corrupt_config(app):
    from ggpost.config import config_dir, config_path

    config_dir().mkdir(parents=True, exist_ok=True)
    config_path().write_text("{kaputt", encoding="utf-8")
    status, state, _ = req(app, "GET", "/api/state")
    assert status == 200 and state["config"]["exists"] is False
