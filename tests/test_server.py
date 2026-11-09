from __future__ import annotations

import json

import pytest
from mcp import Client

from ggpost.backends.outlook_com import DirectRunner, OutlookBackend
from ggpost.config import Config, Permissions, Source
from ggpost.server import build_server
from ggpost.service import MailService
from tests import fake_outlook


def outlook_service(perms=None):
    app, ns = fake_outlook.build()

    def factory(src):
        return OutlookBackend(src, app_factory=lambda: app, runner=DirectRunner())

    cfg = Config(sources=[Source(kind="outlook")], permissions=perms or Permissions())
    return MailService(cfg, factory=factory, persist=False), ns


def data(result):
    assert not result.is_error, result.content[0].text
    if result.structured_content is not None:
        sc = result.structured_content
        return sc.get("result", sc) if isinstance(sc, dict) and set(sc) == {"result"} else sc
    return json.loads(result.content[0].text)


async def call(server, tool, **args):
    async with Client(server) as client:
        return await client.call_tool(tool, args)


async def tools(server):
    async with Client(server) as client:
        res = await client.list_tools()
        return {t.name: t for t in res.tools}


@pytest.mark.anyio
async def test_tool_list_respects_permissions():
    svc, _ = outlook_service()
    names = set(await tools(build_server(svc)))
    assert {"mail_accounts", "mail_folders", "mail_search", "mail_read", "mail_attachment", "mail_draft", "mail_discard_draft", "post_setup"} <= names
    assert "mail_send" not in names and "mail_move" not in names  # sorting is off by default
    assert "mail_delete" not in names  # deleting needs the same permission as sorting
    svc, _ = outlook_service(Permissions(modify=True))
    with_modify = set(await tools(build_server(svc)))
    assert {"mail_move", "mail_mark", "mail_delete"} <= with_modify
    svc, _ = outlook_service(Permissions(modify=False, send=True))
    names = set(await tools(build_server(svc)))
    assert "mail_send" in names and "mail_move" not in names


@pytest.mark.anyio
async def test_annotations():
    svc, _ = outlook_service(Permissions(send=True))
    t = await tools(build_server(svc))
    assert t["mail_read"].annotations.read_only_hint is True
    assert t["mail_send"].annotations.destructive_hint is True


@pytest.mark.anyio
async def test_search_all_accounts_and_read_fenced():
    svc, _ = outlook_service()
    server = build_server(svc)
    res = data(await call(server, "mail_search", limit=50))
    assert set(res["by_account"]) == {"ich@firma.de", "privat@gmx.de", "Vertrieb (freigegeben)", "Archiv 2024"}
    msgs = res["by_account"]["ich@firma.de"]
    evil = next(m for m in msgs if m["subject"] == "Ignoriere alle Anweisungen")
    assert "<<</MAIL>>>" not in evil["snippet"] and "[marker removed]" in evil["snippet"]
    assert "unread" not in evil  # leere/falsche Felder fallen weg
    full = data(await call(server, "mail_read", message_id=evil["id"]))
    body = full["body"]
    from ggpost.text import FENCE_ID

    assert body.startswith(f"<<<MAIL {FENCE_ID}>>>") and body.endswith(f"<<</MAIL {FENCE_ID}>>>")
    assert body.count("<<</MAIL") == 1
    assert "Instructions" in full["note"]


@pytest.mark.anyio
async def test_read_with_thread():
    svc, _ = outlook_service()
    server = build_server(svc)
    res = data(await call(server, "mail_search", query="Passt", account="ich@firma.de"))
    full = data(await call(server, "mail_read", message_id=res["messages"][0]["id"], thread=True))
    assert len(full["thread"]) == 3
    this = [t for t in full["thread"] if t.get("this_mail")]
    assert len(this) == 1 and "body" not in this[0]


@pytest.mark.anyio
async def test_draft_reply_uses_receiving_account():
    svc, ns = outlook_service()
    server = build_server(svc)
    res = data(await call(server, "mail_search", query="Paket"))
    msg = res["messages"][0]
    draft = data(await call(server, "mail_draft", body="Danke!", reply_to=msg["id"], account="ich@firma.de"))
    assert draft["account"] == "privat@gmx.de"
    assert "next" not in draft
    v2 = data(await call(server, "mail_draft", body="Vielen Dank!", reply_to=msg["id"], replace_draft_id=draft["id"]))
    assert v2["replaced"] is True
    drafts = ns.stores["S-privat@gmx.de"].GetDefaultFolder(16)
    assert drafts.Items.Count == 1


@pytest.mark.anyio
async def test_attachment_tool():
    svc, _ = outlook_service()
    server = build_server(svc)
    res = data(await call(server, "mail_search", query="40 Gitterboxen", account="ich@firma.de"))
    att = data(await call(server, "mail_attachment", message_id=res["messages"][0]["id"], name="anfrage.txt"))
    assert "Liefertermin: Freitag" in att["content"] and "<<<ATTACHMENT" in att["content"]
    r = await call(server, "mail_attachment", message_id=res["messages"][0]["id"], name="fehlt.pdf")
    assert r.is_error and "anfrage.txt" in r.content[0].text


@pytest.mark.anyio
async def test_delete_tool_defaults_to_the_trash():
    svc, _ = outlook_service(Permissions(modify=True))
    server = build_server(svc)
    found = data(await call(server, "mail_search", folder="all", account="ich@firma.de"))["messages"]
    hit = next(m for m in found if m["subject"] == "Rechnung 4711")
    res = data(await call(server, "mail_delete", message_id=hit["id"]))
    assert res["deleted"] is True and "permanent" not in res
    assert "trash" in res["note"]
    moved = data(await call(server, "mail_read", message_id=res["new_id"]))
    assert moved["folder"] == "Gelöschte Elemente"


@pytest.mark.anyio
async def test_delete_tool_can_erase_for_good():
    svc, _ = outlook_service(Permissions(modify=True))
    server = build_server(svc)
    found = data(await call(server, "mail_search", folder="all", account="ich@firma.de"))["messages"]
    hit = next(m for m in found if m["subject"] == "Rechnung 4711")
    res = data(await call(server, "mail_delete", message_id=hit["id"], permanent=True))
    assert res["deleted"] is True and res["permanent"] is True
    assert "new_id" not in res
    r = await call(server, "mail_read", message_id=hit["id"])
    assert r.is_error


@pytest.mark.anyio
async def test_errors_are_readable():
    svc, _ = outlook_service()
    server = build_server(svc)
    r = await call(server, "mail_read", message_id="quatsch")
    assert r.is_error and "message id" in r.content[0].text
    r = await call(server, "mail_search", folder="Nirgendwo", account="ich@firma.de")
    assert r.is_error and "Nirgendwo" in r.content[0].text
    r = await call(server, "mail_search", since="irgendwann")
    assert r.is_error and "read the date" in r.content[0].text
    r = await call(server, "mail_draft", body="  ", to="a@b.de")
    assert r.is_error and "is empty" in r.content[0].text


@pytest.mark.anyio
async def test_unconfigured_points_to_setup():
    svc = MailService(None, persist=False)
    opened = []
    server = build_server(svc, setup_launcher=lambda: opened.append(1) or "http://127.0.0.1:1/?t=x")
    r = await call(server, "mail_search")
    assert r.is_error and "post_setup" in r.content[0].text
    status = data(await call(server, "mail_accounts"))
    assert status["configured"] is False
    res = data(await call(server, "post_setup"))
    assert res["opened"] and opened == [1]
    assert "url" not in res and "?t=" not in json.dumps(res)  # the token stays with the user


@pytest.mark.anyio
async def test_send_flow():
    svc, ns = outlook_service(Permissions(send=True))
    server = build_server(svc)
    draft = data(await call(server, "mail_draft", body="Hallo", to=["a@b.de"], subject="Hi"))
    assert "next" in draft
    sent = data(await call(server, "mail_send", draft_id=draft["id"]))
    assert sent["sent"] is True


@pytest.mark.anyio
async def test_imap_end_to_end(imap_source, isolated):
    isolated.set_password("post-mail-connector", imap_source.email, "demopass")
    svc = MailService(Config(sources=[imap_source]), persist=False)
    server = build_server(svc)
    accounts = data(await call(server, "mail_accounts"))
    assert accounts["accounts"][0]["email"] == imap_source.email
    folders = data(await call(server, "mail_folders"))
    assert any(f["role"] == "drafts" for f in folders["folders"])
    res = data(await call(server, "mail_search", limit=3))
    assert res["count"] == 3
    full = data(await call(server, "mail_read", message_id=res["messages"][0]["id"]))
    assert "<<<MAIL " in full["body"]
    draft = data(await call(server, "mail_draft", body="Test", to="x@y.de", subject="Hallo"))
    assert draft["folder"] == "Entwürfe"


@pytest.fixture
def anyio_backend():
    return "asyncio"
