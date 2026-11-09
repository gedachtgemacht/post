"""A real end-to-end test: post. as its own process over stdio, the way Claude Desktop starts it."""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pytest
from mcp import Client, StdioServerParameters

from ggpost.config import Config, save_config

ROOT = Path(__file__).resolve().parents[1]


def _commands():
    cmds = [("python", sys.executable, ["-m", "ggpost", "serve"])]
    binary = os.environ.get("POST_BINARY")
    if binary:
        cmds.append(("binary", binary, ["serve"]))
    return cmds


@pytest.fixture
def anyio_backend():
    return "asyncio"


@pytest.mark.anyio
@pytest.mark.parametrize("kind,command,args", _commands())
async def test_stdio_roundtrip(kind, command, args, imap_source, imap_server, tmp_path):
    from email.message import EmailMessage

    from imapclient import IMAPClient

    from tests.test_extract import make_pdf

    m = EmailMessage()
    m["From"] = "einkauf@kunde.de"
    m["To"] = imap_source.email
    m["Subject"] = f"Stdio-PDF {kind}"
    m.set_content("Bestellung anbei.")
    m.add_attachment(make_pdf("Bestellnummer 4711"), maintype="application", subtype="pdf", filename="bestellung.pdf")
    c = IMAPClient("127.0.0.1", port=imap_server, ssl=False)
    c.login("demouser", "demopass")
    c.append("INBOX", m.as_bytes())
    c.logout()
    cfg_dir = tmp_path / f"cfg-{kind}"
    save_config(Config(sources=[imap_source]), cfg_dir / "config.json")
    env = dict(os.environ)
    env.update(
        {
            "POST_CONFIG_DIR": str(cfg_dir),
            "POST_EMAIL": imap_source.email,
            "POST_PASSWORD": "demopass",
            "PYTHONPATH": str(ROOT / "src"),
        }
    )
    params = StdioServerParameters(command=command, args=args, env=env)
    async with Client(params, read_timeout_seconds=30) as client:
        tools = {t.name for t in (await client.list_tools()).tools}
        assert {"mail_search", "mail_read", "mail_draft"} <= tools
        res = await client.call_tool("mail_search", {"limit": 2})
        assert not res.is_error, res.content[0].text
        payload = json.loads(res.content[0].text)
        assert payload["count"] == 2
        # Read an attachment: checks that PDF support really made it into the build
        res = await client.call_tool("mail_search", {"query": "Stdio-PDF"})
        hit = json.loads(res.content[0].text)["messages"][0]
        res = await client.call_tool("mail_attachment", {"message_id": hit["id"], "name": "bestellung.pdf"})
        assert not res.is_error, res.content[0].text
        assert "Bestellnummer 4711" in json.loads(res.content[0].text)["content"]
    audit = (cfg_dir / "audit.log").read_text(encoding="utf-8")
    assert '"tool": "mail_search"' in audit and "demopass" not in audit
