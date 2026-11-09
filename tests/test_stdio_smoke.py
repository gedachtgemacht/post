"""Smoke test without a mailbox: does the built program run as an MCP server at all?

The big end-to-end test (test_stdio.py) needs an IMAP test server and is
therefore skipped on Windows - which is exactly where post.exe is built. This
test needs no mailbox and so checks, on every platform, what can go wrong with
the build: starting the process, the bundled dependencies (mcp, keyring,
pywin32), the MCP protocol itself, and the gating of the writing tools.

The settings are deliberately saved empty: automatic detection then stays out
of the way, and the test does not start Outlook on a developer machine.
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import pytest
from mcp import Client, StdioServerParameters

from ggpost.config import Config, save_config

ROOT = Path(__file__).resolve().parents[1]

READ_ONLY_TOOLS = {"mail_accounts", "mail_folders", "mail_search", "mail_read", "mail_attachment"}


def _commands():
    cmds = [("python", sys.executable, ["-m", "ggpost", "serve"])]
    binary = os.environ.get("POST_BINARY")
    if binary:
        cmds.append(("binary", binary, ["serve"]))
    return cmds


@pytest.fixture
def anyio_backend():
    return "asyncio"


def _env(cfg_dir: Path, **extra: str) -> dict[str, str]:
    save_config(Config(sources=[]), cfg_dir / "config.json")
    env = dict(os.environ)
    for var in ("POST_EMAIL", "POST_PASSWORD", "POST_ALLOW_SEND", "POST_ALLOW_MODIFY"):
        env.pop(var, None)
    env.update({"POST_CONFIG_DIR": str(cfg_dir), "PYTHONPATH": str(ROOT / "src")})
    env.update(extra)
    return env


@pytest.mark.anyio
@pytest.mark.parametrize("kind,command,args", _commands())
async def test_server_startet_und_meldet_seine_werkzeuge(kind, command, args, tmp_path):
    params = StdioServerParameters(command=command, args=args, env=_env(tmp_path / f"a-{kind}"))
    async with Client(params, read_timeout_seconds=30) as client:
        tools = {t.name for t in (await client.list_tools()).tools}
        assert READ_ONLY_TOOLS <= tools
        assert {"mail_draft", "mail_discard_draft", "post_setup"} <= tools
        # Without the permission, nothing that writes may show up
        assert "mail_send" not in tools
        assert "mail_move" not in tools and "mail_mark" not in tools

        res = await client.call_tool("mail_accounts", {})
        assert not res.is_error, res.content[0].text
        status = json.loads(res.content[0].text)
        assert status["configured"] is False
        assert status["permissions"] == {"read": True, "draft": True, "move_and_mark": False, "send": False}

        # Without a mailbox the error has to point at the next step, not crash
        res = await client.call_tool("mail_search", {})
        assert res.is_error
        assert "post_setup" in res.content[0].text


@pytest.mark.anyio
@pytest.mark.parametrize("kind,command,args", _commands())
async def test_rechte_lassen_sich_ueber_die_umgebung_freischalten(kind, command, args, tmp_path):
    env = _env(tmp_path / f"b-{kind}", POST_ALLOW_MODIFY="1", POST_ALLOW_SEND="1")
    params = StdioServerParameters(command=command, args=args, env=env)
    async with Client(params, read_timeout_seconds=30) as client:
        tools = {t.name for t in (await client.list_tools()).tools}
        assert {"mail_move", "mail_mark", "mail_send"} <= tools
        status = json.loads((await client.call_tool("mail_accounts", {})).content[0].text)
        assert status["permissions"]["move_and_mark"] is True
        assert status["permissions"]["send"] is True


@pytest.mark.anyio
@pytest.mark.parametrize("kind,command,args", _commands())
async def test_nicht_ersetzte_platzhalter_aus_dem_manifest_werden_ignoriert(kind, command, args, tmp_path):
    """Claude Desktop leaves empty fields as ``${user_config.email}``."""
    env = _env(
        tmp_path / f"c-{kind}",
        POST_EMAIL="${user_config.email}",
        POST_PASSWORD="${user_config.password}",
        POST_ALLOW_MODIFY="${user_config.allow_organize}",
    )
    params = StdioServerParameters(command=command, args=args, env=env)
    async with Client(params, read_timeout_seconds=30) as client:
        tools = {t.name for t in (await client.list_tools()).tools}
        assert "mail_move" not in tools  # a placeholder must not grant anything
        status = json.loads((await client.call_tool("mail_accounts", {})).content[0].text)
        assert status["accounts"] == []
