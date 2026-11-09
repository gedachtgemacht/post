from __future__ import annotations

import os
import shutil
import socket
import subprocess
import sys
import time

import keyring
import pytest
from keyring.backend import KeyringBackend

from ggpost.config import ServerSettings, Source


class MemoryKeyring(KeyringBackend):
    priority = 1

    def __init__(self) -> None:
        super().__init__()
        self.store: dict[tuple[str, str], str] = {}

    def get_password(self, service, username):
        return self.store.get((service, username))

    def set_password(self, service, username, password):
        self.store[(service, username)] = password

    def delete_password(self, service, username):
        self.store.pop((service, username), None)


@pytest.fixture(autouse=True)
def isolated(tmp_path, monkeypatch):
    """Every test gets its own settings directory and an in-memory keychain."""
    monkeypatch.setenv("POST_CONFIG_DIR", str(tmp_path / "cfg"))
    for var in ("POST_EMAIL", "POST_PASSWORD", "POST_ALLOW_SEND", "POST_ALLOW_MODIFY"):
        monkeypatch.delenv(var, raising=False)
    ring = MemoryKeyring()
    old = keyring.get_keyring()
    keyring.set_keyring(ring)
    yield ring
    keyring.set_keyring(old)


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


@pytest.fixture(scope="session")
def imap_server():
    """A real IMAP server (pymap, in memory) with demo data."""
    exe = shutil.which("pymap")
    if not exe:
        pytest.skip("pymap is not installed")
    if sys.platform == "win32":
        pytest.skip("pymap needs POSIX signal handlers, which Windows does not have; the IMAP tests run on Linux and macOS")
    port = _free_port()
    proc = subprocess.Popen(
        [exe, "--host", "127.0.0.1", "--port", str(port), "--no-tls", "dict", "--demo-data"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    deadline = time.time() + 15
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                break
        except OSError:
            time.sleep(0.1)
    else:
        proc.kill()
        pytest.skip("pymap does not start")
    # Create a drafts folder with a German name (this tests the name matching)
    from imapclient import IMAPClient

    c = IMAPClient("127.0.0.1", port=port, ssl=False)
    c.login("demouser", "demopass")
    if not any(name == "Entwürfe" for _f, _d, name in c.list_folders()):
        c.create_folder("Entwürfe")
    c.create_folder("Archiv")
    c.logout()
    yield port
    proc.terminate()
    try:
        proc.wait(timeout=5)
    except subprocess.TimeoutExpired:
        proc.kill()


@pytest.fixture
def imap_source(imap_server) -> Source:
    return Source(
        kind="imap",
        email="demouser@example.org",
        label="Demo Nutzer",
        imap=ServerSettings("127.0.0.1", imap_server, "plain", "demouser"),
    )


def pytest_configure(config):
    os.environ.setdefault("PYTHONIOENCODING", "utf-8")
    sys.dont_write_bytecode = True


@pytest.fixture(scope="session")
def imap_server_umlaut():
    """A second IMAP server whose password holds umlauts and a euro sign."""
    exe = shutil.which("pymap")
    if not exe or sys.platform == "win32":
        pytest.skip("pymap is not available")
    port = _free_port()
    proc = subprocess.Popen(
        [exe, "--host", "127.0.0.1", "--port", str(port), "--no-tls", "dict", "--demo-data", "--demo-password", "Grüße€2026"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    deadline = time.time() + 15
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=0.5):
                break
        except OSError:
            time.sleep(0.1)
    yield port
    proc.terminate()


@pytest.fixture
def smtp_server():
    """A small SMTP server (aiosmtpd) that records what it receives."""
    controller_mod = pytest.importorskip("aiosmtpd.controller")

    class Sink:
        def __init__(self):
            self.messages = []

        async def handle_DATA(self, server, session, envelope):  # noqa: N802
            self.messages.append((envelope.mail_from, list(envelope.rcpt_tos), envelope.content))
            return "250 OK"

    sink = Sink()
    port = _free_port()
    controller = controller_mod.Controller(sink, hostname="127.0.0.1", port=port)
    controller.start()
    yield port, sink
    controller.stop()
