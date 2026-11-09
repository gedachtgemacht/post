from __future__ import annotations

import json

import pytest

from ggpost import service as svc_mod
from ggpost.backends.base import BackendError
from ggpost.config import Config, ServerSettings, Source, config_dir, save_config
from ggpost.detect.autoconfig import Discovery
from ggpost.service import MailService


def test_env_address_is_added_to_existing_config(monkeypatch):
    save_config(Config(sources=[Source(kind="imap", email="a@firma.de", imap=ServerSettings("imap.firma.de", 993))]))
    monkeypatch.setenv("POST_EMAIL", "b@gmx.de")
    svc = MailService.from_environment()
    assert [s.email for s in svc.config.sources] == ["a@firma.de", "b@gmx.de"]
    monkeypatch.setenv("POST_EMAIL", "${user_config.email}")
    assert [s.email for s in MailService.from_environment().config.sources] == ["a@firma.de"]


def test_discovery_is_cached(monkeypatch):
    calls = []

    def fake_discover(email):
        calls.append(email)
        return Discovery(email=email, imap=ServerSettings("imap.gmx.net", 993, "ssl", email))

    monkeypatch.setattr("ggpost.detect.autoconfig.discover", fake_discover)
    for _ in range(2):
        src = Source(kind="imap", email="x@gmx.de")
        svc_mod._discover_into(src)
        assert src.imap.host == "imap.gmx.net"
    assert calls == ["x@gmx.de"]
    cache = json.loads((config_dir() / "discovery-cache.json").read_text(encoding="utf-8"))
    assert cache["x@gmx.de"]["imap"]["host"] == "imap.gmx.net"


def test_discovery_unsupported_raises_hint(monkeypatch):
    monkeypatch.setattr(
        "ggpost.detect.autoconfig.discover",
        lambda e: Discovery(email=e, supported=False, hint="Microsoft does not allow that."),
    )
    with pytest.raises(BackendError, match="Microsoft"):
        svc_mod._discover_into(Source(kind="imap", email="x@outlook.de"))


def test_env_permissions_override(monkeypatch):
    monkeypatch.setenv("POST_ALLOW_SEND", "true")
    monkeypatch.setenv("POST_ALLOW_MODIFY", "false")
    svc = MailService(Config(sources=[Source(kind="outlook")]), persist=False)
    assert svc.config.permissions.send is True
    assert svc.config.permissions.modify is False


def test_env_cannot_switch_off_saved_permissions(monkeypatch):
    from ggpost.config import Permissions

    save_config(Config(sources=[Source(kind="outlook")], permissions=Permissions(modify=True)))
    monkeypatch.setenv("POST_ALLOW_MODIFY", "false")
    svc = MailService.from_environment()
    assert svc.config.permissions.modify is True
