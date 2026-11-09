from __future__ import annotations

import json
from pathlib import Path

import pytest

from ggpost.detect import aiapps, autoconfig, clients

ISPDB_GMX = """<?xml version="1.0"?>
<clientConfig version="1.1">
  <emailProvider id="gmx.net">
    <domain>gmx.de</domain>
    <displayName>GMX Freemail</displayName>
    <incomingServer type="pop3"><hostname>pop.gmx.net</hostname><port>995</port><socketType>SSL</socketType></incomingServer>
    <incomingServer type="imap">
      <hostname>imap.gmx.net</hostname><port>143</port><socketType>STARTTLS</socketType>
      <username>%EMAILADDRESS%</username><authentication>password-cleartext</authentication>
    </incomingServer>
    <incomingServer type="imap">
      <hostname>imap.gmx.net</hostname><port>993</port><socketType>SSL</socketType>
      <username>%EMAILADDRESS%</username><authentication>password-cleartext</authentication>
    </incomingServer>
    <outgoingServer type="smtp">
      <hostname>mail.gmx.net</hostname><port>587</port><socketType>STARTTLS</socketType>
      <username>%EMAILLOCALPART%</username><authentication>password-cleartext</authentication>
    </outgoingServer>
  </emailProvider>
</clientConfig>"""

OAUTH_ONLY = """<clientConfig><emailProvider id="x"><incomingServer type="imap">
<hostname>imap.x.com</hostname><port>993</port><socketType>SSL</socketType><authentication>OAuth2</authentication>
</incomingServer></emailProvider></clientConfig>"""


def test_parse_autoconfig_prefers_ssl_and_substitutes():
    d = autoconfig.parse_autoconfig(ISPDB_GMX, "max@gmx.de")
    assert d.imap.host == "imap.gmx.net" and d.imap.port == 993 and d.imap.security == "ssl"
    assert d.imap.username == "max@gmx.de"
    assert d.smtp.username == "max" and d.smtp.security == "starttls"
    assert d.provider == "GMX Freemail"


def test_parse_autoconfig_rejects_oauth_only_and_garbage():
    assert autoconfig.parse_autoconfig(OAUTH_ONLY, "a@x.com") is None
    assert autoconfig.parse_autoconfig("<html>nope", "a@x.com") is None


def _discover(email, pages=None, mx=None, srv=None, guess=None):
    pages = pages or {}
    return autoconfig.discover(
        email,
        http_get=lambda url: pages.get(url),
        mx=lambda d: mx or [],
        srv=lambda d: srv,
        guesser=lambda d, e: guess,
    )


def test_discover_ispdb_and_hint():
    d = _discover("max@gmx.de", {autoconfig.ISPDB + "gmx.de": ISPDB_GMX})
    assert d.source == "ispdb" and d.imap.host == "imap.gmx.net"
    assert "POP3/IMAP Abruf" in d.hint


def test_discover_custom_domain_via_mx():
    d = _discover("info@palettenfirma.de", {autoconfig.ISPDB + "kundenserver.de": ISPDB_GMX.replace("gmx.net", "ionos.de")}, mx=["mx00.kundenserver.de"])
    assert d.source == "mx"
    d = _discover("info@palettenfirma.de", mx=["mx00.kundenserver.de"])
    assert d.source == "builtin" and d.imap.host == "imap.ionos.de"
    assert d.imap.username == "info@palettenfirma.de"


def test_discover_microsoft_is_marked_unsupported():
    d = _discover("max@outlook.de")
    assert d.supported is False and "classic Outlook" in d.hint
    d = _discover("info@firma.de", mx=["firma-de.mail.protection.outlook.com"])
    assert d.supported is False


def test_discover_gmail_app_password():
    d = _discover("max@gmail.com")
    assert d.imap.host == "imap.gmail.com"
    assert "app password" in d.hint and d.help_url.startswith("https://myaccount.google.com")


def test_discover_nothing_found():
    d = _discover("x@unbekannt.example")
    assert d.imap is None and d.source == "none"


def test_base_domain():
    assert autoconfig._base_domain("mx01.mail.example.co.uk") == "example.co.uk"
    assert autoconfig._base_domain("ASPMX.L.GOOGLE.COM.") == "google.com"


# ── Thunderbird ──

PREFS = r'''
user_pref("mail.accountmanager.accounts", "account1,account2,account3");
user_pref("mail.account.account1.server", "server1");
user_pref("mail.account.account1.identities", "id1");
user_pref("mail.server.server1.type", "imap");
user_pref("mail.server.server1.hostname", "imap.ionos.de");
user_pref("mail.server.server1.port", 993);
user_pref("mail.server.server1.socketType", 3);
user_pref("mail.server.server1.userName", "info@firma.de");
user_pref("mail.identity.id1.useremail", "info@firma.de");
user_pref("mail.identity.id1.fullName", "Dennis \"D\" Test");
user_pref("mail.identity.id1.smtpServer", "smtp1");
user_pref("mail.smtpserver.smtp1.hostname", "smtp.ionos.de");
user_pref("mail.smtpserver.smtp1.port", 465);
user_pref("mail.smtpserver.smtp1.try_ssl", 3);
user_pref("mail.account.account2.server", "server2");
user_pref("mail.account.account2.identities", "id2");
user_pref("mail.server.server2.type", "pop3");
user_pref("mail.server.server2.hostname", "pop.gmx.net");
user_pref("mail.identity.id2.useremail", "privat@gmx.de");
user_pref("mail.account.account3.server", "server3");
user_pref("mail.server.server3.type", "none");
'''


def test_thunderbird_accounts():
    accs = clients.thunderbird_accounts(clients.parse_prefs(PREFS))
    assert [a.email for a in accs] == ["info@firma.de", "privat@gmx.de"]
    imap = accs[0]
    assert imap.imap.host == "imap.ionos.de" and imap.imap.security == "ssl"
    assert imap.smtp.host == "smtp.ionos.de" and imap.smtp.port == 465
    assert imap.name == 'Dennis "D" Test'
    assert accs[1].kind == "POP3" and accs[1].imap is None


def test_thunderbird_detect_from_profile(tmp_path):
    root = tmp_path / ".thunderbird"
    prof = root / "abc.default-release"
    prof.mkdir(parents=True)
    (prof / "prefs.js").write_text(PREFS, encoding="utf-8")
    (root / "profiles.ini").write_text("[Profile0]\nName=default\nIsRelative=1\nPath=abc.default-release\nDefault=1\n", encoding="utf-8")
    found = clients.detect_thunderbird({"HOME": str(tmp_path)})
    assert found and found[0].accounts[0].email == "info@firma.de"
    assert "POP3" in found[0].note


# -- Outlook through the registry --

class FakeReg:
    def __init__(self, data):
        self.data = data  # {(hive, path): {name: value}}

    def value(self, hive, path, name):
        return self.data.get((hive, path), {}).get(name)

    def subkeys(self, hive, path):
        prefix = path + "\\"
        out = []
        for (h, p) in self.data:
            if h == hive and p.startswith(prefix):
                sub = p[len(prefix):].split("\\")[0]
                if sub not in out:
                    out.append(sub)
        return out

    def values(self, hive, path):
        return self.data.get((hive, path), {})


def _u16(s):
    return (s + "\x00").encode("utf-16-le")


def _outlook_reg(use_new=False):
    base = r"Software\Microsoft\Office\16.0\Outlook"
    acc = base + r"\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676"
    data = {
        ("HKCR", r"Outlook.Application\CurVer"): {"": "Outlook.Application.16"},
        ("HKCU", base): {"DefaultProfile": "Outlook"},
        ("HKCU", acc + r"\00000002"): {"Account Name": _u16("info@firma.de"), "Email": _u16("info@firma.de")},
        ("HKCU", acc + r"\00000003"): {"Account Name": _u16("privat@gmx.de"), "Email": _u16("privat@gmx.de"), "IMAP Server": _u16("imap.gmx.net")},
        ("HKCU", acc + r"\00000004"): {"Account Name": _u16("Outlook Data File")},
        ("HKCU", r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\mailto\UserChoice"): {"ProgId": "Outlook.URL.mailto.15"},
    }
    if use_new:
        data[("HKCU", base + r"\Preferences")] = {"UseNewOutlook": 1}
    return FakeReg(data)


def test_outlook_classic_from_registry():
    det = clients.detect(reg=_outlook_reg(), env={"HOME": "/nonexistent"})
    oc = det.client("outlook_classic")
    assert oc and oc.in_use and oc.backend == "outlook"
    assert [(a.email, a.kind) for a in oc.accounts] == [("info@firma.de", "Exchange"), ("privat@gmx.de", "IMAP")]
    assert det.default_mail_app == "outlook_classic"
    assert det.client("outlook_new") is None


def test_new_outlook_active():
    det = clients.detect(reg=_outlook_reg(use_new=True), env={"HOME": "/nonexistent"})
    assert det.client("outlook_new").in_use
    oc = det.client("outlook_classic")
    assert "new Outlook" in oc.note


# ── KI-Apps ──

def test_register_claude_desktop_keeps_other_servers(tmp_path):
    appdata = tmp_path / "AppData"
    cfg_dir = appdata / "Claude"
    cfg_dir.mkdir(parents=True)
    cfg = cfg_dir / "claude_desktop_config.json"
    cfg.write_text(json.dumps({"mcpServers": {"other": {"command": "x"}}, "theme": "dark"}), encoding="utf-8")
    env = {"APPDATA": str(appdata), "HOME": str(tmp_path), "PATH": ""}
    app = aiapps.detect_claude_desktop(env)
    assert app.installed and not app.registered
    res = aiapps.register(["claude_desktop"], env=env, command=("C:/post/post.exe", ["serve"]))
    assert res[0].ok and res[0].restart_needed
    data = json.loads(cfg.read_text(encoding="utf-8"))
    assert data["mcpServers"]["other"] == {"command": "x"}
    assert data["mcpServers"]["post"] == {"command": "C:/post/post.exe", "args": ["serve"]}
    assert data["theme"] == "dark"
    assert list(cfg_dir.glob("claude_desktop_config.json.post-backup-*"))
    assert aiapps.detect_claude_desktop(env).registered
    aiapps.unregister(["claude_desktop"], env=env)
    assert "post" not in json.loads(cfg.read_text(encoding="utf-8"))["mcpServers"]


def test_register_refuses_broken_json(tmp_path):
    cfg_dir = tmp_path / "AppData" / "Claude"
    cfg_dir.mkdir(parents=True)
    (cfg_dir / "claude_desktop_config.json").write_text("{kaputt", encoding="utf-8")
    env = {"APPDATA": str(tmp_path / "AppData"), "HOME": str(tmp_path), "PATH": ""}
    res = aiapps.register(["claude_desktop"], env=env, command=("x", []))
    assert not res[0].ok
    assert (cfg_dir / "claude_desktop_config.json").read_text(encoding="utf-8") == "{kaputt"


def test_register_codex_idempotent(tmp_path):
    codex = tmp_path / ".codex"
    codex.mkdir()
    (codex / "config.toml").write_text('model = "gpt-5"\n\n[mcp_servers.other]\ncommand = "o"\n', encoding="utf-8")
    env = {"HOME": str(tmp_path), "PATH": ""}
    cmd = ("C:\\Program Files\\post\\post.exe", ["serve"])
    aiapps.register(["codex"], env=env, command=cmd)
    aiapps.register(["codex"], env=env, command=cmd)
    text = (codex / "config.toml").read_text(encoding="utf-8")
    assert text.count("[mcp_servers.post]") == 1
    assert "[mcp_servers.other]" in text
    tomllib = pytest.importorskip("tomllib")

    data = tomllib.loads(text)
    assert data["mcp_servers"]["post"]["command"] == cmd[0]
    assert aiapps.detect_codex(env).registered
    aiapps.unregister(["codex"], env=env)
    data = tomllib.loads((codex / "config.toml").read_text(encoding="utf-8"))
    assert "post" not in data["mcp_servers"] and data["model"] == "gpt-5"


def test_claude_code_fallback_json(tmp_path):
    (tmp_path / ".claude.json").write_text(json.dumps({"projects": {}}), encoding="utf-8")
    env = {"HOME": str(tmp_path), "PATH": ""}
    res = aiapps.register(["claude_code"], env=env, command=("py", ["-m", "ggpost", "serve"]))
    assert res[0].ok
    data = json.loads((tmp_path / ".claude.json").read_text(encoding="utf-8"))
    assert data["mcpServers"]["post"]["args"] == ["-m", "ggpost", "serve"]
    assert data["projects"] == {}


def test_nothing_installed(tmp_path):
    apps = aiapps.detect_apps({"HOME": str(tmp_path), "PATH": ""})
    assert not any(a.installed for a in apps)
    assert Path(tmp_path).exists()


def test_self_install_copies_exe(tmp_path, monkeypatch):
    import sys

    exe = tmp_path / "Downloads" / "post (1).exe"
    exe.parent.mkdir()
    exe.write_bytes(b"MZ fake")
    monkeypatch.setattr(sys, "frozen", True, raising=False)
    monkeypatch.setattr(sys, "platform", "win32")
    monkeypatch.setattr(sys, "executable", str(exe))
    env = {"LOCALAPPDATA": str(tmp_path / "Local"), "HOME": str(tmp_path)}
    path, msg = aiapps.ensure_installed(env)
    assert path == str(tmp_path / "Local" / "Programs" / "post" / "post.exe")
    assert Path(path).read_bytes() == b"MZ fake" and "delete" in msg
    assert aiapps.server_command(env) == (path, ["serve"])
    monkeypatch.setattr(sys, "executable", path)
    assert aiapps.ensure_installed(env) == (path, "")


def test_codex_removes_subtables(tmp_path):
    codex = tmp_path / ".codex"
    codex.mkdir()
    (codex / "config.toml").write_text(
        '[mcp_servers.post]\ncommand = "alt"\n\n[mcp_servers.post.env]\nX = "1"\n\n[mcp_servers.other]\ncommand = "o"\n',
        encoding="utf-8",
    )
    env = {"HOME": str(tmp_path), "PATH": ""}
    aiapps.register(["codex"], env=env, command=("neu", ["serve"]))
    tomllib = pytest.importorskip("tomllib")

    data = tomllib.loads((codex / "config.toml").read_text(encoding="utf-8"))
    assert data["mcp_servers"]["post"] == {"command": "neu", "args": ["serve"]}
    aiapps.unregister(["codex"], env=env)
    data = tomllib.loads((codex / "config.toml").read_text(encoding="utf-8"))
    assert "post" not in data["mcp_servers"] and data["mcp_servers"]["other"]["command"] == "o"


def test_corrupt_config_is_set_aside():
    from ggpost.config import config_dir, config_path, load_config

    config_dir().mkdir(parents=True, exist_ok=True)
    config_path().write_text("{kaputt", encoding="utf-8")
    assert load_config() is None
    assert not config_path().exists()
    assert list(config_dir().glob("config.json.broken-*"))
