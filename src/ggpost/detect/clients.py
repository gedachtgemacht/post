"""Detects mail programs and their accounts without changing anything.

Read-only access: the registry (Windows) and profile files. Outlook is not
started, and no passwords are read.
"""

from __future__ import annotations

import configparser
import json
import os
import re
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Protocol

from ggpost.config import ServerSettings

HKCU = "HKCU"
HKLM = "HKLM"
HKCR = "HKCR"

OUTLOOK_PROFILE_ACCOUNTS = "9375CFF0413111d3B88A00104B2A6676"
OFFICE_VERSIONS = ("16.0", "15.0")


class Registry(Protocol):
    def value(self, hive: str, path: str, name: str) -> Any: ...

    def subkeys(self, hive: str, path: str) -> list[str]: ...

    def values(self, hive: str, path: str) -> dict[str, Any]: ...


class WinRegistry:
    """A thin, read-only wrapper around winreg."""

    def __init__(self) -> None:
        import winreg  # noqa: PLC0415

        self._w = winreg
        self._hives = {HKCU: winreg.HKEY_CURRENT_USER, HKLM: winreg.HKEY_LOCAL_MACHINE, HKCR: winreg.HKEY_CLASSES_ROOT}

    def _open(self, hive: str, path: str):
        return self._w.OpenKey(self._hives[hive], path, 0, self._w.KEY_READ)

    def value(self, hive: str, path: str, name: str) -> Any:
        try:
            with self._open(hive, path) as key:
                return self._w.QueryValueEx(key, name)[0]
        except OSError:
            return None

    def subkeys(self, hive: str, path: str) -> list[str]:
        out: list[str] = []
        try:
            with self._open(hive, path) as key:
                i = 0
                while True:
                    try:
                        out.append(self._w.EnumKey(key, i))
                    except OSError:
                        break
                    i += 1
        except OSError:
            pass
        return out

    def values(self, hive: str, path: str) -> dict[str, Any]:
        out: dict[str, Any] = {}
        try:
            with self._open(hive, path) as key:
                i = 0
                while True:
                    try:
                        name, data, _t = self._w.EnumValue(key, i)
                    except OSError:
                        break
                    out[name] = data
                    i += 1
        except OSError:
            pass
        return out


class NoRegistry:
    def value(self, hive, path, name):
        return None

    def subkeys(self, hive, path):
        return []

    def values(self, hive, path):
        return {}


@dataclass
class DetectedAccount:
    email: str
    name: str = ""
    kind: str = ""  # IMAP, POP3, Exchange …
    via: str = ""  # outlook_classic, thunderbird
    imap: ServerSettings | None = None
    smtp: ServerSettings | None = None


@dataclass
class DetectedClient:
    id: str
    name: str
    installed: bool = True
    in_use: bool = False
    backend: str = "imap"  # outlook | imap
    accounts: list[DetectedAccount] = field(default_factory=list)
    note: str = ""


@dataclass
class Detection:
    platform: str
    clients: list[DetectedClient] = field(default_factory=list)
    default_mail_app: str = ""

    def to_dict(self) -> dict[str, Any]:
        return json.loads(json.dumps(asdict(self), default=str))

    def client(self, cid: str) -> DetectedClient | None:
        return next((c for c in self.clients if c.id == cid), None)


# ── Outlook ──────────────────────────────────────────────────────────────


def _reg_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        try:
            text = value.decode("utf-16-le")
        except UnicodeDecodeError:
            text = value.decode("latin-1", errors="ignore")
        return text.split("\x00", 1)[0].strip()
    return str(value).strip()


def outlook_profile_count(reg: Registry) -> int:
    for ver in OFFICE_VERSIONS:
        profiles = reg.subkeys(HKCU, rf"Software\Microsoft\Office\{ver}\Outlook\Profiles")
        if profiles:
            return len(profiles)
    return 0


def outlook_profile_accounts(reg: Registry) -> tuple[str, list[DetectedAccount]]:
    """Reads the default profile accounts from the registry (Outlook 2013 to 2024)."""
    for ver in OFFICE_VERSIONS:
        base = rf"Software\Microsoft\Office\{ver}\Outlook"
        profiles = reg.subkeys(HKCU, base + r"\Profiles")
        if not profiles:
            continue
        default = _reg_text(reg.value(HKCU, base, "DefaultProfile")) or profiles[0]
        if default not in profiles:
            default = profiles[0]
        acc_path = rf"{base}\Profiles\{default}\{OUTLOOK_PROFILE_ACCOUNTS}"
        accounts: list[DetectedAccount] = []
        seen: set[str] = set()
        for sub in reg.subkeys(HKCU, acc_path):
            vals = reg.values(HKCU, rf"{acc_path}\{sub}")
            email = _reg_text(vals.get("Email")) or _reg_text(vals.get("Account Name"))
            if "@" not in email or email.lower() in seen:
                continue
            seen.add(email.lower())
            if vals.get("IMAP Server"):
                kind = "IMAP"
            elif vals.get("POP3 Server"):
                kind = "POP3"
            else:
                kind = "Exchange"
            accounts.append(
                DetectedAccount(
                    email=email,
                    name=_reg_text(vals.get("Display Name")),
                    kind=kind,
                    via="outlook_classic",
                )
            )
        return default, accounts
    return "", []


def detect_outlook(reg: Registry, env: dict[str, str] | None = None) -> list[DetectedClient]:
    env = env if env is not None else dict(os.environ)
    out: list[DetectedClient] = []
    curver = _reg_text(reg.value(HKCR, r"Outlook.Application\CurVer", ""))
    app_path = reg.value(HKLM, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE", "")
    classic_installed = bool(curver or app_path)
    new_active = any(
        reg.value(HKCU, rf"Software\Microsoft\Office\{v}\Outlook\Preferences", "UseNewOutlook") == 1
        for v in OFFICE_VERSIONS
    )
    local = env.get("LOCALAPPDATA", "")
    new_installed = new_active or bool(local and Path(local, "Microsoft", "Olk").exists())

    if classic_installed:
        profile, accounts = outlook_profile_accounts(reg)
        note = ""
        if new_active:
            note = (
                "You are currently using the new Outlook. post. can still use classic Outlook "
                "in the background, as long as your accounts are set up there."
            )
        elif not accounts:
            note = "Outlook is installed, but no account was found in its profile."
        out.append(
            DetectedClient(
                id="outlook_classic",
                name="Outlook (classic)" + (f", profile {profile}" if profile and outlook_profile_count(reg) > 1 else ""),
                in_use=not new_active,
                backend="outlook",
                accounts=accounts,
                note=note,
            )
        )
    if new_installed:
        out.append(
            DetectedClient(
                id="outlook_new",
                name="Outlook (new)",
                in_use=new_active,
                backend="imap",
                note=(
                    "The new Outlook cannot be driven from outside, so post. connects to your "
                    "mailbox directly over IMAP. Enter your address below to do that."
                ),
            )
        )
    return out


# ── Thunderbird ──────────────────────────────────────────────────────────


def thunderbird_roots(env: dict[str, str] | None = None) -> list[Path]:
    env = env if env is not None else dict(os.environ)
    home = Path(env.get("HOME") or env.get("USERPROFILE") or Path.home())
    roots: list[Path] = []
    if sys.platform == "win32" or env.get("APPDATA"):
        appdata = env.get("APPDATA")
        if appdata:
            roots.append(Path(appdata) / "Thunderbird")
    roots += [
        home / "Library" / "Thunderbird",
        home / ".thunderbird",
        home / "snap" / "thunderbird" / "common" / ".thunderbird",
        home / ".var" / "app" / "org.mozilla.Thunderbird" / ".thunderbird",
    ]
    return [r for r in roots if (r / "profiles.ini").exists()]


def thunderbird_default_profile(root: Path) -> Path | None:
    ini = configparser.ConfigParser(interpolation=None)
    try:
        ini.read(root / "profiles.ini", encoding="utf-8")
    except Exception:
        return None
    install_default = None
    candidates: list[tuple[bool, Path]] = []
    for section in ini.sections():
        if section.startswith("Install") and ini.has_option(section, "Default"):
            install_default = ini.get(section, "Default")
        if section.startswith("Profile") and ini.has_option(section, "Path"):
            rel = ini.get(section, "IsRelative", fallback="1") == "1"
            raw = ini.get(section, "Path")
            path = (root / raw) if rel else Path(raw)
            candidates.append((ini.get(section, "Default", fallback="0") == "1", path))
    if install_default:
        p = root / install_default
        if (p / "prefs.js").exists():
            return p
        if (Path(install_default) / "prefs.js").exists():
            return Path(install_default)
    for is_default, path in sorted(candidates, key=lambda c: not c[0]):
        if (path / "prefs.js").exists():
            return path
    return None


_PREF_RE = re.compile(r'^user_pref\("([^"]+)",\s*(.+)\);\s*$', re.M)


def parse_prefs(text: str) -> dict[str, Any]:
    prefs: dict[str, Any] = {}
    for key, raw in _PREF_RE.findall(text):
        raw = raw.strip()
        if raw.startswith('"'):
            try:
                prefs[key] = json.loads(raw)
            except json.JSONDecodeError:
                prefs[key] = raw.strip('"')
        elif raw in ("true", "false"):
            prefs[key] = raw == "true"
        else:
            try:
                prefs[key] = int(raw)
            except ValueError:
                prefs[key] = raw
    return prefs


def _tb_security(socket_type: Any) -> str:
    return {3: "ssl", 2: "starttls", 1: "starttls", 0: "plain"}.get(int(socket_type or 0), "plain")


def thunderbird_accounts(prefs: dict[str, Any]) -> list[DetectedAccount]:
    out: list[DetectedAccount] = []
    ids = [a.strip() for a in str(prefs.get("mail.accountmanager.accounts", "")).split(",") if a.strip()]
    for acc in ids:
        server = prefs.get(f"mail.account.{acc}.server")
        identities = [i.strip() for i in str(prefs.get(f"mail.account.{acc}.identities", "")).split(",") if i.strip()]
        if not server or not identities:
            continue
        stype = prefs.get(f"mail.server.{server}.type", "")
        if stype not in ("imap", "pop3"):
            continue
        ident = identities[0]
        email = prefs.get(f"mail.identity.{ident}.useremail", "")
        if not email:
            continue
        host = prefs.get(f"mail.server.{server}.hostname") or prefs.get(f"mail.server.{server}.realhostname")
        sec = _tb_security(prefs.get(f"mail.server.{server}.socketType", 0))
        imap = None
        if stype == "imap" and host:
            port = int(prefs.get(f"mail.server.{server}.port") or (993 if sec == "ssl" else 143))
            imap = ServerSettings(
                host=str(host),
                port=port,
                security=sec,
                username=str(prefs.get(f"mail.server.{server}.userName", "") or email),
            )
        smtp = None
        smtp_id = prefs.get(f"mail.identity.{ident}.smtpServer") or str(prefs.get("mail.smtp.defaultserver", ""))
        if smtp_id:
            shost = prefs.get(f"mail.smtpserver.{smtp_id}.hostname")
            if shost:
                ssec = _tb_security(prefs.get(f"mail.smtpserver.{smtp_id}.try_ssl", 0))
                sport = int(prefs.get(f"mail.smtpserver.{smtp_id}.port") or (465 if ssec == "ssl" else 587))
                smtp = ServerSettings(
                    host=str(shost),
                    port=sport,
                    security=ssec,
                    username=str(prefs.get(f"mail.smtpserver.{smtp_id}.username", "") or email),
                )
        out.append(
            DetectedAccount(
                email=str(email),
                name=str(prefs.get(f"mail.identity.{ident}.fullName", "") or ""),
                kind=stype.upper(),
                via="thunderbird",
                imap=imap,
                smtp=smtp,
            )
        )
    return out


def detect_thunderbird(env: dict[str, str] | None = None) -> list[DetectedClient]:
    out: list[DetectedClient] = []
    for root in thunderbird_roots(env):
        profile = thunderbird_default_profile(root)
        if not profile:
            continue
        try:
            prefs = parse_prefs((profile / "prefs.js").read_text(encoding="utf-8", errors="replace"))
        except OSError:
            continue
        accounts = thunderbird_accounts(prefs)
        pop = [a for a in accounts if a.kind == "POP3"]
        note = ""
        if pop:
            note = "POP3 accounts are connected over IMAP; the server settings for that are looked up automatically."
        out.append(DetectedClient(id="thunderbird", name="Thunderbird", backend="imap", accounts=accounts, note=note))
        break
    return out


# ── Sonstiges ────────────────────────────────────────────────────────────


def default_mail_app(reg: Registry) -> str:
    prog = _reg_text(
        reg.value(HKCU, r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\mailto\UserChoice", "ProgId")
    )
    low = prog.lower()
    if not low:
        return ""
    if "outlook" in low and "olk" not in low:
        return "outlook_classic"
    if "thunderbird" in low:
        return "thunderbird"
    if low.startswith("appx"):
        return "outlook_new"
    return prog


def detect(reg: Registry | None = None, env: dict[str, str] | None = None) -> Detection:
    if reg is None:
        try:
            reg = WinRegistry() if sys.platform == "win32" else NoRegistry()
        except Exception:
            reg = NoRegistry()
    env = env if env is not None else dict(os.environ)
    det = Detection(platform=sys.platform)
    det.clients += detect_outlook(reg, env)
    det.clients += detect_thunderbird(env)
    home = Path(env.get("HOME") or Path.home())
    if sys.platform == "darwin" and (home / "Library" / "Mail").exists():
        det.clients.append(
            DetectedClient(
                id="apple_mail",
                name="Apple Mail",
                backend="imap",
                note="Apple Mail does not expose its accounts. Please enter the address below.",
            )
        )
    det.default_mail_app = default_mail_app(reg)
    if det.default_mail_app:
        for c in det.clients:
            if c.id == det.default_mail_app:
                c.in_use = True
    return det
