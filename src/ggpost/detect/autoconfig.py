"""Finds the server settings for a mail address, the way Thunderbird does.

Reihenfolge:
  1. The provider autoconfig (autoconfig.<domain> or /.well-known/)
  2. Mozilla ISPDB (the database behind Thunderbird)
  3. The MX record of the domain (own domain at IONOS, Strato, Google ...)
  4. DNS SRV records (RFC 6186)
  5. A built-in list of known providers (works offline too)
  6. Guessing: imap.<domain>, mail.<domain> on port 993
For known providers there is also a hint about the password.
"""

from __future__ import annotations

import logging
import socket
import ssl
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor
from dataclasses import asdict, dataclass
from typing import Any, Callable

from ggpost.config import ServerSettings

log = logging.getLogger(__name__)

HTTP_TIMEOUT = 6.0
MAX_XML = 256 * 1024
ISPDB = "https://autoconfig.thunderbird.net/v1.1/"


@dataclass
class Discovery:
    email: str
    provider: str = ""
    imap: ServerSettings | None = None
    smtp: ServerSettings | None = None
    source: str = ""
    hint: str = ""
    help_url: str = ""
    supported: bool = True

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


# ── Bekannte Anbieter ────────────────────────────────────────────────────

_GMAIL = ("Gmail", ("imap.gmail.com", 993, "ssl"), ("smtp.gmail.com", 465, "ssl"))
_GMX = ("GMX", ("imap.gmx.net", 993, "ssl"), ("mail.gmx.net", 587, "starttls"))
_WEBDE = ("WEB.DE", ("imap.web.de", 993, "ssl"), ("smtp.web.de", 587, "starttls"))
_TONLINE = ("Telekom", ("secureimap.t-online.de", 993, "ssl"), ("securesmtp.t-online.de", 465, "ssl"))
_ICLOUD = ("iCloud", ("imap.mail.me.com", 993, "ssl"), ("smtp.mail.me.com", 587, "starttls"))
_YAHOO = ("Yahoo", ("imap.mail.yahoo.com", 993, "ssl"), ("smtp.mail.yahoo.com", 465, "ssl"))
_POSTEO = ("Posteo", ("posteo.de", 993, "ssl"), ("posteo.de", 465, "ssl"))
_MAILBOX = ("mailbox.org", ("imap.mailbox.org", 993, "ssl"), ("smtp.mailbox.org", 465, "ssl"))
_FREENET = ("freenet", ("mx.freenet.de", 993, "ssl"), ("mx.freenet.de", 587, "starttls"))
_IONOS = ("IONOS", ("imap.ionos.de", 993, "ssl"), ("smtp.ionos.de", 465, "ssl"))
_STRATO = ("STRATO", ("imap.strato.de", 993, "ssl"), ("smtp.strato.de", 465, "ssl"))

BUILTIN_DOMAINS = {
    "gmail.com": _GMAIL, "googlemail.com": _GMAIL,
    "gmx.de": _GMX, "gmx.net": _GMX, "gmx.at": _GMX, "gmx.ch": _GMX, "gmx.com": _GMX,
    "web.de": _WEBDE,
    "t-online.de": _TONLINE, "magenta.de": _TONLINE,
    "icloud.com": _ICLOUD, "me.com": _ICLOUD, "mac.com": _ICLOUD,
    "yahoo.com": _YAHOO, "yahoo.de": _YAHOO,
    "posteo.de": _POSTEO, "posteo.net": _POSTEO,
    "mailbox.org": _MAILBOX,
    "freenet.de": _FREENET,
    "ionos.de": _IONOS, "online.de": _IONOS,
    "strato.de": _STRATO,
}
BUILTIN_MX = {
    "kundenserver.de": _IONOS, "ionos.de": _IONOS, "ionos.com": _IONOS,
    "rzone.de": _STRATO, "strato.de": _STRATO,
    "google.com": _GMAIL, "googlemail.com": _GMAIL,
    "gmx.net": _GMX, "web.de": _WEBDE, "t-online.de": _TONLINE,
    "mailbox.org": _MAILBOX, "posteo.de": _POSTEO,
}

_MS_DOMAINS = {"outlook.com", "outlook.de", "hotmail.com", "hotmail.de", "live.com", "live.de", "msn.com"}

_HINTS = [
    (lambda d, mx: d in ("gmail.com", "googlemail.com") or mx in ("google.com", "googlemail.com"),
     "Google requires an app password (Google Account > Security > 2-Step Verification > App passwords). "
     "Your normal password will not work here.",
     "https://myaccount.google.com/apppasswords"),
    (lambda d, mx: d.startswith("gmx.") or mx == "gmx.net",
     "GMX needs IMAP access switched on once: in the webmail under "
     "E-Mail > Einstellungen > POP3/IMAP Abruf.", ""),
    (lambda d, mx: d == "web.de" or mx == "web.de",
     "WEB.DE needs IMAP access switched on once: in the webmail under "
     "E-Mail > Einstellungen > POP3/IMAP Abruf.", ""),
    (lambda d, mx: d in ("t-online.de", "magenta.de") or mx == "t-online.de",
     "Telekom requires a separate e-mail password (customer centre > e-mail settings), "
     "not the password of your Telekom login.", ""),
    (lambda d, mx: d in ("icloud.com", "me.com", "mac.com"),
     "Apple requires an app-specific password (account.apple.com > Sign-In and Security).", ""),
    (lambda d, mx: d.startswith("yahoo.") or d == "aol.com",
     "Yahoo and AOL require an app password from the account settings.", ""),
]


def _base_domain(host: str) -> str:
    labels = host.rstrip(".").lower().split(".")
    if len(labels) >= 3 and len(labels[-1]) == 2 and labels[-2] in {"co", "com", "net", "org", "ac", "gv", "or"}:
        return ".".join(labels[-3:])
    return ".".join(labels[-2:])


def _from_builtin(email: str, entry, source: str) -> Discovery:
    name, imap, smtp = entry
    return Discovery(
        email=email,
        provider=name,
        imap=ServerSettings(imap[0], imap[1], imap[2], email),
        smtp=ServerSettings(smtp[0], smtp[1], smtp[2], email),
        source=source,
    )


# ── XML (Autoconfig / ISPDB) ─────────────────────────────────────────────


def _subst(value: str, email: str) -> str:
    local, _, domain = email.partition("@")
    return (
        value.replace("%EMAILADDRESS%", email)
        .replace("%EMAILLOCALPART%", local)
        .replace("%EMAILDOMAIN%", domain)
    )


def parse_autoconfig(xml_text: str, email: str) -> Discovery | None:
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return None
    provider = root.find("emailProvider")
    if provider is None:
        return None

    def pick(tag: str, type_: str) -> ServerSettings | None:
        ranked: list[tuple[int, ServerSettings]] = []
        for srv in provider.findall(tag):
            if srv.get("type") != type_:
                continue
            auths = [(a.text or "").strip().lower() for a in srv.findall("authentication")]
            if auths and not any(a.startswith("password") for a in auths):
                continue  # OAuth2 only, or similar
            host = (srv.findtext("hostname") or "").strip()
            port = (srv.findtext("port") or "").strip()
            sock = (srv.findtext("socketType") or "").strip().upper()
            if not host or not port.isdigit():
                continue
            sec = {"SSL": "ssl", "STARTTLS": "starttls"}.get(sock, "plain")
            user = _subst((srv.findtext("username") or "%EMAILADDRESS%").strip(), email)
            rank = {"ssl": 0, "starttls": 1}.get(sec, 5)
            ranked.append((rank, ServerSettings(_subst(host, email), int(port), sec, user)))
        ranked.sort(key=lambda r: r[0])
        return ranked[0][1] if ranked else None

    imap = pick("incomingServer", "imap")
    if imap is None:
        return None
    return Discovery(
        email=email,
        provider=(provider.findtext("displayName") or provider.get("id") or "").strip(),
        imap=imap,
        smtp=pick("outgoingServer", "smtp"),
    )


def _http_get(url: str, timeout: float = HTTP_TIMEOUT) -> str | None:
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "post-mail-connector/0.1"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:  # noqa: S310 (https URLs only)
            if resp.status != 200:
                return None
            return resp.read(MAX_XML).decode("utf-8", errors="replace")
    except Exception as exc:
        log.debug("autoconfig %s: %s", url, exc)
        return None


# ── DNS ──────────────────────────────────────────────────────────────────


def mx_hosts(domain: str) -> list[str]:
    try:
        import dns.resolver  # noqa: PLC0415

        answers = dns.resolver.resolve(domain, "MX", lifetime=HTTP_TIMEOUT)
        ranked = sorted(answers, key=lambda r: r.preference)
        return [str(r.exchange).rstrip(".").lower() for r in ranked]
    except Exception as exc:
        log.debug("MX %s: %s", domain, exc)
        return []


def srv_lookup(domain: str) -> Discovery | None:
    try:
        import dns.resolver  # noqa: PLC0415
    except Exception:
        return None

    def one(name: str) -> tuple[str, int] | None:
        try:
            ans = sorted(dns.resolver.resolve(name, "SRV", lifetime=HTTP_TIMEOUT), key=lambda r: (r.priority, -r.weight))
        except Exception:
            return None
        for r in ans:
            target = str(r.target).rstrip(".")
            if target and target != ".":
                return target, int(r.port)
        return None

    imap = smtp = None
    hit = one(f"_imaps._tcp.{domain}")
    if hit:
        imap = ServerSettings(hit[0], hit[1], "ssl")
    else:
        hit = one(f"_imap._tcp.{domain}")
        if hit:
            imap = ServerSettings(hit[0], hit[1], "starttls")
    hit = one(f"_submissions._tcp.{domain}")
    if hit:
        smtp = ServerSettings(hit[0], hit[1], "ssl")
    else:
        hit = one(f"_submission._tcp.{domain}")
        if hit:
            smtp = ServerSettings(hit[0], hit[1], "starttls")
    if not imap:
        return None
    return Discovery(email="", imap=imap, smtp=smtp, source="srv")


def _tls_reachable(host: str, port: int, timeout: float = 4.0) -> bool:
    try:
        ctx = ssl.create_default_context()
        with socket.create_connection((host, port), timeout=timeout) as sock:
            with ctx.wrap_socket(sock, server_hostname=host):
                return True
    except Exception:
        return False


def _tcp_reachable(host: str, port: int, timeout: float = 4.0) -> bool:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True
    except Exception:
        return False


def guess(domain: str, email: str) -> Discovery | None:
    for host in (f"imap.{domain}", f"mail.{domain}", domain):
        if _tls_reachable(host, 993):
            smtp = None
            for shost in (f"smtp.{domain}", f"mail.{domain}", host):
                if _tls_reachable(shost, 465):
                    smtp = ServerSettings(shost, 465, "ssl", email)
                    break
                if _tcp_reachable(shost, 587):
                    smtp = ServerSettings(shost, 587, "starttls", email)
                    break
            return Discovery(email=email, provider=domain, imap=ServerSettings(host, 993, "ssl", email), smtp=smtp, source="guess")
    return None


# ── Hauptfunktion ────────────────────────────────────────────────────────


def provider_hint(domain: str, mx_base: str) -> tuple[str, str, bool]:
    if domain in _MS_DOMAINS or mx_base in ("outlook.com", "hotmail.com") or mx_base.endswith("outlook.com"):
        return (
            "Microsoft no longer allows IMAP logins with a password for Outlook.com and Microsoft 365. "
            "Set the account up in classic Outlook or in Thunderbird; post. detects it there automatically.",
            "",
            False,
        )
    for test, text, url in _HINTS:
        if test(domain, mx_base):
            return text, url, True
    return "", "", True


def discover(
    email: str,
    http_get: Callable[[str], str | None] = _http_get,
    mx: Callable[[str], list[str]] = mx_hosts,
    srv: Callable[[str], Discovery | None] = srv_lookup,
    guesser: Callable[[str, str], Discovery | None] = guess,
) -> Discovery:
    email = email.strip()
    if "@" not in email:
        raise ValueError("Please enter a complete mail address.")
    domain = email.rsplit("@", 1)[1].lower()
    q = urllib.parse.quote(email)
    urls = [
        ("autoconfig", f"https://autoconfig.{domain}/mail/config-v1.1.xml?emailaddress={q}"),
        ("autoconfig", f"https://{domain}/.well-known/autoconfig/mail/config-v1.1.xml?emailaddress={q}"),
        ("ispdb", ISPDB + domain),
    ]
    with ThreadPoolExecutor(max_workers=4) as pool:
        mx_future = pool.submit(mx, domain)
        futures = [(src, pool.submit(http_get, url)) for src, url in urls]
        result: Discovery | None = None
        for src, fut in futures:
            text = fut.result()
            if text:
                found = parse_autoconfig(text, email)
                if found:
                    found.source = src
                    result = found
                    break
        mx_list = mx_future.result()

    mx_base = _base_domain(mx_list[0]) if mx_list else ""

    if result is None and mx_base and mx_base != domain:
        text = http_get(ISPDB + mx_base)
        if text:
            found = parse_autoconfig(text, email)
            if found:
                found.source = "mx"
                result = found
    if result is None:
        found = srv(domain)
        if found:
            found.email = email
            for s in (found.imap, found.smtp):
                if s and not s.username:
                    s.username = email
            result = found
    if result is None and domain in BUILTIN_DOMAINS:
        result = _from_builtin(email, BUILTIN_DOMAINS[domain], "builtin")
    if result is None and mx_base in BUILTIN_MX:
        result = _from_builtin(email, BUILTIN_MX[mx_base], "builtin")
    if result is None:
        result = guesser(domain, email)
    if result is None:
        result = Discovery(email=email, source="none")

    hint, url, supported = provider_hint(domain, mx_base)
    result.hint, result.help_url, result.supported = hint, url, supported
    if not result.provider:
        result.provider = domain
    return result
