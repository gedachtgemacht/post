"""The local setup assistant in the browser.

Runs on 127.0.0.1 only, with a random port and a one-time token. No data
leaves the machine apart from the requests to mail servers and the autoconfig
services that setup needs.
"""

from __future__ import annotations

import json
import logging
from dataclasses import asdict
import secrets
import sys
import threading
import os
import subprocess
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from importlib import resources
from typing import Any, Callable
from urllib.parse import parse_qs, urlparse

from ggpost import __version__, vault
from ggpost.backends.base import BackendError
from ggpost.config import Config, Permissions, ServerSettings, Source, config_path, load_config, save_config
from ggpost.detect import aiapps, autoconfig, clients

log = logging.getLogger(__name__)

STATIC_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".woff2": "font/woff2",
    ".svg": "image/svg+xml",
    ".png": "image/png",
}
MAX_BODY = 64 * 1024


def _static(name: str) -> bytes | None:
    if "/" in name.strip("/").replace("fonts/", "", 1) or ".." in name:
        return None
    try:
        return resources.files("ggpost.setup").joinpath("static", *name.strip("/").split("/")).read_bytes()
    except (FileNotFoundError, IsADirectoryError, OSError):
        return None


def open_browser_quietly(url: str) -> bool:
    """Opens the browser without any helper writing to stdout.
    In server mode stdout is the MCP channel and must see nothing else."""
    try:
        if sys.platform == "win32":
            os.startfile(url)  # type: ignore[attr-defined]  # noqa: S606
            return True
        cmd = ["open", url] if sys.platform == "darwin" else ["xdg-open", url]
        subprocess.Popen(cmd, stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)  # noqa: S603
        return True
    except Exception:
        pass
    try:  # letzter Versuch, stdout dabei umleiten
        import contextlib  # noqa: PLC0415

        with open(os.devnull, "w") as null, contextlib.redirect_stdout(null):
            return bool(webbrowser.open(url))
    except Exception:
        return False


def _settings(d: dict[str, Any] | None) -> ServerSettings | None:
    if not d or not d.get("host"):
        return None
    security = str(d.get("security", "ssl"))
    if security not in ("ssl", "starttls", "plain"):
        raise ValueError("Invalid encryption setting.")
    port = int(d.get("port") or (993 if security == "ssl" else 143))
    if not 0 < port < 65536:
        raise ValueError("Invalid port.")
    return ServerSettings(host=str(d["host"]).strip(), port=port, security=security, username=str(d.get("username", "")).strip())


class SetupApp:
    def __init__(
        self,
        on_saved: Callable[[Config], None] | None = None,
        detect_fn: Callable[[], clients.Detection] = clients.detect,
        apps_fn: Callable[[], list[aiapps.AiApp]] = aiapps.detect_apps,
        discover_fn: Callable[[str], autoconfig.Discovery] = autoconfig.discover,
        test_imap_fn: Callable[[ServerSettings, str, str], str] | None = None,
        test_outlook_fn: Callable[[], dict[str, Any]] | None = None,
        register_fn: Callable[[list[str]], list[aiapps.RegisterResult]] = aiapps.register,
        mode: str = "standalone",  # standalone | embedded (started from the MCP server)
    ) -> None:
        self.token = secrets.token_urlsafe(24)
        self.on_saved = on_saved
        self.detect_fn = detect_fn
        self.apps_fn = apps_fn
        self.discover_fn = discover_fn
        self.test_imap_fn = test_imap_fn or self._default_test_imap
        self.test_outlook_fn = test_outlook_fn or self._default_test_outlook
        self.register_fn = register_fn
        self.mode = mode
        self.finished = threading.Event()
        app = self

        class Handler(_Handler):
            ctx = app

        self.httpd = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self.httpd.daemon_threads = True
        self.port = self.httpd.server_address[1]
        self.url = f"http://127.0.0.1:{self.port}/?t={self.token}"
        self._thread: threading.Thread | None = None

    # ── Lebenszyklus ──

    def start(self, open_browser: bool = True) -> str:
        self._thread = threading.Thread(target=self.httpd.serve_forever, name="post-setup", daemon=True)
        self._thread.start()
        if open_browser:
            open_browser_quietly(self.url)
        return self.url

    def stop(self) -> None:
        self.finished.set()
        threading.Thread(target=self.httpd.shutdown, daemon=True).start()

    def wait(self, timeout: float | None = None) -> bool:
        return self.finished.wait(timeout)

    # ── Standardaktionen ──

    @staticmethod
    def _default_test_imap(settings: ServerSettings, email: str, password: str) -> str:
        from ggpost.backends.imap import verify_login  # noqa: PLC0415

        return verify_login(settings, email, password)

    @staticmethod
    def _default_test_outlook() -> dict[str, Any]:
        if sys.platform != "win32":
            raise BackendError("Classic Outlook only exists on Windows.")
        from ggpost.backends.outlook_com import OutlookBackend  # noqa: PLC0415

        backend = OutlookBackend(Source(kind="outlook"))
        accounts = backend.accounts()
        message = backend.check()
        return {"message": message, "accounts": [a.to_dict() for a in accounts]}

    # ── API ──

    def api_state(self) -> dict[str, Any]:
        det = self.detect_fn()
        apps = self.apps_fn()
        cfg = load_config()
        sources = []
        if cfg:
            for s in cfg.sources:
                d = s.to_dict()
                if s.kind == "imap":
                    d["has_password"] = bool(vault.get_password(s.email))
                sources.append(d)
        detection = det.to_dict()
        for client in detection["clients"]:
            for acc in client["accounts"]:
                domain = acc["email"].rsplit("@", 1)[-1].lower()
                hint, url, supported = autoconfig.provider_hint(domain, "")
                if client["id"] == "outlook_classic":
                    continue  # Outlook needs no password, a hint would only mislead
                acc.update(hint=hint, help_url=url, supported=supported)
        return {
            "version": __version__,
            "mode": self.mode,
            "platform": sys.platform,
            "detection": detection,
            "apps": [a.to_dict() for a in apps],
            "config": {
                "exists": cfg is not None,
                "path": str(config_path()),
                "sources": sources,
                "permissions": asdict(cfg.permissions if cfg else Permissions()),
            },
        }

    def api_discover(self, body: dict[str, Any]) -> dict[str, Any]:
        email = str(body.get("email", "")).strip()
        if "@" not in email or "." not in email.split("@")[-1]:
            raise ValueError("Please enter a complete mail address, for example name@company.com.")
        found = self.discover_fn(email)
        data = found.to_dict()
        data["has_password"] = bool(vault.get_password(email))
        return data

    def api_test(self, body: dict[str, Any]) -> dict[str, Any]:
        email = str(body.get("email", "")).strip()
        settings = _settings(body.get("imap"))
        if not settings:
            raise ValueError("The server settings are missing.")
        password = str(body.get("password") or "")
        if not password:
            # Use a stored password only for exactly the stored server, otherwise it
            # could be sent to a server somebody else named.
            cfg = load_config()
            saved = next((s for s in (cfg.sources if cfg else []) if s.kind == "imap" and s.email.lower() == email.lower()), None)
            if saved and saved.imap and (saved.imap.host, saved.imap.port, saved.imap.username) == (
                settings.host,
                settings.port,
                settings.username,
            ):
                password = vault.get_password(email) or ""
        if not password:
            raise ValueError("Please enter the password.")
        return {"ok": True, "message": self.test_imap_fn(settings, email, password)}

    def api_test_outlook(self, _body: dict[str, Any]) -> dict[str, Any]:
        result = self.test_outlook_fn()
        result["ok"] = True
        return result

    def api_save(self, body: dict[str, Any]) -> dict[str, Any]:
        sources: list[Source] = []
        for raw in body.get("sources", []):
            kind = raw.get("kind")
            if kind == "outlook":
                accounts = raw.get("outlook_accounts")
                sources.append(Source(kind="outlook", label="Outlook", outlook_accounts=[str(a) for a in accounts] if accounts else None))
            elif kind == "imap":
                email = str(raw.get("email", "")).strip()
                if "@" not in email:
                    raise ValueError("Invalid mail address.")
                imap = _settings(raw.get("imap"))
                if not imap:
                    raise ValueError(f"The IMAP settings for {email} are missing.")
                sources.append(
                    Source(
                        kind="imap",
                        email=email,
                        label=str(raw.get("label") or "").strip(),
                        imap=imap,
                        smtp=_settings(raw.get("smtp")),
                        signature=str(raw.get("signature") or "").strip()[:2000],
                    )
                )
        if not sources:
            raise ValueError("Please select at least one mailbox.")
        perms = body.get("permissions") or {}
        # Passwords first, so a keychain failure cannot leave things half saved
        for email, pw in (body.get("passwords") or {}).items():
            if pw:
                vault.set_password(str(email), str(pw))
        for s in sources:
            if s.kind == "imap" and not vault.get_password(s.email):
                raise ValueError(f"The password for {s.email} is missing.")
        cfg = load_config() or Config()
        cfg.sources = sources
        cfg.permissions = Permissions(modify=bool(perms.get("modify", True)), send=bool(perms.get("send", False)))
        path = save_config(cfg)
        if self.on_saved:
            try:
                self.on_saved(cfg)
            except Exception as exc:
                log.warning("Neu laden fehlgeschlagen: %s", exc)
        return {"ok": True, "path": str(path)}

    def api_register(self, body: dict[str, Any]) -> dict[str, Any]:
        ids = [str(a) for a in body.get("apps", [])]
        installed = ""
        if ids and self.register_fn is aiapps.register:
            _path, installed = aiapps.ensure_installed()
        results = self.register_fn(ids)
        return {"results": [r.to_dict() for r in results], "installed": installed}

    def api_finish(self, _body: dict[str, Any]) -> dict[str, Any]:
        self.stop()
        return {"ok": True}


class _Handler(BaseHTTPRequestHandler):
    ctx: SetupApp
    server_version = "post-setup"
    sys_version = ""

    def log_message(self, fmt: str, *args: Any) -> None:  # leise
        log.debug("setup: " + fmt, *args)

    # -- Security --

    def _host_ok(self) -> bool:
        host = (self.headers.get("Host") or "").lower()
        port = self.ctx.port
        return host in (f"127.0.0.1:{port}", f"localhost:{port}")

    def _origin_ok(self) -> bool:
        origin = self.headers.get("Origin")
        if origin is None:
            return True
        port = self.ctx.port
        return origin in (f"http://127.0.0.1:{port}", f"http://localhost:{port}")

    def _token_ok(self) -> bool:
        return secrets.compare_digest(self.headers.get("X-Post-Token", ""), self.ctx.token)

    def _send(self, status: int, body: bytes, ctype: str) -> None:
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header(
            "Content-Security-Policy",
            "default-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self'; "
            "connect-src 'self'; font-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'",
        )
        self.end_headers()
        self.wfile.write(body)

    def _json(self, status: int, data: Any) -> None:
        self._send(status, json.dumps(data, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8")

    # -- Routes --

    def do_GET(self) -> None:  # noqa: N802
        if not self._host_ok():
            self._send(HTTPStatus.FORBIDDEN, b"forbidden", "text/plain")
            return
        url = urlparse(self.path)
        if url.path == "/":
            token = parse_qs(url.query).get("t", [""])[0]
            name = "index.html" if secrets.compare_digest(token, self.ctx.token) else "expired.html"
            body = _static(name) or b""
            self._send(HTTPStatus.OK, body, STATIC_TYPES[".html"])
            return
        if url.path.startswith("/static/"):
            name = url.path[len("/static/") :]
            data = _static(name)
            ext = "." + name.rsplit(".", 1)[-1] if "." in name else ""
            if data is None or ext not in STATIC_TYPES or ext == ".html":
                self._send(HTTPStatus.NOT_FOUND, b"not found", "text/plain")
                return
            self._send(HTTPStatus.OK, data, STATIC_TYPES[ext])
            return
        if url.path == "/api/state":
            if not self._token_ok():
                self._json(HTTPStatus.FORBIDDEN, {"error": "Session expired. Please start setup again."})
                return
            self._dispatch(lambda: self.ctx.api_state())
            return
        self._send(HTTPStatus.NOT_FOUND, b"not found", "text/plain")

    def do_POST(self) -> None:  # noqa: N802
        if not (self._host_ok() and self._origin_ok() and self._token_ok()):
            self._json(HTTPStatus.FORBIDDEN, {"error": "Session expired. Please start setup again."})
            return
        if not (self.headers.get("Content-Type") or "").startswith("application/json"):
            self._json(HTTPStatus.UNSUPPORTED_MEDIA_TYPE, {"error": "JSON erwartet."})
            return
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_BODY:
            self._json(HTTPStatus.REQUEST_ENTITY_TOO_LARGE, {"error": "Request too large."})
            return
        try:
            body = json.loads(self.rfile.read(length) or b"{}")
            if not isinstance(body, dict):
                raise ValueError
        except ValueError:
            self._json(HTTPStatus.BAD_REQUEST, {"error": "Invalid request."})
            return
        routes = {
            "/api/discover": self.ctx.api_discover,
            "/api/test": self.ctx.api_test,
            "/api/test-outlook": self.ctx.api_test_outlook,
            "/api/save": self.ctx.api_save,
            "/api/register": self.ctx.api_register,
            "/api/finish": self.ctx.api_finish,
        }
        handler = routes.get(urlparse(self.path).path)
        if handler is None:
            self._json(HTTPStatus.NOT_FOUND, {"error": "Unbekannt."})
            return
        self._dispatch(lambda: handler(body))

    def _dispatch(self, fn: Callable[[], Any]) -> None:
        try:
            self._json(HTTPStatus.OK, fn())
        except (ValueError, BackendError, vault.SecretStoreError) as exc:
            self._json(HTTPStatus.UNPROCESSABLE_ENTITY, {"error": str(exc)})
        except Exception as exc:  # pragma: no cover
            log.exception("Error inside setup")
            self._json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": f"Unexpected error: {exc}"})
