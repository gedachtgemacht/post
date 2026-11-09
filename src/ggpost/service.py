"""Ties configuration to backends: which account runs through which backend."""

from __future__ import annotations

import logging
import os
import sys
import threading
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from typing import Any, Callable

from ggpost import vault
from ggpost.backends.base import AccountInfo, BackendError, MailBackend, MessageSummary, NotFound, SearchQuery
from ggpost.config import Config, Source, config_path, env_flag, env_str, load_config, save_config
from ggpost.ids import decode_id

log = logging.getLogger(__name__)

BackendFactory = Callable[[Source], MailBackend]


def default_factory(source: Source) -> MailBackend:
    if source.kind == "outlook":
        if sys.platform != "win32":
            raise BackendError("Classic Outlook only exists on Windows.")
        from ggpost.backends.outlook_com import OutlookBackend  # noqa: PLC0415

        return OutlookBackend(source)
    if source.kind == "imap":
        if source.imap is None:
            _discover_into(source)
        from ggpost.backends.imap import ImapBackend  # noqa: PLC0415

        return ImapBackend(source, vault.get_password)
    raise BackendError(f"Unknown source: {source.kind}")


def _cache_path():
    from ggpost.config import config_dir  # noqa: PLC0415

    return config_dir() / "discovery-cache.json"


def _discover_into(source: Source) -> None:
    """For accounts from the Claude Desktop extension: look up the servers.
    The result is cached so the next start does not have to wait."""
    import json  # noqa: PLC0415

    from ggpost.config import ServerSettings  # noqa: PLC0415

    key = source.email.lower()
    path = _cache_path()
    try:
        cache = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    except (OSError, ValueError):
        cache = {}
    hit = cache.get(key)
    if hit and hit.get("imap"):
        source.imap = ServerSettings.from_dict(hit["imap"])
        source.smtp = ServerSettings.from_dict(hit.get("smtp"))
    else:
        from ggpost.detect.autoconfig import discover  # noqa: PLC0415

        found = discover(source.email)
        if not found.supported:
            raise BackendError(found.hint)
        if not found.imap:
            raise BackendError(
                f"No IMAP settings were found for {source.email}. Please run setup (post_setup)."
            )
        source.imap, source.smtp = found.imap, found.smtp
        cache[key] = {"imap": found.imap.__dict__, "smtp": found.smtp.__dict__ if found.smtp else None}
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(cache, indent=2), encoding="utf-8")
        except OSError:
            pass
    if not source.label:
        source.label = source.email


@dataclass
class AutoChoice:
    """What falls out by itself when there is no settings file.

    ``sources`` is ordered by fitness: the first entry is the way post. takes
    when the AI names no account. ``hint`` says in plain words what is still
    missing, so the AI does not leave the user in the dark.
    """

    sources: list[Source] = field(default_factory=list)
    via: str = ""
    hint: str = ""


def _env_email(env: dict[str, str]) -> str:
    """POST_EMAIL from the given environment, ignoring placeholders from
    manifest.json that were never filled in (``${user_config.email}``)."""
    raw = (env.get("POST_EMAIL") or "").strip()
    return "" if raw.startswith("${") else raw


def _thunderbird_sources(client: Any) -> list[Source]:
    """Thunderbird accounts as IMAP sources. Host, port and user name come from
    the profile; only the password is still missing. For POP3 accounts ``imap``
    stays empty and is looked up on first access."""
    out: list[Source] = []
    for acc in getattr(client, "accounts", []) or []:
        if "@" not in (acc.email or ""):
            continue
        out.append(
            Source(kind="imap", email=acc.email, label=acc.name or acc.email, imap=acc.imap, smtp=acc.smtp)
        )
    return out


def choose_sources(det: Any = None, env: dict[str, str] | None = None) -> AutoChoice:
    """Finds the workable way by itself: an explicitly given address, classic
    Outlook, Thunderbird, and guessing last.

    The order follows what works without the user lifting a finger. Classic
    Outlook needs no password and therefore wins as soon as it holds accounts;
    if Thunderbird is the default mail program, it comes first. Nothing is
    probed here: connections are made on first access, so the server starts
    fast and Outlook does not launch without need.
    """
    env = dict(os.environ) if env is None else env
    if det is None:
        try:
            from ggpost.detect.clients import detect  # noqa: PLC0415

            det = detect(env=env)
        except Exception as exc:  # detection is a convenience, not a must
            log.info("Automatic detection failed: %s", exc)

    outlook = det.client("outlook_classic") if det else None
    thunderbird = det.client("thunderbird") if det else None
    new_outlook = det.client("outlook_new") if det else None
    default_app = (getattr(det, "default_mail_app", "") or "") if det else ""

    tb_sources = _thunderbird_sources(thunderbird) if thunderbird else []
    outlook_sources = [Source(kind="outlook", label="Outlook")] if outlook and outlook.accounts else []

    ways: list[tuple[int, str, list[Source]]] = []
    if outlook_sources:
        ways.append((2 if default_app == "thunderbird" and tb_sources else 1, "outlook_classic", outlook_sources))
    if tb_sources:
        ways.append((1 if default_app == "thunderbird" else 2, "thunderbird", tb_sources))
    ways.sort(key=lambda w: w[0])

    sources: list[Source] = []
    via: list[str] = []

    # An explicitly given address (POST_EMAIL, e.g. from the Claude Desktop
    # extension) is the clearest wish and therefore comes first.
    email = _env_email(env)
    if "@" in email:
        sources.append(Source(kind="imap", email=email, label=email))
        via.append("POST_EMAIL")
    for _rank, name, found in ways:
        for src in found:
            if src.kind == "imap" and any(s.email.lower() == src.email.lower() for s in sources):
                continue
            sources.append(src)
        if name not in via:
            via.append(name)

    if sources:
        hint = ""
        if all(s.kind == "imap" for s in sources):
            hint = (
                "The mailboxes are known, but the password is only in the keychain once setup "
                "has run. If post. cannot log in, post_setup will help."
            )
        return AutoChoice(sources=sources, via="+".join(via), hint=hint)

    # Nothing carries by itself: say as precisely as possible what is wrong.
    if outlook and not outlook.accounts:
        hint = (
            "Classic Outlook is installed, but its profile holds no account. Set one up there, "
            "or connect the mailbox directly over IMAP with post_setup."
        )
    elif new_outlook:
        hint = (
            "Only the new Outlook was found, and it cannot be driven from outside. Call "
            "post_setup and enter the mail address; post. then connects over IMAP."
        )
    elif det and det.clients:
        hint = (
            "A mail program was found, but it does not expose its accounts. Call post_setup "
            "and enter the mail address."
        )
    else:
        hint = "No mail program found. Call post_setup and enter the mail address."
    return AutoChoice(hint=hint)


def parse_day(value: str | None) -> datetime | None:
    """'2026-10-01', 'today', 'yesterday', '7d' (7 days ago) to local midnight.

    The German words stay understood, so a German-speaking user can say 'heute'.
    """
    if not value:
        return None
    v = value.strip().lower()
    today = datetime.now().astimezone().replace(hour=0, minute=0, second=0, microsecond=0)
    if v in ("today", "heute"):
        return today
    if v in ("yesterday", "gestern"):
        return today - timedelta(days=1)
    if v.endswith("d") and v[:-1].isdigit():
        return today - timedelta(days=int(v[:-1]))
    if v.endswith("w") and v[:-1].isdigit():
        return today - timedelta(weeks=int(v[:-1]))
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.strptime(v, fmt).astimezone()
        except ValueError:
            continue
    raise BackendError(f"Could not read the date '{value}'. Use YYYY-MM-DD, 'today', 'yesterday' or '7d'.")


class MailService:
    def __init__(self, config: Config | None, factory: BackendFactory = default_factory, persist: bool = True) -> None:
        self._factory = factory
        self._persist = persist
        self._lock = threading.RLock()
        self._backends: list[tuple[Source, MailBackend | None, str]] = []
        self.config = Config()
        self.auto = AutoChoice()  # only filled when we had to pick without a settings file
        self.reload(config)

    @classmethod
    def from_environment(cls) -> MailService:
        cfg = load_config()
        if cfg is None:
            # Nothing set up: pick the most workable way, but write nothing down
            choice = choose_sources()
            service = cls(Config(sources=choice.sources) if choice.sources else None, persist=False)
            service.auto = choice
            return service
        email = env_str("POST_EMAIL")
        if email and "@" in email and not any(s.email.lower() == email.lower() for s in cfg.sources):
            # Address from the Claude Desktop extension, on top of what setup saved
            cfg.sources.append(Source(kind="imap", email=email, label=email))
            return cls(cfg, persist=False)
        return cls(cfg)

    def reload(self, config: Config | None) -> None:
        with self._lock:
            for _src, backend, _err in self._backends:
                if backend is not None:
                    try:
                        backend.close()
                    except Exception:
                        pass
            self.config = config or Config()
            # Switches from the environment (e.g. the Claude Desktop extension): "on" grants.
            # "off" only counts while no settings file exists; otherwise setup decides.
            has_file = config_path().exists()
            for env, attr in (("POST_ALLOW_SEND", "send"), ("POST_ALLOW_MODIFY", "modify")):
                flag = env_flag(env)
                if flag is True or (flag is False and not has_file):
                    setattr(self.config.permissions, attr, flag)
            self._backends = [(s, None, "") for s in self.config.sources]

    @property
    def configured(self) -> bool:
        return bool(self.config.sources)

    def _ensure(self) -> list[tuple[Source, MailBackend | None, str]]:
        with self._lock:
            changed = False
            for i, (src, backend, _err) in enumerate(self._backends):
                if backend is None:
                    had_imap = src.imap is not None
                    try:
                        backend = self._factory(src)
                        self._backends[i] = (src, backend, "")
                        if src.kind == "imap" and not had_imap and src.imap is not None:
                            changed = True
                    except BackendError as exc:
                        self._backends[i] = (src, None, str(exc))
            if changed and self._persist:
                try:
                    save_config(self.config)
                except OSError as exc:
                    log.warning("Configuration not saved: %s", exc)
            return list(self._backends)

    def require_setup(self) -> None:
        if not self.configured:
            msg = (
                "post. is not set up yet. Call the post_setup tool and setup opens in the browser."
            )
            if self.auto.hint:
                msg += " " + self.auto.hint
            raise BackendError(msg)

    def accounts(self) -> tuple[list[AccountInfo], list[str]]:
        self.require_setup()
        accounts: list[AccountInfo] = []
        problems: list[str] = []
        for src, backend, err in self._ensure():
            if backend is None:
                problems.append(err)
                continue
            try:
                accounts.extend(backend.accounts())
            except BackendError as exc:
                problems.append(str(exc))
        return accounts, problems

    def backend_for(self, account: str | None) -> tuple[MailBackend, str]:
        self.require_setup()
        entries = self._ensure()
        problems = [err for _s, b, err in entries if b is None and err]
        if account:
            for _src, backend, _err in entries:
                if backend is not None and backend.owns(account):
                    return backend, account
            msg = f"Account {account} is not set up in post."
            if problems:
                msg += " " + " ".join(problems)
            raise NotFound(msg)
        for _src, backend, _err in entries:
            if backend is None:
                continue
            accs = backend.accounts()
            if accs:
                return backend, accs[0].email
        raise BackendError(" ".join(problems) or "No account reachable.")

    def backend_for_ref(self, message_id: str) -> tuple[MailBackend, dict[str, Any]]:
        ref = decode_id(message_id)
        backend, _ = self.backend_for(ref["a"])
        return backend, ref

    def search_all(self, query: SearchQuery) -> tuple[list[MessageSummary], list[str]]:
        """Searches every account and merges the results by date."""
        accounts, problems = self.accounts()
        results: list[MessageSummary] = []
        for acc in accounts:
            try:
                backend, email = self.backend_for(acc.email)
                results.extend(backend.search(email, query))
            except BackendError as exc:
                problems.append(f"{acc.email}: {exc}")
        results.sort(key=lambda m: m.date, reverse=True)
        return results[: query.limit], problems

    def status(self) -> dict[str, Any]:
        accounts: list[dict[str, Any]] = []
        problems: list[str] = []
        if self.configured:
            accs, problems = self.accounts()
            accounts = [a.to_dict() for a in accs]
        out = {
            "configured": self.configured,
            "accounts": accounts,
            "permissions": {
                "read": True,
                "draft": True,
                "move_and_mark": self.config.permissions.modify,
                "send": self.config.permissions.send,
            },
            "problems": problems,
        }
        if self.auto.via:
            out["detected_via"] = self.auto.via
        if self.auto.hint and (not accounts or problems):
            out["hint"] = self.auto.hint
        return out
