"""post. command line.

  post            Double-click or terminal: open setup. Started by an AI app: serve.
  post serve      MCP server over stdio (this is what the AI app calls)
  post setup      Setup in the browser
  post doctor     Check the connections
  post detect     Detected mail programs and AI apps as JSON
  post register   Register with AI apps (claude_desktop, claude_code, codex)
  post uninstall  Remove from every AI app
"""

from __future__ import annotations

import argparse
import json
import logging
import sys

from ggpost import __version__


def _log_to_stderr(level: int = logging.INFO) -> None:
    logging.basicConfig(level=level, stream=sys.stderr, format="post: %(levelname)s %(message)s")


def cmd_serve(_args: argparse.Namespace) -> int:
    # Important: stdout belongs to the MCP protocol. Everything else goes to stderr.
    _log_to_stderr(logging.WARNING)
    from ggpost.server import build_server  # noqa: PLC0415
    from ggpost.service import MailService  # noqa: PLC0415

    service = MailService.from_environment()
    holder: dict[str, object] = {}

    def launch_setup() -> str:
        from ggpost.backends.base import BackendError  # noqa: PLC0415
        from ggpost.setup.web import SetupApp, open_browser_quietly  # noqa: PLC0415

        old = holder.get("app")
        if old is not None:
            old.stop()  # type: ignore[attr-defined]
        app = SetupApp(on_saved=service.reload, mode="embedded")
        holder["app"] = app
        url = app.start(open_browser=False)
        if not open_browser_quietly(url):
            app.stop()
            raise BackendError("The browser could not be opened. Please run 'post setup' in a terminal.")
        return url

    server = build_server(service, setup_launcher=launch_setup)
    server.run("stdio")
    return 0


def cmd_setup(args: argparse.Namespace) -> int:
    _log_to_stderr()
    from ggpost.setup.web import SetupApp  # noqa: PLC0415

    app = SetupApp(mode="standalone")
    url = app.start(open_browser=not args.no_browser)
    print("post. setup")
    print(f"Opened in the browser: {url}")
    print("If no browser opens, copy the link above. Stop with Ctrl+C.")
    try:
        app.wait(timeout=60 * 60)
    except KeyboardInterrupt:
        pass
    app.stop()
    return 0


def cmd_detect(_args: argparse.Namespace) -> int:
    from ggpost.detect import aiapps, clients  # noqa: PLC0415

    data = {"mail": clients.detect().to_dict(), "ai_apps": [a.to_dict() for a in aiapps.detect_apps()]}
    print(json.dumps(data, indent=2, ensure_ascii=False))
    return 0


def cmd_doctor(_args: argparse.Namespace) -> int:
    _log_to_stderr(logging.WARNING)
    from ggpost.config import config_path, load_config  # noqa: PLC0415
    from ggpost.detect import aiapps  # noqa: PLC0415
    from ggpost.service import MailService  # noqa: PLC0415

    print(f"post. {__version__}")
    cfg = load_config()
    print(f"Settings: {config_path()} ({'present' if cfg else 'missing, using automatic detection'})")
    service = MailService.from_environment()
    if not service.configured:
        print("No mailbox set up. Run: post setup")
        if service.auto.hint:
            print(f"  {service.auto.hint}")
        return 1
    if service.auto.via:
        print(f"Automatically chosen route: {service.auto.via}")
    failures = 0
    for src, backend, err in service._ensure():  # noqa: SLF001
        name = src.email or "Outlook"
        if backend is None:
            print(f"  ✗ {name}: {err}")
            failures += 1
            continue
        try:
            print(f"  ✓ {name}: {backend.check()}")
        except Exception as exc:
            print(f"  ✗ {name}: {exc}")
            failures += 1
    p = service.config.permissions
    rights = ["read", "draft"] + (["move, delete and mark"] if p.modify else []) + (["send"] if p.send else [])
    print("Permissions: " + ", ".join(rights))
    for app in aiapps.detect_apps():
        if app.installed:
            print(f"  {'✓' if app.registered else '·'} {app.name}: {'registered' if app.registered else 'not registered'}")
    return 1 if failures else 0


def cmd_register(args: argparse.Namespace) -> int:
    from ggpost.detect import aiapps  # noqa: PLC0415

    ids = args.apps or [a.id for a in aiapps.detect_apps() if a.installed]
    if not ids:
        print("No AI app found (Claude Desktop, Claude Code, Codex).")
        return 1
    _path, installed = aiapps.ensure_installed()
    if installed:
        print(installed)
    ok = True
    for r in aiapps.register(ids):
        print(f"{'✓' if r.ok else '✗'} {r.app}: {r.message}")
        ok = ok and r.ok
    return 0 if ok else 1


def cmd_uninstall(args: argparse.Namespace) -> int:
    import shutil  # noqa: PLC0415

    from ggpost import vault  # noqa: PLC0415
    from ggpost.config import config_dir, load_config  # noqa: PLC0415
    from ggpost.detect import aiapps  # noqa: PLC0415

    results = aiapps.unregister(["claude_desktop", "claude_code", "codex"])
    for r in results:
        print(f"{'✓' if r.ok else '✗'} {r.app}: {r.message}")
    if not results:
        print("Not registered with any AI app.")
    if not args.all:
        print("Settings and stored passwords are still there. Remove everything: post uninstall --all")
        return 0
    cfg = load_config()
    for src in cfg.sources if cfg else []:
        if src.kind == "imap":
            vault.delete_password(src.email)
            print(f"OK password for {src.email} deleted")
    folder = config_dir()
    if folder.exists():
        shutil.rmtree(folder, ignore_errors=True)
        print(f"OK settings deleted ({folder})")
    target = aiapps.install_target()
    if target and target.exists():
        print(f"The program file is still at {target} and can be deleted once no AI app is running any more.")
    return 0


class _Parser(argparse.ArgumentParser):
    """argparse with our own headings and error wording."""

    def __init__(self, *args, **kwargs):
        kwargs.setdefault("add_help", False)
        super().__init__(*args, **kwargs)
        self._positionals.title = "Commands"
        self._optionals.title = "Options"
        self.add_argument("-h", "--help", action="help", help="show this help")

    def error(self, message: str) -> None:  # type: ignore[override]
        self.print_usage()
        self.exit(2, f"error: {message}\n")


def main(argv: list[str] | None = None) -> int:
    parser = _Parser(prog="post", description="post. Local mail connector for AI apps")
    parser.add_argument("--version", action="version", version=f"post. {__version__}", help="show the version")
    sub = parser.add_subparsers(dest="cmd", metavar="COMMAND", parser_class=_Parser)
    sub.add_parser("serve", help="start the MCP server over stdio (this is what the AI app calls)")
    p_setup = sub.add_parser("setup", help="open setup in the browser")
    p_setup.add_argument("--no-browser", action="store_true", help="only print the link")
    sub.add_parser("doctor", help="check the connections")
    sub.add_parser("detect", help="show detected mail programs and AI apps")
    p_reg = sub.add_parser("register", help="register with AI apps")
    p_reg.add_argument("apps", nargs="*", choices=["claude_desktop", "claude_code", "codex"], help="default: every one found")
    p_un = sub.add_parser("uninstall", help="remove from every AI app")
    p_un.add_argument("--all", action="store_true", help="also delete the settings and stored passwords")
    args = parser.parse_args(argv)

    if args.cmd is None:
        interactive = sys.stdin is not None and sys.stdin.isatty()
        if interactive:
            args = parser.parse_args(["setup"])
        else:
            args = parser.parse_args(["serve"])

    handlers = {
        "serve": cmd_serve,
        "setup": cmd_setup,
        "doctor": cmd_doctor,
        "detect": cmd_detect,
        "register": cmd_register,
        "uninstall": cmd_uninstall,
    }
    return handlers[args.cmd](args)


if __name__ == "__main__":
    raise SystemExit(main())
