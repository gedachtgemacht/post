"""Detects AI apps (Claude Desktop, Claude Code, Codex) and registers post. with them.

Every change to someone else configuration file happens only on request, with
a backup, and in a way that leaves existing entries untouched.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

SERVER_NAME = "post"


@dataclass
class AiApp:
    id: str
    name: str
    installed: bool
    registered: bool = False
    detail: str = ""
    config_path: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class RegisterResult:
    app: str
    ok: bool
    message: str
    restart_needed: bool = False

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


def install_target(env: dict[str, str] | None = None) -> Path | None:
    """A fixed home for post.exe, so AI apps do not point at the downloads folder."""
    env = env if env is not None else dict(os.environ)
    if not getattr(sys, "frozen", False) or sys.platform != "win32":
        return None
    base = env.get("LOCALAPPDATA") or str(_home(env) / "AppData" / "Local")
    return Path(base) / "Programs" / "post" / "post.exe"


def ensure_installed(env: dict[str, str] | None = None) -> tuple[str, str]:
    """Copies the running post.exe to its fixed home. Returns the path and a note."""
    target = install_target(env)
    if target is None:
        return sys.executable, ""
    current = Path(sys.executable).resolve()
    if target.exists() and current == target.resolve():
        return str(target), ""
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_name("post.new.exe")
    try:
        shutil.copy2(current, tmp)
        os.replace(tmp, target)
    except PermissionError:
        # The installed copy is running right now (an AI app started it).
        if tmp.exists():
            tmp.unlink(missing_ok=True)
        if target.exists():
            return str(target), (
                "post. is already installed and currently running inside an AI app. "
                "The new version is picked up once you quit the AI app and start setup again."
            )
        raise
    return str(target), f"post. is now installed at {target}. You can delete the file you downloaded."


def server_command(env: dict[str, str] | None = None) -> tuple[str, list[str]]:
    """How an AI app should start post."""
    if getattr(sys, "frozen", False):  # gebaute post.exe
        path, _msg = ensure_installed(env)
        return path, ["serve"]
    return sys.executable, ["-m", "ggpost", "serve"]


def _home(env: dict[str, str]) -> Path:
    return Path(env.get("HOME") or env.get("USERPROFILE") or Path.home())


def _backup(path: Path) -> Path | None:
    if not path.exists():
        return None
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    dest = path.with_name(f"{path.name}.post-backup-{stamp}")
    shutil.copy2(path, dest)
    return dest


def _atomic_write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix=".post-", suffix=".tmp")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as fh:
            fh.write(text)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)


# ── Claude Desktop ───────────────────────────────────────────────────────


def claude_desktop_dirs(env: dict[str, str]) -> list[Path]:
    home = _home(env)
    dirs: list[Path] = []
    if env.get("APPDATA"):
        dirs.append(Path(env["APPDATA"]) / "Claude")
    if env.get("LOCALAPPDATA"):
        pkgs = Path(env["LOCALAPPDATA"]) / "Packages"
        if pkgs.exists():
            dirs += [p / "LocalCache" / "Roaming" / "Claude" for p in pkgs.glob("Claude_*")]
    dirs.append(home / "Library" / "Application Support" / "Claude")
    dirs.append(Path(env.get("XDG_CONFIG_HOME") or home / ".config") / "Claude")
    return [d for d in dirs if d.exists()]


def _desktop_extension_installed(base: Path) -> bool:
    """Is post. already installed as a Claude Desktop extension (.mcpb)?"""
    for manifest in list(base.glob("*xtensions*/*/manifest.json"))[:200]:
        try:
            if json.loads(manifest.read_text(encoding="utf-8")).get("name") == SERVER_NAME:
                return True
        except Exception:
            continue
    return False


def _read_json(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {}
    text = path.read_text(encoding="utf-8").strip()
    if not text:
        return {}
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError(f"{path} has an unexpected format.")
    return data


def detect_claude_desktop(env: dict[str, str]) -> AiApp:
    dirs = claude_desktop_dirs(env)
    if not dirs:
        return AiApp("claude_desktop", "Claude Desktop", installed=False)
    msix = [d for d in dirs if "Packages" in str(d)]  # the Store build wins, it is the active one
    base = msix[0] if msix else dirs[0]
    cfg = base / "claude_desktop_config.json"
    registered = False
    detail = ""
    try:
        registered = SERVER_NAME in _read_json(cfg).get("mcpServers", {})
    except Exception:
        detail = "configuration file not readable"
    if _desktop_extension_installed(base):
        registered = True
        detail = "installed as an extension"
    return AiApp("claude_desktop", "Claude Desktop", installed=True, registered=registered, detail=detail, config_path=str(cfg))


def register_claude_desktop(app: AiApp, command: tuple[str, list[str]]) -> RegisterResult:
    if not app.installed:
        return RegisterResult(app.id, False, "Claude Desktop was not found.")
    if app.detail == "installed as an extension":
        return RegisterResult(app.id, True, "post. is already installed as an extension in Claude Desktop.")
    path = Path(app.config_path)
    try:
        data = _read_json(path)
    except Exception as exc:
        return RegisterResult(app.id, False, f"claude_desktop_config.json is damaged ({exc}). Nothing changed.")
    _backup(path)
    data.setdefault("mcpServers", {})[SERVER_NAME] = {"command": command[0], "args": command[1]}
    _atomic_write(path, json.dumps(data, indent=2, ensure_ascii=False))
    return RegisterResult(app.id, True, "Registered with Claude Desktop. Quit Claude Desktop completely and open it again.", True)


def unregister_claude_desktop(app: AiApp) -> RegisterResult:
    path = Path(app.config_path)
    try:
        data = _read_json(path)
    except Exception as exc:
        return RegisterResult(app.id, False, f"Not readable: {exc}")
    if SERVER_NAME not in data.get("mcpServers", {}):
        return RegisterResult(app.id, True, "Was not registered.")
    _backup(path)
    del data["mcpServers"][SERVER_NAME]
    _atomic_write(path, json.dumps(data, indent=2, ensure_ascii=False))
    return RegisterResult(app.id, True, "Removed from Claude Desktop.", True)


# ── Claude Code ──────────────────────────────────────────────────────────


def detect_claude_code(env: dict[str, str]) -> AiApp:
    exe = shutil.which("claude", path=env.get("PATH"))
    cfg = _home(env) / ".claude.json"
    if not exe and not cfg.exists():
        return AiApp("claude_code", "Claude Code", installed=False)
    registered = False
    try:
        registered = SERVER_NAME in _read_json(cfg).get("mcpServers", {})
    except Exception:
        pass
    return AiApp("claude_code", "Claude Code", installed=True, registered=registered, detail=exe or "", config_path=str(cfg))


def register_claude_code(app: AiApp, command: tuple[str, list[str]]) -> RegisterResult:
    if not app.installed:
        return RegisterResult(app.id, False, "Claude Code was not found.")
    if app.detail:
        try:
            subprocess.run([app.detail, "mcp", "remove", "--scope", "user", SERVER_NAME], capture_output=True, timeout=30)
            proc = subprocess.run(
                [app.detail, "mcp", "add", "--scope", "user", SERVER_NAME, "--", command[0], *command[1]],
                capture_output=True,
                text=True,
                timeout=60,
            )
            if proc.returncode == 0:
                return RegisterResult(app.id, True, "Registered with Claude Code (for every project).")
        except Exception:
            pass
    path = Path(app.config_path)
    try:
        data = _read_json(path)
    except Exception as exc:
        return RegisterResult(app.id, False, f"~/.claude.json not readable ({exc}). Nothing changed.")
    _backup(path)
    data.setdefault("mcpServers", {})[SERVER_NAME] = {"type": "stdio", "command": command[0], "args": command[1], "env": {}}
    _atomic_write(path, json.dumps(data, indent=2, ensure_ascii=False))
    return RegisterResult(app.id, True, "Registered with Claude Code (for every project).")


def unregister_claude_code(app: AiApp) -> RegisterResult:
    if not app.registered:
        return RegisterResult(app.id, True, "Was not registered.")
    if app.detail:
        try:
            subprocess.run([app.detail, "mcp", "remove", "--scope", "user", SERVER_NAME], capture_output=True, timeout=30)
            return RegisterResult(app.id, True, "Removed from Claude Code.")
        except Exception:
            pass
    return RegisterResult(app.id, False, "Please remove it by hand: claude mcp remove post")


# ── Codex ────────────────────────────────────────────────────────────────


def _codex_config(env: dict[str, str]) -> Path:
    base = Path(env["CODEX_HOME"]) if env.get("CODEX_HOME") else _home(env) / ".codex"
    return base / "config.toml"


_CODEX_HEADER_RE = re.compile(r'^\s*\[\s*(?:mcp_servers|"mcp_servers")\s*\.\s*(?:post|"post")\s*(?:\.[^\]]*)?\]\s*(?:#.*)?$')
_ANY_HEADER_RE = re.compile(r"^\s*\[")


def _codex_split(text: str) -> tuple[str, bool]:
    """Entfernt [mcp_servers.post] samt Untertabellen ([mcp_servers.post.env] …)."""
    out: list[str] = []
    inside = found = False
    for line in text.splitlines(keepends=True):
        if _ANY_HEADER_RE.match(line):
            inside = bool(_CODEX_HEADER_RE.match(line))
            found = found or inside
        if not inside:
            out.append(line)
    return "".join(out), found


def detect_codex(env: dict[str, str]) -> AiApp:
    exe = shutil.which("codex", path=env.get("PATH"))
    cfg = _codex_config(env)
    if not exe and not cfg.parent.exists():
        return AiApp("codex", "Codex", installed=False)
    text = cfg.read_text(encoding="utf-8") if cfg.exists() else ""
    return AiApp("codex", "Codex", installed=True, registered=_codex_split(text)[1], detail=exe or "", config_path=str(cfg))


def register_codex(app: AiApp, command: tuple[str, list[str]]) -> RegisterResult:
    if not app.installed:
        return RegisterResult(app.id, False, "Codex was not found.")
    path = Path(app.config_path)
    text = path.read_text(encoding="utf-8") if path.exists() else ""
    block = (
        f"[mcp_servers.{SERVER_NAME}]\n"
        f"command = {json.dumps(command[0])}\n"
        f"args = {json.dumps(command[1])}\n"
    )
    _backup(path)
    rest, _found = _codex_split(text)
    text = (rest.rstrip() + "\n\n" if rest.strip() else "") + block
    _atomic_write(path, text)
    return RegisterResult(app.id, True, "In Codex eingetragen.", True)


def unregister_codex(app: AiApp) -> RegisterResult:
    path = Path(app.config_path)
    if not path.exists():
        return RegisterResult(app.id, True, "Was not registered.")
    rest, found = _codex_split(path.read_text(encoding="utf-8"))
    if not found:
        return RegisterResult(app.id, True, "Was not registered.")
    _backup(path)
    _atomic_write(path, rest.rstrip() + "\n")
    return RegisterResult(app.id, True, "Removed from Codex.")


# ── Sammelfunktionen ─────────────────────────────────────────────────────

_DETECT = {"claude_desktop": detect_claude_desktop, "claude_code": detect_claude_code, "codex": detect_codex}
_REGISTER = {"claude_desktop": register_claude_desktop, "claude_code": register_claude_code, "codex": register_codex}
_UNREGISTER = {"claude_desktop": unregister_claude_desktop, "claude_code": unregister_claude_code, "codex": unregister_codex}


def detect_apps(env: dict[str, str] | None = None) -> list[AiApp]:
    env = env if env is not None else dict(os.environ)
    return [fn(env) for fn in _DETECT.values()]


def register(app_ids: list[str], env: dict[str, str] | None = None, command: tuple[str, list[str]] | None = None) -> list[RegisterResult]:
    env = env if env is not None else dict(os.environ)
    command = command or server_command()
    out = []
    for app_id in app_ids:
        if app_id not in _DETECT:
            out.append(RegisterResult(app_id, False, "Unbekannte App."))
            continue
        app = _DETECT[app_id](env)
        try:
            out.append(_REGISTER[app_id](app, command))
        except Exception as exc:
            out.append(RegisterResult(app_id, False, f"Error: {exc}"))
    return out


def unregister(app_ids: list[str], env: dict[str, str] | None = None) -> list[RegisterResult]:
    env = env if env is not None else dict(os.environ)
    out = []
    for app_id in app_ids:
        app = _DETECT[app_id](env)
        if not app.installed or not app.registered:
            continue
        try:
            out.append(_UNREGISTER[app_id](app))
        except Exception as exc:
            out.append(RegisterResult(app_id, False, f"Error: {exc}"))
    return out
