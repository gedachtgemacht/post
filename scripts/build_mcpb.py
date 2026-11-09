"""Builds the Claude Desktop extension (.mcpb).

  python scripts/build_mcpb.py windows   # needs dist/post.exe (PyInstaller)
  python scripts/build_mcpb.py uv        # platform independent, Python through uv

Requires the MCPB CLI:  npm install -g @anthropic-ai/mcpb
"""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def version() -> str:
    text = (ROOT / "src" / "ggpost" / "__init__.py").read_text(encoding="utf-8")
    return re.search(r'__version__ = "([^"]+)"', text).group(1)


def stage(kind: str) -> Path:
    out = ROOT / "build" / f"mcpb-{kind}"
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    src_manifest = ROOT / "packaging" / ("mcpb" if kind == "windows" else "mcpb-uv") / "manifest.json"
    manifest = json.loads(src_manifest.read_text(encoding="utf-8"))
    manifest["version"] = version()
    (out / "manifest.json").write_text(json.dumps(manifest, indent=2, ensure_ascii=False), encoding="utf-8")
    shutil.copy2(ROOT / "packaging" / "icon.png", out / "icon.png")
    shutil.copy2(ROOT / "LICENSE", out / "LICENSE")
    shutil.copy2(ROOT / "README.md", out / "README.md")
    if kind == "windows":
        exe = ROOT / "dist" / "post.exe"
        if not exe.exists():
            sys.exit("dist/post.exe is missing. Build it with PyInstaller first (see README).")
        (out / "server").mkdir()
        shutil.copy2(exe, out / "server" / "post.exe")
    else:
        shutil.copy2(ROOT / "pyproject.toml", out / "pyproject.toml")
        shutil.copytree(ROOT / "src", out / "src", ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
    return out


def main() -> None:
    kind = sys.argv[1] if len(sys.argv) > 1 else "windows"
    if kind not in ("windows", "uv"):
        sys.exit("usage: build_mcpb.py windows|uv")
    staged = stage(kind)
    mcpb = shutil.which("mcpb")
    if not mcpb:
        sys.exit("MCPB CLI is missing: npm install -g @anthropic-ai/mcpb")
    dist = ROOT / "dist"
    dist.mkdir(exist_ok=True)
    target = dist / f"post-{version()}-{kind}.mcpb"
    subprocess.run([mcpb, "validate", str(staged / "manifest.json")], check=True)
    subprocess.run([mcpb, "pack", str(staged), str(target)], check=True)
    print(f"Done: {target}")


if __name__ == "__main__":
    main()
