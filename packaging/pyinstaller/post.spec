# PyInstaller recipe for post.exe (one file, a console program because of stdio).
#   pyinstaller packaging/pyinstaller/post.spec --noconfirm
# Result: dist/post.exe (Windows) or dist/post (Linux/macOS, for tests only)

import sys
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

ROOT = Path(SPECPATH).resolve().parents[1]  # noqa: F821 (set by PyInstaller)

datas = collect_data_files("ggpost", includes=["setup/static/**/*"])
datas += copy_metadata("mcp")
hidden = collect_submodules("keyring.backends") + collect_submodules("mcp", filter=lambda n: not n.startswith(("mcp.cli", "mcp.server.fastmcp")))
if sys.platform == "win32":
    hidden += ["win32timezone", "win32com.client.dynamic", "pythoncom", "pywintypes"]

a = Analysis(  # noqa: F821
    [str(ROOT / "src" / "ggpost" / "__main__.py")],
    pathex=[str(ROOT / "src")],
    datas=datas,
    hiddenimports=hidden,
    excludes=[
        "tkinter", "pytest", "pymap", "playwright", "numpy", "PIL", "fontTools", "arabic_reshaper",
        "pygments", "yaml", "reportlab", "docx", "openpyxl", "lxml", "IPython", "matplotlib",
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)  # noqa: F821
exe = EXE(  # noqa: F821
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="post",
    console=True,
    icon=str(ROOT / "packaging" / "icon.ico"),
    version=None,
    upx=False,
)
