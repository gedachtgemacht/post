"""Starts setup with sample data, without touching any real account.

For design work and screenshots:  python scripts/demo_setup.py [--shots DIR]
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from ggpost.config import ServerSettings  # noqa: E402
from ggpost.detect import aiapps, autoconfig  # noqa: E402
from ggpost.setup.web import SetupApp  # noqa: E402
from tests.test_setup_web import fake_apps, fake_detection  # noqa: E402


def discover(email: str) -> autoconfig.Discovery:
    time.sleep(0.6)
    domain = email.split("@")[-1]
    if domain in ("outlook.de", "hotmail.com"):
        hint, _url, ok = autoconfig.provider_hint(domain, "")
        return autoconfig.Discovery(email=email, provider="Outlook.com", hint=hint, supported=ok, source="builtin")
    hint, url, _ok = autoconfig.provider_hint(domain, "")
    return autoconfig.Discovery(
        email=email, provider=domain, imap=ServerSettings(f"imap.{domain}", 993, "ssl", email),
        smtp=ServerSettings(f"smtp.{domain}", 465, "ssl", email), hint=hint, help_url=url, source="ispdb",
    )


def test_imap(_s, _e, password: str) -> str:
    time.sleep(0.5)
    if password != "demo":
        raise ValueError("Login at imap.gmx.net failed. Check the password; some providers require an app password.")
    return "Inbox: 1,284 mails, 7 unread"


def make_app() -> SetupApp:
    return SetupApp(
        detect_fn=fake_detection,
        apps_fn=fake_apps,
        discover_fn=discover,
        test_imap_fn=test_imap,
        test_outlook_fn=lambda: {
            "message": "2 accounts, inbox info@palletcompany.com: 3,912 mails",
            "accounts": [
                {"email": "info@palettenfirma.de", "detail": "Outlook, Exchange"},
                {"email": "vertrieb@palettenfirma.de", "detail": "Outlook, IMAP"},
            ],
        },
        register_fn=lambda ids: [aiapps.RegisterResult(i, True, "Registered with Claude Desktop. Quit Claude Desktop completely and open it again.", True) for i in ids],
    )


def shots(out: Path) -> None:
    from playwright.sync_api import sync_playwright

    out.mkdir(parents=True, exist_ok=True)
    app = make_app()
    url = app.start(open_browser=False)
    with sync_playwright() as p:
        browser = p.chromium.launch()
        for name, size in (("desktop", {"width": 1280, "height": 900}), ("mobile", {"width": 390, "height": 844})):
            page = browser.new_page(viewport=size, device_scale_factor=2)
            page.goto(url)
            page.wait_for_selector(".source")
            page.wait_for_timeout(400)
            page.screenshot(path=str(out / f"{name}-1-postfaecher.png"), full_page=True)
            if name == "desktop":
                page.fill("#add-email", "dennis@gmail.com")
                page.click("#add-form button")
                page.wait_for_selector(".imap-row[data-email=\"dennis@gmail.com\"] .hint:not([hidden])")
                page.wait_for_timeout(300)
                page.screenshot(path=str(out / f"{name}-2-adresse.png"), full_page=True)
                for row in page.query_selector_all(".imap-row.picked"):
                    pw = row.query_selector(".pw")
                    pw.fill("demo")
                page.click("#next-1")
                page.wait_for_selector("#step-2.open")
                page.wait_for_timeout(300)
                page.screenshot(path=str(out / f"{name}-3-permissions.png"), full_page=True)
                page.click("#next-2")
                page.wait_for_selector("#step-3.open")
                page.wait_for_timeout(800)
                page.screenshot(path=str(out / f"{name}-4-apps.png"), full_page=True)
                page.click("#next-3")
                page.wait_for_selector("#step-4.open")
                page.wait_for_timeout(800)
                page.screenshot(path=str(out / f"{name}-5-fertig.png"), full_page=True)
            page.close()
        browser.close()
    app.stop()


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--shots", type=Path, help="write screenshots into this directory and exit")
    args = ap.parse_args()
    if args.shots:
        import os
        import tempfile

        os.environ.setdefault("POST_CONFIG_DIR", tempfile.mkdtemp(prefix="post-demo-"))
        import keyring
        from tests.conftest import MemoryKeyring

        keyring.set_keyring(MemoryKeyring())
        shots(args.shots)
        return
    app = make_app()
    print(app.start())
    try:
        app.wait()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
