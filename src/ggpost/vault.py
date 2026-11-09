"""Passwords in the operating system's keychain.

Windows: Credential Manager, macOS: Keychain, Linux: Secret Service.
If the keychain is unavailable nothing is stored in plain text; the password
then has to come from the environment.
"""

from __future__ import annotations

import logging

from ggpost.config import env_str

SERVICE = "post-mail-connector"
log = logging.getLogger(__name__)


class SecretStoreError(RuntimeError):
    pass


def _keyring():
    try:
        import keyring  # noqa: PLC0415

        return keyring
    except Exception as exc:  # pragma: no cover - only on a broken installation
        raise SecretStoreError(f"Keychain not available: {exc}") from exc


def get_password(email: str) -> str | None:
    env_email = env_str("POST_EMAIL").lower()
    env_pw = env_str("POST_PASSWORD")
    if env_pw and env_email and env_email == email.lower():
        return env_pw
    try:
        return _keyring().get_password(SERVICE, email.lower())
    except Exception as exc:
        log.warning("Could not read password for %s: %s", email, exc)
        return None


def set_password(email: str, password: str) -> None:
    try:
        _keyring().set_password(SERVICE, email.lower(), password)
    except Exception as exc:
        raise SecretStoreError(
            "The password could not be stored in the keychain "
            f"({exc}). Nothing was written in plain text."
        ) from exc


def delete_password(email: str) -> None:
    try:
        _keyring().delete_password(SERVICE, email.lower())
    except Exception:
        pass
