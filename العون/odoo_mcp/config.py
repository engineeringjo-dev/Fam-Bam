"""Configuration for the Odoo MCP server, read from environment variables."""

from __future__ import annotations

import os
from dataclasses import dataclass


class OdooConfigError(RuntimeError):
    """Raised when the server is not configured correctly."""


def _flag(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    if raw is None or raw == "":
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


def _int(name: str, default: int) -> int:
    raw = os.environ.get(name)
    if raw is None or raw == "":
        return default
    try:
        return int(raw)
    except ValueError as exc:
        raise OdooConfigError(f"{name} must be an integer, got {raw!r}") from exc


@dataclass(frozen=True)
class OdooConfig:
    url: str
    db: str
    username: str
    api_key: str
    lang: str | None = None
    timeout: int = 30
    default_limit: int = 80
    max_limit: int = 1000
    allow_write: bool = True
    allow_delete: bool = False

    @classmethod
    def from_env(cls) -> "OdooConfig":
        missing = [
            name
            for name in ("ODOO_URL", "ODOO_DB", "ODOO_USERNAME", "ODOO_API_KEY")
            if not os.environ.get(name)
        ]
        if missing:
            raise OdooConfigError(
                "Missing required environment variable(s): "
                + ", ".join(missing)
                + ". See .env.example for the expected values."
            )

        url = os.environ["ODOO_URL"].strip().rstrip("/")
        if not url.startswith(("http://", "https://")):
            url = "https://" + url
        if url.startswith("http://") and not _flag("ODOO_ALLOW_INSECURE", False):
            raise OdooConfigError(
                f"Refusing to send the API key over plain HTTP ({url}). "
                "Use https://, or set ODOO_ALLOW_INSECURE=1 for a trusted local server."
            )

        lang = os.environ.get("ODOO_LANG", "").strip() or None

        return cls(
            url=url,
            db=os.environ["ODOO_DB"].strip(),
            username=os.environ["ODOO_USERNAME"].strip(),
            api_key=os.environ["ODOO_API_KEY"].strip(),
            lang=lang,
            timeout=_int("ODOO_TIMEOUT", 30),
            default_limit=_int("ODOO_DEFAULT_LIMIT", 80),
            max_limit=_int("ODOO_MAX_LIMIT", 1000),
            allow_write=_flag("ODOO_ALLOW_WRITE", True),
            allow_delete=_flag("ODOO_ALLOW_DELETE", False),
        )

    @property
    def base_context(self) -> dict:
        return {"lang": self.lang} if self.lang else {}
