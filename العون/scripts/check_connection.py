#!/usr/bin/env python3
"""Verify the Odoo connection before wiring the server into Claude.

    uv run python scripts/check_connection.py
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from odoo_mcp.client import OdooClient, OdooError  # noqa: E402
from odoo_mcp.config import OdooConfig, OdooConfigError  # noqa: E402


def load_dotenv(path: Path) -> None:
    """Minimal .env loader so the script works without extra dependencies."""
    if not path.exists():
        return
    import os

    for line in path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip())


def main() -> int:
    load_dotenv(Path(__file__).resolve().parent.parent / ".env")

    try:
        config = OdooConfig.from_env()
    except OdooConfigError as exc:
        print(f"Configuration problem: {exc}")
        return 2

    client = OdooClient(config)
    print(f"Server   : {config.url}")
    print(f"Database : {config.db}")

    try:
        print(f"Version  : {client.version().get('server_version')}")
        user = client.read("res.users", [client.uid], ["name", "login", "company_id"])[0]
        print(f"User     : {user['name']} <{user['login']}>  (uid {client.uid})")
        company = user.get("company_id")
        print(f"Company  : {company[1] if company else '-'}")
        print(f"Partners : {client.search_count('res.partner')} contacts visible")
    except OdooError as exc:
        print(f"\nFailed: {exc}")
        return 1

    print(f"\nWrite allowed : {config.allow_write}")
    print(f"Delete allowed: {config.allow_delete}")
    print("\nConnection OK.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
