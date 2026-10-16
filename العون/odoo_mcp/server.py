"""MCP server exposing an Odoo database to Claude over stdio."""

from __future__ import annotations

import functools
import json
import os
from typing import Any, Callable

from mcp.types import ToolAnnotations

try:  # mcp >= 2.0
    from mcp.server import MCPServer as _Server
except ImportError:  # mcp 1.x
    from mcp.server.fastmcp import FastMCP as _Server

from .client import OdooClient, OdooError, iter_ids
from .config import OdooConfig, OdooConfigError

mcp = _Server(
    "odoo",
    instructions=(
        "Query and update an Odoo ERP database. Call odoo_status first, then "
        "odoo_list_models / odoo_fields to discover the schema before reading or "
        "writing. Prefer odoo_aggregate over reading many records for totals."
    ),
)

_client: OdooClient | None = None

MAX_CHARS = int(os.environ.get("ODOO_MAX_RESPONSE_CHARS", "60000"))


def get_client() -> OdooClient:
    """Build the client on first use so the server still starts unconfigured."""
    global _client
    if _client is None:
        _client = OdooClient(OdooConfig.from_env())
    return _client


def dump(value: Any) -> str:
    """Serialize a result as UTF-8 JSON so Arabic text stays readable."""
    text = json.dumps(value, ensure_ascii=False, indent=2, default=str)
    if len(text) > MAX_CHARS:
        return (
            text[:MAX_CHARS]
            + f"\n\n... truncated at {MAX_CHARS} characters. "
            "Narrow the domain, request fewer fields, or use limit/offset."
        )
    return text


def tool(
    *, read_only: bool = False, destructive: bool = False
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Register an MCP tool that reports Odoo failures as readable text."""

    annotations = ToolAnnotations.model_validate(
        {"readOnlyHint": read_only, "destructiveHint": destructive}
    )

    def decorator(fn: Callable[..., Any]) -> Callable[..., Any]:
        @functools.wraps(fn)
        def wrapper(*args: Any, **kwargs: Any) -> str:
            try:
                return fn(*args, **kwargs)
            except (OdooError, OdooConfigError) as exc:
                return f"ERROR: {exc}"

        return mcp.tool(annotations=annotations)(wrapper)

    return decorator


# -- connection ---------------------------------------------------------------


@tool(read_only=True)
def odoo_status() -> str:
    """Check the Odoo connection and report who this server is logged in as.

    Call this first in a session: it confirms the server URL, database, Odoo
    version, the authenticated user and company, and whether writing and
    deleting are currently allowed.
    """
    client = get_client()
    version = client.version()
    user = client.read(
        "res.users",
        [client.uid],
        ["name", "login", "company_id", "lang", "tz"],
    )
    return dump(
        {
            "url": client.config.url,
            "database": client.config.db,
            "server_version": version.get("server_version"),
            "uid": client.uid,
            "user": user[0] if user else None,
            "write_allowed": client.config.allow_write,
            "delete_allowed": client.config.allow_delete,
            "default_limit": client.config.default_limit,
        }
    )


# -- schema discovery ---------------------------------------------------------


@tool(read_only=True)
def odoo_list_models(search: str = "", limit: int = 60) -> str:
    """List the models available in this database, optionally filtered.

    Use this to find the technical model name behind a business concept, e.g.
    search "pos" for point-of-sale models or "stock" for inventory ones.
    """
    domain: list[Any] = []
    if search:
        domain = ["|", ("model", "ilike", search), ("name", "ilike", search)]
    records = get_client().search_read(
        "ir.model", domain, ["model", "name", "transient"], limit=limit, order="model"
    )
    return dump(records)


@tool(read_only=True)
def odoo_fields(model: str, search: str = "") -> str:
    """Describe a model's fields: technical name, label, type and relation.

    Always check this before writing to an unfamiliar model, so the values you
    send match the real field names and types.
    """
    fields = get_client().fields_get(model)
    if search:
        needle = search.lower()
        fields = {
            name: meta
            for name, meta in fields.items()
            if needle in name.lower() or needle in str(meta.get("string", "")).lower()
        }
    return dump(fields)


# -- reading ------------------------------------------------------------------


@tool(read_only=True)
def odoo_search_read(
    model: str,
    domain: list[Any] | None = None,
    fields: list[str] | None = None,
    limit: int | None = None,
    offset: int = 0,
    order: str | None = None,
) -> str:
    """Search records and read their fields in one call. The main read tool.

    `domain` uses Odoo's search syntax, as a list of triples plus optional
    "&", "|", "!" operators, for example:
        [["state", "=", "sale"], ["date_order", ">=", "2026-01-01"]]
        ["|", ["name", "ilike", "cement"], ["default_code", "=", "CEM50"]]
    Always pass `fields` — omitting it returns every column and wastes context.
    """
    records = get_client().search_read(
        model, domain, fields, limit=limit, offset=offset, order=order
    )
    return dump(records)


@tool(read_only=True)
def odoo_count(model: str, domain: list[Any] | None = None) -> str:
    """Count matching records without fetching them. Cheap way to size a query."""
    return dump({"model": model, "count": get_client().search_count(model, domain)})


@tool(read_only=True)
def odoo_read(model: str, ids: list[int], fields: list[str] | None = None) -> str:
    """Read specific records by id, after another tool returned those ids.

    Pass `fields` to keep the response small; without it Odoo returns every
    column on the model.
    """
    return dump(get_client().read(model, iter_ids(ids), fields))


@tool(read_only=True)
def odoo_name_search(
    model: str, name: str, limit: int = 20, domain: list[Any] | None = None
) -> str:
    """Find records by display name and return [id, name] pairs.

    The reliable way to turn a name the user typed - in Arabic or English -
    into the record id that the other tools need.
    """
    return dump(get_client().name_search(model, name, limit=limit, domain=domain))


@tool(read_only=True)
def odoo_aggregate(
    model: str,
    groupby: list[str],
    aggregates: list[str] | None = None,
    domain: list[Any] | None = None,
    limit: int | None = None,
    order: str | None = None,
) -> str:
    """Group records and compute totals - use this for reporting questions.

    `groupby` takes field names, with an optional granularity for dates, e.g.
    ["date_order:month"] or ["partner_id"]. `aggregates` takes entries like
    "amount_total:sum", "price_unit:avg" or "__count" (the default).

    Far cheaper than reading every record and adding the numbers up yourself.
    """
    return dump(
        get_client().aggregate(
            model, groupby, aggregates, domain=domain, limit=limit, order=order
        )
    )


# -- writing ------------------------------------------------------------------


@tool()
def odoo_create(model: str, values: dict[str, Any]) -> str:
    """Create one record and return its new id.

    Check `odoo_fields` first, and resolve related records to ids with
    `odoo_name_search` - relational fields expect ids, not names.
    """
    created = get_client().create(model, values)
    ids = created if isinstance(created, list) else [created]
    return dump({"model": model, "created_ids": ids})


@tool()
def odoo_write(model: str, ids: list[int], values: dict[str, Any]) -> str:
    """Update fields on existing records.

    This overwrites the given fields on every id passed, so confirm the target
    ids with a read first.
    """
    record_ids = iter_ids(ids)
    result = get_client().write(model, record_ids, values)
    return dump({"model": model, "updated_ids": record_ids, "ok": bool(result)})


@tool(destructive=True)
def odoo_delete(model: str, ids: list[int]) -> str:
    """Permanently delete records. Disabled unless ODOO_ALLOW_DELETE=1.

    Prefer archiving (`odoo_write` with {"active": false}) - Odoo refuses to
    delete anything referenced by accounting or stock moves anyway.
    """
    record_ids = iter_ids(ids)
    result = get_client().unlink(model, record_ids)
    return dump({"model": model, "deleted_ids": record_ids, "ok": bool(result)})


@tool()
def odoo_call(
    model: str,
    method: str,
    args: list[Any] | None = None,
    kwargs: dict[str, Any] | None = None,
) -> str:
    """Call any model method - the escape hatch for business actions.

    Use it for workflow methods the dedicated tools do not cover, such as
    confirming a quotation:
        model="sale.order", method="action_confirm", args=[[42]]
    Anything that is not a known read method counts as a write.
    """
    return dump(get_client().call_method(model, method, args, kwargs))


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
