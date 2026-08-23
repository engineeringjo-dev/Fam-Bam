"""Thin XML-RPC client for the Odoo external API (Odoo 15 through 19)."""

from __future__ import annotations

import http.client
import socket
import xmlrpc.client
from typing import Any, Iterable, Sequence

from .config import OdooConfig

# Methods that only read. Anything not listed here is treated as a write when it
# goes through `call_method`, so the write/delete switches still apply.
READ_ONLY_METHODS = frozenset(
    {
        "check_access_rights",
        "default_get",
        "exists",
        "fields_get",
        "formatted_read_group",
        "get_view",
        "name_get",
        "name_search",
        "read",
        "read_group",
        "search",
        "search_count",
        "search_read",
        "web_read_group",
        "web_search_read",
    }
)


class OdooError(RuntimeError):
    """Base class for every failure raised by this client."""


class OdooAuthError(OdooError):
    """Authentication against the Odoo database failed."""


class OdooPermissionError(OdooError):
    """The operation is blocked by this server's own write/delete switches."""


class OdooCallError(OdooError):
    """Odoo accepted the request but the call itself failed."""


class _TimeoutTransport(xmlrpc.client.SafeTransport):
    """SafeTransport that applies a socket timeout to each connection."""

    def __init__(self, timeout: int) -> None:
        super().__init__()
        self._timeout = timeout

    def make_connection(self, host):  # type: ignore[no-untyped-def]
        connection = super().make_connection(host)
        connection.timeout = self._timeout
        return connection


class _PlainTimeoutTransport(xmlrpc.client.Transport):
    def __init__(self, timeout: int) -> None:
        super().__init__()
        self._timeout = timeout

    def make_connection(self, host):  # type: ignore[no-untyped-def]
        connection = super().make_connection(host)
        connection.timeout = self._timeout
        return connection


def _clean_fault(fault: xmlrpc.client.Fault) -> str:
    """Turn an Odoo server traceback into the one line that actually matters."""
    text = (fault.faultString or "").strip()
    lines = [line for line in text.splitlines() if line.strip()]
    if not lines:
        return f"Odoo fault {fault.faultCode}"
    tail = lines[-1].strip()
    # Server tracebacks end with "odoo.exceptions.UserError: the real message".
    if ": " in tail and tail.split(": ", 1)[0].count(".") >= 1:
        tail = tail.split(": ", 1)[1].strip() or tail
    return tail


class OdooClient:
    """Authenticates once, then forwards calls to Odoo's `execute_kw`."""

    def __init__(self, config: OdooConfig) -> None:
        self.config = config
        self._uid: int | None = None
        self._common = self._proxy("/xmlrpc/2/common")
        self._models = self._proxy("/xmlrpc/2/object")

    def _proxy(self, path: str) -> xmlrpc.client.ServerProxy:
        transport: xmlrpc.client.Transport
        if self.config.url.startswith("https://"):
            transport = _TimeoutTransport(self.config.timeout)
        else:
            transport = _PlainTimeoutTransport(self.config.timeout)
        return xmlrpc.client.ServerProxy(
            self.config.url + path, transport=transport, allow_none=True
        )

    # -- connection ----------------------------------------------------------

    def version(self) -> dict:
        try:
            return self._common.version()
        except (OSError, socket.timeout, http.client.HTTPException) as exc:
            raise OdooError(
                f"Cannot reach {self.config.url}: {type(exc).__name__}: {exc}"
            ) from exc

    @property
    def uid(self) -> int:
        if self._uid is None:
            self._uid = self._authenticate()
        return self._uid

    def _authenticate(self) -> int:
        try:
            uid = self._common.authenticate(
                self.config.db, self.config.username, self.config.api_key, {}
            )
        except xmlrpc.client.Fault as fault:
            raise OdooAuthError(
                f"Odoo rejected the login for database {self.config.db!r}: {_clean_fault(fault)}"
            ) from fault
        except (OSError, socket.timeout, http.client.HTTPException) as exc:
            raise OdooError(
                f"Cannot reach {self.config.url}: {type(exc).__name__}: {exc}"
            ) from exc

        if not uid:
            raise OdooAuthError(
                f"Login failed for {self.config.username!r} on database "
                f"{self.config.db!r}. Check ODOO_USERNAME and ODOO_API_KEY "
                "(the key must come from My Profile > Account Security > New API Key)."
            )
        return int(uid)

    # -- guards --------------------------------------------------------------

    def _require_write(self, what: str) -> None:
        if not self.config.allow_write:
            raise OdooPermissionError(
                f"{what} is blocked: this server runs read-only. "
                "Set ODOO_ALLOW_WRITE=1 to enable writing."
            )

    def _require_delete(self) -> None:
        self._require_write("Deleting records")
        if not self.config.allow_delete:
            raise OdooPermissionError(
                "Deleting records is blocked. Set ODOO_ALLOW_DELETE=1 to enable it. "
                "Most Odoo records should be archived (active=false) instead of deleted."
            )

    # -- core call -----------------------------------------------------------

    def execute_kw(
        self,
        model: str,
        method: str,
        args: Sequence[Any] | None = None,
        kwargs: dict | None = None,
    ) -> Any:
        payload_kwargs = dict(kwargs or {})
        context = {**self.config.base_context, **(payload_kwargs.get("context") or {})}
        if context:
            payload_kwargs["context"] = context

        try:
            return self._models.execute_kw(
                self.config.db,
                self.uid,
                self.config.api_key,
                model,
                method,
                list(args or []),
                payload_kwargs,
            )
        except xmlrpc.client.Fault as fault:
            raise OdooCallError(
                f"{model}.{method} failed: {_clean_fault(fault)}"
            ) from fault
        except (OSError, socket.timeout, http.client.HTTPException) as exc:
            raise OdooError(
                f"{model}.{method} could not be sent: {type(exc).__name__}: {exc}"
            ) from exc

    # -- reads ---------------------------------------------------------------

    def _capped(self, limit: int | None) -> int:
        if not limit or limit <= 0:
            return self.config.default_limit
        return min(limit, self.config.max_limit)

    def search_read(
        self,
        model: str,
        domain: list | None = None,
        fields: Sequence[str] | None = None,
        limit: int | None = None,
        offset: int = 0,
        order: str | None = None,
        context: dict | None = None,
    ) -> list[dict]:
        kwargs: dict[str, Any] = {"limit": self._capped(limit), "offset": offset}
        if fields:
            kwargs["fields"] = list(fields)
        if order:
            kwargs["order"] = order
        if context:
            kwargs["context"] = context
        return self.execute_kw(model, "search_read", [domain or []], kwargs)

    def read(
        self,
        model: str,
        ids: Sequence[int],
        fields: Sequence[str] | None = None,
        context: dict | None = None,
    ) -> list[dict]:
        kwargs: dict[str, Any] = {}
        if fields:
            kwargs["fields"] = list(fields)
        if context:
            kwargs["context"] = context
        return self.execute_kw(model, "read", [list(ids)], kwargs)

    def search_count(self, model: str, domain: list | None = None) -> int:
        return self.execute_kw(model, "search_count", [domain or []])

    def name_search(
        self, model: str, name: str, limit: int | None = None, domain: list | None = None
    ) -> list:
        kwargs: dict[str, Any] = {"name": name, "limit": self._capped(limit)}
        if domain:
            kwargs["args"] = domain
        return self.execute_kw(model, "name_search", [], kwargs)

    def fields_get(self, model: str, attributes: Sequence[str] | None = None) -> dict:
        return self.execute_kw(
            model,
            "fields_get",
            [],
            {"attributes": list(attributes or ["string", "type", "required", "relation", "selection", "readonly", "help"])},
        )

    def aggregate(
        self,
        model: str,
        groupby: Sequence[str],
        aggregates: Sequence[str] | None = None,
        domain: list | None = None,
        limit: int | None = None,
        order: str | None = None,
    ) -> list:
        """Grouped totals, across Odoo versions.

        Odoo 18+ exposes `formatted_read_group`; older releases only have
        `read_group`. Try the modern one first and fall back on the legacy call.
        """
        wanted = list(aggregates or ["__count"])
        try:
            kwargs: dict[str, Any] = {
                "domain": domain or [],
                "groupby": list(groupby),
                "aggregates": wanted,
                "limit": self._capped(limit),
            }
            if order:
                kwargs["order"] = order
            return self.execute_kw(model, "formatted_read_group", [], kwargs)
        except OdooCallError as modern_error:
            legacy_fields = [item for item in wanted if item != "__count"]
            try:
                kwargs = {
                    "domain": domain or [],
                    "fields": legacy_fields,
                    "groupby": list(groupby),
                    "limit": self._capped(limit),
                    "lazy": False,
                }
                if order:
                    kwargs["orderby"] = order
                return self.execute_kw(model, "read_group", [], kwargs)
            except OdooCallError:
                raise modern_error

    # -- writes --------------------------------------------------------------

    def create(self, model: str, values: dict | list[dict]) -> Any:
        self._require_write("Creating records")
        payload = values if isinstance(values, list) else [values]
        return self.execute_kw(model, "create", [payload])

    def write(self, model: str, ids: Sequence[int], values: dict) -> bool:
        self._require_write("Updating records")
        if not ids:
            raise OdooCallError("write needs at least one record id.")
        return self.execute_kw(model, "write", [list(ids), values])

    def unlink(self, model: str, ids: Sequence[int]) -> bool:
        self._require_delete()
        if not ids:
            raise OdooCallError("unlink needs at least one record id.")
        return self.execute_kw(model, "unlink", [list(ids)])

    def call_method(
        self,
        model: str,
        method: str,
        args: Sequence[Any] | None = None,
        kwargs: dict | None = None,
    ) -> Any:
        if method not in READ_ONLY_METHODS:
            self._require_write(f"Calling {model}.{method}")
        if method == "unlink":
            self._require_delete()
        return self.execute_kw(model, method, args, kwargs)


def iter_ids(value: Iterable[Any]) -> list[int]:
    """Coerce a loose list of ids into ints, rejecting anything else."""
    result = []
    for item in value:
        try:
            result.append(int(item))
        except (TypeError, ValueError) as exc:
            raise OdooCallError(f"{item!r} is not a valid record id.") from exc
    return result
