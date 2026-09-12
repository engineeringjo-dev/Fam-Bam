"""Smoke tests: configuration, guardrails and error reporting.

Everything here runs offline except the tests guarded by ODOO_SMOKE_URL, which
authenticate against a real server with deliberately wrong credentials to check
that the failure message is useful.
"""

from __future__ import annotations

import os

import pytest

from odoo_mcp import server
from odoo_mcp.client import OdooClient, OdooPermissionError, iter_ids
from odoo_mcp.config import OdooConfig, OdooConfigError

ENV_KEYS = (
    "ODOO_URL",
    "ODOO_DB",
    "ODOO_USERNAME",
    "ODOO_API_KEY",
    "ODOO_ALLOW_WRITE",
    "ODOO_ALLOW_DELETE",
    "ODOO_ALLOW_INSECURE",
    "ODOO_LANG",
)


@pytest.fixture(autouse=True)
def clean_env(monkeypatch):
    for key in ENV_KEYS:
        monkeypatch.delenv(key, raising=False)
    server._client = None
    yield
    server._client = None


def configure(monkeypatch, **overrides):
    values = {
        "ODOO_URL": "https://example.odoo.com",
        "ODOO_DB": "example",
        "ODOO_USERNAME": "user@example.com",
        "ODOO_API_KEY": "not-a-real-key",
    }
    values.update(overrides)
    for key, value in values.items():
        monkeypatch.setenv(key, value)


def test_missing_config_names_the_variables(monkeypatch):
    with pytest.raises(OdooConfigError) as excinfo:
        OdooConfig.from_env()
    message = str(excinfo.value)
    for key in ("ODOO_URL", "ODOO_DB", "ODOO_USERNAME", "ODOO_API_KEY"):
        assert key in message


def test_url_gets_https_scheme_and_loses_trailing_slash(monkeypatch):
    configure(monkeypatch, ODOO_URL="alawn.odoo.com/")
    assert OdooConfig.from_env().url == "https://alawn.odoo.com"


def test_plain_http_is_refused_unless_explicitly_allowed(monkeypatch):
    configure(monkeypatch, ODOO_URL="http://intranet.local")
    with pytest.raises(OdooConfigError, match="plain HTTP"):
        OdooConfig.from_env()

    configure(monkeypatch, ODOO_URL="http://intranet.local", ODOO_ALLOW_INSECURE="1")
    assert OdooConfig.from_env().url == "http://intranet.local"


def test_limits_are_capped_at_max(monkeypatch):
    configure(monkeypatch)
    monkeypatch.setenv("ODOO_MAX_LIMIT", "50")
    client = OdooClient(OdooConfig.from_env())
    assert client._capped(10_000) == 50
    assert client._capped(None) == client.config.default_limit


def test_writes_blocked_in_read_only_mode(monkeypatch):
    configure(monkeypatch, ODOO_ALLOW_WRITE="0")
    client = OdooClient(OdooConfig.from_env())
    for call in (
        lambda: client.create("res.partner", {"name": "x"}),
        lambda: client.write("res.partner", [1], {"name": "x"}),
        lambda: client.unlink("res.partner", [1]),
        lambda: client.call_method("sale.order", "action_confirm", [[1]]),
    ):
        with pytest.raises(OdooPermissionError):
            call()


def test_delete_blocked_even_when_writes_are_allowed(monkeypatch):
    configure(monkeypatch)
    client = OdooClient(OdooConfig.from_env())
    assert client.config.allow_write is True
    with pytest.raises(OdooPermissionError, match="ODOO_ALLOW_DELETE"):
        client.unlink("res.partner", [1])
    with pytest.raises(OdooPermissionError, match="ODOO_ALLOW_DELETE"):
        client.call_method("res.partner", "unlink", [[1]])


def test_read_methods_pass_the_write_guard(monkeypatch):
    """A read through call_method must not be rejected by the write switch."""
    configure(monkeypatch, ODOO_ALLOW_WRITE="0")
    client = OdooClient(OdooConfig.from_env())
    sent = {}

    def fake_execute(model, method, args=None, kwargs=None):
        sent["call"] = (model, method)
        return []

    client.execute_kw = fake_execute  # type: ignore[assignment]
    client.call_method("res.partner", "search_read", [[]])
    assert sent["call"] == ("res.partner", "search_read")


def test_language_context_is_merged_into_every_call(monkeypatch):
    configure(monkeypatch, ODOO_LANG="ar_001")
    client = OdooClient(OdooConfig.from_env())
    captured = {}

    class FakeModels:
        def execute_kw(self, db, uid, key, model, method, args, kwargs):
            captured.update(kwargs)
            return []

    client._uid = 2
    client._models = FakeModels()
    client.search_read("res.partner", [], ["name"], context={"active_test": False})
    assert captured["context"] == {"lang": "ar_001", "active_test": False}


def test_dump_keeps_arabic_readable():
    text = server.dump({"name": "محلات العون لمواد البناء"})
    assert "محلات العون لمواد البناء" in text
    assert "\\u" not in text


def test_dump_truncates_oversized_payloads(monkeypatch):
    monkeypatch.setattr(server, "MAX_CHARS", 200)
    text = server.dump([{"note": "x" * 50} for _ in range(50)])
    assert "truncated at 200 characters" in text


def test_iter_ids_rejects_junk():
    assert iter_ids(["3", 4]) == [3, 4]
    with pytest.raises(Exception):
        iter_ids(["not-an-id"])


def test_tools_report_config_errors_as_text(monkeypatch):
    """An unconfigured server answers with ERROR: text, it does not crash."""
    result = server.odoo_status()
    assert result.startswith("ERROR:")
    assert "ODOO_URL" in result


# -- live checks against a real server (no valid credentials needed) ----------

live = pytest.mark.skipif(
    not os.environ.get("ODOO_SMOKE_URL"),
    reason="set ODOO_SMOKE_URL to run checks against a real Odoo server",
)


@live
def test_live_server_reports_its_version(monkeypatch):
    configure(monkeypatch, ODOO_URL=os.environ["ODOO_SMOKE_URL"])
    version = OdooClient(OdooConfig.from_env()).version()
    assert "server_version" in version


@live
def test_live_bad_credentials_give_an_actionable_message(monkeypatch):
    configure(
        monkeypatch,
        ODOO_URL=os.environ["ODOO_SMOKE_URL"],
        ODOO_DB=os.environ.get("ODOO_SMOKE_DB", "example"),
    )
    server._client = None
    result = server.odoo_status()
    assert result.startswith("ERROR:")
    assert "ODOO_API_KEY" in result or "database" in result.lower()


# -- error message shaping ----------------------------------------------------


def make_fault(text: str):
    import xmlrpc.client

    return xmlrpc.client.Fault(1, text)


def test_fault_strips_the_exception_class_name():
    from odoo_mcp.client import _clean_fault

    fault = make_fault(
        "Traceback (most recent call last):\n"
        "  File ...\n"
        "odoo.exceptions.UserError: You cannot delete a posted entry."
    )
    assert _clean_fault(fault) == "You cannot delete a posted entry."

    plain = make_fault("ValueError: something went wrong")
    assert _clean_fault(plain) == "something went wrong"


def test_invalid_field_error_points_at_fields_get():
    from odoo_mcp.client import _clean_fault

    fault = make_fault("ValueError: Invalid field 'mobile' on 'res.partner'")
    message = _clean_fault(fault)
    assert message.startswith("Invalid field 'mobile' on 'res.partner'")
    assert "fields_get" in message and "res.partner" in message


def test_empty_fault_still_produces_a_message():
    from odoo_mcp.client import _clean_fault

    assert "Odoo fault" in _clean_fault(make_fault(""))


def test_void_method_fault_is_treated_as_success(monkeypatch):
    """Odoo commits before marshalling; a None return faults after the work is done."""
    import xmlrpc.client

    from odoo_mcp.client import OdooClient
    from odoo_mcp.config import OdooConfig

    configure(monkeypatch)
    client = OdooClient(OdooConfig.from_env())
    client._uid = 2

    class FakeModels:
        def execute_kw(self, *args):
            raise xmlrpc.client.Fault(1, "cannot marshal None unless allow_none is enabled")

    client._models = FakeModels()
    assert client.execute_kw("account.move", "button_cancel", [[1]]) is None
