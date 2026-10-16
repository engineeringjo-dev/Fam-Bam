#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""كشف حساب مشروع — per-project statement of account.

Follows the paper form used at the store: goods sold, returns, payments,
running balance and aging for one project partner. The query shapes come from
the verified plan spec (panel step 22).

Usage:
    uv run python scripts/project_statement.py <partner_id> [date_from] [date_to]
    # dates as YYYY-MM-DD; omit for all time

Writes an RTL Arabic HTML file next to the console summary.
"""

from __future__ import annotations

import sys
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from scripts.check_connection import load_dotenv  # noqa: E402

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

from odoo_mcp.client import OdooClient  # noqa: E402
from odoo_mcp.config import OdooConfig  # noqa: E402


def date_domain(field: str, dfrom: str | None, dto: str | None) -> list:
    dom: list = []
    if dfrom:
        dom.append((field, ">=", dfrom))
    if dto:
        dom.append((field, "<=", dto))
    return dom


def fetch(c: OdooClient, pid: int, dfrom: str | None, dto: str | None) -> dict:
    partner = c.read("res.partner", [pid], ["name", "ref", "vat", "phone", "email"])[0]

    goods = c.search_read(
        "account.move.line",
        [("move_id.partner_id", "=", pid), ("move_id.move_type", "=", "out_invoice"),
         ("display_type", "=", "product"), ("parent_state", "=", "posted")]
        + date_domain("date", dfrom, dto),
        ["date", "move_name", "name", "quantity", "price_unit", "price_total"],
        limit=1000, order="date,id",
    )
    returns = c.search_read(
        "account.move.line",
        [("move_id.partner_id", "=", pid), ("move_id.move_type", "=", "out_refund"),
         ("display_type", "=", "product"), ("parent_state", "=", "posted")]
        + date_domain("date", dfrom, dto),
        ["date", "move_name", "name", "quantity", "price_unit", "price_total"],
        limit=1000, order="date,id",
    )
    # Point-of-sale tickets settled on account never become invoices, so their
    # goods detail lives only here — without this the statement shows a balance
    # with nothing behind it.
    pos_lines = c.search_read(
        "pos.order.line",
        [("order_id.partner_id", "=", pid), ("order_id.state", "in", ["paid", "done", "invoiced"])],
        ["order_id", "full_product_name", "qty", "price_unit", "price_subtotal_incl"],
        limit=1000, order="id",
    )
    pos_dates = {}
    order_ids = list({l["order_id"][0] for l in pos_lines})
    if order_ids:
        for o in c.read("pos.order", order_ids, ["date_order", "name", "account_move"]):
            pos_dates[o["id"]] = (o["date_order"][:10], o["name"])
    for line in pos_lines:
        oid = line["order_id"][0]
        stamp, ticket = pos_dates.get(oid, ("", line["order_id"][1]))
        row = {"date": stamp, "move_name": ticket, "name": line["full_product_name"],
               "quantity": abs(line["qty"]), "price_unit": line["price_unit"],
               "price_total": abs(line["price_subtotal_incl"])}
        if dfrom and stamp < dfrom:
            continue
        if dto and stamp > dto:
            continue
        (returns if line["qty"] < 0 else goods).append(row)
    goods.sort(key=lambda r: (r["date"], r["move_name"]))
    returns.sort(key=lambda r: (r["date"], r["move_name"]))

    payments = c.search_read(
        "account.payment",
        [("partner_id", "=", pid), ("state", "not in", ["draft", "canceled", "rejected"])]
        + date_domain("date", dfrom, dto),
        ["date", "name", "amount", "payment_type", "journal_id", "move_id"],
        limit=1000, order="date,id",
    )

    # Receivable ledger with a running balance (all time, so the balance is true).
    ledger = c.search_read(
        "account.move.line",
        [("partner_id", "=", pid), ("account_id.account_type", "=", "asset_receivable"),
         ("parent_state", "=", "posted")],
        ["date", "move_name", "move_id", "debit", "credit", "amount_residual"],
        limit=2000, order="date,id",
    )
    # كل سطر يروح لعموده: فاتورة، مرتجع، أو دفعة من المقاول — بدل ما تنحط
    # الدفعة بالسالب داخل عمود واحد.
    move_types = {}
    move_ids = list({l["move_id"][0] for l in ledger if l["move_id"]})
    for i in range(0, len(move_ids), 200):
        for mv in c.read("account.move", move_ids[i:i + 200], ["move_type"]):
            move_types[mv["id"]] = mv["move_type"]
    running = 0.0
    for line in ledger:
        mtype = move_types.get(line["move_id"][0] if line["move_id"] else 0, "entry")
        if not line["credit"]:
            line["kind"] = "invoice"
        elif mtype == "out_refund":
            line["kind"] = "return"
        else:
            line["kind"] = "payment"
        running += line["debit"] - line["credit"]
        line["balance_running"] = running

    today = date.today()
    aging = {"0-30": 0.0, "31-60": 0.0, "61-90": 0.0, "90+": 0.0}
    for line in ledger:
        residual = line.get("amount_residual") or 0.0
        if not residual:
            continue
        days = (today - date.fromisoformat(line["date"])).days
        key = "0-30" if days <= 30 else "31-60" if days <= 60 else "61-90" if days <= 90 else "90+"
        aging[key] += residual

    return {"partner": partner, "goods": goods, "returns": returns,
            "payments": payments, "ledger": ledger, "aging": aging, "balance": running}


def html_report(data: dict, dfrom: str | None, dto: str | None) -> str:
    p = data["partner"]
    period = f"{dfrom or 'البداية'} — {dto or 'اليوم'}"

    def rows_goods(lines, sign=1):
        out = []
        for l in lines:
            out.append(
                f"<tr><td>{l['date']}</td><td>{l['move_name']}</td><td>{l['name']}</td>"
                f"<td>{l['quantity']:g}</td><td>{l['price_unit']:.2f}</td>"
                f"<td>{sign * l['price_total']:.2f}</td></tr>"
            )
        return "\n".join(out) or "<tr><td colspan='6'>لا يوجد</td></tr>"

    pay_rows = []
    for l in data["payments"]:
        flag = "" if l["move_id"] else " ⚠ غير مرحّلة"
        pay_rows.append(
            f"<tr><td>{l['date']}</td><td>{l['name']}</td>"
            f"<td>{l['journal_id'][1] if l['journal_id'] else ''}</td>"
            f"<td>{l['amount']:.2f}{flag}</td></tr>"
        )
    ledger_rows = []
    led_inv = led_ret = led_pay = 0.0
    for l in data["ledger"]:
        kind = l.get("kind", "invoice")
        inv = f"{l['debit']:.2f}" if kind == "invoice" and l["debit"] else ""
        ret = f"{l['credit']:.2f}" if kind == "return" and l["credit"] else ""
        # الدفعة بالسالب وبخط عريض حتى تنفرز عن أسطر الفواتير
        pay = (f"<strong>-{l['credit']:.2f}</strong>"
               if kind == "payment" and l["credit"] else "")
        led_inv += l["debit"] if kind == "invoice" else 0.0
        led_ret += l["credit"] if kind == "return" else 0.0
        led_pay += l["credit"] if kind == "payment" else 0.0
        cls = ' class="pay"' if kind == "payment" else ""
        ledger_rows.append(
            f"<tr{cls}><td>{l['date']}</td><td>{l['move_name']}</td>"
            f"<td>{inv}</td><td>{ret}</td><td>{pay}</td>"
            f"<td>{l['balance_running']:.2f}</td></tr>"
        )
    aging = data["aging"]
    total_goods = sum(l["price_total"] for l in data["goods"])
    total_returns = sum(l["price_total"] for l in data["returns"])
    total_pay = sum(l["amount"] for l in data["payments"])

    return f"""<meta charset="utf-8">
<title>كشف حساب — {p['name']}</title>
<style>
 @page {{ size: A4; margin: 14mm; }}
 body {{ font-family: "Noto Naskh Arabic","Noto Sans Arabic","Segoe UI",Tahoma,sans-serif;
         direction: rtl; margin: 0; color: #222; font-size: 12px; }}
 table {{ font-variant-numeric: tabular-nums; }}
 thead {{ display: table-header-group; }}
 tr {{ break-inside: avoid; }}
 h1 {{ font-size: 1.4em; }} h2 {{ font-size: 1.1em; margin-top: 1.4em; }}
 table {{ border-collapse: collapse; width: 100%; font-size: 0.9em; }}
 th, td {{ border: 1px solid #999; padding: 4px 8px; text-align: right; }}
 th {{ background: #eee; }}
 .totals td {{ font-weight: bold; background: #f6f6f6; }}
 .pay td {{ background: #eef7ee; }}
</style>
<h1>كشف حساب مشروع — {p['name']}</h1>
<p>محلات العون لمواد البناء · الفترة: {period}
 · هاتف: {p.get('phone') or '-'} · مرجع: {p.get('ref') or '-'}</p>

<h2>١. البضائع المباعة</h2>
<table><tr><th>التاريخ</th><th>الفاتورة</th><th>الصنف</th><th>الكمية</th><th>الإفرادي</th><th>الإجمالي</th></tr>
{rows_goods(data['goods'])}
<tr class="totals"><td colspan="5">مجموع المبيعات</td><td>{total_goods:.2f}</td></tr></table>

<h2>٢. المرتجعات</h2>
<table><tr><th>التاريخ</th><th>الإشعار</th><th>الصنف</th><th>الكمية</th><th>الإفرادي</th><th>الإجمالي</th></tr>
{rows_goods(data['returns'], sign=-1)}
<tr class="totals"><td colspan="5">مجموع المرتجعات</td><td>-{total_returns:.2f}</td></tr></table>

<h2>٣. الدفعات</h2>
<table><tr><th>التاريخ</th><th>السند</th><th>الصندوق/البنك</th><th>المبلغ</th></tr>
{''.join(pay_rows) or "<tr><td colspan='4'>لا يوجد</td></tr>"}
<tr class="totals"><td colspan="3">مجموع الدفعات</td><td>{total_pay:.2f}</td></tr></table>

<h2>٤. الحركة والرصيد الجاري</h2>
<table><tr><th>التاريخ</th><th>المستند</th><th>فواتير (عليه)</th><th>مرتجعات</th>
<th>دفعات المقاول</th><th>الرصيد</th></tr>
{''.join(ledger_rows) or "<tr><td colspan='6'>لا يوجد</td></tr>"}
<tr class="totals"><td colspan="2">المجموع</td><td>{led_inv:.2f}</td><td>{led_ret:.2f}</td>
<td><strong>-{led_pay:.2f}</strong></td><td>{data['balance']:.2f}</td></tr></table>

<h2>٥. أعمار الذمم (على المتبقي)</h2>
<table><tr><th>٠-٣٠ يوم</th><th>٣١-٦٠</th><th>٦١-٩٠</th><th>أكثر من ٩٠</th><th>الرصيد المستحق</th></tr>
<tr><td>{aging['0-30']:.2f}</td><td>{aging['31-60']:.2f}</td><td>{aging['61-90']:.2f}</td>
<td>{aging['90+']:.2f}</td><td>{data['balance']:.2f}</td></tr></table>
"""


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    pid = int(sys.argv[1])
    dfrom = sys.argv[2] if len(sys.argv) > 2 else None
    dto = sys.argv[3] if len(sys.argv) > 3 else None

    client = OdooClient(OdooConfig.from_env())
    data = fetch(client, pid, dfrom, dto)
    p = data["partner"]

    print(f"كشف حساب: {p['name']}  (partner {pid})")
    print(f"  بضائع مباعة : {sum(l['price_total'] for l in data['goods']):>10.2f}  ({len(data['goods'])} بند)")
    print(f"  مرتجعات     : {sum(l['price_total'] for l in data['returns']):>10.2f}  ({len(data['returns'])} بند)")
    print(f"  دفعات       : {sum(l['amount'] for l in data['payments']):>10.2f}  ({len(data['payments'])} سند)")
    unposted = [l for l in data["payments"] if not l["move_id"]]
    if unposted:
        print(f"  ⚠ دفعات غير مرحّلة: {len(unposted)} — لا تُحتسب في الرصيد الدفتري")
    print(f"  الرصيد المستحق: {data['balance']:>10.2f} د.أ")

    out = Path(f"statement_{pid}.html")
    out.write_text(html_report(data, dfrom, dto), encoding="utf-8")
    print(f"\nHTML: {out.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
