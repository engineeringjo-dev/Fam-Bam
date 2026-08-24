# -*- coding: utf-8 -*-
"""Print the shop's paper order book (نموذج الطلبية) from live Odoo data.

The shop runs projects off a spiral notebook whose columns are
    #  ·  الطلبية  ·  مطلوب  ·  خارج  ·  متبقي  ·  التاريخ  ·  مرتجع  ·  الفاتورة
and whose one rule is that what a project actually took is (خارج − مرتجع).

Odoo carries every one of those numbers natively on sale.order.line, but its
printed quotation is a QWeb template — editing it would count as customisation
and put the database back on the Custom plan. So the sheet is rendered here
instead, straight from the same fields, and nothing in Odoo is touched.

    python scripts/order_sheet.py "مشروع المدرسة" [out.html]
"""
import sys
from collections import defaultdict
from html import escape
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from scripts.check_connection import load_dotenv

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

from odoo_mcp.client import OdooClient      # noqa: E402
from odoo_mcp.config import OdooConfig      # noqa: E402

OPEN_STATES = ("draft", "sent", "sale")


def collect(c, project):
    partners = c.search_read(
        "res.partner", ["|", ("name", "=", project), ("name", "ilike", project)],
        ["name"], limit=5)
    if not partners:
        raise SystemExit(f"ما لقينا شريك اسمه «{project}»")
    pid = partners[0]["id"]

    orders = c.search_read(
        "sale.order", [("partner_id", "=", pid), ("state", "in", OPEN_STATES)],
        ["name", "state", "date_order", "amount_total"], limit=200, order="date_order,id")
    if not orders:
        return partners[0], [], []

    lines = c.search_read(
        "sale.order.line", [("order_id", "in", [o["id"] for o in orders])],
        ["order_id", "name", "product_id", "product_uom_qty", "qty_delivered",
         "qty_to_deliver", "qty_invoiced", "price_unit", "price_subtotal"],
        limit=2000, order="order_id,sequence,id")

    # qty_delivered is already net of returns, so a return has to be read off the
    # incoming stock moves or the notebook's مرتجع column stays empty.
    returned, moved_on = defaultdict(float), {}
    if lines:
        moves = c.search_read(
            "stock.move",
            [("sale_line_id", "in", [l["id"] for l in lines]), ("state", "=", "done")],
            ["sale_line_id", "quantity", "date", "location_id", "location_dest_id",
             "picking_code"], limit=4000, order="date")
        for m in moves:
            lid = m["sale_line_id"][0]
            if m.get("picking_code") == "incoming" or "Customers" in str(m["location_id"]):
                returned[lid] += m["quantity"]
            else:
                moved_on[lid] = m["date"][:10]
    return partners[0], orders, [dict(l, _returned=returned.get(l["id"], 0.0),
                                      _date=moved_on.get(l["id"], "")) for l in lines]


def clean(name):
    """Odoo prefixes the line description with [code]; the notebook never did."""
    return name.split("]", 1)[1].strip() if name.startswith("[") and "]" in name else name.strip()


def num(v, dash="—"):
    if not v:
        return dash
    return f"{v:,.0f}" if float(v).is_integer() else f"{v:,.2f}"


def render(partner, orders, lines):
    rows, totals = [], defaultdict(float)
    for i, l in enumerate(lines, start=1):
        net = l["qty_delivered"] - l["_returned"]
        for k, v in (("req", l["product_uom_qty"]), ("out", l["qty_delivered"]),
                     ("left", l["qty_to_deliver"]), ("ret", l["_returned"]),
                     ("inv", l["qty_invoiced"]), ("net", net),
                     ("val", l["price_subtotal"])):
            totals[k] += v
        rows.append(f"""<tr>
    <td class="n">{i}</td>
    <td class="d">{escape(clean(l['name']))}</td>
    <td class="q">{num(l['product_uom_qty'])}</td>
    <td class="q out">{num(l['qty_delivered'])}</td>
    <td class="q{' owed' if l['qty_to_deliver'] else ''}">{num(l['qty_to_deliver'])}</td>
    <td class="t">{escape(l['_date']) or '—'}</td>
    <td class="q{' ret' if l['_returned'] else ''}">{num(l['_returned'])}</td>
    <td class="q">{num(l['qty_invoiced'])}</td>
  </tr>""")

    refs = " · ".join(f"{o['name']}" for o in orders) or "—"
    span = f"{orders[0]['date_order'][:10]} ← {orders[-1]['date_order'][:10]}" if orders else "—"

    return f"""<!doctype html>
<html lang="ar" dir="rtl"><head><meta charset="utf-8">
<title>نموذج الطلبية — {escape(partner['name'])}</title>
<style>
  @page {{ size: A4; margin: 12mm; }}
  :root {{ --ink:#111; --line:#111; --soft:#6b7280; --tint:#f4f1ea;
           --owed:#fdf0e3; --ret:#fdeaea; --out:#eef4ee; }}
  * {{ box-sizing:border-box; }}
  body {{ font-family:"Tahoma","Arial",sans-serif; color:var(--ink);
          margin:0; padding:6mm; background:#fff; font-size:11.5px; }}
  .head {{ display:flex; justify-content:space-between; align-items:flex-start;
           border:2px solid var(--line); padding:6px 10px; margin-bottom:8px; }}
  .head .name {{ font-size:20px; font-weight:bold; }}
  .head .meta {{ font-size:10.5px; color:var(--soft); line-height:1.7; text-align:left; }}
  .shop {{ font-size:12px; font-weight:bold; margin-bottom:4px; }}
  table {{ width:100%; border-collapse:collapse; }}
  th, td {{ border:1px solid var(--line); padding:4px 5px; vertical-align:middle; }}
  thead th {{ background:var(--tint); font-weight:bold; text-align:center; font-size:11px; }}
  td.n {{ text-align:center; width:26px; color:var(--soft); }}
  td.d {{ text-align:right; }}
  td.q {{ text-align:center; width:52px; font-variant-numeric:tabular-nums; }}
  td.t {{ text-align:center; width:74px; font-size:10px; color:var(--soft); }}
  td.out {{ background:var(--out); }}
  td.owed {{ background:var(--owed); font-weight:bold; }}
  td.ret {{ background:var(--ret); font-weight:bold; }}
  tfoot td {{ background:var(--tint); font-weight:bold; text-align:center; }}
  tfoot td.d {{ text-align:left; }}
  .rule {{ margin-top:8px; font-size:10.5px; color:var(--soft);
           border-right:3px solid var(--line); padding:3px 8px; }}
  .sign {{ margin-top:14mm; display:flex; gap:14mm; font-size:11px; }}
  .sign div {{ flex:1; border-top:1px solid var(--line); padding-top:4px; text-align:center; }}
  @media print {{ body {{ padding:0; }} }}
</style></head><body>

<div class="head">
  <div>
    <div class="shop">محلات العون لمواد البناء</div>
    <div>اسم المشروع: <span class="name">{escape(partner['name'])}</span></div>
  </div>
  <div class="meta">
    الطلبيات: {escape(refs)}<br>الفترة: {escape(span)}<br>عدد البنود: {len(lines)}
  </div>
</div>

<table>
  <thead><tr>
    <th>#</th><th>الطلبية</th><th>مطلوب</th><th>خارج</th>
    <th>متبقي</th><th>التاريخ</th><th>مرتجع</th><th>الفاتورة</th>
  </tr></thead>
  <tbody>
{chr(10).join(rows) if rows else '  <tr><td colspan="8" style="text-align:center;padding:14px">لا توجد بنود</td></tr>'}
  </tbody>
  <tfoot><tr>
    <td colspan="2" class="d">الإجمالي</td>
    <td>{num(totals['req'])}</td><td>{num(totals['out'])}</td>
    <td>{num(totals['left'])}</td><td>—</td>
    <td>{num(totals['ret'])}</td><td>{num(totals['inv'])}</td>
  </tr></tfoot>
</table>

<div class="rule">
  صافي المستلم = (خارج − مرتجع) = <b>{num(totals['net'], '0')}</b> &nbsp;·&nbsp;
  قيمة الطلبية: <b>{totals['val']:,.2f}</b> د.ا &nbsp;·&nbsp;
  متبقي للتسليم: <b>{num(totals['left'], '0')}</b>
</div>

<div class="sign"><div>المستلم</div><div>أمين المستودع</div><div>المحاسب</div></div>
</body></html>"""


def main():
    project = sys.argv[1] if len(sys.argv) > 1 else "مشروع المدرسة"
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(f"طلبية_{project}.html")
    c = OdooClient(OdooConfig.from_env())
    partner, orders, lines = collect(c, project)
    out.write_text(render(partner, orders, lines), encoding="utf-8")
    print(f"المشروع : {partner['name']}")
    print(f"الطلبيات: {[o['name'] for o in orders]}")
    print(f"البنود  : {len(lines)}")
    for l in lines:
        print(f"   {clean(l['name'])[:34]:<36} مطلوب={l['product_uom_qty']:>5.0f} "
              f"خارج={l['qty_delivered']:>5.0f} متبقي={l['qty_to_deliver']:>5.0f} "
              f"مرتجع={l['_returned']:>5.0f} فاتورة={l['qty_invoiced']:>5.0f}")
    print(f"\n✔ {out}")


if __name__ == "__main__":
    main()
