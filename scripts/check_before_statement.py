# -*- coding: utf-8 -*-
"""فحص إلزامي قبل إصدار أي كشف حساب أو تقرير لأي جهة.

قاعدة صاحب المحل (19/09): ممنوع إعطاء كشف حساب لأي مشروع فيه بند بسعر صفر
بدون تنبيه صريح. هذا الفحص يُشغَّل قبل كل كشف، ويُوقف الإصدار إذا طلع شي.

الاستعمال:
    python3 scripts/check_before_statement.py            # كل الجهات
    python3 scripts/check_before_statement.py 19         # جهة واحدة
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'odoo_templates'))
from jrpc import x
from collections import defaultdict

# حالات معروفة ومعتمدة من صاحب المحل — تُعرض كملاحظة لا كمانع
ALLOW = {
    'BILL/2026/08/0001': 'رصيد مدوّر لكتانة — سطر بلا صنف بشكل مقصود',
}

SALES = ['out_invoice', 'out_refund']
BILLS = ['in_invoice', 'in_refund']


def _lines(dom, fields):
    return x('account.move.line', 'search_read', dom, fields, order='date')


def scan(partner_id=None):
    pdom = [('partner_id', '=', partner_id)] if partner_id else []
    findings = defaultdict(list)

    # ١) بنود بسعر صفر على فواتير مرحّلة
    for l in _lines([('display_type', '=', 'product'), ('parent_state', '=', 'posted'),
                     ('move_id.move_type', 'in', SALES + BILLS),
                     ('price_unit', '=', 0)] + pdom,
                    ['move_id', 'product_id', 'quantity', 'partner_id', 'date']):
        findings[l['partner_id'][1]].append(
            ('سعر صفر', l['date'], l['move_id'][1],
             l['product_id'][1] if l['product_id'] else '(بلا صنف)', l['quantity']))

    # ٢) بنود بلا صنف مربوط (ما بتظهر بتقارير المخزون ولا بتحديث الكلفة)
    for l in _lines([('display_type', '=', 'product'), ('parent_state', '=', 'posted'),
                     ('move_id.move_type', 'in', SALES + BILLS),
                     ('product_id', '=', False)] + pdom,
                    ['move_id', 'name', 'quantity', 'partner_id', 'date', 'price_unit']):
        findings[l['partner_id'][1]].append(
            ('بلا صنف', l['date'], l['move_id'][1], (l['name'] or '')[:40], l['quantity']))

    # ٣) مسودات معلّقة — ما بتدخل بالكشف فبيطلع الرصيد ناقص
    for m in x('account.move', 'search_read',
               [('state', '=', 'draft'), ('move_type', 'in', SALES + BILLS)] + pdom,
               ['name', 'partner_id', 'invoice_date', 'amount_total'], order='invoice_date'):
        if not m['partner_id']:
            continue
        findings[m['partner_id'][1]].append(
            ('مسودة معلّقة', m['invoice_date'], m['name'] or '(بلا رقم)',
             'إجمالي %.3f' % m['amount_total'], ''))
    return findings


def main():
    pid = int(sys.argv[1]) if len(sys.argv) > 1 else None
    f = scan(pid)
    blocking, known = {}, []
    for partner, rows in f.items():
        keep = []
        for r in rows:
            hit = next((k for k in ALLOW if k in (r[2] or '')), None)
            if hit:
                known.append((partner, hit, ALLOW[hit]))
            else:
                keep.append(r)
        if keep:
            blocking[partner] = keep
    if known:
        print('ℹ️  حالات معروفة ومعتمدة (مش مانع):')
        for partner, doc, why in known:
            print('   %-28s %-18s %s' % (partner, doc, why))
        print()
    if not blocking:
        print('✅ نظيف — ما في مانع من إصدار الكشف.')
        return 0
    print('⛔ توقّف — لازم تنبيه صاحب المحل قبل إصدار أي كشف:\n')
    for partner in sorted(blocking):
        print('■ %s' % partner)
        for kind, date, doc, what, qty in blocking[partner]:
            print('   [%-12s] %-11s %-18s %-42s %s' % (kind, date, doc, what, qty))
        print()
    return 1


if __name__ == '__main__':
    sys.exit(main())
