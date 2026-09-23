# -*- coding: utf-8 -*-
"""مدير الكشوفات — محلات العون لمواد البناء

البوّابة الوحيدة لإصدار كشوف الحساب. أي كشف بيمرّ من هون أولاً:
المدير بفحص الحساب، وإذا طلع شي بيوقف الإصدار ويقول شو المطلوب.
ما بيطلع كشف إلا بعد ما يعتمده.

  python3 scripts/مدير_الكشوفات.py                    # فحص كل الجهات
  python3 scripts/مدير_الكشوفات.py 19                  # فحص جهة
  python3 scripts/مدير_الكشوفات.py 19 --اصدار          # فحص ثم إصدار PDF
  python3 scripts/مدير_الكشوفات.py "مشروع المدرسة" --اصدار

قاعدة صاحب المحل (19/09): ممنوع تسليم كشف حساب فيه بند بسعر صفر
أو بند بلا صنف أو مسودة معلّقة، بدون تنبيه صريح ومعالجة.
وزيادة (23/09): يفحص كمان إذا الفاتورة **اندبلت وانرحّلت مرتين**.
"""
import sys, os, io
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'odoo_templates'))
from jrpc import x
from collections import defaultdict

SALES = ['out_invoice', 'out_refund']
BILLS = ['in_invoice', 'in_refund']
SERVICES_CATEG = 28          # فئة الخدمات — أصنافها بتتكرر بشكل مشروع

# حالات فحصها صاحب المحل واعتمدها — تُعرض كملاحظة لا كمانع
ALLOW = {
    'BILL/2026/08/0001': 'رصيد مدوّر لكتانة — سطر بلا صنف بشكل مقصود',
}

# التقرير المناسب حسب طبيعة التعامل مع الجهة
REPORTS = {
    'both':     ('alawn.report_joint_statement',    'كشف حساب موحّد'),
    'customer': ('alawn.report_customer_statement', 'كشف حساب العميل'),
    'vendor':   ('alawn.report_vendor_statement',   'كشف حساب المورّد'),
}

BAR = '─' * 78


def _resolve(token):
    """رقم الجهة أو اسمها ← (id, name)."""
    if token.isdigit():
        r = x('res.partner', 'read', [int(token)], ['id', 'name'])
        if not r:
            raise SystemExit('⛔ مدير الكشوفات: ما في جهة بالرقم %s' % token)
        return r[0]['id'], r[0]['name']
    hits = x('res.partner', 'search_read', [('name', 'ilike', token)], ['id', 'name'])
    if not hits:
        raise SystemExit('⛔ مدير الكشوفات: ما لقيت جهة اسمها «%s»' % token)
    if len(hits) > 1:
        raise SystemExit('⛔ مدير الكشوفات: «%s» بتطابق %d جهات: %s'
                         % (token, len(hits), ' · '.join('%d %s' % (h['id'], h['name']) for h in hits)))
    return hits[0]['id'], hits[0]['name']


def فحص(partner_id=None):
    pdom = [('partner_id', '=', partner_id)] if partner_id else []
    out = defaultdict(list)

    def lines(extra, fields):
        return x('account.move.line', 'search_read',
                 [('display_type', '=', 'product'), ('parent_state', '=', 'posted'),
                  ('move_id.move_type', 'in', SALES + BILLS)] + extra + pdom,
                 fields, order='date')

    for l in lines([('price_unit', '=', 0)],
                   ['move_id', 'product_id', 'quantity', 'partner_id', 'date']):
        out[l['partner_id'][1]].append(
            ('سعر صفر', l['date'], l['move_id'][1],
             l['product_id'][1] if l['product_id'] else '(بلا صنف)', l['quantity']))

    for l in lines([('product_id', '=', False)],
                   ['move_id', 'name', 'quantity', 'partner_id', 'date']):
        out[l['partner_id'][1]].append(
            ('بلا صنف', l['date'], l['move_id'][1], (l['name'] or '')[:42], l['quantity']))

    # سطر قيد شريكه جهة مؤرشفة — بيختفي من رصيد المشروع بصمت
    # (لغم 19/09: commercial_partner_id بايت من زلّة parent_id القديمة)
    for l in x('account.move.line', 'search_read',
               [('parent_state', '=', 'posted'), ('partner_id.active', '=', False),
                ('account_id.account_type', 'in', ['asset_receivable', 'liability_payable'])]
               + ([('move_id.partner_id', '=', partner_id)] if partner_id else []),
               ['move_id', 'partner_id', 'debit', 'credit', 'date']):
        mv = x('account.move', 'read', [l['move_id'][0]], ['partner_id'])[0]
        owner = mv['partner_id'][1] if mv['partner_id'] else '(بلا جهة)'
        out[owner].append(
            ('شريك مؤرشف', l['date'], l['move_id'][1],
             'السطر على «%s»' % l['partner_id'][1], '%.3f' % (l['debit'] - l['credit'])))

    # فاتورة مكرّرة — نفس الجهة ونفس اليوم ونفس الصنف بنفس الكمية على مستندين
    # (لغم 23/09: نموذج الورشة بيغطي الورشة من أولها، فبنوده القديمة بتنفوتر مرتين)
    # أصناف الخدمة العامة بتتكرر بشكل مشروع (فاتورة يدوية سابقة · هدية · خصم · تبرّع)
    عام = {p['id'] for p in x('product.product', 'search_read',
           [('categ_id', '=', SERVICES_CATEG)], ['id'])}
    مفاتيح = defaultdict(list)
    for l in x('account.move.line', 'search_read',
               [('display_type', '=', 'product'), ('parent_state', '=', 'posted'),
                ('move_id.move_type', 'in', SALES), ('product_id', '!=', False)] + pdom,
               ['move_id', 'product_id', 'quantity', 'price_unit', 'partner_id', 'date']):
        if l['product_id'][0] in عام:
            continue
        مفاتيح[(l['partner_id'][0], l['date'], l['product_id'][0],
                l['quantity'], l['price_unit'])].append(l)
    for (pid, d, prod, qty, pu), ls in مفاتيح.items():
        مستندات = {i['move_id'][1] for i in ls}
        if len(مستندات) > 1:
            out[ls[0]['partner_id'][1]].append(
                ('فاتورة مكرّرة', d, ' + '.join(sorted(مستندات)),
                 ls[0]['product_id'][1][:42], qty))

    for m in x('account.move', 'search_read',
               [('state', '=', 'draft'), ('move_type', 'in', SALES + BILLS)] + pdom,
               ['name', 'partner_id', 'invoice_date', 'amount_total'], order='invoice_date'):
        if m['partner_id']:
            out[m['partner_id'][1]].append(
                ('مسودة معلّقة', m['invoice_date'], m['name'] or '(بلا رقم)',
                 'إجمالي %.3f' % m['amount_total'], ''))
    return out


def افرز(findings):
    """يفصل الموانع الحقيقية عن الحالات المعتمدة."""
    blocking, known = {}, []
    for partner, rows in findings.items():
        keep = []
        for r in rows:
            hit = next((k for k in ALLOW if k in (r[2] or '')), None)
            (known.append((partner, hit, ALLOW[hit])) if hit else keep.append(r))
        if keep:
            blocking[partner] = keep
    return blocking, known


def طبيعة_التعامل(partner_id):
    has_sale = x('account.move', 'search_count',
                 [('partner_id', '=', partner_id), ('move_type', 'in', SALES), ('state', '=', 'posted')])
    has_bill = x('account.move', 'search_count',
                 [('partner_id', '=', partner_id), ('move_type', 'in', BILLS), ('state', '=', 'posted')])
    if has_sale and has_bill:
        return 'both'
    return 'vendor' if has_bill else 'customer'


def اصدار(partner_id, name):
    from render import render
    kind = طبيعة_التعامل(partner_id)
    report, label = REPORTS[kind]
    safe = ''.join(c for c in name if c not in '/\\:*?"<>|').strip().replace(' ', '_')
    path = '/home/user/Fam-Bam/العون/%s_%s.pdf' % (label.replace(' ', '_'), safe)
    render(report, partner_id, path)
    print('📄 %s → %s' % (label, path))
    return path


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    do_issue = '--اصدار' in sys.argv or '--issue' in sys.argv

    pid, pname = (None, 'كل الجهات')
    if args:
        pid, pname = _resolve(args[0])

    print(BAR)
    print('  مدير الكشوفات — محلات العون لمواد البناء')
    print('  الطلب: كشف حساب · %s' % pname)
    print(BAR)

    blocking, known = افرز(فحص(pid))

    if known:
        print('\nℹ️  حالات معروفة ومعتمدة سابقاً (مش مانع):')
        for partner, doc, why in known:
            print('    %-26s %-18s %s' % (partner, doc, why))

    if blocking:
        print('\n⛔ الكشف موقوف — في ملاحظات لازم تتعالج أو يوافق عليها صاحب المحل:\n')
        for partner in sorted(blocking):
            print('  ■ %s' % partner)
            for kind, date, doc, what, qty in blocking[partner]:
                print('     [%-12s] %-11s %-18s %-42s %s' % (kind, date, doc, what, qty))
            print()
        print('  لا يُسلَّم الكشف قبل معالجة اللي فوق.')
        print(BAR)
        return 1

    print('\n✅ الحساب نظيف — ما في بند بسعر صفر ولا بند بلا صنف ولا مسودة معلّقة.')
    if do_issue and pid:
        print('   الكشف معتمد للتسليم.\n')
        اصدار(pid, pname)
    elif do_issue:
        print('   (الإصدار بدّه جهة محددة — مرّر رقمها أو اسمها)')
    else:
        print('   الكشف معتمد للتسليم. للإصدار: أضف --اصدار')
    print(BAR)
    return 0


if __name__ == '__main__':
    sys.exit(main())
