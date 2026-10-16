# -*- coding: utf-8 -*-
"""ترحيل مسودة ثروت المصري [44] — بأمر صاحب المحل «رحل ثروت» 07/10
من drive/مسودات/ثروت_مسودة_يومية.json · كل تاريخ = فاتورة · الفواتير اليدوية كل وحدة لحالها
المرتجع = إشعار دائن واحد بتاريخه · كل بند بسطر مستقل (ممنوع الدمج)."""
import sys, json, collections
sys.path.insert(0, 'odoo_templates')
from jrpc import x
P = 44; TAX = [(6, 0, [29])]
L = json.load(open('drive/مسودات/ثروت_مسودة_يومية.json', encoding='utf-8'))
assert not any(l['معلّق'] for l in L)
log = {'moves': [], 'payments': []}
LOG = 'odoo_backup/ROLLBACK_tharwat_post_07-10.json'
def save(): json.dump(log, open(LOG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
ids = {p['default_code']: p['id'] for p in x('product.product', 'search_read',
       [['default_code', 'in', list({l['كود'] for l in L if l['كود']})]], fields=['default_code'])}
def اسم(l):
    if l.get('وصف'): return l['وصف']
    if l['كود'] == '900900': return 'فاتورة يدوية سابقة — ' + l['ورقة']
    if l['كود'] == '702554': return f"{l['صنف']} — {l['ورقة']}"
    return l['صنف']
groups = collections.OrderedDict()
for l in sorted(L, key=lambda l: l['تاريخ']):
    if l['نوع'] in ('بند', 'فاتورة يدوية'):
        k = ('out_invoice', l['تاريخ'], l['رقم'] if l['نوع'] == 'فاتورة يدوية' else '')
    elif l['نوع'] == 'مرتجع': k = ('out_refund', l['تاريخ'], '')
    else: continue
    groups.setdefault(k, []).append(l)
# 1) إنشاء كل المستندات كمسودة
for (t, d, sep), G in groups.items():
    nums = ' '.join(g['رقم'] for g in G)
    vals = {'move_type': t, 'partner_id': P, 'invoice_date': d, 'date': d,
            'ref': f'ورشة ثروت المصري — بنود الدرايف {nums}' + (' (مرتجع)' if t == 'out_refund' else ''),
            'invoice_line_ids': [(0, 0, {'product_id': ids[g['كود']], 'name': اسم(g), 'quantity': g['كمية'],
                                          'price_unit': g['سعر'], 'tax_ids': TAX}) for g in G]}
    mid = x('account.move', 'create', vals)
    exp = round(sum(g['كمية'] * g['سعر'] for g in G), 3)
    got = x('account.move', 'read', [mid], fields=['amount_total'])[0]['amount_total']
    assert abs(got - exp) < 0.001, (d, got, exp)
    log['moves'].append({'id': mid, 'type': t, 'date': d, 'amount': exp, 'rows': nums}); save()
print('مسودات:', len(log['moves']))
# 2) ترحيل
for m in log['moves']:
    x('account.move', 'action_post', [m['id']])
    m['name'] = x('account.move', 'read', [m['id']], fields=['name'])[0]['name']; save()
# 3) الدفعات والصرف
pay = {'125': (15, 6, 'inbound', 'دفعة نقدية من ثروت صباح (26/4) — ملف الدرايف'),
       '126': (15, 6, 'inbound', 'دفعة نقدية من ثروت صباح (30/5) — ملف الدرايف'),
       '127': (13, 3, 'inbound', 'دفعة تحويل كليك من (عيسى مضاعين) الى (م.عمر)'),
       '50':  (15, 5, 'outbound', 'صرف نقدي الى ثروت — كاش 30 دينار حسب ورقة الدرايف')}
for l in L:
    if l['رقم'] not in pay: continue
    j, ml, tp, memo = pay[l['رقم']]
    pid = x('account.payment', 'create', {'partner_id': P, 'partner_type': 'customer', 'payment_type': tp,
            'journal_id': j, 'payment_method_line_id': ml, 'amount': l['كمية'] * l['سعر'], 'date': l['تاريخ'], 'memo': memo})
    x('account.payment', 'action_post', [pid])
    log['payments'].append({'id': pid, 'row': l['رقم'], 'amount': l['سعر'], 'date': l['تاريخ'],
                            'name': x('account.payment', 'read', [pid], fields=['name'])[0]['name']}); save()
print('رصيد [44]:', x('res.partner', 'read', [P], fields=['credit'])[0]['credit'])
