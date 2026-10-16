# -*- coding: utf-8 -*-
"""ترحيل ورشة دريد الماضي — بأمر صاحب المحل «رحل دريد» 08/10
حساب جديد بتصنيف الورشات (زي طارق) · كل تاريخ = فاتورة بضاعة · سعر الورقة · المرجع فاضي · المستلم ما بينزل
اسم البند = اسم الصنف · الدفعة 15 كاش «دفعة نقدية» بلا من/الى."""
import sys, json, collections
sys.path.insert(0, 'odoo_templates')
from jrpc import x
TAX = [(6, 0, [29])]
d = json.load(open('drive/مسودات/ورشة دريد الماضي.json', encoding='utf-8'))
L = [l for l in d['أسطر'] if l['#'] != 25]; PAY = [l for l in d['أسطر'] if l['#'] == 25][0]
assert all(l['كود'] for l in L)
log = {'partner': None, 'moves': [], 'payments': []}
LOG = 'odoo_backup/ROLLBACK_dureid_post_08-10.json'
def save(): json.dump(log, open(LOG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
assert not x('res.partner', 'search', [['name', '=', 'ورشة دريد الماضي']])
P = x('res.partner', 'create', {'name': 'ورشة دريد الماضي', 'category_id': [(6, 0, [15])], 'customer_rank': 1})
log['partner'] = P; save()
prod = {p['default_code']: p for p in x('product.product', 'search_read', [['default_code', 'in', list({l['كود'] for l in L})]], fields=['default_code', 'name'])}
G = collections.OrderedDict()
for l in sorted(L, key=lambda l: (l['تاريخ'], l['#'])): G.setdefault(l['تاريخ'], []).append(l)
for dt, ls in G.items():
    mid = x('account.move', 'create', {'move_type': 'out_invoice', 'partner_id': P, 'invoice_date': dt, 'date': dt, 'journal_id': 8,
          'invoice_line_ids': [(0, 0, {'product_id': prod[l['كود']]['id'], 'name': prod[l['كود']]['name'], 'quantity': l['كمية'],
                                       'price_unit': l['سعر'], 'tax_ids': TAX}) for l in ls]})
    exp = round(sum(l['كمية'] * l['سعر'] for l in ls), 3)
    got = x('account.move', 'read', [mid], fields=['amount_total'])[0]['amount_total']
    assert abs(got - exp) < 0.001, (dt, got, exp)
    log['moves'].append({'id': mid, 'date': dt, 'amount': exp, 'rows': [l['#'] for l in ls]}); save()
for m in log['moves']:
    x('account.move', 'action_post', [m['id']])
    m['name'] = x('account.move', 'read', [m['id']], fields=['name'])[0]['name']; save()
    print('✓', m['name'], m['date'], '%.3f' % m['amount'], 'بنود', m['rows'])
pid = x('account.payment', 'create', {'partner_id': P, 'partner_type': 'customer', 'payment_type': 'inbound', 'journal_id': 15,
        'payment_method_line_id': 6, 'amount': PAY['كمية'] * PAY['سعر_الورقة'], 'date': PAY['تاريخ'], 'memo': 'دفعة نقدية'})
x('account.payment', 'action_post', [pid])
log['payments'].append({'id': pid, 'name': x('account.payment', 'read', [pid], fields=['name'])[0]['name'], 'amount': 15.0, 'date': PAY['تاريخ']}); save()
print('✓', log['payments'][0]['name'], PAY['تاريخ'], '15.000')
print('حساب [%d] ورشة دريد الماضي · الرصيد:' % P, x('res.partner', 'read', [P], fields=['credit'])[0]['credit'])
