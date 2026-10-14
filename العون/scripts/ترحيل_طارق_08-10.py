# -*- coding: utf-8 -*-
"""ترحيل طارق الماضي [48] #56–#58 (08/10/2026) — بأمر صاحب المحل «رحل طارق» 08/10 · الرول @2 بقراره"""
import sys, json
sys.path.insert(0, 'odoo_templates')
from jrpc import x
P = 48; p = 'drive/مسودات/طارق_08-10.json'
d = json.load(open(p, encoding='utf-8'))
for l in d['أسطر']:
    if l['#'] == 58: l['سعر'] = 2.0; l['مصدر_السعر'] = 'قرار صاحب المحل 08/10 (2 دينار)'
json.dump(d, open(p, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
assert x('account.move', 'search_count', [['partner_id', '=', P], ['invoice_date', '=', d['تاريخ']], ['state', '!=', 'cancel']]) == 0, 'في فاتورة بنفس التاريخ'
before = x('res.partner', 'read', [P], fields=['credit'])[0]['credit']
pr = {q['default_code']: q for q in x('product.product', 'search_read', [['default_code', 'in', [l['كود'] for l in d['أسطر']]]], fields=['id', 'name', 'default_code'])}
mid = x('account.move', 'create', {'move_type': 'out_invoice', 'partner_id': P, 'invoice_date': d['تاريخ'], 'date': d['تاريخ'], 'journal_id': 8,
      'invoice_line_ids': [(0, 0, {'product_id': pr[l['كود']]['id'], 'name': pr[l['كود']]['name'], 'quantity': l['كمية'], 'price_unit': l['سعر'],
                                   'tax_ids': [(6, 0, [29])]}) for l in d['أسطر']]})
exp = round(sum(l['كمية'] * l['سعر'] for l in d['أسطر']), 3)
assert abs(x('account.move', 'read', [mid], fields=['amount_total'])[0]['amount_total'] - exp) < 0.001
x('account.move', 'action_post', [mid])
m = x('account.move', 'read', [mid], fields=['id', 'name', 'amount_total', 'date'])[0]
after = x('res.partner', 'read', [P], fields=['credit'])[0]['credit']
json.dump({'move': m, 'before': before, 'after': after}, open('odoo_backup/ROLLBACK_tarek_post_08-10.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('✓', m['name'], m['date'], m['amount_total'], '· رصيد [48]', before, '←', after)
