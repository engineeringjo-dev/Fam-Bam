# -*- coding: utf-8 -*-
"""ترحيل بضاعة أبو محمود — شخصي [51] — بأمر صاحب المحل «تمام ورحل بسعر الورقه» 08/10
05/06/2025: لمبة ×9 @0.6 ← `600003` · سوكة ×9 @0.6 ← `600019` · فاتورة وحدة · المرجع فاضي."""
import sys, json
sys.path.insert(0, 'odoo_templates')
from jrpc import x
P = 51; LOG = 'odoo_backup/ROLLBACK_abumahmoud_post_08-10.json'
before = x('res.partner', 'read', [P], fields=['credit'])[0]['credit']
pr = {p['default_code']: p for p in x('product.product', 'search_read', [['default_code', 'in', ['600003', '600019']]], fields=['id', 'name', 'default_code'])}
mid = x('account.move', 'create', {'move_type': 'out_invoice', 'partner_id': P, 'invoice_date': '2025-06-05', 'date': '2025-06-05', 'journal_id': 8,
      'invoice_line_ids': [(0, 0, {'product_id': pr[c]['id'], 'name': pr[c]['name'], 'quantity': 9, 'price_unit': 0.6, 'tax_ids': [(6, 0, [29])]}) for c in ('600003', '600019')]})
assert abs(x('account.move', 'read', [mid], fields=['amount_total'])[0]['amount_total'] - 10.8) < 0.001
x('account.move', 'action_post', [mid])
m = x('account.move', 'read', [mid], fields=['id', 'name', 'amount_total', 'date'])[0]
after = x('res.partner', 'read', [P], fields=['credit'])[0]['credit']
json.dump({'move': m, 'before': before, 'after': after}, open(LOG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('✓', m['name'], m['date'], m['amount_total'], '· رصيد [51]', before, '←', after)
