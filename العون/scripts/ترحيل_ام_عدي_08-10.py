# -*- coding: utf-8 -*-
"""ترحيل «ام عدي جميل مصلح» — بأمر صاحب المحل «ام عدي رحل وبسعر الورقه» 08/10
بند واحد: سكينة جيب ×1 @4.000 (الورقة) ← `700826` · التاريخ المخزّن بالخلية 03/10/2026 · المرجع فاضي · المستلم ما بينزل."""
import sys, json
sys.path.insert(0, 'odoo_templates')
from jrpc import x
LOG = 'odoo_backup/ROLLBACK_umodai_post_08-10.json'; log = {}
assert not x('res.partner', 'search', [['name', '=', 'ام عدي جميل مصلح']])
P = x('res.partner', 'create', {'name': 'ام عدي جميل مصلح', 'category_id': [(6, 0, [15])], 'customer_rank': 1}); log['partner'] = P
pr = x('product.product', 'search_read', [['default_code', '=', '700826']], fields=['id', 'name'])[0]
mid = x('account.move', 'create', {'move_type': 'out_invoice', 'partner_id': P, 'invoice_date': '2026-10-03', 'date': '2026-10-03', 'journal_id': 8,
      'invoice_line_ids': [(0, 0, {'product_id': pr['id'], 'name': pr['name'], 'quantity': 1, 'price_unit': 4.0, 'tax_ids': [(6, 0, [29])]})]})
assert abs(x('account.move', 'read', [mid], fields=['amount_total'])[0]['amount_total'] - 4.0) < 0.001
x('account.move', 'action_post', [mid])
log['move'] = x('account.move', 'read', [mid], fields=['id', 'name', 'amount_total', 'date'])[0]
json.dump(log, open(LOG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('✓ [%d] ام عدي جميل مصلح ·' % P, log['move']['name'], log['move']['date'], log['move']['amount_total'], '· الرصيد', x('res.partner', 'read', [P], fields=['credit'])[0]['credit'])
