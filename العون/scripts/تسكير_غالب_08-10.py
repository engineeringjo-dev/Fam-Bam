# -*- coding: utf-8 -*-
"""تسكير ورشة د.غالب صويص [49] — بأمر صاحب المحل 08/10
«دفع دفعة تحويل الى المهندس عمر بقيمة 395 دينار والباقي لتسكير الفاتورة سجّله مرتجع بضاعة وهيك بصير حسابها الورشة تسكر»
1) دفعة تحويل 395.000 (بنك 13 · كليك / حوالة واردة) بتاريخ 08/10/2026
2) مرتجع بضاعة بالباقي (427.700 − 395 = 32.700) · بند 900900 باسم «مرتجع بضاعة» · مرجع فاضي
3) مرجع RINV/2026/00007 فاضي (قواعد الصياغة) · ومطابقة كل الحركات ← الرصيد 0
"""
import sys, json
sys.path.insert(0, 'odoo_templates')
from jrpc import x
P, DATE, PAY = 49, '2026-10-08', 395.0
M = x('account.move', 'search_read', [['partner_id', '=', P], ['state', '=', 'posted']], fields=['id', 'name', 'ref', 'amount_total', 'amount_residual'])
bal = x('res.partner', 'read', [P], fields=['credit'])[0]['credit']
json.dump({'moves': M, 'credit': bal}, open('odoo_backup/ROLLBACK_ghaleb_close_08-10.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
assert abs(bal - 427.7) < 0.001, bal
rest = round(bal - PAY, 3)
# 1) مرجع المرتجع القديم فاضي
x('account.move', 'write', [m['id'] for m in M if m['name'] == 'RINV/2026/00007'], {'ref': False})
# 2) الدفعة
pid = x('account.payment', 'create', {'partner_id': P, 'partner_type': 'customer', 'payment_type': 'inbound', 'amount': PAY, 'date': DATE,
                                      'journal_id': 13, 'payment_method_line_id': 3, 'memo': 'دفعة تحويل من (د.غالب صويص) الى (م.عمر)'})
x('account.payment', 'action_post', [pid])
# 3) مرتجع البضاعة بالباقي
prod = x('product.product', 'search', [['default_code', '=', '900900']])[0]
rid = x('account.move', 'create', {'move_type': 'out_refund', 'partner_id': P, 'invoice_date': DATE, 'date': DATE, 'journal_id': 8,
                                   'invoice_line_ids': [(0, 0, {'product_id': prod, 'name': 'مرتجع بضاعة', 'quantity': 1, 'price_unit': rest, 'tax_ids': [(6, 0, [29])]})]})
x('account.move', 'action_post', [rid])
r = x('account.move', 'read', [rid], fields=['name', 'amount_total', 'state'])[0]
assert abs(r['amount_total'] - rest) < 0.001 and r['state'] == 'posted', r
# 4) مطابقة كل الحركات المفتوحة على الذمم
ids = x('account.move.line', 'search', [['partner_id', '=', P], ['account_id', '=', 55], ['reconciled', '=', False], ['parent_state', '=', 'posted']])
try: x('account.move.line', 'reconcile', ids)
except Exception as e: print('ملاحظة مطابقة:', e)
pay = x('account.payment', 'read', [pid], fields=['name', 'state'])[0]
print('✓ الدفعة', pay['name'], PAY, '· المرتجع', r['name'], rest)
for m in x('account.move', 'search_read', [['partner_id', '=', P], ['state', '=', 'posted'], ['move_type', '!=', 'entry']], fields=['name', 'amount_total', 'payment_state']):
    print('  ', m['name'], m['amount_total'], m['payment_state'])
print('رصيد [49]:', x('res.partner', 'read', [P], fields=['credit'])[0]['credit'])
