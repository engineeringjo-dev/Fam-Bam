# -*- coding: utf-8 -*-
"""تاريخ دفعة غالب PBNK1/2026/00035 ← 30/09/2026 (صاحب المحل 08/10) + إعادة المطابقة"""
import sys, json
sys.path.insert(0, 'odoo_templates')
from jrpc import x
P = 49
p = x('account.payment', 'search_read', [['name', '=', 'PBNK1/2026/00035']], fields=['id', 'name', 'date', 'amount', 'state', 'partner_id'])[0]
assert p['partner_id'][0] == P and abs(p['amount'] - 395) < 0.001, p
json.dump(p, open('odoo_backup/ROLLBACK_ghaleb_paydate_08-10.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
x('account.payment', 'action_draft', [p['id']])
x('account.payment', 'write', [p['id']], {'date': '2026-09-30'})
x('account.payment', 'action_post', [p['id']])
ids = x('account.move.line', 'search', [['partner_id', '=', P], ['account_id', '=', 55], ['reconciled', '=', False], ['parent_state', '=', 'posted']])
if ids: x('account.move.line', 'reconcile', ids)
q = x('account.payment', 'read', [p['id']], fields=['name', 'date', 'amount', 'state', 'move_id'])[0]
mv = x('account.move', 'read', [q['move_id'][0]], fields=['date'])[0]
print('✓', q['name'], q['date'], mv['date'], q['amount'], q['state'])
for m in x('account.move', 'search_read', [['partner_id', '=', P], ['state', '=', 'posted'], ['move_type', '!=', 'entry']], fields=['name', 'payment_state']): print('  ', m['name'], m['payment_state'])
print('رصيد [49]:', x('res.partner', 'read', [P], fields=['credit'])[0]['credit'])
