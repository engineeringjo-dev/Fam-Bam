# -*- coding: utf-8 -*-
"""تطبيق قواعد الصياغة (07/10) على مشروع المدرسة [19] — بأمر صاحب المحل 08/10
- كل المراجع فاضية، إلا INV/2026/00159 «محيسن» (ملاحظة على كل الفاتورة → بالمرجع مرة وحدة، مش على كل بند)
- اسم البند = اسم الصنف فقط (بدون [كود] وبدون الأسماء القديمة)
- الفواتير اليدوية: «(فاتورة يدوية سابقة رقم N)»
- ملاحظات موجودة أصلاً بالبند (فصول · صحراوي · ازرق) → سطر ثاني جوّا نفس الخانة
"""
import sys, json, re
sys.path.insert(0, 'odoo_templates')
from jrpc import x
P = 19
M = x('account.move', 'search_read', [['partner_id', '=', P], ['state', '=', 'posted'], ['move_type', 'in', ['out_invoice', 'out_refund']]],
      fields=['id', 'name', 'ref', 'narration', 'move_type', 'amount_total'])
L = x('account.move.line', 'search_read', [['move_id', 'in', [m['id'] for m in M]], ['display_type', '=', 'product']],
      fields=['id', 'move_id', 'name', 'product_id', 'sequence'])
json.dump({'moves': M, 'lines': L}, open('odoo_backup/ROLLBACK_school_siyagha_08-10.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
pname = {p['id']: p['name'] for p in x('product.product', 'read', list({l['product_id'][0] for l in L}), fields=['name'])}
REF = {'INV/2026/00159': 'محيسن'}
NOTE = {('INV/2026/00048', '[500533] ماسورة بلاستيك 4" ضغط اردني (فصول)'): '(فصول)',
        ('INV/2026/00035', '[200210] صبغة هندي صحراوي'): 'صحراوي',
        ('INV/2026/00007', '[100129] خيط بناء سماكة0.85 ملي ازرق'): 'ازرق'}
bymove = {}
for l in L: bymove.setdefault(l['move_id'][0], []).append(l)
n_ref = n_mv = 0
for m in M:
    want_ref = REF.get(m['name'], False)
    if (m['ref'] or False) != want_ref or m['narration']:
        x('account.move', 'write', [m['id']], {'ref': want_ref, 'narration': False}); n_ref += 1
    cmds = []
    for l in bymove.get(m['id'], []):
        nm = l['name'] or ''; new = pname[l['product_id'][0]]
        mm = re.match(r'فاتورة يدوية سابقة رقم (\d+)$', nm)
        if mm: new = '(فاتورة يدوية سابقة رقم %s)' % mm.group(1)
        elif (m['name'], nm) in NOTE: new += '\n' + NOTE[(m['name'], nm)]
        if new != nm: cmds.append((1, l['id'], {'name': new}))
    if not cmds: continue
    before = m['amount_total']
    x('account.move', 'button_draft', [m['id']])
    x('account.move', 'write', [m['id']], {'invoice_line_ids': cmds})
    x('account.move', 'action_post', [m['id']])
    after = x('account.move', 'read', [m['id']], fields=['amount_total', 'name', 'state'])[0]
    assert abs(after['amount_total'] - before) < 0.001 and after['name'] == m['name'] and after['state'] == 'posted', after
    n_mv += 1
    print('✓', m['name'], len(cmds), 'بند', before)
print('مراجع انمسحت/تعدّلت:', n_ref, '· فواتير تعدّلت بنودها:', n_mv)
print('رصيد [19]:', x('res.partner', 'read', [P], fields=['credit'])[0]['credit'])
