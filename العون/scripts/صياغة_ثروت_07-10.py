# -*- coding: utf-8 -*-
"""تطبيق قواعد الصياغة (07/10) على ورشة ثروت المصري [44] — بموافقة صاحب المحل"""
import sys, json, re
sys.path.insert(0, 'odoo_templates')
from jrpc import x
P = 44
M = x('account.move', 'search_read', [['partner_id', '=', P], ['state', '=', 'posted'], ['move_type', 'in', ['out_invoice', 'out_refund']]],
      fields=['id', 'name', 'ref', 'narration', 'move_type', 'amount_total'])
L = x('account.move.line', 'search_read', [['move_id', 'in', [m['id'] for m in M]], ['display_type', 'in', ['product', 'line_note']]],
      fields=['id', 'move_id', 'name', 'product_id', 'display_type', 'sequence'])
json.dump({'moves': M, 'lines': L}, open('odoo_backup/ROLLBACK_tharwat_siyagha_07-10.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
pname = {p['id']: p['name'] for p in x('product.product', 'read', list({l['product_id'][0] for l in L if l['product_id']}), fields=['name'])}
REF = {'INV/2026/00139': '(عيسى مضاعين شخصي)'}
# ملاحظات صاحب المحل ← سطر ملاحظة تحت البند (المفتاح: الفاتورة + نص الاسم الحالي)
NOTE = {('INV/2026/00126', 'براغي بالوزن — سلف دريل'): 'سلف دريل',
        ('INV/2026/00086', 'معجونة سيليكون تركي 450 غم — اسود'): 'اسود',
        ('INV/2026/00086', 'معجونة سيليكون تركي 450 غم — بيج'): 'بيج',
        ('INV/2026/00086', 'معجونة سيليكون تركي 450 غم — ابيض'): 'ابيض',
        ('INV/2026/00086', 'إيجار عدة — كمبريسة هواء 50 لتر — يومين × 5.000'): 'كمبريسة هواء 50 لتر — يومين × 5.000'}
bymove = {}
for l in L: bymove.setdefault(l['move_id'][0], []).append(l)
for m in M:
    # 1) المرجع والملاحظة الداخلية (بدون ما يرجع مسودة)
    x('account.move', 'write', [m['id']], {'ref': REF.get(m['name'], False), 'narration': False})
    lines = sorted(bymove.get(m['id'], []), key=lambda l: (l['sequence'], l['id']))
    plan = []
    for l in lines:
        if l['display_type'] != 'product': continue
        pn = pname[l['product_id'][0]]; nm = l['name'] or ''; new = pn; note = None
        mm = re.match(r'فاتورة يدوية سابقة رقم \((\d+)\)', nm)
        if mm: new, note = '(فاتورة يدوية سابقة رقم %s)' % mm.group(1), 'واتس اب'
        elif m['name'] == 'INV/2026/00139': new = 'فاتورة بضاعة (عيسى مضاعين شخصي)'
        elif (m['name'], nm) in NOTE: note = NOTE[(m['name'], nm)]
        plan.append((l, new, note))
    if not any(new != (l['name'] or '') or note for l, new, note in plan): continue
    before = m['amount_total']
    x('account.move', 'button_draft', [m['id']])
    cmds, seq = [], 10
    for l, new, note in plan:
        cmds.append((1, l['id'], {'name': new, 'sequence': seq}))
        if note: cmds.append((0, 0, {'display_type': 'line_note', 'name': note, 'sequence': seq + 1}))
        seq += 10
    x('account.move', 'write', [m['id']], {'invoice_line_ids': cmds})
    x('account.move', 'action_post', [m['id']])
    after = x('account.move', 'read', [m['id']], fields=['amount_total', 'name', 'state'])[0]
    assert abs(after['amount_total'] - before) < 0.001 and after['name'] == m['name'] and after['state'] == 'posted', after
    print('✓', m['name'], before)
print('رصيد [44]:', x('res.partner', 'read', [P], fields=['credit'])[0]['credit'])
