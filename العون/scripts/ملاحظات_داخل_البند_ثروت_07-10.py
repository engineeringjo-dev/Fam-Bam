# -*- coding: utf-8 -*-
"""ثروت [44] — الملاحظة جوّا خانة البند (سطر ثاني) بدل سطر ملاحظة مستقل · توحيد «براغي وزن» (صاحب المحل 07/10)"""
import sys, json
sys.path.insert(0, 'odoo_templates')
from jrpc import x
# 1) توحيد التسمية: بالوزن ← وزن (البراغي)
ren = {'702554': 'براغي وزن', '800050': 'براغي 5 سم مع اسافين (وزن)', '800051': 'براغي 3.5 سم (وزن)'}
old = {}
for c, n in ren.items():
    p = x('product.product', 'search_read', [['default_code', '=', c]], fields=['name', 'product_tmpl_id'])[0]
    old[c] = p['name']; x('product.template', 'write', [p['product_tmpl_id'][0]], {'name': n})
M = {m['name']: m for m in x('account.move', 'search_read', [['partner_id', '=', 44], ['state', '=', 'posted']], fields=['id', 'name', 'amount_total'])}
L = x('account.move.line', 'search_read', [['move_id', 'in', [m['id'] for m in M.values()]], ['display_type', 'in', ['product', 'line_note']]],
      fields=['id', 'move_id', 'name', 'product_id', 'display_type', 'sequence'])
json.dump({'products': old, 'lines': L}, open('odoo_backup/ROLLBACK_tharwat_inline_notes_07-10.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
pn = {p['id']: p['name'] for p in x('product.product', 'read', list({l['product_id'][0] for l in L if l['product_id']}), fields=['name'])}
# ملاحظات للبنود اللي ما إلها سطر ملاحظة: (الفاتورة, كود) ← قائمة ملاحظات بالترتيب
EXTRA = {('INV/2026/00146', '702554'): ['(سلف)'], ('INV/2026/00148', '702554'): ['(سلف درل)'],
         ('INV/2026/00149', '702554'): ['(سن خشن 6ملم)'], ('INV/2026/00126', '200209'): ['اسود'],
         ('INV/2026/00129', '200137'): ['ابيض'], ('INV/2026/00081', '200209'): ['صحراوي'], ('INV/2026/00082', '200209'): ['صحراوي'],
         ('INV/2026/00089', '200209'): ['صحراوي'], ('INV/2026/00089', '200210'): ['صحراوي'], ('INV/2026/00084', '900921'): ['ابيض']}
code = {p['id']: p['default_code'] for p in x('product.product', 'read', list(pn), fields=['default_code'])}
bym = {}
for l in L: bym.setdefault(l['move_id'][0], []).append(l)
for name, m in sorted(M.items()):
    rows = sorted(bym.get(m['id'], []), key=lambda l: (l['sequence'], l['id']))
    cmds, last = [], None
    for l in rows:
        if l['display_type'] == 'line_note':
            if last is None: continue
            note = l['name'].strip()
            if last[0] == '702554' and not note.startswith('('): note = '(%s)' % note
            last[1].append(note); cmds.append((2, l['id'], 0)); continue
        c = code[l['product_id'][0]]
        first = (l['name'] or '').split('\n')[0]
        if c != '900900': first = pn[l['product_id'][0]]
        last = (c, list(EXTRA.get((name, c), [])), l['id'], first, l['name'])
        cmds.append(last)
    upd = []
    for c in cmds:
        if isinstance(c, tuple) and len(c) == 5:
            newname = '\n'.join([c[3]] + c[1])
            if newname != c[4]: upd.append((1, c[2], {'name': newname}))
        elif c[0] == 2: upd.append(c)
    if not upd: continue
    before = m['amount_total']
    x('account.move', 'button_draft', [m['id']]); x('account.move', 'write', [m['id']], {'invoice_line_ids': upd}); x('account.move', 'action_post', [m['id']])
    a = x('account.move', 'read', [m['id']], fields=['amount_total', 'name', 'state'])[0]
    assert abs(a['amount_total'] - before) < 0.001 and a['name'] == name and a['state'] == 'posted', a
    print('✓', name)
print('رصيد', x('res.partner', 'read', [44], fields=['credit'])[0]['credit'])
