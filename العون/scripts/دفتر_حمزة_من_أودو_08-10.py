# -*- coding: utf-8 -*-
"""دفتر المهندس حمزة شخصي — تعبئة النموذج المعتمد بحركات أودو اللي مش فيه + ترتيب حسب التاريخ (صاحب المحل 08/10)
python3 scripts/دفتر_حمزة_من_أودو_08-10.py <الدفتر.xlsx> <الناتج.xlsx>
- أسطر الدفتر زي ما هي · الجديد من أودو: المدقق فاضي و«من أودو — رقم المستند» بالملاحظات
- الدفعة (دفعة نقدية من محمد جهاد الصعايدة) ← شيت «الدفعات» الجديد"""
import sys, os, datetime, subprocess, openpyxl
sys.path.insert(0, 'odoo_templates')
from jrpc import x
src, out = sys.argv[1], sys.argv[2]
tmp = out + '.tmp.xlsx'
subprocess.run([sys.executable, 'scripts/شيت_الدفعات_للنموذج.py', src, tmp], check=True)
wb = openpyxl.load_workbook(tmp); ws = wb['الورشة']
D = lambda s: datetime.datetime.strptime(s, '%Y-%m-%d')
rows = []
for r in range(7, 1007):
    v = [ws.cell(r, c).value for c in range(2, 10)]
    if v[0] is not None: rows.append((v[4], 0, len(rows), v))
n_old = len(rows)
# أودو: قيود الكاش اللي مش بالدفتر (الثمانية — قرار 08/10 تضل) — بنطابق بالتاريخ والمبلغ
book = {}
for _, _, _, v in rows: book.setdefault((v[4].date(), float(v[3] or 0)), 0); book[(v[4].date(), float(v[3] or 0))] += 1
L = x('account.move.line', 'search_read', [['partner_id', '=', 52], ['account_id', '=', 55], ['parent_state', '=', 'posted'], ['journal_id', '=', 10]],
      fields=['date', 'move_id', 'name', 'debit', 'credit'], order='date,id')
extra = []
for l in L:
    if not l['debit']: continue
    k = (D(l['date']).date(), l['debit'])
    if book.get(k): book[k] -= 1; continue
    doc = l['move_id'][1].split(' ')[0]; det = (l['name'] or '').split(' — ', 1)[1] if ' — ' in (l['name'] or '') else ''
    det = det.replace('ملاحظة: ', '')
    extra.append((D(l['date']), 1, len(rows) + len(extra), ['كاش', 1, None, l['debit'], D(l['date']), None, None, ' · '.join(filter(None, [det, 'من أودو — ' + doc]))]))
# أودو: بنود الفواتير
for m in x('account.move', 'search_read', [['partner_id', '=', 52], ['move_type', '=', 'out_invoice'], ['state', '=', 'posted']], fields=['id', 'name', 'invoice_date', 'ref']):
    for i in x('account.move.line', 'search_read', [['move_id', '=', m['id']], ['display_type', '=', 'product']], fields=['name', 'quantity', 'price_unit', 'product_id']):
        nm = i['name'].split('\n'); note = ' · '.join(nm[1:])
        extra.append((D(m['invoice_date']), 1, len(rows) + len(extra),
                      [i['product_id'][1].split('] ', 1)[-1], i['quantity'], i['price_unit'], None, D(m['invoice_date']), m['ref'] or None, None,
                       ' · '.join(filter(None, [note, 'من أودو — ' + m['name']]))]))
allr = sorted(rows + extra, key=lambda t: (t[0], t[1], t[2]))
for r in range(7, 1007):
    for c in range(2, 10): ws.cell(r, c).value = None
for i, (_, _, _, v) in enumerate(allr):
    for c, val in enumerate(v, 2): ws.cell(7 + i, c).value = val
# الدفعات
P = wb['الدفعات']
for l in L:
    if l['credit']:
        det = (l['name'] or '')
        frm = det.split(' من ', 1)[1].strip('() ') if ' من ' in det else ''
        doc = l['move_id'][1].split(' ')[0]
        P['B7'], P['C7'], P['D7'], P['E7'], P['H7'] = D(l['date']), 'كاش نقدي', l['credit'], frm, 'من أودو — ' + doc
wb.save(out); os.remove(tmp)
print('الدفتر %d سطر + من أودو %d = %d · مجموع الكاش %.3f · البضاعة %.3f' % (n_old, len(extra), len(allr),
      sum(v[3] or 0 for *_, v in allr), sum((v[1] or 0) * (v[2] or 0) for *_, v in allr)))
for _, _, _, v in extra: print('  +', v[4].strftime('%d/%m/%Y'), v[0], v[1], v[2] or v[3], '|', v[5] or '', '|', v[7])
