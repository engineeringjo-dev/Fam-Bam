# -*- coding: utf-8 -*-
"""مسودة ثروت المصري [44] الكاملة — محلياً بس (ولا إشي على أودو)
بتدمج: drive/بنود/ثروت_06-10_مسا.tsv (التواريخ) + ثروت_اختيارات_06-10.json + ثروت_مقترحات_استكمال_06-10.json
وبتطلع: drive/مسودات/ثروت_مسودة_يومية.json + ملخص يوم يوم + الرصيد المتوقع."""
import sys, os, json, re, csv, collections
sys.path.insert(0, 'odoo_templates')
from jrpc import x
D = 'drive/مسودات/'
tsv = {r['#']: r for r in csv.DictReader(open('drive/بنود/ثروت_06-10_مسا.tsv', encoding='utf-8'), delimiter='\t')}
a = json.load(open(D + 'ثروت_اختيارات_06-10.json', encoding='utf-8'))
b = {r['#']: r for r in json.load(open(D + 'ثروت_مقترحات_استكمال_06-10.json', encoding='utf-8'))}
def num(s):
    m = re.search(r'\d+(?:\.\d+)?', str(s)); return float(m.group()) if m else 1.0
def iso(mdy):
    m, d, y = mdy.split('/'); return f'{y}-{int(m):02d}-{int(d):02d}'
اسماء = {}
بنود = []
for k, v in a.items():
    t = tsv[k]; r = b.get(k)
    كود = r['كود'] if r else v['كود']
    سعر = (r['سعر'] if r else v['سعر_معتمد'])
    كمية = num(r.get('كمية', t['الكمية'])) if r else float(v.get('كمية_معتمدة') or num(t['الكمية']))
    بنود.append(dict(رقم=k, تاريخ=iso(t['التاريخ']), نوع=v['نوع'], ورقة=t['البند'], كود=كود,
                     كمية=كمية, سعر=float(سعر or 0), معلّق=bool(r and r['حالة'] != '✅')))
    if v['نوع'] == 'فاتورة يدوية' and k in ('120', '121', '122', '123', '124'):
        بنود[-1]['وصف'] = t['البند'] + ' - واتس اب'   # بأمر صاحب المحل 06/10
# أسماء الأصناف
كودات = sorted({l['كود'] for l in بنود if l['كود']})
for p in x('product.product', 'search_read', [['default_code', 'in', كودات], ['active', 'in', [True, False]]], fields=['default_code', 'name']):
    اسماء[p['default_code']] = p['name']
# فحص سعر المرتجع = آخر سعر بيع لثروت لنفس الصنف قبل تاريخ الرجوع (مرحّل + مسودة)
تاريخ = [(l['date'], l['product_id'][1].split(']')[0].strip('['), l['price_unit']) for l in
          x('account.move.line', 'search_read', [['move_id.partner_id', '=', 44], ['display_type', '=', 'product'], ['parent_state', '=', 'posted'], ['move_id.move_type', '=', 'out_invoice']], fields=['date', 'product_id', 'price_unit'])]
تاريخ += [(l['تاريخ'], l['كود'], l['سعر']) for l in بنود if l['نوع'] == 'بند']
تنبيهات = []
for l in بنود:
    if l['نوع'] != 'مرتجع': continue
    سابق = sorted([t for t in تاريخ if t[1] == l['كود'] and t[0] <= l['تاريخ']])
    if سابق and abs(سابق[-1][2] - l['سعر']) > 1e-6:
        تنبيهات.append(f"{l['رقم']} {l['ورقة']}: سعر المسودة {l['سعر']:.3f} ← آخر بيع لثروت {سابق[-1][2]:.3f}")
        l['سعر'] = سابق[-1][2]
    if not سابق: l['ملاحظة'] = 'ما انباع لثروت قبل — سعر الكتالوج/قراره'
# أسعار المرتجع اليدوي بقرار صاحب المحل (06/10) — للمرتجعات اللي مش مبيّنة ببضاعته الصادرة
يدوي = json.load(open(D + 'ثروت_مرتجع_يدوي_06-10.json', encoding='utf-8'))
for l in بنود:
    if l['رقم'] in يدوي: l['سعر'] = يدوي[l['رقم']]; l['ملاحظة'] = 'مرتجع يدوي — سعر بقرار صاحب المحل'
# تجميع
ايام = collections.defaultdict(list)
for l in بنود: ايام[(l['تاريخ'], l['نوع'] if l['نوع'] != 'فاتورة يدوية' else 'بند', l['رقم'] if l['نوع'] == 'فاتورة يدوية' else '')].append(l)
رصيد_اودو = x('res.partner', 'read', [44], fields=['credit'])[0]['credit']
فواتير = مرتجع = دفعات = صرف = 0
print(f'رصيد ثروت على أودو هلّق: {رصيد_اودو:.3f}\n')
for (d, n, sep), L in sorted(ايام.items()):
    s = sum(l['كمية'] * l['سعر'] for l in L)
    if n == 'بند': فواتير += s
    elif n == 'مرتجع': مرتجع += s
    elif n == 'دفعة': دفعات += s
    elif n == 'صرف نقدي': صرف += s
    علم = ' ❓' if any(l['معلّق'] for l in L) else ''
    print(f"{d} · {n}{' (لحالها)' if sep else ''} · {len(L)} سطر · {s:,.3f}{علم}  [{' '.join(l['رقم'] for l in L)}]")
print(f'\nفواتير {فواتير:,.3f} · مرتجعات −{مرتجع:,.3f} · دفعات −{دفعات:,.3f} · صرف نقدي +{صرف:,.3f}')
print(f'الرصيد المتوقع بعد الترحيل: {رصيد_اودو + فواتير - مرتجع - دفعات + صرف:,.3f}')
if تنبيهات: print('\nتصحيح أسعار مرتجعات:\n  ' + '\n  '.join(تنبيهات))
for l in بنود: l['صنف'] = اسماء.get(l['كود'], '')
json.dump(بنود, open(D + 'ثروت_مسودة_يومية.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
