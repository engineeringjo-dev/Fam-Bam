# -*- coding: utf-8 -*-
"""الترحيل الموحّد — من مهام مدير الورشات (صاحب المحل 09/10: «كل ما أقول رحّل هو ببلّش»)

سكربت واحد لكل الورش بدل سكربت لكل ورشة. بيقرا مسودة JSON موحّدة، وبيفحص قبل ما يرحّل،
وبيرحّل بقواعد الصياغة، وبعدين بيشغّل مدير الكشوفات على الجهة.

  python3 scripts/ترحيل_موحّد.py drive/مسودات/<المسودة>.json            # تجربة: عرض + فحص (بلا تغيير)
  python3 scripts/ترحيل_موحّد.py drive/مسودات/<المسودة>.json --اكد      # الترحيل (بعد «رحّل» من صاحب المحل)

شكل المسودة:
  {"شريك": 48  أو  null,
   "شريك_جديد": {"اسم": "ورشة …", "تصنيف": 13},          ← بس إذا ما إله حساب · 13 «بدون مقاول» (الافتراضي) · 15 «طارق» = مقاول طارق بس (13/10)
   "مراجع": {"2026-10-08": "محيسن"},                         ← اختياري: ملاحظة على الفاتورة كلها (بحكي صاحب المحل)
   "أسطر": [
     {"نوع": "بند",   "تاريخ": "2026-10-08", "كود": "200083", "كمية": 5, "سعر": 0.2, "ملاحظة": "اسود"},
     {"نوع": "مرتجع", "تاريخ": "2026-10-08", "كود": "200083", "كمية": 1, "سعر": 0.2},
     {"نوع": "دفعة",  "تاريخ": "2026-10-08", "مبلغ": 15, "دفتر": "نقدي" | "كليك", "ملاحظة": "دفعة نقدية من (…) الى (…)"},
     {"نوع": "صرف",   "تاريخ": "2026-10-08", "مبلغ": 20, "دفتر": "نقدي" | "كليك", "ملاحظة": "صرف نقدي الى (…)"}]}

🔴 الحماية (أي واحد منهم بيوقف الترحيل كله — ولا إشي بينرحّل):
  1. **التكرار**: نفس الجهة · نفس التاريخ · نفس الصنف والكمية مرحّل من قبل ← ⛔ (درس أبو محمود 08/10)
     وأي فاتورة/مرتجع مرحّل بنفس التاريخ للجهة ← ⛔ (كل تاريخ = فاتورة وحدة) إلا مع --نفس_اليوم بأمره.
     دفعة بنفس التاريخ والمبلغ والدفتر ← ⛔.
  2. الصنف لازم يكون موجود وفعّال · الكمية والسعر أكبر من صفر.
  3. صيغة الدفعة لازم تطابق الصيغة الموحّدة (من مدير الكشوفات) · ممنوع كلمة «الدرايف» بأي مكان.
القواعد: اسم البند = اسم الصنف (+ الملاحظة بسطر ثاني جوّا الخانة) · المرجع فاضي إلا اللي بالمسودة ·
كل تاريخ = فاتورة وحدة · مرتجعات نفس اليوم = إشعار دائن واحد · نسخة تراجع أول بأول بـ odoo_backup/.
"""
import sys, os, json, re, datetime, collections, importlib.util
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.join(HERE, '..')
sys.path.insert(0, os.path.join(ROOT, 'odoo_templates'))
from jrpc import x

TAX = [(6, 0, [29])]; SALES_J = 8
PAYJ = {('نقدي', 'inbound'): (15, 6), ('نقدي', 'outbound'): (15, 5), ('كليك', 'inbound'): (13, 3), ('كليك', 'outbound'): (13, 4)}
BAR = '─' * 78

spec = importlib.util.spec_from_file_location('مك', os.path.join(HERE, 'مدير_الكشوفات.py'))
مك = importlib.util.module_from_spec(spec); _argv = sys.argv; sys.argv = ['x']; spec.loader.exec_module(مك); sys.argv = _argv


def افحص(d, نفس_اليوم=False):
    """بيرجّع (موانع, ملاحظات, خطة)"""
    موانع, ملاحظات = [], []
    P = d.get('شريك')
    if not P and not d.get('شريك_جديد'): موانع.append('ما في شريك ولا «شريك_جديد»')
    if P and not x('res.partner', 'search_count', [('id', '=', P)]): موانع.append('الشريك %s مش موجود' % P)
    if d.get('شريك_جديد') and x('res.partner', 'search', [('name', '=', d['شريك_جديد']['اسم'])]):
        موانع.append('في حساب اسمه «%s» من قبل — استعمل رقمه' % d['شريك_جديد']['اسم'])
    L = d['أسطر']
    if 'الدرايف' in json.dumps(d, ensure_ascii=False): موانع.append('كلمة «الدرايف» موجودة بالمسودة — ممنوعة')
    codes = sorted({l['كود'] for l in L if l['نوع'] in ('بند', 'مرتجع') and l.get('كود')})
    prod = {p['default_code']: p for p in x('product.product', 'search_read', [('default_code', 'in', codes)], ['id', 'name', 'default_code'])}
    for l in L:
        if l['نوع'] in ('بند', 'مرتجع'):
            if not l.get('كود'): موانع.append('بلا صنف: «%s» (%s) — بدها قرار' % (l.get('ورقة', '?'), l['تاريخ']))
            elif l['كود'] not in prod: موانع.append('الصنف %s مش موجود/مش فعّال (%s)' % (l['كود'], l['تاريخ']))
            if not (l.get('كمية') or 0) > 0 or not (l.get('سعر') or 0) > 0: موانع.append('كمية/سعر صفر: %s %s' % (l['كود'], l['تاريخ']))
        elif l['نوع'] in ('دفعة', 'صرف'):
            if l.get('دفتر') not in ('نقدي', 'كليك'): موانع.append('الدفتر لازم «نقدي» أو «كليك»: %s' % l)
            نوع = 'in' if l['نوع'] == 'دفعة' else 'out'
            if not مك.صيغ_الدفعات[نوع].match(' '.join((l.get('ملاحظة') or '').split())):
                موانع.append('صيغة الدفعة مش موحّدة: «%s»' % l.get('ملاحظة'))
        else: موانع.append('نوع مش معروف: %s' % l['نوع'])
    # الخطة: كل تاريخ فاتورة · مرتجعات نفس اليوم إشعار واحد
    G = collections.OrderedDict()
    for l in sorted([l for l in L if l['نوع'] in ('بند', 'مرتجع')], key=lambda l: l['تاريخ']):
        G.setdefault(('out_invoice' if l['نوع'] == 'بند' else 'out_refund', l['تاريخ']), []).append(l)
    # 🔴 فحص التكرار مقابل أودو
    if P:
        for (t, dt), ls in G.items():
            old = x('account.move', 'search_read', [('partner_id', '=', P), ('move_type', '=', t), ('invoice_date', '=', dt), ('state', '!=', 'cancel')], ['id', 'name'])
            if not old: continue
            ol = x('account.move.line', 'search_read', [('move_id', 'in', [m['id'] for m in old]), ('display_type', '=', 'product')], ['move_id', 'product_id', 'quantity'])
            have = {(o['product_id'] and o['product_id'][0], round(o['quantity'], 3)): o['move_id'][1] for o in ol}
            for l in ls:
                k = (prod.get(l['كود'], {}).get('id'), round(l['كمية'], 3))
                if k in have: موانع.append('🔁 مكرّر: %s ×%g بتاريخ %s مرحّل من قبل بـ %s' % (l['كود'], l['كمية'], dt, have[k]))
            if not نفس_اليوم:
                موانع.append('في %s مرحّل للجهة بتاريخ %s (%s) — كل تاريخ فاتورة وحدة · بدها قرار (--نفس_اليوم)' %
                             ('فاتورة' if t == 'out_invoice' else 'مرتجع', dt, ' · '.join(m['name'] for m in old)))
        for l in L:
            if l['نوع'] not in ('دفعة', 'صرف'): continue
            j, _ = PAYJ[(l['دفتر'], 'inbound' if l['نوع'] == 'دفعة' else 'outbound')]
            same = x('account.payment', 'search_read', [('partner_id', '=', P), ('date', '=', l['تاريخ']), ('amount', '=', l['مبلغ']),
                                                         ('journal_id', '=', j), ('state', '!=', 'cancel')], ['name'])
            if same: موانع.append('🔁 دفعة مكرّرة: %.3f بتاريخ %s موجودة (%s)' % (l['مبلغ'], l['تاريخ'], same[0]['name']))
    return موانع, ملاحظات, G, prod


def اعرض(d, G, prod):
    tot = 0
    for (t, dt), ls in G.items():
        s = round(sum(l['كمية'] * l['سعر'] for l in ls), 3); tot += s if t == 'out_invoice' else -s
        print('  %s %s — %.3f%s' % ('فاتورة بضاعة' if t == 'out_invoice' else 'مرتجع بضاعة', dt, s,
                                    ('  · مرجع: %s' % d.get('مراجع', {}).get(dt)) if d.get('مراجع', {}).get(dt) else ''))
        for l in ls:
            print('      [%s] %s ×%g @%.3f%s' % (l['كود'], prod.get(l['كود'], {}).get('name', '؟'), l['كمية'], l['سعر'], ('  ⏎ ' + l['ملاحظة']) if l.get('ملاحظة') else ''))
    for l in d['أسطر']:
        if l['نوع'] in ('دفعة', 'صرف'):
            tot += -l['مبلغ'] if l['نوع'] == 'دفعة' else l['مبلغ']
            print('  %s %s — %.3f · %s' % (l['نوع'], l['تاريخ'], l['مبلغ'], l.get('ملاحظة')))
    print('  صافي اللي رح ينضاف على الرصيد: %.3f' % tot)
    return tot


def رحّل(path, d, G, prod):
    stamp = datetime.datetime.now().strftime('%d-%m_%H%M%S')
    LOG = os.path.join(ROOT, 'odoo_backup', 'ROLLBACK_post_%s_%s.json' % (re.sub(r'\W+', '_', os.path.basename(path))[:30], stamp))
    log = {'مسودة': path, 'partner': None, 'moves': [], 'payments': []}
    save = lambda: json.dump(log, open(LOG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    P = d.get('شريك')
    if not P:
        nw = d['شريك_جديد']
        P = x('res.partner', 'create', {'name': nw['اسم'], 'category_id': [(6, 0, [nw.get('تصنيف', 13)])], 'customer_rank': 1})
        log['partner_created'] = True
    log['partner'] = P; log['before'] = x('res.partner', 'read', [P], ['credit'])[0]['credit']; save()
    for (t, dt), ls in G.items():
        vals = {'move_type': t, 'partner_id': P, 'invoice_date': dt, 'date': dt, 'journal_id': SALES_J,
                'invoice_line_ids': [(0, 0, {'product_id': prod[l['كود']]['id'],
                                             'name': prod[l['كود']]['name'] + (('\n' + l['ملاحظة']) if l.get('ملاحظة') else ''),
                                             'quantity': l['كمية'], 'price_unit': l['سعر'], 'tax_ids': TAX}) for l in ls]}
        ref = d.get('مراجع', {}).get(dt)
        if ref and t == 'out_invoice': vals['ref'] = ref
        mid = x('account.move', 'create', vals)
        exp = round(sum(l['كمية'] * l['سعر'] for l in ls), 3)
        got = x('account.move', 'read', [mid], ['amount_total'])[0]['amount_total']
        assert abs(got - exp) < 0.001, (dt, got, exp)
        x('account.move', 'action_post', [mid])
        log['moves'].append({'id': mid, 'name': x('account.move', 'read', [mid], ['name'])[0]['name'], 'date': dt, 'amount': exp}); save()
        print('  ✓', log['moves'][-1]['name'], dt, '%.3f' % exp)
    for l in d['أسطر']:
        if l['نوع'] not in ('دفعة', 'صرف'): continue
        tp = 'inbound' if l['نوع'] == 'دفعة' else 'outbound'; j, ml = PAYJ[(l['دفتر'], tp)]
        pid = x('account.payment', 'create', {'partner_id': P, 'partner_type': 'customer', 'payment_type': tp, 'journal_id': j,
                                              'payment_method_line_id': ml, 'amount': l['مبلغ'], 'date': l['تاريخ'], 'memo': l['ملاحظة']})
        x('account.payment', 'action_post', [pid])
        log['payments'].append({'id': pid, 'name': x('account.payment', 'read', [pid], ['name'])[0]['name'], 'date': l['تاريخ'], 'amount': l['مبلغ']}); save()
        print('  ✓', log['payments'][-1]['name'], l['تاريخ'], '%.3f' % l['مبلغ'])
    log['after'] = x('res.partner', 'read', [P], ['credit'])[0]['credit']; save()
    print('  رصيد [%d]: %.3f ← %.3f · تراجع: %s' % (P, log['before'], log['after'], os.path.relpath(LOG, ROOT)))
    return P


def main():
    a = [s for s in sys.argv[1:] if not s.startswith('--')]
    if len(a) != 1: raise SystemExit(__doc__)
    path = a[0]; d = json.load(open(path, encoding='utf-8'))
    print(BAR); print('  مدير الورشات — الترحيل الموحّد · %s' % os.path.basename(path)); print(BAR)
    موانع, ملاحظات, G, prod = افحص(d, '--نفس_اليوم' in sys.argv)
    اعرض(d, G, prod)
    if موانع:
        print('\n⛔ الترحيل موقوف — ولا إشي انرحّل:'); [print('   •', m) for m in موانع]; print(BAR); return 1
    if '--اكد' not in sys.argv:
        print('\n✅ الفحص نظيف (ولا تكرار) — للترحيل: أضف --اكد بعد «رحّل» من صاحب المحل'); print(BAR); return 0
    P = رحّل(path, d, G, prod)
    print('\n── مدير الكشوفات ──'); os.system('python3 "%s" %d | tail -4' % (os.path.join(HERE, 'مدير_الكشوفات.py'), P))
    print(BAR); return 0


if __name__ == '__main__':
    sys.exit(main())
