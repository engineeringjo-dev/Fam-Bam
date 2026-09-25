# -*- coding: utf-8 -*-
"""مدير الورشات — محلات العون لمواد البناء

مهمّته الأولى بأمر صاحب المحل (25/09):
«**دائماً بس أعطيك ورشة جديدة، شيك من قبل إذا في بنود أُدخلت قبل بحيث ما تتدبّل**».

  python3 scripts/مدير_الورشات.py --كل                         # جرد كل الورشات وحالتها
  python3 scripts/مدير_الورشات.py --ورشة "ثروت"                # حالة ورشة وحدة
  python3 scripts/مدير_الورشات.py --فحص ملف.xlsx --ورشة "العجرمي"   # ملف جديد مقابل المرحّل

المدير ما بيخلّي ولا بند ينزل مرتين، وما بيخلّي ورشة تنفتح بلا مقاول ولا بلا رصيد مفهوم.
"""
import sys, os, re, datetime
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
sys.path.insert(0, os.path.join(ROOT, 'odoo_templates'))
sys.path.insert(0, os.path.join(ROOT, 'scripts'))
from jrpc import x
import importlib.util
_s = importlib.util.spec_from_file_location('mgr', os.path.join(ROOT, 'scripts', 'مدير_الأصناف.py'))
_m = importlib.util.module_from_spec(_s); _s.loader.exec_module(_m)
وحّد = _m.وحّد

BAR = '─' * 82
SERVICES = 28


def ورشات():
    """كل الجهات اللي إلها فواتير بيع — مع رصيدها ومقاوليها ومدى تواريخها."""
    ps = {p['id']: p for p in x('res.partner', 'search_read',
          [('customer_rank', '>', 0)], ['name', 'total_due', 'category_id'])}
    mv = x('account.move', 'search_read',
           [('move_type', 'in', ['out_invoice', 'out_refund']), ('state', '=', 'posted')],
           ['partner_id', 'invoice_date', 'amount_total'])
    tags = {t['id']: t['name'] for t in x('res.partner.category', 'search_read', [], ['name'])}
    out = {}
    for m in mv:
        pid = m['partner_id'][0]
        d = out.setdefault(pid, {'اسم': m['partner_id'][1], 'عدد': 0, 'مبلغ': 0.0, 'تواريخ': []})
        d['عدد'] += 1; d['مبلغ'] += m['amount_total']
        if m['invoice_date']: d['تواريخ'].append(m['invoice_date'])
    for pid, d in out.items():
        p = ps.get(pid, {})
        d['رصيد'] = p.get('total_due', 0.0)
        d['مقاولون'] = [tags.get(t, str(t)) for t in (p.get('category_id') or [])]
    return out


def اسرد_الورشات(فلتر=None):
    o = ورشات()
    بنود = [(d['اسم'], d) for d in o.values()
            if not فلتر or وحّد(فلتر) in وحّد(d['اسم'])]
    if not بنود:
        print('— ما في ورشة بهذا الاسم. الموجود:')
        for d in sorted(o.values(), key=lambda d: d['اسم']): print('   ·', d['اسم'])
        return
    print(BAR)
    print('%-32s %5s %10s %10s  %-22s %s' % ('الورشة', 'فواتير', 'المرحّل', 'الرصيد', 'المقاولون', 'المدى'))
    print(BAR)
    for اسم, d in sorted(بنود, key=lambda z: -z[1]['مبلغ']):
        t = sorted(d['تواريخ'])
        مدى = '%s → %s' % (t[0], t[-1]) if t else '—'
        مقاول = '، '.join(x for x in d['مقاولون'] if x != 'مشروع') or '🔴 بلا مقاول'
        print('%-32s %5d %10.3f %10.3f  %-22s %s' % (اسم[:32], d['عدد'], d['مبلغ'], d['رصيد'], مقاول[:22], مدى))
    print(BAR)


# ═══════════ فحص الازدواج لملف جديد ═══════════
def اقرأ_الملف(path):
    """يقرأ أي ورقة: يمشي على خلايا كل صف، وكل ما يلاقي تاريخ يقفل سجلّاً
    بآخر نص وآخر رقم شافهم — فيمسك النماذج ذات العمود الواحد والقوائم المزدوجة."""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    سجلات = []
    for ws in wb:
        for row in ws.iter_rows(values_only=True):
            نص, رقم = None, None
            for v in row:
                if v is None: continue
                if isinstance(v, datetime.datetime):
                    if نص: سجلات.append({'تاريخ': v.date().isoformat(), 'اسم': نص.strip(), 'كمية': رقم, 'ورقة': ws.title})
                    نص, رقم = None, None
                elif isinstance(v, (int, float)):
                    # أول رقم **بعد** اسم البند هو الكمية — لا آخر رقم (اللي بيطلع السعر)
                    # ورقم التسلسل اللي قبل الاسم بينشطب لأن النص بيصفّر الرقم
                    if رقم is None and نص is not None: رقم = float(v)
                else:
                    s = str(v).strip()
                    m = re.fullmatch(r'(\d{1,2})\s*/+\s*(\d{1,2})\s*/+\s*(\d{4})', s)
                    if m:                       # تاريخ مكتوب نصاً ولو بشرطات زايدة
                        try:
                            dd = datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1))).isoformat()
                            if نص: سجلات.append({'تاريخ': dd, 'اسم': نص.strip(), 'كمية': رقم, 'ورقة': ws.title})
                        except ValueError: pass
                        نص, رقم = None, None
                    elif re.fullmatch(r'[\d.]+\s*(م|متر|ربطة|كغم)?', s):
                        if رقم is None and نص is not None:
                            try: رقم = float(re.sub(r'[^\d.]', '', s) or 0)
                            except ValueError: pass
                    elif len(s) > 2:
                        نص, رقم = s, None
    return سجلات


def أسطر_مرحّلة(pid):
    mv = {m['id']: m for m in x('account.move', 'search_read',
          [('partner_id', '=', pid), ('move_type', 'in', ['out_invoice', 'out_refund']),
           ('state', '=', 'posted')], ['name', 'invoice_date'])}
    if not mv: return []
    ls = x('account.move.line', 'search_read',
           [('move_id', 'in', list(mv)), ('display_type', '=', 'product')],
           ['move_id', 'name', 'quantity', 'price_unit'])
    for l in ls:
        l['تاريخ'] = mv[l['move_id'][0]]['invoice_date']
        l['فاتورة'] = mv[l['move_id'][0]]['name']
        l['ن'] = set(وحّد(re.sub(r'^\[\d+\]\s*', '', l['name'])).split())
    return ls


def فحص_ملف(path, اسم_ورشة):
    o = {d['اسم']: pid for pid, d in ورشات().items()}
    pid = None
    for nm, i in o.items():
        if وحّد(اسم_ورشة) in وحّد(nm): pid, اسم = i, nm; break
    if pid is None:
        ps = x('res.partner', 'search_read', [('name', 'ilike', اسم_ورشة)], ['id', 'name'])
        if ps: pid, اسم = ps[0]['id'], ps[0]['name']
    print(BAR)
    if pid is None:
        print('🆕 ورشة جديدة تماماً — ما في شريك باسم «%s»' % اسم_ورشة)
        print('   كل بنود الملف جديدة. **افتح الشريك أولاً وحدّد المقاول.**')
        اسم, مرحّل = اسم_ورشة, []
    else:
        مرحّل = أسطر_مرحّلة(pid)
        print('الورشة: %s (شريك %d) · أسطر مرحّلة: %d' % (اسم, pid, len(مرحّل)))
    print(BAR)
    سجلات = اقرأ_الملف(path)
    جديد, مكرر, مشتبه = [], [], []
    for r in سجلات:
        ن = set(وحّد(r['اسم']).split())
        أفضل, س = None, 0.0
        for l in مرحّل:
            if l['تاريخ'] != r['تاريخ']: continue
            if r['كمية'] is not None and abs(l['quantity'] - r['كمية']) > 0.001: continue
            # تغطية كلمات الورقة داخل اسم الكتالوج — لا تقاطع/اتحاد،
            # لأن أسماء الكتالوج أطول بكثير («اجو 1/2ك» ⊂ «اجو (اسود) 1/2 ك شنل امريكي»)
            ت = len(ن & l['ن']) / max(1, len(ن))
            if ت > س: أفضل, س = l, ت
        if س >= 0.5: مكرر.append((r, أفضل, س))
        elif س >= 0.3: مشتبه.append((r, أفضل, س))
        else: جديد.append((r, أفضل, س))
    if مكرر:
        print('\n⛔ **بنود مرحّلة سابقاً — ممنوع إعادة إدخالها (%d)**' % len(مكرر))
        for r, l, s in مكرر[:80]:
            print('   %s  %-34s ×%-6s ← %s  %s' % (r['تاريخ'], r['اسم'][:34],
                  ('%g' % r['كمية']) if r['كمية'] is not None else '?', l['فاتورة'], l['name'][:34]))
    if مشتبه:
        print('\n🟡 **مشتبه بها — نفس التاريخ ونفس الكمية بس الاسم مش مطابق (%d)**' % len(مشتبه))
        for r, l, s in مشتبه:
            print('   %s  %-34s ×%-6s ≈ %s  %s' % (r['تاريخ'], r['اسم'][:34],
                  ('%g' % r['كمية']) if r['كمية'] is not None else '?', l['فاتورة'], l['name'][:38]))
        print('   ← راجعها بالعين قبل الترحيل.')
    print('\n🆕 **بنود جديدة (%d)**' % len(جديد))
    for r, _, _ in جديد:
        print('   %s  %-40s ×%s' % (r['تاريخ'], r['اسم'][:40],
              ('%g' % r['كمية']) if r['كمية'] is not None else '🔴 بلا كمية'))
    print('\n' + BAR)
    print('الحصيلة: %d بند بالملف · **%d مرحّل سابقاً** · %d مشتبه · **%d جديد**'
          % (len(سجلات), len(مكرر), len(مشتبه), len(جديد)))
    if مكرر or مشتبه: print('🔴 لا تُرحّل الملف كما هو — رحّل الجديد فقط بعد مراجعة المشتبه.')
    print(BAR)
    return جديد, مكرر, مشتبه


def main(argv):
    if not argv: raise SystemExit(__doc__)
    if '--فحص' in argv:
        path = argv[argv.index('--فحص') + 1]
        ورشة = argv[argv.index('--ورشة') + 1] if '--ورشة' in argv else ''
        فحص_ملف(path, ورشة); return
    if argv[0] == '--كل': اسرد_الورشات(); return
    if argv[0] == '--ورشة': اسرد_الورشات(argv[1]); return
    اسرد_الورشات(' '.join(argv))


if __name__ == '__main__':
    main(sys.argv[1:])
