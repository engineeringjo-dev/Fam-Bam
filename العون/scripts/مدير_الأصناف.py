# -*- coding: utf-8 -*-
"""مدير الأصناف — محلات العون لمواد البناء

البوّابة الوحيدة لأي صنف جديد. قاعدة صاحب المحل (23/09):
«ما يعطيني صنف إلا بصنف شبيه له، وما يخليني أنشئ صنف جديد حي الله».

  python3 scripts/مدير_الأصناف.py "نقاصة نحاس ايطالي 1×11/4"      # فحص
  python3 scripts/مدير_الأصناف.py --عائلة "نقاصة نحاس"             # اسرد العائلة مرتّبة
  python3 scripts/مدير_الأصناف.py --رتب "نقاصة نحاس"               # أعد الترتيب بالمقاس
  python3 scripts/مدير_الأصناف.py --انشاء "…" --سعر 2 --فئة 5 --اكد

المدير ما بيقول «مش موجود» أبداً بدون ما يسرد العائلة كاملة قدّامك.
"""
import sys, os, re, unicodedata
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'odoo_templates'))
from jrpc import x, setname

BAR = '─' * 82

# إملاءات الكتالوج اللي وقعنا فيها فعلياً — أي اسم بينعمله توحيد قبل البحث
مرادفات = [
    ('ادابتر', 'ادبتر'), ('أدابتر', 'ادبتر'),
    ('فلتراب', 'فولتراب'), ('فيلتراب', 'فولتراب'),
    ('فرشاة', 'فرشاي'), ('فرشاه', 'فرشاي'), ('فرشاية', 'فرشاي'),
    ('انجاص', 'نجاص'), ('إنجاص', 'نجاص'),
    ('جلسمبورد', 'جبسمبورد'), ('جبسم بورد', 'جبسمبورد'),
    ('كيزر', 'كيزر'), ('جيزر', 'كيزر'), ('قيزر', 'كيزر'),
    ('ص.م', 'صغير ومرار'), ('ص م', 'صغير ومرار'),
]

def وحّد(s):
    s = unicodedata.normalize('NFKC', (s or '').strip())
    s = re.sub(r'[ً-ْـ]', '', s)
    s = (s.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
           .replace('ى', 'ي').replace('ة', 'ه')
           .replace('”', '"').replace('“', '"').replace('*', '×').replace('x', '×').replace('X', '×'))
    for a, b in مرادفات:
        s = s.replace(وحّد_خام(a), وحّد_خام(b))
    return re.sub(r'\s+', ' ', s)

def وحّد_خام(s):
    s = unicodedata.normalize('NFKC', (s or '').strip())
    s = re.sub(r'[ً-ْـ]', '', s)
    return (s.replace('أ', 'ا').replace('إ', 'ا').replace('آ', 'ا')
             .replace('ى', 'ي').replace('ة', 'ه'))

# ── قراءة المقاس: بالكتالوج 11/4 يعني ١¼ و11/2 يعني ١½ ──
def مقاس(name):
    """يرجّع قائمة أرقام للمقارنة — أول رقم هو الأساس."""
    n = وحّد(name)
    out = []
    for tok in re.findall(r'\d+\s*/\s*\d+|\d+(?:\.\d+)?', n):
        tok = tok.replace(' ', '')
        if '/' in tok:
            a, b = tok.split('/')
            if len(a) == 2 and a[0] == '1':        # 11/4 · 11/2 · 21/2
                out.append(float(a[0]) + int(a[1]) / float(b))
            else:
                out.append(int(a) / float(b))
        else:
            out.append(float(tok))
    return out or [0.0]

def مفتاح_الترتيب(p):
    m = مقاس(p['name'])
    return (m + [0.0] * 4)[:4]

# ── جذر العائلة: أول كلمتين ذات معنى ──
def جذر(name):
    ws = [w for w in وحّد(name).split() if not re.match(r'^[\d/\.×"]+$', w)]
    return ws[:2]

_مخزون = []

def كل_الأصناف():
    """الكتالوج كامل مرة وحدة — الفلترة بتصير محلياً عشان الإملاء ما يضيّع إشي."""
    global _مخزون
    if not _مخزون:
        off = 0
        while True:
            b = x('product.template', 'search_read', [],
                  ['id', 'default_code', 'name', 'list_price', 'categ_id', 'sequence'],
                  limit=1000, offset=off, order='id', context={'active_test': False})
            if not b:
                break
            _مخزون += b
            off += len(b)
            if len(b) < 1000:
                break
        for p in _مخزون:
            p['_ن'] = وحّد(p['name'])
    return _مخزون

def عائلة(name):
    """كل الأصناف اللي بتشارك جذر الاسم — مرتّبة بالمقاس."""
    ws = جذر(name)
    if not ws:
        return []
    كل = [p for p in كل_الأصناف() if ws[0] in p['_ن']]
    if len(ws) > 1:
        ضيّق = [p for p in كل if ws[1] in p['_ن']]
        if ضيّق:
            كل = ضيّق
    return sorted(كل, key=مفتاح_الترتيب)

def اسرد(rows, عنوان):
    print('\n%s\n%s (%d صنف)\n%s' % (BAR, عنوان, len(rows), BAR))
    for p in rows:
        print('  %-8s %-54s @%-7s %s' % (p['default_code'] or '—', p['name'][:54],
              p['list_price'], p['categ_id'][1] if p['categ_id'] else ''))
    if not rows:
        print('  — العائلة فاضية —')

def استخدام(fam, حد=8):
    """سجلّ البيع الفعلي لكل مرشّح — الحكم عند تكافؤ المواصفة والسعر (قاعدة 23/09)."""
    out = []
    for p in fam[:حد]:
        v = x('product.product', 'search_read', [('product_tmpl_id', '=', p['id'])], ['id'])
        if not v:
            continue
        ls = x('account.move.line', 'search_read',
               [('product_id', '=', v[0]['id']), ('parent_state', '=', 'posted'),
                ('move_id.move_type', '=', 'out_invoice')],
               ['partner_id', 'quantity'])
        out.append((p, len(ls), sum(l['quantity'] for l in ls),
                    len({l['partner_id'][0] for l in ls})))
    return sorted(out, key=lambda z: (-z[2], -z[1]))


def تدقيق(name, سعر=None, فئة=None, وحدة=None):
    """تحذيرات مبنية على أخطاء وقعنا فيها فعلاً — بتنطبع قبل أي إنشاء."""
    fam = عائلة(name)
    تحذير = []
    if not fam:
        return تحذير
    # (أ) نفس المقاس باسم مكتوب بطريقة ثانية = تكرار شبه مؤكد
    # — بس للأصناف اللي إلها مقاس رقمي فعلاً؛ اللي بلا مقاس (تب · معجونة …)
    #   كلها بتطلع صفر فبتبيّن «توائم» وهي مش كذلك
    بصمة = مفتاح_الترتيب({'name': name})
    توأم = [p for p in fam if مفتاح_الترتيب(p) == بصمة] if any(بصمة) else []
    if توأم:
        تحذير.append('🔴 نفس المقاس موجود بالعائلة: ' +
                     ' · '.join('%s %s' % (p['default_code'], p['name']) for p in توأم[:3]))
    # (ب) التكرارات الداخلية بالعائلة نفسها — تنظيف لاحق
    from collections import Counter
    ك = Counter(tuple(مفتاح_الترتيب(p)) for p in fam if any(مفتاح_الترتيب(p)))
    مكرر = [m for m, c in ك.items() if c > 1]
    if مكرر:
        تحذير.append('⚠️  العائلة نفسها فيها %d مقاس مكرّر بإملاءين — بدها تنظيف.' % len(مكرر))
    # (ج) الفئة والوحدة والضريبة — لازم تطابق العائلة
    from collections import Counter as C2
    ف = C2(p['categ_id'][0] for p in fam if p['categ_id'])
    if ف and فئة is not None:
        غالب = ف.most_common(1)[0]
        if فئة != غالب[0]:
            اسم = next(p['categ_id'][1] for p in fam if p['categ_id'] and p['categ_id'][0] == غالب[0])
            تحذير.append('⚠️  فئة العائلة الغالبة «%s» (%d من %d) — وانت حاطط فئة %d.'
                         % (اسم, غالب[1], len(fam), فئة))
    ids = [p['id'] for p in fam]
    تفاصيل = x('product.template', 'search_read', [('id', 'in', ids[:80])],
                ['uom_id', 'taxes_id', 'list_price'], context={'active_test': False})
    و = C2(d['uom_id'][0] for d in تفاصيل if d['uom_id'])
    if و and وحدة is not None and وحدة != و.most_common(1)[0][0]:
        اسم = next(d['uom_id'][1] for d in تفاصيل if d['uom_id'] and d['uom_id'][0] == و.most_common(1)[0][0])
        تحذير.append('⚠️  وحدة العائلة الغالبة «%s» — وانت حاطط وحدة %d.' % (اسم, وحدة))
    بلا_ضريبة = [d for d in تفاصيل if not d['taxes_id']]
    if بلا_ضريبة:
        تحذير.append('ℹ️  %d صنف بالعائلة بلا ضريبة شراء — انتبه للاتساق.' % len(بلا_ضريبة))
    # (د) السعر برّا مدى العائلة
    أسعار = [d['list_price'] for d in تفاصيل if d['list_price']]
    if أسعار and سعر:
        lo, hi = min(أسعار), max(أسعار)
        if سعر < lo * 0.5 or سعر > hi * 2:
            تحذير.append('⚠️  السعر %.3f برّا مدى العائلة (%.3f – %.3f) — تأكّد.' % (سعر, lo, hi))
    # (هـ) سعر صفر + نقطة بيع = نفس مشكلة الـ364 صنف
    if سعر is not None and float(سعر) == 0:
        تحذير.append('🔴 سعر صفر — بينقفل عن نقطة البيع تلقائياً لحد ما ينتسعّر.')
    return تحذير


def فحص(name, صامت=False):
    """يرجّع (مطابق, العائلة). ما بيحكي «مش موجود» إلا والعائلة مسرودة."""
    n = وحّد(name)
    مطابق = [p for p in كل_الأصناف() if p['_ن'] == n]
    fam = عائلة(name)
    if not صامت:
        اسرد(fam, 'عائلة «%s»' % ' '.join(جذر(name)))
        print(BAR)
        if مطابق:
            p = مطابق[0]
            print('✅ موجود: %s — %s @%s' % (p['default_code'], p['name'], p['list_price']))
        elif fam:
            print('⚠️  ما لقيت مطابق تماماً — بس العائلة فوق فيها %d صنف.' % len(fam))
            print('    دقّق فيها قبل ما تفكّر بصنف جديد (الإملاء بيختلف: ادبتر/ادابتر · فولتراب/فلتراب · فرشاية/فرشاة).')
        else:
            print('🔴 لا مطابق ولا عائلة. تأكّد من الاسم قبل الإنشاء.')
        for w in تدقيق(name):
            print('    ' + w)
        if len(fam) > 1:
            print('\n  ── سجلّ البيع (الأكثر استخداماً أولاً) ──')
            for p, n, q, k in استخدام(fam):
                print('    %-8s %-44s %d سطر · %g حبة · %d مشروع'
                      % (p['default_code'] or '—', p['name'][:44], n, q, k))
    return (مطابق[0] if مطابق else None), fam

def انشاء(name, سعر, فئة, وحدة=1, ضريبة=29, اكد=False):
    موجود, fam = فحص(name)
    if موجود:
        raise SystemExit('\n⛔ مدير الأصناف: الصنف موجود أصلاً (%s) — ممنوع التكرار.' % موجود['default_code'])
    ت = تدقيق(name, float(سعر), فئة, وحدة)
    if ت:
        print('\n── تدقيق مدير الأصناف ──')
        for w in ت:
            print('  ' + w)
    if any(w.startswith('🔴 نفس المقاس') for w in ت) and '--تجاهل-التوأم' not in sys.argv:
        raise SystemExit('\n⛔ مدير الأصناف: في صنف بنفس المقاس. لو متأكّد إنه مختلف، أضف --تجاهل-التوأم')
    if not اكد:
        raise SystemExit('\n⛔ مدير الأصناف: راجع العائلة فوق. لو فعلاً مش موجود، أعد الأمر مع --اكد')
    # الكود: البادئة الغالبة بالعائلة (لا أول صنف — قد يكون شاذاً)، ثم أول رقم شاغر
    from collections import Counter
    بوادئ = Counter((p['default_code'] or '')[:3] for p in fam
                    if re.fullmatch(r'\d{6}', p['default_code'] or ''))
    بادئة = بوادئ.most_common(1)[0][0] if بوادئ else '900'
    محجوز = {c['default_code'] for c in كل_الأصناف() if c['default_code']}
    محجوز |= {b['barcode'] for b in x('product.template', 'search_read',
              [('barcode', '!=', False)], ['barcode'], limit=6000,
              context={'active_test': False})}
    أرقام = sorted(int(c) for c in محجوز if re.fullmatch(r'%s\d{3}' % بادئة, c or ''))
    كود = (أرقام[-1] if أرقام else int(بادئة + '000')) + 1
    while str(كود) in محجوز:
        كود += 1
    كود = str(كود)
    tid = x('product.template', 'create', [{
        'name': name, 'default_code': كود, 'barcode': كود,
        'categ_id': فئة, 'uom_id': وحدة, 'list_price': float(سعر), 'standard_price': 0.0,
        'taxes_id': [[6, 0, [ضريبة]]], 'available_in_pos': float(سعر) > 0,
        'type': 'consu', 'is_storable': False}])[0]
    setname(tid, name)
    _مخزون.clear()
    print('\n✅ أُنشئ: %s — %s @%s' % (كود, name, سعر))
    رتّب(name)
    return tid

def رتّب(name):
    """يعيد ترتيب العائلة من الأصغر مقاساً للأكبر عبر حقل sequence."""
    fam = عائلة(name)
    if len(fam) < 2:
        print('ℹ️  العائلة أقل من صنفين — ما في ترتيب.')
        return
    for i, p in enumerate(fam, start=1):
        if p['sequence'] != i:
            x('product.template', 'write', [p['id']], {'sequence': i})
    print('\n🔢 انرتّبت العائلة من الأصغر للأكبر (%d صنف):' % len(fam))
    for i, p in enumerate(fam, start=1):
        print('   %2d. %-8s %s' % (i, p['default_code'] or '—', p['name']))

def main(argv):
    if not argv:
        raise SystemExit(__doc__)
    if argv[0] == '--عائلة':
        اسرد(عائلة(argv[1]), 'عائلة «%s»' % argv[1]); return
    if argv[0] == '--رتب':
        رتّب(argv[1]); return
    if argv[0] == '--انشاء':
        g = lambda k, d=None: argv[argv.index(k) + 1] if k in argv else d
        انشاء(argv[1], float(g('--سعر', 0)), int(g('--فئة', 25)),
              int(g('--وحدة', 1)), int(g('--ضريبة', 29)), '--اكد' in argv); return
    فحص(' '.join(argv))

if __name__ == '__main__':
    main(sys.argv[1:])
