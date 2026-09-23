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
    return (مطابق[0] if مطابق else None), fam

def انشاء(name, سعر, فئة, وحدة=1, ضريبة=29, اكد=False):
    موجود, fam = فحص(name)
    if موجود:
        raise SystemExit('\n⛔ مدير الأصناف: الصنف موجود أصلاً (%s) — ممنوع التكرار.' % موجود['default_code'])
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
        'taxes_id': [[6, 0, [ضريبة]]], 'available_in_pos': True,
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
