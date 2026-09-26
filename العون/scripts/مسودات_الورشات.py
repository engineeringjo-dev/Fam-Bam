# -*- coding: utf-8 -*-
"""مسودات الورشات — مسودة وحدة لكل ورشة، بتتحدّث بنداً بنداً من ملف الدرايف (صاحب المحل 26/09)

«بدي التشييك يكون دقيق — شو في تعديل وإضافات عالإكسل كبنود، وضيفهم على مسودة شغلنا لتلك الورشة قبل ما نرحّل كلشي»

  python3 scripts/مسودات_الورشات.py --قارن "ورشة طارق الماضي.xlsx" < نص.txt   # نص read_file_content ← فرق بنود + تحديث المسودة
  python3 scripts/مسودات_الورشات.py --اعرض "ورشة طارق الماضي.xlsx"            # المسودة كاملة: جاهز · معلّق · مرتجعات · فواتير حسب التاريخ
  python3 scripts/مسودات_الورشات.py --كود  "ملف" مفتاح كود [سعر]              # مطابقة سطر بصنف (وسعر إذا بدك)
  python3 scripts/مسودات_الورشات.py --سعر  "ملف" مفتاح سعر                     # سعر يدوي لسطر
  python3 scripts/مسودات_الورشات.py --أجّل "ملف" مفتاح "سبب"                   # تأجيل سطر (أو --ارجع لإلغاء التأجيل)
  python3 scripts/مسودات_الورشات.py --رحّل "ملف" [--اكد]                        # فاتورة لكل تاريخ للجاهز فقط (بلا --اكد = تجربة)

القواعد:
- كل بند بسطره — **ممنوع الدمج** · كل تاريخ فاتورة مستقلة · المرتجع إشعار دائن بتاريخ الرجوع.
- السعر = **سعر الورقة** إذا مكتوب (معتمد لكل الورشات) · وإلا سعر يدوي بقرار صاحب المحل.
- المطابقة التلقائية بس إذا الاسم **مطابق تماماً** بالكتالوج (مدير الأصناف) — غير هيك «بده مطابقة».
- سطر انحذف من الورقة ما بينمسح من المسودة — بينعلّم 🗑️ لحد ما صاحب المحل يقرّر.
- سطر انعدّل بعد ما انرحّل → ⚠️ بينعلّم وما بينلمس.
"""
import sys, os, re, json, csv, io, datetime
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
sys.path.insert(0, os.path.join(ROOT, 'odoo_templates'))
from jrpc import x

لقطات = os.path.join(ROOT, 'drive', 'بنود')
مسودات = os.path.join(ROOT, 'drive', 'مسودات')
JOURNAL = 8
BAR = '─' * 90


# ═══════════ قراءة نص read_file_content ═══════════
def _تاريخ(s):
    m = re.fullmatch(r'\s*(\d{1,2})/(\d{1,2})/(\d{4})\s*', s or '')
    if not m: return None
    try: return datetime.date(int(m.group(3)), int(m.group(1)), int(m.group(2))).isoformat()   # m/d/yyyy
    except ValueError: return None

def _رقم(s):
    s = (s or '').strip()
    m = re.fullmatch(r'([\d.]+)\s*(م|متر|ربطة|حبة|جوز|لفة|ك|باكيت)?', s)
    if not m: return None
    try: return float(m.group(1))
    except ValueError: return None

def _حقول(text):
    """حقول CSV مع احترام الاقتباس — الفاصلة بس."""
    out, cur, q, i = [], [], False, 0
    while i < len(text):
        c = text[i]
        if c == '"':
            if q and i + 1 < len(text) and text[i + 1] == '"': cur.append('"'); i += 1
            else: q = not q
        elif c == ',' and not q:
            out.append(''.join(cur)); cur = []
        else: cur.append(c)
        i += 1
    out.append(''.join(cur))
    return out

def _جدول(text, N):
    """صفوف بعدد أعمدة N — الصف بينفصل عن اللي بعده بمسافة مدموجة بآخر حقل."""
    f = _حقول(text)
    rows, row = [], []
    for tok in f:
        row.append(tok)
        if len(row) == N:
            last = row[-1]
            k = last.rfind(' ')
            if k < 0: rows.append(row); row = []; continue
            row[-1] = last[:k]
            rows.append(row); row = [last[k + 1:]]
    if any(c.strip() for c in row): rows.append(row + [''] * (N - len(row)))
    return rows

def اقرأ_نص(text):
    """يرجّع {'نموذج': 'قالب'|'قديم', 'أسطر': [...]} — كل سطر بمفتاح ثابت."""
    text = text.replace('\\#', '#').replace('\\*', '*').replace('\\_', '_')
    رؤوس = [m for m in re.finditer(r'#,البند,الكمية,الإفرادي,التاريخ,(المستلم|المُرجِع),المدقق ', text)]
    أسطر = []
    if رؤوس:
        for i, h in enumerate(رؤوس):
            نهاية = رؤوس[i + 1].start() if i + 1 < len(رؤوس) else len(text)
            for علامة in ('لا يُرسَل الملف', 'الإفرادي هنا ='):
                k = text.find(علامة, h.end())
                if 0 <= k < نهاية: نهاية = k
            مرتجع = h.group(1) == 'المُرجِع'
            for r in _جدول(text[h.end():نهاية], 7):
                n, بند, كم, سعر, ت, شخص, مدقق = [c.strip() for c in r[:7]]
                if not n.isdigit() or not بند: continue
                أسطر.append({'مفتاح': ('R' if مرتجع else '') + n, 'مرتجع': مرتجع, '#': int(n),
                             'بند': re.sub(r'\s+', ' ', بند), 'كمية_نص': كم, 'كمية': _رقم(كم),
                             'سعر_الورقة': _رقم(سعر), 'تاريخ_نص': ت, 'تاريخ': _تاريخ(ت),
                             'شخص': شخص, 'مدقق': مدقق})
        return {'نموذج': 'قالب', 'أسطر': أسطر}
    # النموذج القديم (المدرسة…): # · البند · … · خارج · التاريخ — بلوكات جنب بعض
    عدّ = {}
    for m in re.finditer(r'(?:(?<=,)|(?<=^)|(?<= ))(\d+),("(?:[^"]|"")*"|[^,"]*),,,([^,]*),(\d{1,2}/\d{1,2}/\d{4})', text):
        بند = re.sub(r'\s+', ' ', m.group(2).strip('"').replace('""', '"')).strip()
        if not بند: continue
        ت = _تاريخ(m.group(4)); كم = m.group(3).strip()
        ك = (ت, بند, كم); عدّ[ك] = عدّ.get(ك, 0) + 1
        أسطر.append({'مفتاح': 'g|%s|%s|%s|%d' % (ت, بند, كم, عدّ[ك]), 'مرتجع': False, '#': int(m.group(1)),
                     'بند': بند, 'كمية_نص': كم, 'كمية': _رقم(كم), 'سعر_الورقة': None,
                     'تاريخ_نص': m.group(4), 'تاريخ': ت, 'شخص': '', 'مدقق': '—'})
    return {'نموذج': 'قديم', 'أسطر': أسطر}


# ═══════════ التخزين ═══════════
def _مسار(مجلد, ملف):
    os.makedirs(مجلد, exist_ok=True)
    return os.path.join(مجلد, re.sub(r'\.xlsx$', '', ملف) + '.json')

def حمّل(مجلد, ملف, افتراضي=None):
    p = _مسار(مجلد, ملف)
    return json.load(open(p)) if os.path.exists(p) else افتراضي

def احفظ(مجلد, ملف, d):
    json.dump(d, open(_مسار(مجلد, ملف), 'w'), ensure_ascii=False, indent=1)

def _الآن():
    try:
        from zoneinfo import ZoneInfo
        return datetime.datetime.now(ZoneInfo('Asia/Amman')).strftime('%Y-%m-%d %H:%M')
    except Exception:
        return (datetime.datetime.utcnow() + datetime.timedelta(hours=3)).strftime('%Y-%m-%d %H:%M')


# ═══════════ الكتالوج ═══════════
_كت = {}
def صنف(كود):
    if كود not in _كت:
        r = x('product.product', 'search_read', [('default_code', '=', كود)], ['id', 'default_code', 'name', 'list_price'])
        _كت[كود] = r[0] if r else None
    return _كت[كود]

_مدير = []
def طابق(بند):
    """مطابقة تامة بس — عبر مدير الأصناف (الكتالوج بيتحمّل مرة وحدة)."""
    try:
        if not _مدير:
            import importlib.util, contextlib
            s = importlib.util.spec_from_file_location('mgr', os.path.join(ROOT, 'scripts', 'مدير_الأصناف.py'))
            m = importlib.util.module_from_spec(s)
            with contextlib.redirect_stdout(io.StringIO()): s.loader.exec_module(m)
            m.كل_الأصناف(); _مدير.append(m)
        m = _مدير[0]; n = m.وحّد(بند)
        ك = [p for p in m.كل_الأصناف() if p['_ن'] == n and p.get('default_code')]
        return ك[0]['default_code'] if len(ك) == 1 else None
    except Exception:
        return None


# ═══════════ الفرق وتحديث المسودة ═══════════
الحقول = ('بند', 'كمية_نص', 'سعر_الورقة', 'تاريخ_نص', 'شخص', 'مدقق')

def قارن(ملف, text):
    جديد = اقرأ_نص(text)
    قديم = حمّل(لقطات, ملف, {'أسطر': []})
    ق = {r['مفتاح']: r for r in قديم['أسطر']}
    ج = {r['مفتاح']: r for r in جديد['أسطر']}
    مضاف = [ج[k] for k in ج if k not in ق]
    محذوف = [ق[k] for k in ق if k not in ج]
    معدّل = [(ق[k], ج[k]) for k in ج if k in ق and any(ق[k].get(f) != ج[k].get(f) for f in الحقول)]

    مس = حمّل(مسودات, ملف) or {'ملف': ملف, 'شريك': None, 'أسطر': []}
    مسم = {l['مفتاح']: l for l in مس['أسطر']}
    أول_مرة = not مس['أسطر']
    for r in مضاف:
        l = dict(r); l.update({'كود': None, 'سعر': r['سعر_الورقة'], 'حالة': '', 'ملاحظة': '',
                                'أُضيف': _الآن(), 'جديد': not أول_مرة})
        l['كود'] = طابق(r['بند'])
        if l['كود'] and l['سعر'] is None:
            p = صنف(l['كود']); l['سعر_مقترح'] = p['list_price'] if p else None
        مس['أسطر'].append(l); مسم[l['مفتاح']] = l
    for a, b in معدّل:
        l = مسم.get(b['مفتاح'])
        if not l: continue
        if l.get('حالة') == 'مرحّل':
            l['تنبيه'] = '⚠️ انعدّل بالورقة بعد الترحيل (%s)' % _الآن(); continue
        if a['بند'] != b['بند']: l['كود'] = طابق(b['بند'])
        if a.get('سعر_الورقة') != b.get('سعر_الورقة'): l['سعر'] = b['سعر_الورقة']
        for f in ('بند', 'كمية_نص', 'كمية', 'سعر_الورقة', 'تاريخ_نص', 'تاريخ', 'شخص', 'مدقق'): l[f] = b[f]
        l['عُدّل'] = _الآن()
    for r in محذوف:
        l = مسم.get(r['مفتاح'])
        if l: l['محذوف_من_الورقة'] = _الآن()
    for r in جديد['أسطر']:            # سطر رجع للورقة بعد ما انحذف
        l = مسم.get(r['مفتاح'])
        if l and l.get('محذوف_من_الورقة'): l.pop('محذوف_من_الورقة')
    احفظ(لقطات, ملف, {'نموذج': جديد['نموذج'], 'وقت': _الآن(), 'أسطر': جديد['أسطر']})
    احفظ(مسودات, ملف, مس)

    def و(r):
        return '%s %s ×%s%s · %s%s' % (r['مفتاح'], r['بند'], r['كمية_نص'] or '؟',
                                       (' @%.3f' % r['سعر_الورقة']) if r.get('سعر_الورقة') is not None else ' بلا سعر',
                                       r['تاريخ_نص'] or '؟', '' if r.get('مدقق') else ' · ⛔ مش مدقّق')
    print(BAR); print('فرق البنود — %s  (%s · %d سطر بالملف)' % (ملف, جديد['نموذج'], len(جديد['أسطر']))); print(BAR)
    if أول_مرة: print('📸 أول لقطة للملف — كل البنود دخلت المسودة كأساس (مش «جديد»)')
    for r in مضاف if not أول_مرة else []: print('🆕 ' + و(r))
    for a, b in معدّل:
        فروق = ['%s: «%s» ← «%s»' % (f, a.get(f), b.get(f)) for f in الحقول if a.get(f) != b.get(f)]
        print('✏️ %s %s — %s' % (b['مفتاح'], b['بند'], ' · '.join(فروق)))
    for r in محذوف: print('🗑️ ' + و(r) + ' — انحذف من الورقة')
    if not (مضاف or معدّل or محذوف): print('✅ ولا بند تغيّر (تعديل شكلي بالملف)')
    print(BAR)
    return مضاف, معدّل, محذوف


# ═══════════ عرض المسودة ═══════════
def حالة(l):
    if l.get('حالة') == 'مرحّل': return 'مرحّل'
    if l.get('محذوف_من_الورقة'): return '🗑️ انحذف من الورقة'
    if l.get('حالة') == 'مؤجّل': return '⏸️ مؤجّل'
    if not l.get('مدقق'): return '⛔ مش مدقّق'
    if not l.get('تاريخ'): return '📅 بلا تاريخ'
    if not l.get('كمية'): return '🔢 بلا كمية'
    if not l.get('كود'): return '🔎 بده مطابقة'
    if l.get('سعر') is None: return '💲 بلا سعر'
    return 'جاهز'

def اعرض(ملف):
    مس = حمّل(مسودات, ملف)
    if not مس: raise SystemExit('— ما في مسودة لـ«%s»' % ملف)
    print(BAR); print('مسودة %s  ·  الشريك: %s' % (ملف, مس.get('شريك') or '🔴 ما إلها شريك')); print(BAR)
    for مرتجع in (False, True):
        ls = [l for l in مس['أسطر'] if l['مرتجع'] == مرتجع]
        if not ls: continue
        print('\n%s' % ('↩️ المرتجعات' if مرتجع else '📦 البنود'))
        print('| مفتاح | التاريخ | الكود | الصنف | كمية | إفرادي | إجمالي | الحالة |')
        print('|--:|---|---|---|--:|--:|--:|---|')
        ك = 0.0; أيام = {}
        for l in sorted(ls, key=lambda l: (l['تاريخ'] or '9', l['#'])):
            ح = حالة(l)
            p = صنف(l['كود']) if l.get('كود') else None
            اسم = p['name'] if p else '«%s»' % l['بند']
            if p and l.get('سعر') is not None and abs(p['list_price'] - l['سعر']) > 0.0005: اسم += ' *(كتالوج %.3f)*' % p['list_price']
            if l.get('سعر') is None and l.get('سعر_مقترح') is not None: اسم += ' *(مقترح كتالوج %.3f)*' % l['سعر_مقترح']
            مج = (l['كمية'] or 0) * (l['سعر'] or 0)
            if ح == 'جاهز': ك += مج; أيام[l['تاريخ']] = أيام.get(l['تاريخ'], 0) + مج
            ت = l['تاريخ'] and '%s/%s' % (l['تاريخ'][8:], l['تاريخ'][5:7]) or l['تاريخ_نص']
            علم = ' 🆕' if l.get('جديد') and ح != 'مرحّل' else ''
            print('| %s%s | %s | %s | %s | %s | %s | %s | %s%s |' % (
                l['مفتاح'], علم, ت, ('`%s`' % l['كود']) if l.get('كود') else '—', اسم, l['كمية_نص'],
                ('%.3f' % l['سعر']) if l.get('سعر') is not None else '—', ('**%.3f**' % مج) if مج else '—',
                ح, (' — ' + l['ملاحظة']) if l.get('ملاحظة') else ''))
        print('\n**الجاهز%s = %.3f**' % (' للإشعار الدائن' if مرتجع else '', ك))
        for d in sorted(أيام): print('  %s/%s  %.3f' % (d[8:], d[5:7], أيام[d]))
        غير = [l for l in ls if حالة(l) not in ('جاهز', 'مرحّل')]
        if غير:
            print('غير جاهز: %d سطر — %s' % (len(غير), ' · '.join(sorted({حالة(l) for l in غير}))))


# ═══════════ تعديلات يدوية ═══════════
def _سطر(مس, مفتاح):
    for l in مس['أسطر']:
        if l['مفتاح'] == مفتاح: return l
    raise SystemExit('— ما في سطر بالمفتاح %s' % مفتاح)

def عدّل(ملف, مفتاح, **kw):
    مس = حمّل(مسودات, ملف); l = _سطر(مس, مفتاح)
    if l.get('حالة') == 'مرحّل': raise SystemExit('⛔ السطر مرحّل — ما بيتعدّل من هون')
    if 'كود' in kw and not صنف(kw['كود']): raise SystemExit('⛔ الكود %s مش موجود' % kw['كود'])
    l.update(kw); احفظ(مسودات, ملف, مس)
    print('✅ %s: %s' % (مفتاح, kw))


# ═══════════ الترحيل ═══════════
def رحّل(ملف, اكد=False):
    مس = حمّل(مسودات, ملف)
    pid = مس.get('شريك')
    if not pid: raise SystemExit('⛔ ما في شريك للورشة')
    جاهز = [l for l in مس['أسطر'] if حالة(l) == 'جاهز']
    موجود = x('account.move', 'search_read', [('partner_id', '=', pid), ('move_type', 'in', ['out_invoice', 'out_refund']),
                                              ('state', '!=', 'cancel')], ['name', 'invoice_date', 'move_type', 'state'])
    مجموعات = {}
    for l in جاهز: مجموعات.setdefault(('out_refund' if l['مرتجع'] else 'out_invoice', l['تاريخ']), []).append(l)
    تعارض = [k for k in مجموعات if any(m['move_type'] == k[0] and m['invoice_date'] == k[1] for m in موجود)]
    if تعارض: raise SystemExit('⛔ فحص الازدواج: في فاتورة بنفس النوع والتاريخ أصلاً: %s' % تعارض)
    print(BAR); print('%s — %s · %d فاتورة/إشعار · %d سطر' % ('ترحيل' if اكد else 'تجربة (بلا --اكد)', ملف, len(مجموعات), len(جاهز)))
    ك = 0.0; سجل = []
    for (نوع, d), ls in sorted(مجموعات.items(), key=lambda z: (z[0][1], z[0][0])):
        مج = round(sum(l['كمية'] * l['سعر'] for l in ls), 3)
        ك += مج * (-1 if نوع == 'out_refund' else 1)
        print('  %s %s  %d سطر  %.3f' % ('↩️ إشعار دائن' if نوع == 'out_refund' else '🧾 فاتورة', d, len(ls), مج))
        if not اكد: continue
        lines = [[0, 0, {'product_id': صنف(l['كود'])['id'], 'name': صنف(l['كود'])['name'],
                         'quantity': l['كمية'], 'price_unit': l['سعر'], 'tax_ids': [[6, 0, []]]}]
                 for l in sorted(ls, key=lambda l: l['#'])]
        mid = x('account.move', 'create', [{'move_type': نوع, 'partner_id': pid, 'journal_id': JOURNAL,
                                            'invoice_date': d, 'date': d, 'invoice_line_ids': lines,
                                            'ref': ('مرتجع بضاعة — ' if نوع == 'out_refund' else '') + re.sub(r'\.xlsx$', '', ملف)}])[0]
        فعلي = x('account.move', 'read', [mid], ['amount_total'])[0]['amount_total']
        assert abs(فعلي - مج) < 0.001, 'خلل بالمجموع %s: %s ≠ %s' % (d, فعلي, مج)
        x('account.move', 'action_post', [[mid]])
        m = x('account.move', 'read', [mid], ['name', 'state'])[0]
        assert m['state'] == 'posted'
        سجل.append({'id': mid, 'name': m['name'], 'date': d, 'type': نوع, 'total': مج})
        for l in ls: l['حالة'] = 'مرحّل'; l['فاتورة'] = m['name']; l.pop('جديد', None)
        احفظ(مسودات, ملف, مس)
        print('     ✅ %s' % m['name'])
    print('الصافي %.3f' % ك); print(BAR)
    if اكد and سجل:
        os.makedirs(os.path.join(ROOT, 'odoo_backup'), exist_ok=True)
        json.dump(سجل, open(os.path.join(ROOT, 'odoo_backup', 'ROLLBACK_%s_%s.json' % (
            re.sub(r'\.xlsx$', '', ملف).replace(' ', '_'), _الآن()[:10])), 'w'), ensure_ascii=False, indent=1)


def main(a):
    if len(a) < 2: raise SystemExit(__doc__)
    cmd, ملف = a[0], a[1]
    if cmd == '--قارن': قارن(ملف, sys.stdin.read()); return
    if cmd == '--اعرض': اعرض(ملف); return
    if cmd == '--كود':
        kw = {'كود': a[3]}
        if len(a) > 4: kw['سعر'] = float(a[4])
        عدّل(ملف, a[2], **kw); return
    if cmd == '--سعر': عدّل(ملف, a[2], سعر=float(a[3])); return
    if cmd == '--أجّل': عدّل(ملف, a[2], حالة='مؤجّل', ملاحظة=' '.join(a[3:])); return
    if cmd == '--ارجع': عدّل(ملف, a[2], حالة=''); return
    if cmd == '--شريك':
        مس = حمّل(مسودات, ملف); مس['شريك'] = int(a[2]); احفظ(مسودات, ملف, مس); print('✅ الشريك', a[2]); return
    if cmd == '--رحّل': رحّل(ملف, '--اكد' in a); return
    raise SystemExit(__doc__)


if __name__ == '__main__':
    main(sys.argv[1:])
