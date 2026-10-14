# -*- coding: utf-8 -*-
"""ترحيل نموذج «التكاليف التشغيلية» على أودو (صاحب المحل 09/10: «نموذج متكامل أودو يقراه ويوزّع كلشي ع حساباته»)

  python3 scripts/ترحيل_التكاليف.py <الملف.xlsx> [--من 2026-10-01] [--الى 2026-10-31]          # تجربة: عرض + فحص (بلا تغيير)
  python3 scripts/ترحيل_التكاليف.py <الملف.xlsx> [--من …] [--الى …] --اكد                     # الترحيل (بعد «رحّل» من صاحب المحل)

كل سطر = قيد لحاله بدفتر العمليات المتنوعة (10) بتاريخه — شيت واحد «التكاليف» (10/10: المواد والأجور سوا):
  مادة (أي غرض غير «أجور عمال»): مدين حساب الغرض (أو الأصل إذا 20+) · دائن: المواد (صيانة المحل · أثاث وستاندات · توالف)
    من البضاعة 500101 · الباقي الصندوق 100901 · «ملاحظات إضافية»: «مشترى/من برا/كاش» ← الصندوق · «تحويل/كليك/شيك» ← البنك 100001.
  أجرة (الغرض «أجور عمال» — العامل والشغلة بالبند زي أي بند · 10/10): مدين 500312 (أو الأصل إذا ورشة 20+) ·
    دائن الصندوق (أو البنك من الملاحظات) · البيان «صرف نقدي — <البند>».
  الورشة (شغل داخلي): مجموعها (مواد + أجور) 20 وفوق ← كل أسطرها مدين 100604 (أصل) · أقل ← حسابها العادي ·
    نفس الاسم بأكثر من كتابة (ة/ه · أ/ا) ← ⛔ وحّد الاسم · اسم ورشة فيه صيانة/تعزيل ← ⛔ ·
    🔴 أسطر الورشة ما بتنرحّل إلا لما يقول «الورشة خلصت» (بتنضاف لـ drive/ورشات_خلصت.json) ← ⏸ لحد هيك.
  الغرض المكتوب بالإيد (مش من القائمة) = «أخرى» · الحساب الدقيق بينختار من البند والملاحظات (REFINE).
الحساب بيتحدّد من جدول «الغرض» بشيت «القوائم» بنفس الملف (مش من المعادلات — بتنحسب هون من جديد).
🔴 بس الأسطر المدقّقة (المدقق معبّى) بتنزل · السطر الناقص بيوقف ⛔ · المرحّل من قبل (نفس الحساب والتاريخ والمبلغ والبيان)
   بيتخطّى ↩ — فتشغيله على الملف كله كل شهر آمن · نسخة تراجع بـ odoo_backup/.
"""
import sys, os, json, re, datetime, openpyxl
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.join(HERE, '..')
sys.path.insert(0, os.path.join(ROOT, 'odoo_templates'))
from jrpc import x

ASSET, MISC_J, WS_ACC = 20, 10, '100604'          # حد الأصل 20 (قراره 10/10)
DONE_FILE = os.path.join(ROOT, 'drive', 'ورشات_خلصت.json')   # الورشات اللي قال إنها خلصت — بس أسطرها بتنرحّل
DONE = set(json.load(open(DONE_FILE, encoding='utf-8'))) if os.path.exists(DONE_FILE) else set()
BAD_WS = ('صيانة', 'تعزيل', 'تنظيف', 'تصليح')                 # الصيانة والتعزيل بلا ورشة (مصروف دايماً)
CASH, BANK, STOCK = '100901', '100001', '500101'
REF = 'مصاريف المحل'
# الدفع (بلا عمود — صاحب المحل 10/10 · وكلمة «من المحل» انلغت 10/10): المواد (صيانة المحل · أثاث وستاندات · توالف بضاعة)
#   افتراضياً من بضاعة المحل 500101 · الباقي نقدي من الصندوق · والملاحظات بتغيّر: «مشترى/من برا/كاش» ← الصندوق · «تحويل/كليك/شيك» ← البنك
STOCK_DEFAULT = {'صيانة المحل', 'أثاث وستاندات', 'توالف بضاعة'}
PAY_WORDS = [(('تحويل', 'كليك', 'شيك', 'من البنك'), BANK, 'بنك'),
             (('مشترى', 'مشتراه', 'من برا', 'من برّا', 'من الخارج', 'كاش', 'نقدي'), CASH, 'نقدي'),
             (('من المحل', 'من البضاعة', 'من الرف'), STOCK, 'من بضاعة المحل')]
# الحساب الدقيق من البند والملاحظات
REFINE = {'فاتورة كهرباء': [(('مي ', 'مياه', 'ماء'), '500403')],
          'إيجار': [(('إيجار المحل', 'ايجار المحل', 'المحل نفسه'), '500401')],
          'صيانة المحل': [(('تنظيف', 'منظف', 'مكنسة', 'مماسح', 'كلور', 'صابون'), '501105')],
          'أخرى': [(('نت ', 'إنترنت', 'انترنت', 'هاتف', 'رصيد', 'اتصالات'), '501101'),
                   (('تنظيف', 'منظف', 'مكنسة', 'مماسح'), '501105'),
                   (('أكياس', 'اكياس', 'كراتين', 'كرتون', 'لاصق', 'قرطاسية', 'دفاتر'), '501106')]}
BAR = '─' * 78

args = sys.argv[1:]
if not args or args[0].startswith('--'): sys.exit(__doc__)
PATH, CONFIRM = args[0], '--اكد' in args
def opt(k):
    return datetime.date.fromisoformat(args[args.index(k) + 1]) if k in args else None
D_FROM, D_TO = opt('--من'), opt('--الى')


def as_date(v):
    if isinstance(v, datetime.datetime): return v.date()
    if isinstance(v, datetime.date): return v
    m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})\s*$', str(v or '').strip())
    return datetime.date(int(m[3]), int(m[2]), int(m[1])) if m else None


def num(v):
    try: return float(v) if v not in (None, '') else None
    except (TypeError, ValueError): return None


wb = openpyxl.load_workbook(PATH); WV = openpyxl.load_workbook(PATH, data_only=True)   # WV: القيم المحسوبة اللي إكسل حفظها
Q = wb['القوائم']
GROUPS = {}
for r in range(2, Q.max_row + 1):
    g = Q.cell(r, 1).value
    if not g: break
    GROUPS[str(g).strip()] = (str(Q.cell(r, 2).value or ''), str(Q.cell(r, 3).value or ''), str(Q.cell(r, 4).value or ''))

rows, errs, held = [], [], []          # (المصدر, التاريخ, الحساب المدين, الحساب الدائن, المبلغ, البيان, النوع)
M = wb['التكاليف']; MV = WV['التكاليف']
# A # · B التاريخ · C اليوم · D الورشة · E البند · F الكمية · G الإفرادي · H الإجمالي · I الغرض · J أصل؟ · K المدقق · L ملاحظات (10/10)
LINES = [r for r in range(7, M.max_row + 1) if M.cell(r, 5).value not in (None, '')]
val = lambda r, c: str(M.cell(r, c).value or '').strip()

# الورشات (شغل داخلي): مجموع كل ورشة (مواد + أجور) · 20 وفوق = كل أسطرها أصل على 100604
norm = lambda w: re.sub(r'\s+', ' ', w.strip().translate(str.maketrans('أإآةى', 'اااهي')))
WS_TOT, WS_NAMES = {}, {}
for r in LINES:
    w = val(r, 4)
    if w:
        WS_TOT[w] = round(WS_TOT.get(w, 0) + (num(M.cell(r, 6).value) or 0) * (num(M.cell(r, 7).value) or 0), 3)
        WS_NAMES.setdefault(norm(w), set()).add(w)
for k, v in WS_NAMES.items():
    if len(v) > 1: errs.append('الورشة مكتوبة بأكثر من شكل: %s — وحّد الاسم' % ' · '.join('«%s»' % n for n in sorted(v)))
WS_ASSET = lambda w: bool(w) and WS_TOT.get(w, 0) >= ASSET

for r in LINES:
    item, w = val(r, 5), val(r, 4)
    where = 'التكاليف سطر %d «%s»' % (r - 6, item)
    q, p, d = num(M.cell(r, 6).value), num(M.cell(r, 7).value), as_date(M.cell(r, 2).value)
    g, chk, note = val(r, 9), val(r, 11), val(r, 12)
    if g.startswith('='):                                  # الغرض التلقائي: ورشة «ستاند» ← أثاث وستاندات
        g = str(MV.cell(r, 9).value or '').strip() or ('أثاث وستاندات' if 'ستاند' in w else '')
    worker = g == 'أجور عمال'                              # الأجرة = غرضها «أجور عمال» (العامل بالبند)
    if not chk: continue                                   # مش مدقّق ← ما بينزل
    if w and any(k in w for k in BAD_WS): errs.append('%s: ورشة «%s» — الصيانة والتعزيل بلا ورشة' % (where, w)); continue
    if w and w not in DONE: held.append((where, w)); continue      # الورشة لسا شغّالة ← بتستنّى لحد ما يقول خلصت
    miss = [n for n, v in (('الكمية', q), ('الإفرادي', p), ('التاريخ', d), ('الغرض', g)) if not v]
    if miss: errs.append('%s: ناقص %s' % (where, ' · '.join(miss))); continue
    if g == 'أخرى' and not note and not w: errs.append('%s: «أخرى» بلا ملاحظات إضافية' % where); continue
    typed = g not in GROUPS                                # غرض مكتوب بالإيد (مثلاً «أكياس للمحل») = أخرى
    key = 'أخرى' if typed else g
    amt = round(q * p, 3); exp, ast, cat = GROUPS[key]; txt = ' %s %s %s ' % (item, g if typed else '', note)
    exp = next((acc for keys, acc in REFINE.get(key, []) if any(k in txt for k in keys)), exp)
    paid = next(((acc, h) for keys, acc, h in PAY_WORDS if any(k in note for k in keys)), None)
    if worker:                                             # أجرة ← من الصندوق (أو البنك من الملاحظات)
        cr, how = paid if paid and paid[0] != STOCK else (CASH, 'نقدي')
        label = 'صرف %s — %s' % ('تحويل كليك' if cr == BANK else 'نقدي', item)
    else:
        cr, how = (STOCK, 'من بضاعة المحل') if key == 'توالف بضاعة' else \
            (paid or ((STOCK, 'من بضاعة المحل') if key in STOCK_DEFAULT else (CASH, 'نقدي')))
        label = item
    if w:
        kind = 'أصل' if WS_ASSET(w) else 'مصروف'; acc = WS_ACC if kind == 'أصل' else exp
    else:
        kind = 'زكاة وصدقة' if cat == 'زكاة وصدقة' else ('أصل' if ast and amt >= ASSET else 'مصروف'); acc = ast if kind == 'أصل' else exp
    rows.append((where, d, acc, cr, amt, label + (' — %s' % w if w else ''),
                 '%s%s · %s · %s%s' % (('ورشة %s %.3f · ' % (w, WS_TOT[w])) if w else '', g, kind, how, (' · ' + note) if note else '')))

rows = [x_ for x_ in rows if (not D_FROM or x_[1] >= D_FROM) and (not D_TO or x_[1] <= D_TO)]
for x_ in rows:
    if 'الدرايف' in x_[5]: errs.append('%s: كلمة «الدرايف» ممنوعة' % x_[0])

codes = sorted({c for x_ in rows for c in x_[2:4]})
ACC = {a['code']: a['id'] for a in x('account.account', 'search_read', [['code', 'in', codes]], fields=['code'])}
for c in codes:
    if c not in ACC: errs.append('الحساب %s مش موجود على أودو' % c)

todo, done = [], []
for x_ in rows:
    if x_[2] in ACC and x('account.move.line', 'search_count', [['account_id', '=', ACC[x_[2]]], ['date', '=', str(x_[1])],
                         ['debit', '=', x_[4]], ['name', '=', x_[5]], ['parent_state', '=', 'posted']]):
        done.append(x_)
    else: todo.append(x_)

print(BAR); print('📄 %s' % os.path.basename(PATH) + ('  (من %s الى %s)' % (D_FROM or '…', D_TO or '…') if D_FROM or D_TO else ''))
print(BAR)
tot = {}
for x_ in sorted(todo, key=lambda y: y[1]):
    print('%s  %10.3f  مدين %s ← دائن %s  %s  [%s]' % (x_[1].strftime('%d/%m/%Y'), x_[4], x_[2], x_[3], x_[5], x_[6]))
    tot[x_[2]] = round(tot.get(x_[2], 0) + x_[4], 3)
print(BAR)
for c, v in sorted(tot.items()): print('  %s  %10.3f' % (c, v))
print('  المجموع  %10.3f  · %d قيد جديد · %d مرحّل من قبل ↩' % (sum(tot.values()), len(todo), len(done)))
for k in sorted({w for _, w in held}): print('⏸ ورشة «%s» شغّالة — أسطرها (%d) بتستنّى لحد ما تقول خلصت' % (k, sum(1 for _, x in held if x == k)))
for e in errs: print('⛔ ' + e)
if errs: sys.exit('⛔ في مشاكل — ولا إشي انرحّل. صلّح الملف وارجع شغّل.')
if not CONFIRM: sys.exit('🔎 تجربة بس — للترحيل: --اكد (بعد «رحّل» من صاحب المحل)')
if not todo: sys.exit('✓ ولا إشي جديد')

stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
LOG = os.path.join(ROOT, 'odoo_backup', 'ROLLBACK_costs_%s.json' % stamp); made = []
for x_ in todo:
    mid = x('account.move', 'create', [{'move_type': 'entry', 'journal_id': MISC_J, 'date': str(x_[1]), 'ref': REF, 'line_ids': [
        (0, 0, {'account_id': ACC[x_[2]], 'name': x_[5], 'debit': x_[4], 'credit': 0}),
        (0, 0, {'account_id': ACC[x_[3]], 'name': x_[5], 'debit': 0, 'credit': x_[4]})]}])
    mid = mid[0] if isinstance(mid, list) else mid
    x('account.move', 'action_post', [mid]); made.append(mid)
    json.dump({'الملف': PATH, 'القيود': made, 'تراجع': 'button_draft ثم button_cancel'}, open(LOG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('✓ انرحّل %d قيد · تراجع %s' % (len(made), os.path.relpath(LOG, ROOT)))
