# -*- coding: utf-8 -*-
"""قارئ ملف «نظام الموظفين» السنوي الأسبوعي (النسخة 13) — لما صاحب المحل يرفع الملف (بأمره 30/09 و01/10)

  python3 scripts/قارئ_نظام_الموظفين.py "ملف.xlsx"                 # كل الأسابيع لحد اليوم — مسودة بس
  python3 scripts/قارئ_نظام_الموظفين.py "ملف.xlsx" --شهر 10        # أسابيع شهر 10 (حسب الجمعة = يوم القبض)
  python3 scripts/قارئ_نظام_الموظفين.py "ملف.xlsx" --لحد-أسبوع 41  # لحد الأسبوع 41

شو بيطلّع:
  1) فواتير بيع (البضاعة): كل سطر «النوع = بضاعة» بورقة «القروض والديون» ← فاتورة بيع باسم الموظف
     بتاريخ السطر (البيان فيه الصنف والكمية ← بيتطابق مع مدير الأصناف قبل الترحيل).
  2) التحصيل: أقساط كل أسبوع (من «دفتر الأسابيع» عمود «قسط الدين») بتاريخ جمعة الأسبوع + «دفع كاش من جيبته» —
     بتتوزّع على ديون الموظف **الأقدم فالأحدث**: اللي بيسدّ بضاعة = تحصيل مبيعات،
     اللي بيسدّ قرض كاش = استرداد قرض (مش مبيعات)، واللي بيسدّ دين قديم = تحصيل دين قديم.
  3) أرقام الشهر من ملخص الأيدي العاملة للعلم.

القواعد:
  • ولا إشي بينرحّل على أودو بلا «رحّل» من صاحب المحل.
  • تشييك التكرار: كل سطر إله مفتاح (سنة · أسبوع · موظف · رقم السطر · مبلغ) بيتسجّل بـ drive/ذمم_الموظفين.json
    بعد الترحيل — نفس الملف بينرفع كل شهر، فاللي انرحّل ما بينعاد. ولو أسبوع مرحّل تغيّر رقمه ← تنبيه ⚠️.
  • الموظف لازم يكون شريك على أودو (عميل) — أول مرة بينفتح بموافقته.
"""
import sys, os, json, shutil, subprocess, tempfile, datetime as dt
import openpyxl

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
REG = os.path.join(ROOT, 'drive', 'ذمم_الموظفين.json')
EMPS = ['عمر المصري', 'عبدالعزيز']
RECALC = '/mnt/skills/public/xlsx/scripts/recalc.py'
BAR = '─' * 78
NWMAX = 53


def حمّل_السجل():
    if os.path.exists(REG): return json.load(open(REG))
    return {'مرحّل': {}, 'أقساط_مرحّلة': {}, 'ديون': {e: [] for e in EMPS}}


def افتح(path):
    """القيم المحسوبة — لو الملف ما انحفظ من إكسل (بلا قيم) بنحسبه بنسخة مؤقتة."""
    wb = openpyxl.load_workbook(path, data_only=True)
    if wb['القوائم']['G2'].value is None and os.path.exists(RECALC):
        tmp = os.path.join(tempfile.mkdtemp(), 'x.xlsx'); shutil.copy(path, tmp)
        subprocess.run([sys.executable, RECALC, tmp, '180'], capture_output=True)
        wb = openpyxl.load_workbook(tmp, data_only=True)
    return wb


def تاريخ(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    if isinstance(v, (int, float)) and v > 30000: return (dt.datetime(1899, 12, 30) + dt.timedelta(days=v)).date()
    return None


def رقم(v):
    try: return round(float(v), 3)
    except (TypeError, ValueError): return 0.0


def اقرأ(path):
    wb = افتح(path)
    H = wb['القوائم']
    if H['P1'].value != 'ALAWN_PAYROLL' or H['P3'].value != 'سنوي':
        sys.exit('⛔ هاد مش ملف «نظام الموظفين» السنوي المعتمد (علامة النموذج ناقصة أو نسخة شهرية قديمة).')
    سنة = int(H['G4'].value); nw = int(H['G2'].value); حالي = int(H['G3'].value)
    جمعة = {k + 1: تاريخ(H['I%d' % (11 + k)].value) for k in range(NWMAX) if تاريخ(H['I%d' % (11 + k)].value)}
    L = wb['القروض والديون']
    سطور = []
    for r in range(11, L.max_row + 1):
        emp, typ, amt, rep = L['C%d' % r].value, L['D%d' % r].value, رقم(L['E%d' % r].value), رقم(L['G%d' % r].value)
        if emp not in EMPS or (not amt and not rep): continue
        wk = L['L%d' % r].value
        سطور.append({'صف': r, 'تاريخ': تاريخ(L['B%d' % r].value), 'أسبوع': int(wk) if isinstance(wk, (int, float)) else None,
                     'موظف': emp, 'نوع': typ or '⚠️ بلا نوع', 'مبلغ': amt, 'من_جيبته': rep, 'بيان': (L['H%d' % r].value or '').strip()})
    # الأقساط من دفتر الأسابيع: عمود A رقم الأسبوع · G القسط · I المدفوع — جدول لكل موظف تحت عنوانه
    W = wb['دفتر الأسابيع']; أقساط = {e: {} for e in EMPS}; مدفوع = {e: {} for e in EMPS}; cur = None
    for r in range(1, W.max_row + 1):
        a = W['A%d' % r].value
        if a in EMPS: cur = a; continue
        if cur and isinstance(a, (int, float)) and 1 <= a <= nw:
            أقساط[cur][int(a)] = رقم(W['G%d' % r].value); مدفوع[cur][int(a)] = رقم(W['I%d' % r].value)
    Bs = wb['بداية السنة']
    افتتاح = {e: {'بداية': int(Bs['B%d' % (6 + i)].value or 1), 'دين': رقم(Bs['D%d' % (6 + i)].value)} for i, e in enumerate(EMPS)}
    M = wb['ملخص الأيدي العاملة']
    return سنة, nw, حالي, جمعة, سطور, أقساط, مدفوع, افتتاح, M


def وزّع(ديون, مبلغ):
    out = []
    for d in ديون:
        if مبلغ <= 0.0005: break
        x = min(d['باقي'], مبلغ)
        if x > 0.0005:
            d['باقي'] = round(d['باقي'] - x, 3); مبلغ = round(مبلغ - x, 3); out.append((d['نوع'], x, d.get('مرجع', '')))
    if مبلغ > 0.0005: out.append(('⚠️ زيادة عن الدين', مبلغ, ''))
    return out


def main():
    if len(sys.argv) < 2: sys.exit(__doc__)
    path = sys.argv[1]; a = sys.argv[2:]
    شهر = int(a[a.index('--شهر') + 1]) if '--شهر' in a else None
    حد = int(a[a.index('--لحد-أسبوع') + 1]) if '--لحد-أسبوع' in a else None
    سنة, nw, حالي, جمعة, سطور, أقساط, مدفوع, افتتاح, M = اقرأ(path)
    reg = حمّل_السجل()
    if حد is None: حد = حالي
    أسابيع = [w for w in range(1, حد + 1) if w in جمعة and (شهر is None or جمعة[w].month == شهر)]
    print(BAR); print('ملف نظام الموظفين — سنة %d · %d أسبوع · الأسبوع الحالي %d' % (سنة, nw, حالي))
    print('النطاق: ' + ('شهر %02d (الأسابيع اللي جمعتها فيه: %s)' % (شهر, ', '.join(map(str, أسابيع))) if شهر else 'الأسابيع 1–%d' % حد)); print(BAR)

    # 1) البضاعة ← فواتير بيع
    print('\n🧾 بضاعة ← فواتير بيع باسم الموظف')
    بضاعة = [s for s in سطور if s['نوع'] == 'بضاعة' and s['مبلغ'] and s['أسبوع'] in أسابيع]
    for s in بضاعة:
        key = '%d|%d|%s|%d|%.3f' % (سنة, s['أسبوع'], s['موظف'], s['صف'], s['مبلغ'])
        حالة = '✔ مرحّل من قبل' if key in reg['مرحّل'] else ('⚠️ البيان فاضي — بدنا الصنف والكمية' if not s['بيان'] else 'جاهز للمطابقة')
        print('  أسبوع %2d · %s · %-11s · %8.3f · %s  ← %s' % (s['أسبوع'], s['تاريخ'].strftime('%d/%m'), s['موظف'], s['مبلغ'], s['بيان'] or '—', حالة))
    if not بضاعة: print('  — ما في بضاعة بهالنطاق')
    for s in سطور:
        if s['نوع'] == '⚠️ بلا نوع' and s['مبلغ']:
            print('  ⚠️ صف %d: %s %.3f بلا «النوع» — لازم يتحدّد قبل الترحيل' % (s['صف'], s['موظف'], s['مبلغ']))
    for s in سطور:
        if s['مبلغ'] and s['أسبوع'] is not None and (s['أسبوع'] < 1 or s['أسبوع'] > nw):
            print('  ⛔ صف %d: تاريخه برّا السنة (أسبوع %s)' % (s['صف'], s['أسبوع']))

    # 2) التحصيل — كل الأحداث من أول السنة (عشان الترتيب الأقدم فالأحدث يضبط)، وبنطبع اللي بالنطاق بس
    print('\n💵 التحصيل (أقساط بتاريخ جمعة الأسبوع + دفع من جيبته) — الأقدم فالأحدث')
    for e in EMPS:
        ديون = [dict(d) for d in reg['ديون'].get(e, [])]
        if not ديون and افتتاح[e]['دين'] > 0.0005: ديون = [{'نوع': 'دين قديم', 'باقي': افتتاح[e]['دين'], 'مرجع': 'رصيد افتتاح'}]
        أحداث = [(s['تاريخ'], 0, 'دين', s) for s in سطور if s['موظف'] == e and s['مبلغ'] and s['تاريخ']]
        أحداث += [(s['تاريخ'], 1, 'جيبته', s) for s in سطور if s['موظف'] == e and s['من_جيبته'] and s['تاريخ']]
        أحداث += [(جمعة[w], 2, 'قسط', (w, v)) for w, v in أقساط[e].items() if v and w in جمعة]
        أحداث.sort(key=lambda t: (t[0], t[1]))
        مجاميع = {}; باقي_قبل = None
        for d, _, kind, obj in أحداث:
            if kind == 'دين':
                ديون.append({'نوع': obj['نوع'], 'باقي': obj['مبلغ'], 'مرجع': 'صف %d' % obj['صف']}); continue
            wk = obj[0] if kind == 'قسط' else obj['أسبوع']
            مبلغ, وصف = (obj['من_جيبته'], 'دفع من جيبته (صف %d)' % obj['صف']) if kind == 'جيبته' else (obj[1], 'قسط الأسبوع %d' % obj[0])
            داخل = wk in أسابيع
            key = '%d|%d|%s|%s|%.3f' % (سنة, wk or 0, e, وصف, مبلغ)
            مرحّل = key in reg['أقساط_مرحّلة']
            for typ, x, ref in وزّع(ديون, مبلغ):
                اسم = {'بضاعة': 'تحصيل مبيعات', 'قرض كاش': 'استرداد قرض (مش مبيعات)', 'دين قديم': 'تحصيل دين قديم'}.get(typ, typ)
                if داخل:
                    print('  %s · %-11s · %-26s · %8.3f ← %s %s%s' % (d.strftime('%d/%m'), e, وصف, x, اسم, ('(' + ref + ')') if ref else '', '  ✔ مرحّل' if مرحّل else ''))
                    مجاميع[اسم] = round(مجاميع.get(اسم, 0) + x, 3)
        باقي = round(sum(d['باقي'] for d in ديون), 3)
        if مجاميع or باقي:
            print('  ⟵ %s: %s · الدين الباقي آخر السنة/اليوم %.3f' % (e, ' · '.join('%s %.3f' % kv for kv in مجاميع.items()) or 'ما في تحصيل بالنطاق', باقي))
    # تنبيه: أسبوع مرحّل تغيّر
    for e in EMPS:
        for w, v in أقساط[e].items():
            old = reg['أقساط_مرحّلة'].get('%d|%d|%s' % (سنة, w, e))
            if old is not None and abs(old - v) > 0.0005:
                print('  ⚠️ %s: قسط الأسبوع %d تغيّر بعد الترحيل (%.3f ← %.3f)' % (e, w, old, v))

    # 3) ملخص الشهر
    if شهر:
        print('\n👷 ملخص الأيدي العاملة — شهر %02d (من الملف)' % شهر)
        r = 5 + شهر
        heads = ['عدد الجمع', 'رواتب عمر', 'عبدالعزيز', 'سلف كاش', 'عمال المهام', 'المجموع كاش', 'كلفة الشغل', 'بضاعة', 'قروض كاش']
        print('  %-14s %s → %s' % ('الفترة', تاريخ(M.cell(r, 2).value), تاريخ(M.cell(r, 3).value)))
        for i, h in enumerate(heads):
            v = M.cell(r, 4 + i).value
            print('  %-14s %10s' % (h, ('%d' % v) if i == 0 else ('%.3f' % رقم(v))))
    print('\n⏸️ مسودة بس — الترحيل على أودو بعد «رحّل» (الموظف شريك عميل · البضاعة بمطابقة مدير الأصناف · الأقساط قيد خصم راتب بتاريخ الجمعة).')


if __name__ == '__main__':
    main()
