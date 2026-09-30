# -*- coding: utf-8 -*-
"""قارئ ملف «نظام الموظفين» الشهري — لما صاحب المحل يرفع الملف (بأمره 30/09)

  python3 scripts/قارئ_نظام_الموظفين.py "ملف.xlsx"        # مسودة بس — ما بيلمس أودو

شو بيطلّع:
  1) فواتير بيع (البضاعة): كل سطر «النوع = بضاعة» بورقة «القروض والديون» ← فاتورة بيع باسم الموظف
     (البيان فيه الصنف والكمية ← بيتطابق مع مدير الأصناف قبل الترحيل).
  2) التحصيل: أقساط كل أسبوع (من «التسوية» عمود «قسط الدين») + «دفع كاش من جيبته» —
     بتتوزّع على ديون الموظف **الأقدم فالأحدث**: اللي بيسدّ بضاعة = تحصيل مبيعات،
     اللي بيسدّ قرض كاش = استرداد قرض (مش مبيعات)، واللي بيسدّ دين قديم = تحصيل دين قديم.
  3) أرقام ملخص الأيدي العاملة للعلم.

القواعد:
  • ولا إشي بينرحّل على أودو بلا «رحّل» من صاحب المحل.
  • تشييك التكرار: كل سطر إله مفتاح (شهر · موظف · تاريخ · رقم السطر · مبلغ) بيتسجّل بـ
    drive/ذمم_الموظفين.json بعد الترحيل — لو انرفع نفس الملف مرة ثانية ما بينعاد.
  • الموظف لازم يكون شريك على أودو (عميل) — أول مرة بينفتح بموافقته.
"""
import sys, os, json, shutil, subprocess, tempfile, datetime as dt
import openpyxl

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
REG = os.path.join(ROOT, 'drive', 'ذمم_الموظفين.json')
EMPS = ['عمر المصري', 'عبدالعزيز']
RECALC = '/mnt/skills/public/xlsx/scripts/recalc.py'
BAR = '─' * 78


def حمّل_السجل():
    if os.path.exists(REG): return json.load(open(REG))
    return {'مرحّل': {}, 'ديون': {e: [] for e in EMPS}}   # ديون = دفعات دين مفتوحة (النوع، الباقي) من الأشهر المرحّلة


def افتح(path):
    """القيم المحسوبة — لو الملف ما انحفظ من إكسل (بلا قيم) بنحسبه بنسخة مؤقتة."""
    wb = openpyxl.load_workbook(path, data_only=True)
    if wb['التسوية']['F9'].value is None and os.path.exists(RECALC):
        tmp = os.path.join(tempfile.mkdtemp(), 'x.xlsx'); shutil.copy(path, tmp)
        subprocess.run([sys.executable, RECALC, tmp, '90'], capture_output=True)
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
    if H['P1'].value != 'ALAWN_PAYROLL':
        sys.exit('⛔ هاد مش ملف «نظام الموظفين» المعتمد (علامة النموذج ناقصة).')
    S = wb['التسوية']; L = wb['القروض والديون']; B = wb['بداية ونهاية الشهر']
    شهر, سنة = int(S['C3'].value), int(S['E3'].value)
    أسابيع = [(تاريخ(H['H%d' % (11 + k)].value), تاريخ(H['I%d' % (11 + k)].value)) for k in range(6)]

    # سطور القروض والديون (من صف 11)
    سطور = []
    for r in range(11, L.max_row + 1):
        emp, typ, amt, rep = L['B%d' % r].value, L['C%d' % r].value, رقم(L['D%d' % r].value), رقم(L['F%d' % r].value)
        if emp not in EMPS or (not amt and not rep): continue
        سطور.append({'صف': r, 'تاريخ': تاريخ(L['K%d' % r].value), 'موظف': emp, 'نوع': typ or '⚠️ بلا نوع',
                     'مبلغ': amt, 'من_جيبته': rep, 'بيان': (L['G%d' % r].value or '').strip()})

    # الأقساط الأسبوعية من جداول الشهر بالتسوية: «الأسبوع …» تحت عنوان كل موظف
    أقساط = {e: [] for e in EMPS}; cur = None
    for r in range(12, S.max_row + 1):
        a = S['A%d' % r].value
        if isinstance(a, str):
            for e in EMPS:
                if a.startswith(e + ' —'): cur = e; k = 0
            if cur and a.startswith('الأسبوع'):
                v = رقم(S['E%d' % r].value)
                if v and k < len(أسابيع) and أسابيع[k][1]: أقساط[cur].append((أسابيع[k][1], v, a))
                k += 1
    افتتاح = {e: رقم(B['C%d' % (6 + i)].value) for i, e in enumerate(EMPS)}
    M = wb['ملخص الأيدي العاملة']
    return شهر, سنة, سطور, أقساط, افتتاح, M


def وزّع(ديون, مبلغ):
    """التحصيل على الديون الأقدم فالأحدث ← [(النوع، المبلغ)]"""
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
    شهر, سنة, سطور, أقساط, افتتاح, M = اقرأ(sys.argv[1])
    reg = حمّل_السجل(); tag = '%04d-%02d' % (سنة, شهر)
    print(BAR); print('ملف نظام الموظفين — شهر %02d/%d' % (شهر, سنة)); print(BAR)

    # 1) البضاعة ← فواتير بيع
    print('\n🧾 بضاعة ← فواتير بيع باسم الموظف')
    بضاعة = [s for s in سطور if s['نوع'] == 'بضاعة' and s['مبلغ']]
    for s in بضاعة:
        key = '%s|%s|%s|%d|%.3f' % (tag, s['موظف'], s['تاريخ'], s['صف'], s['مبلغ'])
        s['مفتاح'] = key; حالة = '✔ مرحّل من قبل' if key in reg['مرحّل'] else ('⚠️ البيان فاضي — بدنا الصنف والكمية' if not s['بيان'] else 'جاهز للمطابقة')
        print('  %s · %-11s · %8.3f · %s  ← %s' % (s['تاريخ'].strftime('%d/%m') if s['تاريخ'] else '??/??', s['موظف'], s['مبلغ'], s['بيان'] or '—', حالة))
    if not بضاعة: print('  — ما في بضاعة هالشهر')
    for s in سطور:
        if s['نوع'] == '⚠️ بلا نوع' and s['مبلغ']:
            print('  ⚠️ صف %d: %s %.3f بلا «النوع» — لازم يتحدّد قبل الترحيل' % (s['صف'], s['موظف'], s['مبلغ']))

    # 2) التحصيل — الأقدم فالأحدث
    print('\n💵 التحصيل (أقساط + دفع من جيبته) — بيتوزّع على الأقدم فالأحدث')
    for e in EMPS:
        ديون = [dict(d) for d in reg['ديون'].get(e, [])]
        if not ديون and افتتاح[e] > 0.0005: ديون = [{'نوع': 'دين قديم', 'باقي': افتتاح[e], 'مرجع': 'رصيد افتتاح'}]
        أحداث = [(s['تاريخ'], 'دين', s) for s in سطور if s['موظف'] == e and s['مبلغ']]
        أحداث += [(s['تاريخ'], 'جيبته', s) for s in سطور if s['موظف'] == e and s['من_جيبته']]
        أحداث += [(d, 'قسط', (v, w)) for d, v, w in أقساط[e]]
        أحداث.sort(key=lambda t: (t[0] or dt.date.min, {'دين': 0, 'جيبته': 1, 'قسط': 2}[t[1]]))
        مجاميع = {}
        for d, kind, obj in أحداث:
            if kind == 'دين':
                ديون.append({'نوع': obj['نوع'], 'باقي': obj['مبلغ'], 'مرجع': 'صف %d' % obj['صف']}); continue
            مبلغ, وصف = (obj['من_جيبته'], 'دفع من جيبته') if kind == 'جيبته' else (obj[0], 'قسط ' + obj[1])
            for typ, x, ref in وزّع(ديون, مبلغ):
                اسم = {'بضاعة': 'تحصيل مبيعات', 'قرض كاش': 'استرداد قرض (مش مبيعات)', 'دين قديم': 'تحصيل دين قديم'}.get(typ, typ)
                print('  %s · %-11s · %-18s · %8.3f ← %s %s' % (d.strftime('%d/%m') if d else '??/??', e, وصف, x, اسم, ('(' + ref + ')') if ref else ''))
                مجاميع[اسم] = round(مجاميع.get(اسم, 0) + x, 3)
        باقي = round(sum(d['باقي'] for d in ديون), 3)
        if مجاميع or باقي:
            print('  ⟵ %s: %s · الدين الباقي %.3f' % (e, ' · '.join('%s %.3f' % kv for kv in مجاميع.items()) or 'ما في تحصيل', باقي))

    # 3) ملخص الأيدي العاملة
    print('\n👷 ملخص الأيدي العاملة (من الملف)')
    for r in range(1, M.max_row + 1):
        a, f = M['A%d' % r].value, M['F%d' % r].value
        if isinstance(a, str) and isinstance(f, (int, float)) and ('دفعته' in a or 'كلفة' in a or 'انخصم' in a or 'بضاعة' in a):
            print('  %-62s %10.3f' % (a, f))
    print('\n⏸️ مسودة بس — الترحيل على أودو بعد «رحّل» (الموظف شريك عميل · البضاعة بمطابقة مدير الأصناف · الأقساط قيد خصم راتب).')


if __name__ == '__main__':
    main()
