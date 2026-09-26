# -*- coding: utf-8 -*-
"""نظام الموظفين الشهري — محلات العون لمواد البناء
ملف لكل شهر: تختار رقم الشهر والسنة، والأسابيع وتواريخها بتطلع لحالها.

القواعد (26/09/2026 من صاحب المحل):
  • المدير: 125 دينار كل 10 أيام، بلا ساعات — ورقته لحال.
  • عمر المصري: السبت–الخميس 7:00 AM–6:00 PM @1.250/ساعة · بعد 6:00 PM @1.500 · الجمعة 2:00 PM–11:00 PM حسب الساعات @1.250.
  • عبدالعزيز: 1.000/ساعة بغض النظر عن اليوم · يقبض نهاية كل يوم.
  • تحميل الجبسمبورد: ساعاته @2.500 — بتنكتب مرة وحدة لكل أسبوع (مش لكل يوم).
  • عمال المهام: بيندفعلهم مباشرة، كل مهمة بسعرها.
  • المستحق بيتجبر لربع دينار (عمر: مستحق الأسبوع · عبدالعزيز: مستحق اليوم لأنه بيقبض يومياً).

الأسبوع: السبت → الجمعة. **الملف فيه أيام الشهر بس.** الأسبوع اللي بيقطع بين شهرين
بينقسم: أيامه بكل شهر بتنحسب بملف شهرها، واللي ما انقبض بينتقل بـ«إقفال الشهر» ← «افتتاح الشهر الجاي».
صيغة التاريخ بكل الملف: يوم/شهر/سنة (dd/mm/yyyy).
الخلايا المحسوبة محمية (بلا كلمة سر) — بس الصفراء بتنكتب.
"""
import openpyxl, datetime as dt
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName

OUT = '/home/user/Fam-Bam/العون/نظام_الموظفين.xlsx'
NAVY, TEAL, GOLD, PLUM, RUST = '1F4E6B', '0F6E6E', '8A6A00', '5B3A6B', '8B3A2F'
EDIT, AUTO, WEEKBAR, SUBT, OUTM = 'FFF6CC', 'EDEDED', 'DCE6EE', 'E3EFE3', 'D9D9D9'
LINE = Side(style='thin', color='9AA5AD'); THICK = Side(style='medium', color=NAVY)
BOX = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
DATE = 'dd/mm/yyyy'; MONEY = '#,##0.000;[Red]-#,##0.000;-'; HRS = '0.00'
NWMAX = 6                      # أقصى عدد أسابيع (كاملة أو مقطوعة) بالشهر
BLOCK = 10                     # عنوان الأسبوع + 7 أيام + مجموع الأيام + سطر الجبسمبورد
B0 = 6
LOG_ROWS, TASK_ROWS, MLOG_ROWS = 120, 60, 25
ORD = ['الأول', 'الثاني', 'الثالث', 'الرابع', 'الخامس', 'السادس']
ORDF = ['الأولى', 'الثانية', 'الثالثة', 'الرابعة']
YEARS = list(range(2026, 2036))

def F(sz=11, b=False, c='000000'): return Font(name='Arial', size=sz, bold=b, color=c)
def C(h='center'): return Alignment(horizontal=h, vertical='center', wrap_text=True)
def fill(c): return PatternFill('solid', fgColor=c)
def col(i): return openpyxl.utils.get_column_letter(i)

wb = openpyxl.Workbook()
INPUTS = {}                                   # ورقة → خلايا مفتوحة للكتابة
def name(nm, ref): wb.defined_names[nm] = DefinedName(nm, attr_text=ref)
def dv(ws, formula, rng, prompt=None):
    d = DataValidation(type='list', formula1=formula, allow_blank=True)
    if prompt: d.prompt = prompt; d.showInputMessage = True
    ws.add_data_validation(d); d.add(rng)
def dv_date(ws, rng, lo='MSTART', hi='MEND', msg='اكتب التاريخ يوم/شهر/سنة — ومن أيام الشهر المختار'):
    d = DataValidation(type='date', operator='between', formula1=lo, formula2=hi, allow_blank=True,
                       showErrorMessage=True, errorTitle='تاريخ غلط', error=msg,
                       showInputMessage=True, promptTitle='التاريخ', prompt='يوم/شهر/سنة  مثل 05/10/2026')
    ws.add_data_validation(d); d.add(rng)
def cellfmt(c, color='FFFFFF', bold=False, fmt=None, sz=11, h='center', fc='000000'):
    c.border = BOX; c.fill = fill(color); c.font = F(sz, bold, fc); c.alignment = C(h)
    if fmt: c.number_format = fmt
def edit(ws, ref, fmt=None, bold=False, sz=11):
    c = ws[ref]; cellfmt(c, EDIT, bold, fmt, sz); INPUTS.setdefault(ws.title, []).append(ref)
def auto(ws, ref, fmt=None, bold=False, sz=11, color=AUTO, fc='000000'):
    cellfmt(ws[ref], color, bold, fmt, sz, fc=fc)
def title(ws, text, color, last, sub=None):
    ws.sheet_view.rightToLeft = True; ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:%s1' % last); ws['A1'] = text; cellfmt(ws['A1'], color, True, sz=16, fc='FFFFFF')
    ws.row_dimensions[1].height = 32
    if sub:
        ws.merge_cells('A2:%s2' % last); ws['A2'] = sub; ws['A2'].font = F(12, True, color); ws['A2'].alignment = C()
        ws.row_dimensions[2].height = 22
def note(ws, ref, text, rng=None, sz=10, color='555555', h=None):
    if rng: ws.merge_cells(rng)
    ws[ref] = text; ws[ref].font = F(sz, False, color); ws[ref].alignment = C('right')
    if h: ws.row_dimensions[ws[ref].row].height = h
def header(ws, row, heads, color, sz=10, height=34):
    for i, h in enumerate(heads, start=1): cellfmt(ws.cell(row, i, h), color, True, sz=sz, fc='FFFFFF')
    ws.row_dimensions[row].height = height

DAYNAME = 'CHOOSE(WEEKDAY({d},1),"الأحد","الاثنين","الثلاثاء","الأربعاء","الخميس","الجمعة","السبت")'
def TXT(x): return 'TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00")&"/"&YEAR(%s)' % (x, x, x)
def RND(x): return 'IF(ROUND_MODE="للأعلى",CEILING(ROUND((%s),3),0.25),ROUND((%s)*4,0)/4)' % (x, x)

# ═══════════════════ ورقة مخفية: القوائم والحسابات ═══════════════════
H = wb.create_sheet('القوائم')
def label(m):
    h, mi = divmod(m, 60)
    return '%d:%02d %s' % ((h % 12) or 12, mi, 'AM' if h < 12 else 'PM')
for i, m in enumerate([(6 * 60 + 15 * k) % 1440 for k in range(96)], start=1):   # تبدأ 6:00 AM
    H.cell(i, 1, label(m)); H.cell(i, 2, dt.time(m // 60, m % 60)).number_format = 'h:mm AM/PM'
for i in range(12): H.cell(i + 1, 4, i + 1)
for i, y in enumerate(YEARS, start=1): H.cell(i, 5, y)
H['G1'] = "=DATE('التسوية'!$E$3,'التسوية'!$C$3,1)"; H['G2'] = '=EOMONTH(G1,0)'
H['G3'] = '=G1-MOD(WEEKDAY(G1,1),7)'; H['G4'] = '=INT((G2-G3)/7)+1'
for c in ('G1', 'G2', 'G3'): H[c].number_format = DATE
name('MSTART', "'القوائم'!$G$1"); name('MEND', "'القوائم'!$G$2"); name('WFIRST', "'القوائم'!$G$3"); name('NW', "'القوائم'!$G$4")
for k in range(NWMAX):     # صف لكل أسبوع: بدايته ونهايته داخل الشهر
    r = 11 + k; sat = 'WFIRST+%d' % (7 * k)
    H['H%d' % r] = '=IF(%d<=NW,MAX(%s,MSTART),"")' % (k + 1, sat)
    H['I%d' % r] = '=IF(%d<=NW,MIN(%s+6,MEND),"")' % (k + 1, sat)
    H['J%d' % r] = '=IF(%d<=NW,"الأسبوع %s","")' % (k + 1, ORD[k])
    H['K%d' % r] = 'الأسبوع %s' % ORD[k]
    H['L%d' % r] = '=IF(%d<=NW,IF(%s<MSTART,"بدأ بالشهر الماضي — أيامه قبل "&%s&" بملف الشهر الماضي",IF(%s+6>MEND,"بيكمل بالشهر الجاي — أيامه بعد "&%s&" بملف الشهر الجاي","أسبوع كامل")),"")' % (
        k + 1, sat, TXT('MSTART'), sat, TXT('MEND'))
    for c in 'HI': H['%s%d' % (c, r)].number_format = DATE
H.sheet_state = 'hidden'
TIME_LIST = "='القوائم'!$A$1:$A$96"
def TV(ref): return "INDEX('القوائم'!$B$1:$B$96,MATCH(%s,'القوائم'!$A$1:$A$96,0))" % ref
WSTART = "'القوائم'!$H$11:$H$16"; WEND = "'القوائم'!$I$11:$I$16"

# ═══════════════════ القواعد ═══════════════════
K = wb.create_sheet('القواعد')
title(K, 'القواعد والأسعار — كل الحسابات بتقرأ من هون', GOLD, 'D')
for k, w in zip('ABCD', [4, 44, 16, 58]): K.column_dimensions[k].width = w
header(K, 3, ['#', 'القاعدة', 'القيمة', 'الشرح'], GOLD, 11, 26)
rules = [
    ('OMR_RATE', 'عمر المصري — أجر الساعة العادية (دينار)', 1.25, '0.000', 'السبت–الخميس لحد بداية الإضافي، والجمعة كل ساعاته.'),
    ('OMR_OT', 'عمر المصري — أجر ساعة الإضافي (دينار)', 1.5, '0.000', 'السبت–الخميس بعد الوقت اللي تحت.'),
    ('OMR_OT_AT', 'عمر المصري — بداية الإضافي', dt.time(18, 0), 'h:mm AM/PM', 'الجمعة ما إلها إضافي.'),
    ('OMR_FRI_IN', 'عمر المصري — دوام الجمعة من', dt.time(14, 0), 'h:mm AM/PM', 'للتنبيه بس — الحساب حسب الساعات الفعلية.'),
    ('OMR_FRI_OUT', 'عمر المصري — دوام الجمعة إلى', dt.time(23, 0), 'h:mm AM/PM', ''),
    ('AZZ_RATE', 'عبدالعزيز — أجر الساعة (دينار)', 1.0, '0.000', 'أي يوم. بيقبض نهاية كل يوم.'),
    ('GYP_RATE', 'تحميل الجبسمبورد — أجر الساعة (دينار)', 2.5, '0.000', 'ساعات التحميل جزء من ساعات الدوام، بس بتنحسب @2.5 بدل الأجر العادي.'),
    ('ROUND_MODE', 'جبر المستحق', 'لأقرب ربع', '@', 'لأقرب ربع: 18.300←18.250 و18.400←18.500 · للأعلى: أي كسر بيطلع للربع اللي فوقه.'),
    ('ADV_CAP', 'سقف السلف بالأسبوع (نسبة من المستحق)', 0.5, '0%', 'تنبيه لو السلف تجاوزت. 0 = بلا سقف.'),
    ('MGR_SAL', 'المدير — راتب كل 10 أيام (دينار)', 125, '0.000', ''),
    ('MGR_START', 'المدير — بداية أول فترة', dt.date(2026, 9, 26), DATE, 'الفترات بتتسلسل من هالتاريخ كل 10 أيام. ثابت لكل الشهور.'),
]
for i, (nm, rule, val, fmt, desc) in enumerate(rules, start=1):
    r = 3 + i
    K.cell(r, 1, i); K.cell(r, 2, rule); K.cell(r, 3, val); K.cell(r, 4, desc)
    for cc in (1, 2, 4): cellfmt(K.cell(r, cc), h='right' if cc > 1 else 'center')
    edit(K, 'C%d' % r, fmt, True); K.cell(r, 4).font = F(10, False, '555555'); K.row_dimensions[r].height = 30
    name(nm, "'القواعد'!$C$%d" % r)
    if nm == 'ROUND_MODE': dv(K, '"لأقرب ربع,للأعلى"', 'C%d' % r)

# ═══════════════════ أوراق الساعات ═══════════════════
def hours_sheet(ttl, banner, color, heads, widths, text):
    ws = wb.create_sheet(ttl)
    title(ws, banner, color, 'I', '="شهر "&TEXT(\'التسوية\'!$C$3,"00")&" / "&\'التسوية\'!$E$3&"   ·   من "&%s&" إلى "&%s' % (TXT('MSTART'), TXT('MEND')))
    note(ws, 'A3', text, 'A3:I3', h=36)
    for k, w in enumerate(widths, start=1): ws.column_dimensions[col(k)].width = w
    header(ws, 5, heads, color, 11, 32); ws.freeze_panes = 'A6'
    ws.column_dimensions['K'].hidden = True; ws.column_dimensions['L'].hidden = True
    return ws

def week_bar(ws, k, color):
    r = B0 + BLOCK * k; sat = 'WFIRST+%d' % (7 * k)
    ws.merge_cells('A%d:I%d' % (r, r))
    ws['A%d' % r] = ('=IF(%d<=NW,"الأسبوع %s   ·   من "&%s&" "&%s&"  إلى  "&%s&" "&%s'
                     '&IF(%s<MSTART,"   (بدأ بالشهر الماضي)",IF(%s+6>MEND,"   (بيكمل بالشهر الجاي)","")),"")') % (
        k + 1, ORD[k], DAYNAME.format(d='MAX(%s,MSTART)' % sat), TXT('MAX(%s,MSTART)' % sat),
        DAYNAME.format(d='MIN(%s+6,MEND)' % sat), TXT('MIN(%s+6,MEND)' % sat), sat, sat)
    cellfmt(ws['A%d' % r], WEEKBAR, True, sz=14, fc=color, h='right'); ws.row_dimensions[r].height = 28
    return r

def day_date(ws, r, k, d):
    ws['A%d' % r] = '=IF(AND(%d<=NW,WFIRST+%d>=MSTART,WFIRST+%d<=MEND),WFIRST+%d,"")' % (k + 1, 7 * k + d, 7 * k + d, 7 * k + d)
    ws['B%d' % r] = '=IF(A{r}="","",%s)'.replace('{r}', str(r)) % DAYNAME.format(d='A%d' % r)
    ws['K%d' % r] = '=IF(OR(A{r}="",C{r}=""),"",%s)'.replace('{r}', str(r)) % TV('C%d' % r)
    ws['L%d' % r] = '=IF(OR(A{r}="",D{r}=""),"",%s)'.replace('{r}', str(r)) % TV('D%d' % r)
    auto(ws, 'A%d' % r, DATE); auto(ws, 'B%d' % r)
    for c in 'KL': ws['%s%d' % (c, r)].number_format = 'h:mm AM/PM'
    if d == 6:
        for c in 'AB': ws['%s%d' % (c, r)].font = F(11, True, NAVY)

# ---- عمر ----
O = hours_sheet('ساعات عمر', 'ساعات عمر المصري', NAVY,
    ['التاريخ', 'اليوم', 'دخول', 'خروج', 'الساعات', 'ساعات عادية', 'ساعات إضافي', 'المستحق (دينار)', 'ملاحظات'],
    [13, 10, 12, 12, 10, 11, 11, 14, 30],
    'اختر الدخول والخروج من القائمة (كل ربع ساعة). السبت–الخميس: بعد 6:00 PM إضافي @1.500 · الجمعة كل الساعات @1.250. '
    'ساعات تحميل الجبسمبورد بتنكتب مرة وحدة تحت كل أسبوع. مستحق الأسبوع بيتجبر لربع دينار.')
OMR = []   # (سطر مجموع الأيام, سطر الجبسمبورد)
for k in range(NWMAX):
    hr = week_bar(O, k, NAVY)
    for d in range(7):
        r = hr + 1 + d; day_date(O, r, k, d)
        O['E%d' % r] = '=IF(OR(K{r}="",L{r}=""),"",ROUND(MOD(L{r}-K{r},1)*24,2))'.replace('{r}', str(r))
        O['F%d' % r] = ('=IF(E{r}="","",IF(WEEKDAY(A{r},1)=6,E{r},'
                        'ROUND(MAX(0,MIN(IF(L{r}<K{r},L{r}+1,L{r}),OMR_OT_AT)-K{r})*24,2)))').replace('{r}', str(r))
        O['G%d' % r] = '=IF(E{r}="","",ROUND(E{r}-F{r},2))'.replace('{r}', str(r))
        O['H%d' % r] = '=IF(E{r}="","",ROUND(F{r}*OMR_RATE+G{r}*OMR_OT,3))'.replace('{r}', str(r))
        for c in 'CD': edit(O, '%s%d' % (c, r))
        edit(O, 'I%d' % r)
        for c, fm in zip('EFGH', [HRS, HRS, HRS, MONEY]): auto(O, '%s%d' % (c, r), fm)
    s1, s2 = hr + 8, hr + 9; lo, hi = hr + 1, hr + 7; on = '%d<=NW' % (k + 1)
    O.merge_cells('A%d:D%d' % (s1, s1)); O['A%d' % s1] = '=IF(%s,"مجموع أيام الأسبوع %s","")' % (on, ORD[k])
    for c, fm in zip('EFGH', [HRS, HRS, HRS, MONEY]):
        O['%s%d' % (c, s1)] = '=IF(%s,SUM(%s%d:%s%d),"")' % (on, c, lo, c, hi)
    for c in 'ABCDEFGHI': auto(O, '%s%d' % (c, s1), None, True, color=SUBT)
    for c, fm in zip('EFGH', [HRS, HRS, HRS, MONEY]): O['%s%d' % (c, s1)].number_format = fm
    O.merge_cells('A%d:D%d' % (s2, s2)); O['A%d' % s2] = '=IF(%s,"ساعات تحميل الجبسمبورد بهالأسبوع ←","")' % on
    auto(O, 'A%d' % s2, None, True, color=SUBT); edit(O, 'E%d' % s2, HRS, True)
    O.merge_cells('F%d:G%d' % (s2, s2)); O['F%d' % s2] = '=IF(%s,"مستحق الأسبوع (دينار) ←","")' % on
    auto(O, 'F%d' % s2, None, True, color=SUBT)
    O['H%d' % s2] = '=IF(%s,%s,"")' % (on, RND('N(H%d)+N(E%d)*(GYP_RATE-OMR_RATE)' % (s1, s2)))
    auto(O, 'H%d' % s2, MONEY, True, 13, color='FFE9A8')
    O['I%d' % s2] = '=IF(%s,"قبل الجبر: "&TEXT(N(H%d)+N(E%d)*(GYP_RATE-OMR_RATE),"0.000"),"")' % (on, s1, s2)
    auto(O, 'I%d' % s2, None, color=SUBT); O['I%d' % s2].font = F(9, False, '555555')
    OMR.append((s1, s2))
OEND = B0 + BLOCK * NWMAX - 1
dv(O, TIME_LIST, 'C%d:D%d' % (B0, OEND))
O.conditional_formatting.add('C%d:D%d' % (B0, OEND), FormulaRule(
    formula=['AND(ISNUMBER($A%d),$K%d<>"",$L%d<>"",WEEKDAY($A%d,1)=6,OR($K%d<OMR_FRI_IN,$L%d>OMR_FRI_OUT))' % ((B0,) * 6)], fill=fill('FDE2C8')))

# ---- عبدالعزيز ----
Z = hours_sheet('ساعات عبدالعزيز', 'ساعات عبدالعزيز — بيقبض نهاية كل يوم', TEAL,
    ['التاريخ', 'اليوم', 'دخول', 'خروج', 'الساعات', 'مستحق اليوم (دينار)', 'قبض؟', 'المقبوض (دينار)', 'ملاحظات'],
    [13, 10, 12, 12, 10, 14, 8, 13, 30],
    'دينار للساعة أي يوم. مستحق كل يوم مجبور لربع دينار. لما يقبض آخر اليوم: «نعم» + المبلغ. '
    'ساعات الجبسمبورد تحت كل أسبوع — فرقها بينضاف لمستحق الأسبوع وبيندفع آخره.')
AZZ = []
for k in range(NWMAX):
    hr = week_bar(Z, k, TEAL)
    for d in range(7):
        r = hr + 1 + d; day_date(Z, r, k, d)
        Z['E%d' % r] = '=IF(OR(K{r}="",L{r}=""),"",ROUND(MOD(L{r}-K{r},1)*24,2))'.replace('{r}', str(r))
        Z['F%d' % r] = '=IF(E%d="","",%s)' % (r, RND('E%d*AZZ_RATE' % r))
        for c in 'CDGI': edit(Z, '%s%d' % (c, r))
        edit(Z, 'H%d' % r, MONEY)
        auto(Z, 'E%d' % r, HRS); auto(Z, 'F%d' % r, MONEY)
    s1, s2 = hr + 8, hr + 9; lo, hi = hr + 1, hr + 7; on = '%d<=NW' % (k + 1)
    Z.merge_cells('A%d:D%d' % (s1, s1)); Z['A%d' % s1] = '=IF(%s,"مجموع أيام الأسبوع %s","")' % (on, ORD[k])
    for c in 'EFH': Z['%s%d' % (c, s1)] = '=IF(%s,SUM(%s%d:%s%d),"")' % (on, c, lo, c, hi)
    for c in 'ABCDEFGHI': auto(Z, '%s%d' % (c, s1), None, True, color=SUBT)
    Z['E%d' % s1].number_format = HRS
    for c in 'FH': Z['%s%d' % (c, s1)].number_format = MONEY
    Z.merge_cells('A%d:D%d' % (s2, s2)); Z['A%d' % s2] = '=IF(%s,"ساعات تحميل الجبسمبورد بهالأسبوع ←","")' % on
    auto(Z, 'A%d' % s2, None, True, color=SUBT); edit(Z, 'E%d' % s2, HRS, True)
    Z['F%d' % s2] = '=IF(%s,N(F%d)+%s,"")' % (on, s1, RND('N(E%d)*(GYP_RATE-AZZ_RATE)' % s2))
    auto(Z, 'F%d' % s2, MONEY, True, 13, color='FFE9A8')
    Z.merge_cells('G%d:I%d' % (s2, s2)); Z['G%d' % s2] = '=IF(%s,"← مستحق الأسبوع كامل (دينار) مع فرق الجبسمبورد","")' % on
    auto(Z, 'G%d' % s2, None, color=SUBT); Z['G%d' % s2].font = F(10, True, TEAL)
    AZZ.append((s1, s2))
ZEND = B0 + BLOCK * NWMAX - 1
dv(Z, TIME_LIST, 'C%d:D%d' % (B0, ZEND)); dv(Z, '"نعم,لا"', 'G%d:G%d' % (B0, ZEND))
Z.conditional_formatting.add('H%d:H%d' % (B0, ZEND), FormulaRule(formula=['AND($G%d="نعم",ISNUMBER($F%d),N($H%d)<>$F%d)' % ((B0,) * 4)], fill=fill('FDE2C8')))
for ws, end in ((O, OEND), (Z, ZEND)):   # أيام خارج الشهر: رمادي
    ws.conditional_formatting.add('C%d:D%d' % (B0, end), FormulaRule(formula=['AND($A%d="",$B%d="")' % (B0, B0)], fill=fill(OUTM)))

# ═══════════════════ السلف والديون ═══════════════════
L = wb.create_sheet('السلف والديون')
title(L, 'السلف والديون — عمر المصري وعبدالعزيز', RUST, 'F',
      '="شهر "&TEXT(\'التسوية\'!$C$3,"00")&" / "&\'التسوية\'!$E$3')
for k, w in zip('ABCDEF', [13, 16, 14, 16, 16, 44]): L.column_dimensions[k].width = w
expl = [
    '• سلفة: كاش بياخذه وسط الأسبوع ← بتنخصم كاملة من صافي نفس الأسبوع.',
    '• دين أو بضاعة: مبلغ كاش كبير أو بضاعة لورشته (بالسعر اللي بتحدده) ← ما بينخصم مرة وحدة، بينخصم «القسط الأسبوعي» كل أسبوع لحد ما يخلص.',
    '• تسديد مباشر: لو دفع من جيبته جزء من الدين (مش من راتبه).',
    '• راتبه اللي بتدفعه آخر الأسبوع ما بينكتب هون — بينكتب بخانة «المدفوع» بورقة التسوية.',
]
for i, t in enumerate(expl): note(L, 'A%d' % (3 + i), t, 'A%d:F%d' % (3 + i, 3 + i), 10, '333333', 18)
header(L, 8, ['التاريخ', 'الموظف', 'سلفة (دينار)', 'دين أو بضاعة (دينار)', 'تسديد مباشر (دينار)', 'البيان'], RUST, 11, 30)
LR0, LE = 9, 9 + LOG_ROWS - 1
for r in range(LR0, LE + 1):
    edit(L, 'A%d' % r, DATE); edit(L, 'B%d' % r)
    for c in 'CDE': edit(L, '%s%d' % (c, r), MONEY)
    edit(L, 'F%d' % r); L['F%d' % r].alignment = C('right')
dv(L, '"عمر المصري,عبدالعزيز"', 'B%d:B%d' % (LR0, LE)); dv_date(L, 'A%d:A%d' % (LR0, LE))
L.freeze_panes = 'A9'
for nm, c in (('LG_D', 'A'), ('LG_E', 'B'), ('LG_ADV', 'C'), ('LG_DEBT', 'D'), ('LG_REP', 'E')):
    name(nm, "'السلف والديون'!$%s$%d:$%s$%d" % (c, LR0, c, LE))
note(L, 'A%d' % (LE + 2), 'مثال: 05/10/2026 · عمر المصري · سلفة 20.000 · «كاش من الصندوق»   |   07/10/2026 · عمر المصري · دين أو بضاعة 150.000 · «بضاعة لورشة بيته بسعر الجملة»', 'A%d:F%d' % (LE + 2, LE + 2))

# ═══════════════════ التسوية ═══════════════════
S = wb['Sheet']; S.title = 'التسوية'
title(S, 'نظام الموظفين — التسوية الأسبوعية والشهرية', NAVY, 'L')
for k, w in zip('ABCDEFGHIJKL', [15, 13, 13, 11, 12, 12, 12, 12, 12, 12, 12, 26]): S.column_dimensions[k].width = w
# الشهر والسنة وأوله وآخره — جنب بعض
S['B3'] = 'الشهر'; S['D3'] = 'السنة'; S['F3'] = 'أول الشهر'; S['H3'] = 'آخر الشهر'
for c in ('B3', 'D3', 'F3', 'H3'): S[c].font = F(12, True, NAVY); S[c].alignment = C('left')
S['C3'] = 10; edit(S, 'C3', '00', True, 14); dv(S, "='القوائم'!$D$1:$D$12", 'C3', 'رقم الشهر 1–12')
S['E3'] = 2026; edit(S, 'E3', '0', True, 14); dv(S, "='القوائم'!$E$1:$E$10", 'E3')
S['G3'] = '=MSTART'; auto(S, 'G3', DATE, True, 12); S['I3'] = '=MEND'; auto(S, 'I3', DATE, True, 12)
S.row_dimensions[3].height = 30
# الأسبوع المختار
S['B5'] = 'الأسبوع'; S['B5'].font = F(12, True, NAVY); S['B5'].alignment = C('left')
S.merge_cells('C5:D5'); S['C5'] = 'الأسبوع الأول'; edit(S, 'C5', None, True, 13)
dv(S, "='القوائم'!$J$11:$J$16", 'C5', 'القائمة بتتغيّر حسب الشهر')
S['E5'] = 'من'; S['G5'] = 'إلى'
for c in ('E5', 'G5'): S[c].font = F(12, True, NAVY); S[c].alignment = C('left')
S['F5'] = '=WS'; auto(S, 'F5', DATE, True, 12); S['H5'] = '=WE'; auto(S, 'H5', DATE, True, 12)
S.merge_cells('I5:L5'); S['I5'] = "=INDEX('القوائم'!$L$11:$L$16,WK)"; S['I5'].font = F(10, True, RUST); S['I5'].alignment = C('right')
S.row_dimensions[5].height = 30
S['N5'] = "=IFERROR(MIN(MATCH($C$5,'القوائم'!$K$11:$K$16,0),NW),1)"; S['N5'].font = F(8, False, 'BBBBBB')
name('WK', "'التسوية'!$N$5"); name('WS', "INDEX(%s,'التسوية'!$N$5)" % WSTART); name('WE', "INDEX(%s,'التسوية'!$N$5)" % WEND)
S['F5'] = '=INDEX(%s,WK)' % WSTART; S['H5'] = '=INDEX(%s,WK)' % WEND

# افتتاح الشهر
S['A7'] = 'افتتاح الشهر — انقل الأرقام من «إقفال الشهر» بملف الشهر الماضي'; S['A7'].font = F(12, True, RUST); S.merge_cells('A7:L7'); S['A7'].alignment = C('right')
header(S, 8, ['الموظف', 'مستحق له لسا ما قبضه (دينار)', 'دين أو بضاعة متبقي عليه (دينار)', 'القسط الأسبوعي (دينار)'], RUST, 10, 34)
S.merge_cells('E8:L8'); note(S, 'E8', '«مستحق له» = أيام اشتغلها آخر الشهر الماضي وما انقبضت (مثلاً أسبوع مقطوع). بينضاف لصافي أول أسبوع. · «دين متبقي» = اللي ضل عليه من ديون وبضاعة. بينخصم منه القسط الأسبوعي لحد ما يخلص.', sz=9)
OPEN = {}
for i, nm in enumerate(['عمر المصري', 'عبدالعزيز']):
    r = 9 + i; S['A%d' % r] = nm; auto(S, 'A%d' % r, None, True, color='FFFFFF')
    for c in 'BCD': S['%s%d' % (c, r)] = 0; edit(S, '%s%d' % (c, r), MONEY, True)
    OPEN[nm] = r

# جدول عمر للشهر
def month_table(top, emp, color, rows_ref, is_azz):
    S['A%d' % top] = emp + ' — كل أسابيع الشهر'; S['A%d' % top].font = F(13, True, color); S.merge_cells('A%d:L%d' % (top, top)); S['A%d' % top].alignment = C('right')
    if is_azz:
        heads = ['الأسبوع', 'من', 'إلى', 'الساعات', 'المستحق (دينار)', '+ مدوّر', '− مقبوض يومياً', '− السلف', '− القسط', 'الصافي للدفع', 'المدفوع (دينار)', 'الباقي له']
    else:
        heads = ['الأسبوع', 'من', 'إلى', 'الساعات', 'المستحق (دينار)', '+ مدوّر', '− السلف', '− القسط', 'الصافي للدفع', 'المدفوع (دينار)', 'الباقي له', 'ملاحظة']
    header(S, top + 1, heads, color, 10, 34)
    r0 = top + 2; orow = OPEN[emp]
    for k in range(NWMAX):
        r = r0 + k; on = '%d<=NW' % (k + 1); s1, s2 = rows_ref[k]; sh = "'ساعات عبدالعزيز'" if is_azz else "'ساعات عمر'"
        S['A%d' % r] = '=IF(%s,"الأسبوع %s","")' % (on, ORD[k])
        S['B%d' % r] = "=IF(%s,'القوائم'!H%d,\"\")" % (on, 11 + k); S['C%d' % r] = "=IF(%s,'القوائم'!I%d,\"\")" % (on, 11 + k)
        S['D%d' % r] = '=IF(%s,N(%s!E%d),"")' % (on, sh, s1)
        S['E%d' % r] = '=IF(%s,N(%s!%s%d),"")' % (on, sh, 'F' if is_azz else 'H', s2)
        S['F%d' % r] = '=IF(%s,%s,"")' % (on, '$B$%d' % orow if k == 0 else '0')
        adv = 'SUMIFS(LG_ADV,LG_E,"%s",LG_D,">="&B%d,LG_D,"<="&C%d)' % (emp, r, r)
        prev_inst_col = 'I' if is_azz else 'H'
        prev = '0' if k == 0 else 'SUM(%s%d:%s%d)' % (prev_inst_col, r0, prev_inst_col, r - 1)
        inst = ('IF(N(E{r})=0,0,MIN($D${o},MAX(0,$C${o}+SUMIFS(LG_DEBT,LG_E,"{e}",LG_D,"<="&C{r})'
                '-SUMIFS(LG_REP,LG_E,"{e}",LG_D,"<="&C{r})-{p})))').format(r=r, o=orow, e=emp, p=prev)
        if is_azz:
            S['G%d' % r] = '=IF(%s,N(%s!H%d),"")' % (on, sh, s1)
            S['H%d' % r] = '=IF(%s,%s,"")' % (on, adv)
            S['I%d' % r] = '=IF(%s,%s,"")' % (on, inst)
            S['J%d' % r] = '=IF(%s,E%d+F%d-G%d-H%d-I%d,"")' % ((on,) + (r,) * 5)
            edit(S, 'K%d' % r, MONEY)
            S['L%d' % r] = '=IF(%s,J%d-N(K%d),"")' % (on, r, r)
            auto_cols, money = 'ABCDEFGHIJL', 'EFGHIJL'
        else:
            S['G%d' % r] = '=IF(%s,%s,"")' % (on, adv)
            S['H%d' % r] = '=IF(%s,%s,"")' % (on, inst)
            S['I%d' % r] = '=IF(%s,E%d+F%d-G%d-H%d,"")' % ((on,) + (r,) * 4)
            edit(S, 'J%d' % r, MONEY)
            S['K%d' % r] = '=IF(%s,I%d-N(J%d),"")' % (on, r, r)
            S['L%d' % r] = '=IF(%s,IF(AND(ADV_CAP>0,E%d>0,G%d>ADV_CAP*E%d),"⚠️ السلف فوق السقف",IF(I%d<0,"⚠️ السلف أكثر من المستحق","")),"")' % ((on,) + (r,) * 4)
            auto_cols, money = 'ABCDEFGHIKL', 'EFGHIK'
        for c in auto_cols: auto(S, '%s%d' % (c, r))
        for c in 'BC': S['%s%d' % (c, r)].number_format = DATE
        S['D%d' % r].number_format = HRS
        for c in money: S['%s%d' % (c, r)].number_format = MONEY
        S['A%d' % r].font = F(11, True, color)
        if not is_azz: S['L%d' % r].font = F(9, True, 'C00000')
    tr = r0 + NWMAX
    S['A%d' % tr] = 'مجموع الشهر'
    sumcols = 'DEFGHIJKL' if is_azz else 'DEFGHIJK'
    for c in sumcols: S['%s%d' % (c, tr)] = '=SUM(%s%d:%s%d)' % (c, r0, c, tr - 1)
    for c in 'ABCDEFGHIJKL': auto(S, '%s%d' % (c, tr), None, True, color=SUBT)
    S['D%d' % tr].number_format = HRS
    for c in sumcols[1:]: S['%s%d' % (c, tr)].number_format = MONEY
    S.conditional_formatting.add('A%d:L%d' % (r0, tr - 1), FormulaRule(formula=['AND($A%d<>"",ROW()-%d+1=WK)' % (r0, r0)], fill=fill('FFF2B3')))
    return r0, tr

OR0, OTR = month_table(20, 'عمر المصري', NAVY, OMR, False)
ZR0, ZTR = month_table(OTR + 2, 'عبدالعزيز', TEAL, AZZ, True)

# الأسبوع المختار (فوق الجداول)
S['A12'] = '="تسوية الأسبوع المختار   ·   "&$C$5&"   ·   من "&%s&" إلى "&%s' % (TXT('WS'), TXT('WE'))
S['A12'].font = F(13, True, NAVY); S.merge_cells('A12:L12'); S['A12'].alignment = C('right'); S.row_dimensions[12].height = 24
header(S, 13, ['الموظف', 'الساعات', 'المستحق (دينار)', '+ مدوّر', '− مقبوض يومياً', '− السلف', '− القسط', 'الصافي للدفع', 'المدفوع (دينار)', 'الباقي له', 'ملاحظة', ''], NAVY, 10, 34)
S.merge_cells('K13:L13')
sel = [('عمر المصري', OR0, {'D': 'D', 'E': 'E', 'F': 'F', 'G': None, 'H': 'G', 'I': 'H', 'J': 'I', 'P': 'J', 'K': 'K'}),
       ('عبدالعزيز', ZR0, {'D': 'D', 'E': 'E', 'F': 'F', 'G': 'G', 'H': 'H', 'I': 'I', 'J': 'J', 'P': 'K', 'K': 'L'})]
for i, (nm, r0, m) in enumerate(sel):
    r = 14 + i
    def pick(c): return '=N(INDEX(%s%d:%s%d,WK))' % (c, r0, c, r0 + NWMAX - 1)
    S['A%d' % r] = nm
    S['B%d' % r] = pick(m['D']); S['C%d' % r] = pick(m['E']); S['D%d' % r] = pick(m['F'])
    S['E%d' % r] = pick(m['G']) if m['G'] else 0
    S['F%d' % r] = pick(m['H']); S['G%d' % r] = pick(m['I']); S['H%d' % r] = pick(m['J'])
    S['I%d' % r] = pick(m['P']); S['J%d' % r] = pick(m['K'])
    S['K%d' % r] = '=IF(AND(ADV_CAP>0,C{r}>0,F{r}>ADV_CAP*C{r}),"⚠️ السلف فوق السقف",IF(H{r}<0,"⚠️ السلف أكثر من المستحق",IF(AND(H{r}>0,J{r}=0),"✔ مدفوع",IF(H{r}>0,"لسا ما اندفع",""))))'.replace('{r}', str(r))
    S.merge_cells('K%d:L%d' % (r, r))
    for c in 'ABCDEFGHIJK': auto(S, '%s%d' % (c, r))
    S['A%d' % r].font = F(11, True)
    S['B%d' % r].number_format = HRS
    for c in 'CDEFGHIJ': S['%s%d' % (c, r)].number_format = MONEY
    S['H%d' % r].font = F(13, True); S['K%d' % r].font = F(10, True, 'C00000')
    S.row_dimensions[r].height = 24
note(S, 'A17', 'المدفوع بينكتب بجداول الشهر تحت (سطر الأسبوع) — وهون بيطلع لحاله. المدير بورقته لحال · عمال المهام بورقتهم.', 'A17:L17', 9)

# إقفال الشهر
CL = ZTR + 3
S['A%d' % CL] = 'إقفال الشهر — هالأرقام بتنكتب بـ«افتتاح الشهر» بملف الشهر الجاي'; S['A%d' % CL].font = F(13, True, RUST)
S.merge_cells('A%d:L%d' % (CL, CL)); S['A%d' % CL].alignment = C('right')
header(S, CL + 1, ['الموظف', 'مستحق له لسا ما قبضه (دينار)', 'دين أو بضاعة متبقي عليه (دينار)', 'القسط الأسبوعي (دينار)'], RUST, 10, 34)
for i, (nm, tr, rest_c, inst_c) in enumerate([('عمر المصري', OTR, 'K', 'H'), ('عبدالعزيز', ZTR, 'L', 'I')]):
    r = CL + 2 + i; o = OPEN[nm]
    S['A%d' % r] = nm
    S['B%d' % r] = '=%s%d' % (rest_c, tr)
    S['C%d' % r] = '=$C$%d+SUMIFS(LG_DEBT,LG_E,A%d)-SUMIFS(LG_REP,LG_E,A%d)-%s%d' % (o, r, r, inst_c, tr)
    S['D%d' % r] = '=IF(C%d>0,$D$%d,0)' % (r, o)
    for c in 'ABCD': auto(S, '%s%d' % (c, r), MONEY if c != 'A' else None, True, 12)
    S.merge_cells('E%d:L%d' % (r, r))
    S['E%d' % r] = '=IF(B{r}>0.0005,"المحل لسا مدين إله بـ "&TEXT(B{r},"0.000"),IF(B{r}<-0.0005,"قبض زيادة "&TEXT(-B{r},"0.000")&" — بتنخصم الشهر الجاي","ما إله شي"))&IF(C{r}>0.0005,"  ·  وعليه دين "&TEXT(C{r},"0.000"),"")'.replace('{r}', str(r))
    S['E%d' % r].font = F(10, True, RUST); S['E%d' % r].alignment = C('right')
S.freeze_panes = 'A6'

# ═══════════════════ المدير (ورقته لحال) ═══════════════════
M = wb.create_sheet('المدير')
title(M, 'المدير — 125 دينار كل 10 أيام', GOLD, 'J',
      '="شهر "&TEXT(\'التسوية\'!$C$3,"00")&" / "&\'التسوية\'!$E$3&"   ·   الفترات اللي بتخلص بهالشهر"')
for k, w in zip('ABCDEFGHIJ', [14, 13, 13, 12, 12, 12, 13, 13, 13, 24]): M.column_dimensions[k].width = w
note(M, 'A3', 'الفترات بتطلع لحالها من «بداية أول فترة» بورقة القواعد، وكل فترة بتنحسب بملف الشهر اللي بتخلص فيه. '
     'سلفه وديونه بجدول تحت بهالورقة نفسها.', 'A3:J3', h=32)
header(M, 5, ['افتتاح الشهر', 'مستحق له لسا ما قبضه', 'دين متبقي عليه', 'القسط كل فترة'], GOLD, 10, 30)
M['A6'] = 'من الشهر الماضي'; auto(M, 'A6', None, True, color='FFFFFF')
for c in 'BCD': M['%s6' % c] = 0; edit(M, '%s6' % c, MONEY, True)
header(M, 8, ['الفترة', 'من', 'إلى', 'الراتب (دينار)', '+ مدوّر', '− السلف', '− القسط', 'الصافي', 'المدفوع (دينار)', 'الباقي له'], GOLD, 10, 30)
M['L8'] = '=MAX(0,ROUNDUP((MSTART-MGR_START-9)/10,0))'; M['L8'].font = F(8, False, 'BBBBBB')
MR0 = 9; ML0, MLE = 18, 18 + MLOG_ROWS - 1
for i in range(4):
    r = MR0 + i; n = '($L$8+%d)' % i
    ok = 'AND(MGR_START+10*{n}+9>=MSTART,MGR_START+10*{n}+9<=MEND)'.format(n=n)
    M['B%d' % r] = '=IF(%s,MGR_START+10*%s,"")' % (ok, n)
    M['A%d' % r] = '=IF(B%d="","","الفترة %s")' % (r, ORDF[i])
    M['C%d' % r] = '=IF(B%d="","",B%d+9)' % (r, r)
    M['D%d' % r] = '=IF(B%d="","",MGR_SAL)' % r
    M['E%d' % r] = '=IF(B%d="","",%s)' % (r, '$B$6' if i == 0 else '0')
    M['F%d' % r] = "=IF(B%d=\"\",\"\",SUMIFS($B$%d:$B$%d,$A$%d:$A$%d,\">=\"&B%d,$A$%d:$A$%d,\"<=\"&C%d))" % (r, ML0, MLE, ML0, MLE, r, ML0, MLE, r)
    prev = '0' if i == 0 else 'SUM(G%d:G%d)' % (MR0, r - 1)
    M['G%d' % r] = ('=IF(B{r}="","",MIN($D$6,MAX(0,$C$6+SUMIFS($C${a}:$C${b},$A${a}:$A${b},"<="&C{r})'
                    '-SUMIFS($D${a}:$D${b},$A${a}:$A${b},"<="&C{r})-{p})))').format(r=r, a=ML0, b=MLE, p=prev)
    M['H%d' % r] = '=IF(B%d="","",D%d+E%d-F%d-G%d)' % ((r,) * 5)
    edit(M, 'I%d' % r, MONEY)
    M['J%d' % r] = '=IF(B%d="","",H%d-N(I%d))' % (r, r, r)
    for c in 'ABCDEFGHJ': auto(M, '%s%d' % (c, r))
    for c in 'BC': M['%s%d' % (c, r)].number_format = DATE
    for c in 'DEFGHJ': M['%s%d' % (c, r)].number_format = MONEY
M['A13'] = 'مجموع الشهر'
for c in 'DEFGHIJ': M['%s13' % c] = '=SUM(%s%d:%s%d)' % (c, MR0, c, MR0 + 3)
for c in 'ABCDEFGHIJ': auto(M, '%s13' % c, MONEY if c in 'DEFGHIJ' else None, True, color=SUBT)
M['A15'] = 'إقفال الشهر ← افتتاح الشهر الجاي:'; M['A15'].font = F(11, True, RUST)
M['C15'] = 'مستحق له'; M['D15'] = '=J13'; M['E15'] = 'دين عليه'
M['F15'] = '=C6+SUM(C%d:C%d)-SUM(D%d:D%d)-G13' % (ML0, MLE, ML0, MLE)
for c in ('D15', 'F15'): auto(M, c, MONEY, True, 12)
header(M, 17, ['التاريخ', 'سلفة (دينار)', 'دين أو بضاعة (دينار)', 'تسديد مباشر (دينار)', 'البيان'], RUST, 10, 28)
M.merge_cells('E17:J17')
for r in range(ML0, MLE + 1):
    edit(M, 'A%d' % r, DATE)
    for c in 'BCD': edit(M, '%s%d' % (c, r), MONEY)
    M.merge_cells('E%d:J%d' % (r, r)); edit(M, 'E%d' % r); M['E%d' % r].alignment = C('right')
dv_date(M, 'A%d:A%d' % (ML0, MLE), 'MGR_START', 'MEND', 'اكتب التاريخ يوم/شهر/سنة')

# ═══════════════════ عمال المهام ═══════════════════
T = wb.create_sheet('عمال المهام')
title(T, 'عمال المهام — بيندفعلهم مباشرة', PLUM, 'E', '="شهر "&TEXT(\'التسوية\'!$C$3,"00")&" / "&\'التسوية\'!$E$3')
for k, w in zip('ABCDE', [13, 20, 40, 16, 32]): T.column_dimensions[k].width = w
note(T, 'A3', 'كل مهمة سطر: التاريخ، العامل، شو عمل، وقديش أخذ. بلا أسعار ثابتة.', 'A3:E3')
header(T, 5, ['التاريخ', 'اسم العامل', 'المهمة', 'المبلغ المدفوع (دينار)', 'ملاحظات'], PLUM, 11, 30)
TE = 6 + TASK_ROWS - 1
for r in range(6, TE + 1):
    edit(T, 'A%d' % r, DATE); edit(T, 'B%d' % r); edit(T, 'C%d' % r); edit(T, 'D%d' % r, MONEY); edit(T, 'E%d' % r)
    for c in 'CE': T['%s%d' % (c, r)].alignment = C('right')
dv_date(T, 'A6:A%d' % TE)
T['C%d' % (TE + 1)] = 'مجموع الشهر'; auto(T, 'C%d' % (TE + 1), None, True, color=SUBT)
T['D%d' % (TE + 1)] = '=SUM(D6:D%d)' % TE; auto(T, 'D%d' % (TE + 1), MONEY, True, 12, color=SUBT)
T.freeze_panes = 'A6'
note(T, 'A%d' % (TE + 3), 'مثال: 07/10/2026 · محمد · تنزيل طبلية جبسمبورد · 10.000', 'A%d:E%d' % (TE + 3, TE + 3))

# ═══════════════════ التعليمات ═══════════════════
I = wb.create_sheet('التعليمات')
I.sheet_view.rightToLeft = True; I.sheet_view.showGridLines = False
I.column_dimensions['A'].width = 5; I.column_dimensions['B'].width = 118
I['B1'] = 'كيف يشتغل الملف'; I['B1'].font = F(16, True, NAVY)
steps = [
    ('أول كل شهر', None),
    ('1', 'خذ نسخة من الملف وسمّيها برقم الشهر. بورقة «التسوية» اختر رقم الشهر والسنة — أول الشهر وآخره وكل الأسابيع بتطلع لحالها.'),
    ('2', 'انقل أرقام «إقفال الشهر» من ملف الشهر الماضي لـ«افتتاح الشهر»: مستحق له لسا ما قبضه · دين متبقي عليه · القسط الأسبوعي. (وكمان للمدير بورقته.)'),
    ('الأسبوع اللي بيقطع بين شهرين', None),
    ('3', 'الملف فيه أيام الشهر بس. مثلاً شهر 10/2026 بيبدأ الخميس 01/10: «الأسبوع الأول» = الخميس والجمعة بس، والسبت 26/09 لـالأربعاء 30/09 كانوا بملف شهر 9.'),
    ('4', 'وآخر الشهر نفس الشي: السبت 31/10 بس بملف شهر 10، والباقي بملف شهر 11. اللي اشتغله وما انقبض بيطلع بـ«إقفال الشهر» ← «مستحق له» بافتتاح الشهر الجاي، وبينضاف لصافي أول أسبوع.'),
    ('كل يوم', None),
    ('5', 'بورقة الساعات: اختر الدخول والخروج من القائمة (كل ربع ساعة، AM/PM). عبدالعزيز: آخر اليوم «قبض؟ نعم» + المبلغ.'),
    ('6', 'ساعات تحميل الجبسمبورد: مرة وحدة تحت كل أسبوع (مجموع ساعات التحميل بالأسبوع).'),
    ('7', 'سلفة أو دين أو بضاعة: سطر بورقة «السلف والديون» — كل نوع إله عمود لحاله.'),
    ('آخر الأسبوع', None),
    ('8', 'بورقة «التسوية» اختر الأسبوع — بيطلعلك لكل واحد المستحق (مجبور لربع دينار) ناقص السلف والقسط = الصافي للدفع.'),
    ('9', 'لما تدفع: اكتب المبلغ بخانة «المدفوع» بسطر الأسبوع بجدول الشهر.'),
    ('الحماية', None),
    ('10', 'الخلايا الصفراء بس بتنكتب. الباقي محمي عشان المعادلات ما تنخرب (الحماية بلا كلمة سر — من «مراجعة ← إلغاء حماية الورقة» إذا لزم).'),
    ('التاريخ', None),
    ('11', 'كل التواريخ بالملف يوم/شهر/سنة (05/10/2026). خانات التاريخ ما بتقبل إلا تاريخ من الشهر المختار.'),
]
r = 3
for a, b in steps:
    if b is None: I['B%d' % r] = a; I['B%d' % r].font = F(12, True, TEAL); r += 1; continue
    I['A%d' % r] = a; I['A%d' % r].font = F(11, True); I['A%d' % r].alignment = C()
    I['B%d' % r] = b; I['B%d' % r].font = F(11); I['B%d' % r].alignment = C('right'); I.row_dimensions[r].height = 34; r += 1

# ═══════════════════ الترتيب والحماية ═══════════════════
order = ['التسوية', 'ساعات عمر', 'ساعات عبدالعزيز', 'السلف والديون', 'المدير', 'عمال المهام', 'القواعد', 'التعليمات', 'القوائم']
wb._sheets = [wb[n] for n in order]
wb.active = 0
for ws in wb.worksheets:
    for ref in INPUTS.get(ws.title, []):
        ws[ref].protection = Protection(locked=False)
    ws.protection.sheet = True
    ws.protection.formatColumns = False; ws.protection.formatRows = False
wb.save(OUT); print('saved', OUT)
