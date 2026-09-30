# -*- coding: utf-8 -*-
"""نظام الموظفين الشهري — محلات العون لمواد البناء
ملف لكل شهر: تختار رقم الشهر والسنة، والأسابيع وتواريخها بتطلع لحالها.

القواعد (26/09/2026 من صاحب المحل):
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
DUR = '[h]" س "mm" د"'   # مدة: 11 س 30 د (مش ساعة حائط)
DIN = '[<=31]"يوم "0;[<=3112]00\\/00;dd/mm/yyyy'   # خانة تاريخ بتقبل: 5 · 510 · 05/10/2026
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
def dv_date(ws, rng, *_):
    # قائمة أيام الشهر المختار — أو اكتب رقم اليوم لحاله (5) أو يوم وشهر (510) أو التاريخ كامل
    d = DataValidation(type='list', formula1="='القوائم'!$N$1:$N$31", allow_blank=True, showErrorMessage=False,
                       showInputMessage=True, promptTitle='التاريخ',
                       prompt='اختار من القائمة، أو اكتب رقم اليوم لحاله (مثل 5) — بينحسب من الشهر المختار')
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
def DVAL(x):   # رقم اليوم لحاله ← تاريخ بالشهر المختار · 4 أرقام يوم+شهر (0510) · أو تاريخ كامل
    return ('IF({x}="","",IF(ISNUMBER({x}),IF({x}<=31,DATE(YEAR(MSTART),MONTH(MSTART),{x}),'
            'IF({x}<=3112,DATE(YEAR(MSTART),MOD({x},100),INT({x}/100)),{x})),IFERROR(DATEVALUE({x}),"")))').format(x=x)
def RND(x): return 'IF(ROUND_MODE="للأعلى",CEILING(ROUND((%s),3),0.25),ROUND((%s)*4,0)/4)' % (x, x)

# ═══════════════════ ورقة مخفية: القوائم والحسابات ═══════════════════
H = wb.create_sheet('القوائم')
def label(m):
    h, mi = divmod(m, 60)
    return '%d:%02d %s' % ((h % 12) or 12, mi, 'AM' if h < 12 else 'PM')
NT = 288                        # كل 5 دقائق (بأمره 30/09) — 24 ساعة
for i, m in enumerate([(6 * 60 + 5 * k) % 1440 for k in range(NT)], start=1):   # تبدأ 6:00 AM
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
for i in range(31):
    H['N%d' % (i + 1)] = '=IF(%d<=DAY(MEND),MSTART+%d,"")' % (i + 1, i); H['N%d' % (i + 1)].number_format = DATE
H.sheet_state = 'hidden'
TIME_LIST = "='القوائم'!$A$1:$A$%d" % NT
def TV(ref):   # من القائمة (نص) · أو وقت مكتوب (رقم) · أو نص وقت ثاني — دايماً وقت صافي
    return ("IF(ISNUMBER({r}),MOD({r},1),IFERROR(INDEX('القوائم'!$B$1:$B${n},MATCH(TRIM({r}),'القوائم'!$A$1:$A${n},0)),"
            "IFERROR(MOD(TIMEVALUE({r}),1),\"\")))").format(r=ref, n=NT)
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
]
for i, (nm, rule, val, fmt, desc) in enumerate(rules, start=1):
    r = 3 + i
    K.cell(r, 1, i); K.cell(r, 2, rule); K.cell(r, 3, val); K.cell(r, 4, desc)
    for cc in (1, 2, 4): cellfmt(K.cell(r, cc), h='right' if cc > 1 else 'center')
    edit(K, 'C%d' % r, fmt, True); K.cell(r, 4).font = F(10, False, '555555'); K.row_dimensions[r].height = 30
    name(nm, "'القواعد'!$C$%d" % r)
    if nm == 'ROUND_MODE': dv(K, '"لأقرب ربع,للأعلى"', 'C%d' % r)

# ═══════════════════ أوراق الساعات ═══════════════════
FINAL, FINAL_C, FINAL_BG = [], 'C55A11', 'ED7D31'   # المستحق النهائي: خلفية برتقالية وخط أبيض عريض
def hours_sheet(ttl, banner, color, heads, widths, text, last='I'):
    ws = wb.create_sheet(ttl)
    title(ws, banner, color, last, '="شهر "&TEXT(\'التسوية\'!$C$3,"00")&" / "&\'التسوية\'!$E$3&"   (الشهر والسنة بتنختار من ورقة التسوية)"')
    note(ws, 'A3', text, 'A3:%s3' % last, h=36)
    for k, w in enumerate(widths, start=1): ws.column_dimensions[col(k)].width = w
    header(ws, 5, heads, color, 11, 32); ws.freeze_panes = 'A6'
    ws.column_dimensions['K'].hidden = True; ws.column_dimensions['L'].hidden = True; ws.column_dimensions['M'].hidden = True
    return ws

def week_bar(ws, k, color, last='I'):
    r = B0 + BLOCK * k; sat = 'WFIRST+%d' % (7 * k)
    ws.merge_cells('A%d:%s%d' % (r, last, r))
    ws['A%d' % r] = '=IF(%d<=NW,"الأسبوع %s","")' % (k + 1, ORD[k])
    cellfmt(ws['A%d' % r], WEEKBAR, True, sz=14, fc=color, h='right'); ws.row_dimensions[r].height = 28
    return r

def day_date(ws, r, k, d):
    ws['A%d' % r] = '=IF(AND(%d<=NW,WFIRST+%d>=MSTART,WFIRST+%d<=MEND),WFIRST+%d,"")' % (k + 1, 7 * k + d, 7 * k + d, 7 * k + d)   # برّا الشهر = فاضي (بأمره 29/09)
    ws['M%d' % r] = '=IF(A{r}="",0,IF(AND(A{r}>=MSTART,A{r}<=MEND),1,0))'.replace('{r}', str(r))   # 1 = من أيام الشهر
    ws['B%d' % r] = '=IF(A{r}="","",%s)'.replace('{r}', str(r)) % DAYNAME.format(d='A%d' % r)
    ws['K%d' % r] = '=IF(OR(M{r}=0,C{r}=""),"",%s)'.replace('{r}', str(r)) % TV('C%d' % r)
    ws['L%d' % r] = '=IF(OR(M{r}=0,D{r}=""),"",%s)'.replace('{r}', str(r)) % TV('D%d' % r)
    auto(ws, 'A%d' % r, DATE); auto(ws, 'B%d' % r)
    for c in 'KL': ws['%s%d' % (c, r)].number_format = 'h:mm AM/PM'
    if d == 6:
        for c in 'AB': ws['%s%d' % (c, r)].font = F(11, True, NAVY)

# ---- عمر ----
O = hours_sheet('ساعات عمر', 'ساعات عمر المصري', NAVY,
    ['التاريخ', 'اليوم', 'دخول', 'خروج', 'عدد الساعات', 'منها عادي', 'منها إضافي', 'المستحق (دينار)', 'ملاحظات'],
    [13, 10, 12, 12, 10, 11, 11, 14, 30],
    'اختر الدخول والخروج (AM/PM) — عدد الساعات والعادي والإضافي والمستحق بيطلعوا لحالهم. السبت–الخميس: بعد 6:00 PM إضافي @1.500 · الجمعة كل الساعات @1.250. '
    'ساعات تحميل الجبسمبورد بتنكتب مرة وحدة تحت كل أسبوع. مستحق الأسبوع بيتجبر لربع دينار.')
OMR = []   # (سطر مجموع الأيام, سطر الجبسمبورد)
for k in range(NWMAX):
    hr = week_bar(O, k, NAVY)
    for d in range(7):
        r = hr + 1 + d; day_date(O, r, k, d)
        O['E%d' % r] = '=IF(OR(M{r}=0,K{r}="",L{r}=""),"",MOD(L{r}-K{r},1))'.replace('{r}', str(r))
        O['F%d' % r] = ('=IF(OR(E{r}="",M{r}=0),"",IF(WEEKDAY(A{r},1)=6,E{r},'
                        'MIN(E{r},MAX(0,MIN(IF(L{r}<K{r},L{r}+1,L{r}),OMR_OT_AT)-K{r}))))').replace('{r}', str(r))
        O['G%d' % r] = '=IF(OR(E{r}="",M{r}=0),"",MAX(0,E{r}-F{r}))'.replace('{r}', str(r))
        O['H%d' % r] = '=IF(OR(E{r}="",M{r}=0),"",ROUND((F{r}*OMR_RATE+G{r}*OMR_OT)*24,3))'.replace('{r}', str(r))
        for c in 'CD': edit(O, '%s%d' % (c, r))
        edit(O, 'I%d' % r)
        for c, fm in zip('EFGH', [DUR, DUR, DUR, MONEY]): auto(O, '%s%d' % (c, r), fm)
    s1, s2 = hr + 8, hr + 9; lo, hi = hr + 1, hr + 7; on = '%d<=NW' % (k + 1)
    O.merge_cells('A%d:D%d' % (s1, s1)); O['A%d' % s1] = '=IF(%s,"مجموع أيام الأسبوع %s","")' % (on, ORD[k])
    for c in 'EFGH':
        O['%s%d' % (c, s1)] = '=IF(%s,SUM(%s%d:%s%d),"")' % (on, c, lo, c, hi)
    for c in 'ABCDEFGHI': auto(O, '%s%d' % (c, s1), None, True, color=SUBT)
    for c, fm in zip('EFGH', [DUR, DUR, DUR, MONEY]): O['%s%d' % (c, s1)].number_format = fm
    O.merge_cells('A%d:D%d' % (s2, s2)); O['A%d' % s2] = '=IF(%s,"ساعات تحميل الجبسمبورد بهالأسبوع (رقم: 2 أو 1.5) ←","")' % on
    auto(O, 'A%d' % s2, None, True, color=SUBT); edit(O, 'E%d' % s2, HRS, True)
    O.merge_cells('F%d:G%d' % (s2, s2)); O['F%d' % s2] = '=IF(%s,"المستحق النهائي للأسبوع (دينار) ←","")' % on
    auto(O, 'F%d' % s2, None, True, 12, color=SUBT, fc=FINAL_C)
    O['H%d' % s2] = '=IF(%s,%s,"")' % (on, RND('N(H%d)+N(E%d)*(GYP_RATE-OMR_RATE)' % (s1, s2)))
    auto(O, 'H%d' % s2, MONEY, True, 16, color=SUBT); FINAL.append((O, 'H%d' % s2)); O.row_dimensions[s2].height = 32
    O['I%d' % s2] = '=IF(%s,"قبل الجبر: "&TEXT(N(H%d)+N(E%d)*(GYP_RATE-OMR_RATE),"0.000"),"")' % (on, s1, s2)
    auto(O, 'I%d' % s2, None, color=SUBT); O['I%d' % s2].font = F(9, False, '555555')
    OMR.append((s1, s2))
OEND = B0 + BLOCK * NWMAX - 1
dv(O, TIME_LIST, 'C%d:D%d' % (B0, OEND))
O.conditional_formatting.add('C%d:D%d' % (B0, OEND), FormulaRule(
    formula=['AND($M%d=1,$K%d<>"",$L%d<>"",WEEKDAY($A%d,1)=6,OR($K%d<OMR_FRI_IN,$L%d>OMR_FRI_OUT))' % ((B0,) * 6)], fill=fill('FDE2C8')))

# ---- عبدالعزيز ----
Z = hours_sheet('ساعات عبدالعزيز', 'ساعات عبدالعزيز — بيقبض نهاية كل يوم', TEAL,
    ['التاريخ', 'اليوم', 'دخول', 'خروج', 'عدد الساعات', 'مستحق اليوم (دينار)', 'قبض؟', 'ملاحظات'],
    [13, 10, 12, 12, 10, 14, 9, 34],
    'دينار للساعة أي يوم. مستحق كل يوم مجبور لربع دينار. لما يقبض آخر اليوم: «قبض؟ نعم» — المبلغ = مستحق اليوم لحاله. '
    'ساعات الجبسمبورد تحت كل أسبوع — فرقها بينضاف لمستحق الأسبوع وبيندفع آخره.', 'H')   # 30/09 بأمره: عمود «المقبوض (دينار)» انشال
AZZ = []
for k in range(NWMAX):
    hr = week_bar(Z, k, TEAL, 'H')
    for d in range(7):
        r = hr + 1 + d; day_date(Z, r, k, d)
        Z['E%d' % r] = '=IF(OR(M{r}=0,K{r}="",L{r}=""),"",MOD(L{r}-K{r},1))'.replace('{r}', str(r))
        Z['F%d' % r] = '=IF(OR(E%d="",M%d=0),"",%s)' % (r, r, RND('E%d*24*AZZ_RATE' % r))
        for c in 'CDGH': edit(Z, '%s%d' % (c, r))
        auto(Z, 'E%d' % r, DUR); auto(Z, 'F%d' % r, MONEY)
    s1, s2 = hr + 8, hr + 9; lo, hi = hr + 1, hr + 7; on = '%d<=NW' % (k + 1)
    Z.merge_cells('A%d:D%d' % (s1, s1)); Z['A%d' % s1] = '=IF(%s,"مجموع أيام الأسبوع %s","")' % (on, ORD[k])
    for c in 'EF': Z['%s%d' % (c, s1)] = '=IF(%s,SUM(%s%d:%s%d),"")' % (on, c, lo, c, hi)
    Z['G%d' % s1] = '=IF(%s,SUMIFS(F%d:F%d,G%d:G%d,"نعم"),"")' % (on, lo, hi, lo, hi)   # اللي قبضه يومياً بالأسبوع
    for c in 'ABCDEFGH': auto(Z, '%s%d' % (c, s1), None, True, color=SUBT)
    Z['E%d' % s1].number_format = DUR; Z['F%d' % s1].number_format = MONEY
    Z['G%d' % s1].number_format = '"قبض "#,##0.000'
    Z.merge_cells('A%d:D%d' % (s2, s2)); Z['A%d' % s2] = '=IF(%s,"ساعات تحميل الجبسمبورد بهالأسبوع (رقم: 2 أو 1.5) ←","")' % on
    auto(Z, 'A%d' % s2, None, True, color=SUBT); edit(Z, 'E%d' % s2, HRS, True)
    Z['F%d' % s2] = '=IF(%s,N(F%d)+%s,"")' % (on, s1, RND('N(E%d)*(GYP_RATE-AZZ_RATE)' % s2))
    auto(Z, 'F%d' % s2, MONEY, True, 16, color=SUBT); FINAL.append((Z, 'F%d' % s2)); Z.row_dimensions[s2].height = 32
    Z.merge_cells('G%d:H%d' % (s2, s2)); Z['G%d' % s2] = '=IF(%s,"← المستحق النهائي للأسبوع (دينار) مع فرق الجبسمبورد","")' % on
    auto(Z, 'G%d' % s2, None, True, 12, color=SUBT, fc=FINAL_C)
    AZZ.append((s1, s2))
ZEND = B0 + BLOCK * NWMAX - 1
dv(Z, TIME_LIST, 'C%d:D%d' % (B0, ZEND)); dv(Z, '"نعم,لا"', 'G%d:G%d' % (B0, ZEND))
for ws, end in ((O, OEND), (Z, ZEND)):   # أيام برّا الشهر: فاضية ورمادي خفيف بخانات الدخول والخروج
    ws.conditional_formatting.add('C%d:D%d' % (B0, end), FormulaRule(formula=['AND($A%d="",$B%d="")' % (B0, B0)], fill=fill(OUTM)))

for ws_, ref in FINAL:   # البرتقالي بيطلع بس لما في رقم (الأسابيع اللي برّا الشهر بتضل هادية)
    ws_.conditional_formatting.add(ref, FormulaRule(formula=['%s<>""' % ref], fill=fill(FINAL_BG), font=Font(color='FFFFFF', bold=True)))
    ws_[ref].border = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)

# ═══════════════════ السلف والقروض ═══════════════════
# 29/09 بأمره: السلف الأسبوعية بورقة لحال، والقروض والديون وأقساطها بورقة لحال
# ---- السلف الأسبوعية ----
A = wb.create_sheet('السلف')
title(A, 'السلف الأسبوعية — بتنخصم كاملة من راتب نفس الأسبوع', RUST, 'D',
      '="شهر "&TEXT(\'التسوية\'!$C$3,"00")&" / "&\'التسوية\'!$E$3')
for k, w in zip('ABCD', [15, 18, 18, 40]): A.column_dimensions[k].width = w
header(A, 3, ['الموظف', 'سلف هالشهر (دينار)', 'سلف الأسبوع المختار (دينار)', ''], RUST, 10, 32)
note(A, 'A7', '• كاش بياخذه وسط الأسبوع ← بينخصم كامل من صافي نفس الأسبوع. القروض والبضاعة (اللي بتنخصم أقساط) بورقة «القروض والديون».', 'A7:D7', 10, '333333', 30)
header(A, 9, ['التاريخ', 'الموظف', 'المبلغ (دينار)', 'البيان'], RUST, 11, 30)
AR0, AE = 10, 10 + LOG_ROWS - 1
for r in range(AR0, AE + 1):
    edit(A, 'A%d' % r, DIN); edit(A, 'B%d' % r); edit(A, 'C%d' % r, MONEY); edit(A, 'D%d' % r); A['D%d' % r].alignment = C('right')
    A['F%d' % r] = '=' + DVAL('A%d' % r)
dv(A, '"عمر المصري,عبدالعزيز"', 'B%d:B%d' % (AR0, AE)); dv_date(A, 'A%d:A%d' % (AR0, AE))
A.column_dimensions['F'].hidden = True; A.freeze_panes = 'A10'
for nm, c in (('ADV_D', 'F'), ('ADV_E', 'B'), ('ADV_AMT', 'C')):
    name(nm, "'السلف'!$%s$%d:$%s$%d" % (c, AR0, c, AE))
for i, nm in enumerate(['عمر المصري', 'عبدالعزيز']):
    r = 4 + i; A['A%d' % r] = nm
    A['B%d' % r] = '=SUMIFS(ADV_AMT,ADV_E,A%d)' % r
    A['C%d' % r] = '=SUMIFS(ADV_AMT,ADV_E,A%d,ADV_D,">="&WS,ADV_D,"<="&WE)' % r
    for c in 'ABC': auto(A, '%s%d' % (c, r), MONEY if c != 'A' else None, True, 12)
note(A, 'A%d' % (AE + 2), 'مثال: 5 · عمر المصري · 20.000 · «كاش من الصندوق»', 'A%d:D%d' % (AE + 2, AE + 2))

# ---- القروض والديون ----
L = wb.create_sheet('القروض والديون')
title(L, 'القروض والديون — بتنخصم أقساط كل أسبوع', RUST, 'G',
      '="شهر "&TEXT(\'التسوية\'!$C$3,"00")&" / "&\'التسوية\'!$E$3')
for k, w in zip('ABCDEFG', [14, 16, 18, 16, 15, 16, 22]): L.column_dimensions[k].width = w
header(L, 3, ['الموظف', 'القروض والديون (مع الماضي)', 'انخصم أقساط هالشهر', 'دفع كاش من جيبته', 'الدين المتبقي', 'القسط الأسبوعي الحالي', 'قديش ضل'], RUST, 10, 34)
expl = [
    '• قرض أو دين أو بضاعة: اكتب المبلغ والقسط الأسبوعي بنفس السطر ← القسط بينخصم كل أسبوع لحاله لحد ما يخلص (آخر قسط = الباقي بس).',
    '• بدك تغيّر القسط؟ سطر جديد فيه القسط الجديد بس. · «دفع كاش من جيبته»: لو جاب مصاري من برّا الراتب (جزء أو كل الدين) ← بينقص من الدين فوراً، وإذا سدّ كله الأقساط بتوقف لحالها.',
]
for i, tx in enumerate(expl): note(L, 'A%d' % (7 + i), tx, 'A%d:G%d' % (7 + i, 7 + i), 10, '333333', 22)
header(L, 10, ['التاريخ', 'الموظف', 'قرض / دين / بضاعة (دينار)', 'القسط الأسبوعي (دينار)', 'دفع كاش من جيبته (دينار)', 'البيان', ''], RUST, 11, 34)
L.merge_cells('F10:G10')
LR0, LE = 11, 11 + LOG_ROWS - 1
for r in range(LR0, LE + 1):
    edit(L, 'A%d' % r, DIN); edit(L, 'B%d' % r)
    L['K%d' % r] = '=' + DVAL('A%d' % r)
    for c in 'CDE': edit(L, '%s%d' % (c, r), MONEY)
    L.merge_cells('F%d:G%d' % (r, r)); edit(L, 'F%d' % r); L['F%d' % r].alignment = C('right')
    L['J%d' % r] = r - LR0 + 1                      # ترتيب السطر (للقسط الأحدث)
dv(L, '"عمر المصري,عبدالعزيز"', 'B%d:B%d' % (LR0, LE)); dv_date(L, 'A%d:A%d' % (LR0, LE))
L.column_dimensions['J'].hidden = True; L.column_dimensions['K'].hidden = True
L.freeze_panes = 'A11'
for nm, c in (('LG_D', 'K'), ('LG_E', 'B'), ('LG_DEBT', 'C'), ('LG_INST', 'D'), ('LG_REP', 'E'), ('LG_IDX', 'J')):
    name(nm, "'القروض والديون'!$%s$%d:$%s$%d" % (c, LR0, c, LE))
note(L, 'A%d' % (LE + 2), 'أمثلة: 7 · عمر المصري · 497.000 · القسط 30.000 · «قرض»   |   20 · عمر المصري · دفع كاش من جيبته 50.000', 'A%d:G%d' % (LE + 2, LE + 2))

def INST(emp, d, opening):
    """القسط الساري بتاريخ معيّن = آخر قسط انكتب لهالموظف لحد هالتاريخ، وإلا قسط الافتتاح."""
    m = 'SUMPRODUCT(MAX((LG_E="%s")*(LG_INST<>"")*(LG_D<=%s)*LG_IDX))' % (emp, d)   # بلا MAXIFS — بيشتغل على إكسل 2010+
    return 'IF(%s=0,%s,INDEX(LG_INST,%s))' % (m, opening, m)

# ═══════════════════ التسوية (مبسّطة 29/09) ═══════════════════
S = wb['Sheet']; S.title = 'التسوية'
title(S, 'التسوية — قديش بدفع لكل موظف', NAVY, 'I')
for k, w in zip('ABCDEFGHI', [16, 14, 13, 12, 13, 17, 13, 13, 14]): S.column_dimensions[k].width = w
S['B3'] = 'اختر الشهر'; S['D3'] = 'اختر السنة'
for c in ('B3', 'D3'): S[c].font = F(13, True, NAVY); S[c].alignment = C('left')
S['C3'] = 10; edit(S, 'C3', '00', True, 14); dv(S, "='القوائم'!$D$1:$D$12", 'C3', 'رقم الشهر 1–12')
S['E3'] = 2026; edit(S, 'E3', '0', True, 14); dv(S, "='القوائم'!$E$1:$E$10", 'E3')
S.merge_cells('F3:I3'); note(S, 'F3', '← أول شي اختار الشهر والسنة — كل الملف بيتبرمج عليهم', sz=10, color=RUST)
S.row_dimensions[3].height = 30
S['B5'] = 'اختر الأسبوع'; S['B5'].font = F(13, True, NAVY); S['B5'].alignment = C('left')
S.merge_cells('C5:D5'); S['C5'] = 'الأسبوع الأول'; edit(S, 'C5', None, True, 13)
dv(S, "='القوائم'!$J$11:$J$16", 'C5', 'القائمة بتتغيّر حسب الشهر')
S.row_dimensions[5].height = 30
S['N5'] = "=IFERROR(MIN(MATCH($C$5,'القوائم'!$K$11:$K$16,0),NW),1)"; S['N5'].font = F(8, False, 'FFFFFF')
name('WK', "'التسوية'!$N$5"); name('WS', "INDEX(%s,'التسوية'!$N$5)" % WSTART); name('WE', "INDEX(%s,'التسوية'!$N$5)" % WEND)

# بداية ونهاية الشهر (ورقة لحال — المدوّر من الشهر الماضي وللشهر الجاي)
B = wb.create_sheet('بداية ونهاية الشهر'); BS = "'بداية ونهاية الشهر'"
title(B, 'بداية ونهاية الشهر — اللي بينتقل بين ملف شهر وملف الشهر اللي بعده', RUST, 'H',
      '="شهر "&TEXT(\'التسوية\'!$C$3,"00")&" / "&\'التسوية\'!$E$3')
for k, w in zip('ABCDEFGH', [16, 18, 18, 16, 14, 14, 14, 14]): B.column_dimensions[k].width = w
B['A4'] = 'بداية الشهر — انسخ الأرقام من «نهاية الشهر» بملف الشهر الماضي'; B['A4'].font = F(13, True, RUST); B.merge_cells('A4:H4'); B['A4'].alignment = C('right')
header(B, 5, ['الموظف', 'مستحق له من الشهر الماضي', 'دين متبقي عليه', 'القسط الأسبوعي'], RUST, 10, 34)
OPEN = {}
for i, nm in enumerate(['عمر المصري', 'عبدالعزيز']):
    r = 6 + i; B['A%d' % r] = nm; auto(B, 'A%d' % r, None, True, color='FFFFFF')
    for c in 'BCD': B['%s%d' % (c, r)] = 0; edit(B, '%s%d' % (c, r), MONEY, True)
    OPEN[nm] = r
note(B, 'A8', '«مستحق له» = أيام اشتغلها آخر الشهر الماضي وما انقبضت ← بينضاف لأول أسبوع. · «دين متبقي» والقسط ← بيكمّل ينخصم كل أسبوع.', 'A8:H8', 9)
def OP(c, nm): return '%s!$%s$%d' % (BS, c, OPEN[nm])

# الأعمدة الأساسية بس (30/09 بأمره «خلي الأشياء الأساسية») — نفسها للموظفين والجدولين
#  المدوّر من الشهر الماضي بينضاف لمستحق الأسبوع الأول · عبدالعزيز: اللي قبضه يومياً بينخصم من الصافي لحاله
COLS = ['الأسبوع', 'المستحق', 'إكرامية', 'سلف', 'قسط الدين', 'الصافي للدفع', 'المدفوع', 'الباقي له']
def month_table(top, emp, color, rows_ref, is_azz):
    S['A%d' % top] = emp + ' — كل أسابيع الشهر  (الإكرامية والمدفوع بالخانات الصفراء)'; S['A%d' % top].font = F(12, True, color)
    S.merge_cells('A%d:I%d' % (top, top)); S['A%d' % top].alignment = C('right')
    header(S, top + 1, COLS, color, 11, 30)
    r0 = top + 2; sh = "'ساعات عبدالعزيز'" if is_azz else "'ساعات عمر'"
    for k in range(NWMAX):
        r = r0 + k; on = '%d<=NW' % (k + 1); s1, s2 = rows_ref[k]
        we_ = "'القوائم'!$I$%d" % (11 + k); ws_ = "'القوائم'!$H$%d" % (11 + k)
        daily = 'N(%s!G%d)' % (sh, s1) if is_azz else '0'
        S['A%d' % r] = '=IF(%s,"الأسبوع %s","")' % (on, ORD[k])
        S['B%d' % r] = '=IF(%s,N(%s!%s%d)%s,"")' % (on, sh, 'F' if is_azz else 'H', s2, ('+' + OP('B', emp)) if k == 0 else '')
        edit(S, 'C%d' % r, MONEY)
        S['D%d' % r] = '=IF(%s,SUMIFS(ADV_AMT,ADV_E,"%s",ADV_D,">="&%s,ADV_D,"<="&%s),"")' % (on, emp, ws_, we_)
        owed = '{o}+SUMIFS(LG_DEBT,LG_E,"{e}",LG_D,"<="&{we})-SUMIFS(LG_REP,LG_E,"{e}",LG_D,"<="&{we})'.format(o=OP('C', emp), e=emp, we=we_)
        prev = '0' if k == 0 else 'SUM(E%d:E%d)' % (r0, r - 1)
        avail = 'N(B{r})+N(C{r})-N(D{r})-{d}'.format(r=r, d=daily)
        S['E%d' % r] = '=IF(%s,IF(N(B%d)=0,0,MAX(0,MIN(%s,%s-%s,%s))),"")' % (on, r, INST(emp, we_, OP('D', emp)), owed, prev, avail)
        S['F%d' % r] = '=IF(%s,N(B{r})+N(C{r})-N(D{r})-N(E{r})-%s,"")'.replace('{r}', str(r)) % (on, daily)
        edit(S, 'G%d' % r, MONEY)
        S['H%d' % r] = '=IF(%s,F%d-N(G%d),"")' % (on, r, r)
        S['M%d' % r] = '=IF(%s,MAX(0,%s-SUM(E%d:E%d)),"")' % (on, owed, r0, r)   # الدين المتبقي (مخفي — بيطلع فوق)
        for c in 'ABDEFH': auto(S, '%s%d' % (c, r))
        for c in 'BCDEFGH': S['%s%d' % (c, r)].number_format = MONEY
        S['M%d' % r].number_format = MONEY
        S['A%d' % r].font = F(11, True, color); S['F%d' % r].font = F(12, True, FINAL_C)
    tr = r0 + NWMAX
    S['A%d' % tr] = 'مجموع الشهر'
    for c in 'BCDEFGH': S['%s%d' % (c, tr)] = '=SUM(%s%d:%s%d)' % (c, r0, c, tr - 1)
    S['M%d' % tr] = '=MAX(0,%s+SUMIFS(LG_DEBT,LG_E,"%s")-SUMIFS(LG_REP,LG_E,"%s")-E%d)' % (OP('C', emp), emp, emp, tr)
    for c in 'ABCDEFGH': auto(S, '%s%d' % (c, tr), MONEY if c != 'A' else None, True, color=SUBT)
    S.conditional_formatting.add('A%d:H%d' % (r0, tr - 1), FormulaRule(formula=['AND($A%d<>"",ROW()-%d+1=WK)' % (r0, r0)], fill=fill('FFF2B3')))
    return r0, tr

# الأسبوع المختار — فوق، و«الصافي للدفع» مربّع برتقالي بارز
S['A7'] = '="تسوية  «"&$C$5&"»"'; S['A7'].font = F(14, True, NAVY); S.merge_cells('A7:I7'); S['A7'].alignment = C('right'); S.row_dimensions[7].height = 26
header(S, 8, ['الموظف'] + COLS[1:] + ['الدين المتبقي'], NAVY, 11, 30)
OR0, OTR = month_table(13, 'عمر المصري', NAVY, OMR, False)
ZR0, ZTR = month_table(OTR + 2, 'عبدالعزيز', TEAL, AZZ, True)
for i, (nm, r0) in enumerate([('عمر المصري', OR0), ('عبدالعزيز', ZR0)]):
    r = 9 + i; S['A%d' % r] = nm; auto(S, 'A%d' % r, None, True, 12)
    for c, src in zip('BCDEFGHI', 'BCDEFGHM'):
        S['%s%d' % (c, r)] = '=N(INDEX(%s%d:%s%d,WK))' % (src, r0, src, r0 + NWMAX - 1)
        auto(S, '%s%d' % (c, r), MONEY, False, 12)
    auto(S, 'F%d' % r, MONEY, True, 16, color=FINAL_BG, fc='FFFFFF')
    S['F%d' % r].border = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)
    S['I%d' % r].font = F(11, True, RUST); S.row_dimensions[r].height = 32
note(S, 'A11', 'المستحق = من ورقة الساعات (+ المدوّر من الشهر الماضي بالأسبوع الأول) · عبدالعزيز: الصافي بعد ما انخصم اللي قبضه يومياً · الإكرامية والمدفوع بتنكتب بجداول الشهر تحت.', 'A11:I11', 9)
S.column_dimensions['M'].hidden = True
S.freeze_panes = 'A6'

# نهاية الشهر
B['A11'] = 'نهاية الشهر — هالأرقام بتنكتب بـ«بداية الشهر» بملف الشهر الجاي'; B['A11'].font = F(13, True, RUST); B.merge_cells('A11:H11'); B['A11'].alignment = C('right')
header(B, 12, ['الموظف', 'مستحق له لسا ما قبضه', 'دين متبقي عليه', 'القسط الأسبوعي'], RUST, 10, 34)
for i, (nm, tr) in enumerate([('عمر المصري', OTR), ('عبدالعزيز', ZTR)]):
    r = 13 + i
    B['A%d' % r] = nm
    B['B%d' % r] = "='التسوية'!H%d" % tr
    B['C%d' % r] = "='التسوية'!M%d" % tr
    B['D%d' % r] = '=IF(C%d>0,%s,0)' % (r, INST(nm, 'MEND', '$D$%d' % OPEN[nm]))
    for c in 'ABCD': auto(B, '%s%d' % (c, r), MONEY if c != 'A' else None, True, 12)
    B.merge_cells('E%d:H%d' % (r, r))
    B['E%d' % r] = '=IF(B{r}>0.0005,"المحل لسا مدين إله بـ "&TEXT(B{r},"0.000"),IF(B{r}<-0.0005,"قبض زيادة "&TEXT(-B{r},"0.000")&" — بتنخصم الشهر الجاي","ما إله شي"))&IF(C{r}>0.0005,"  ·  وعليه دين "&TEXT(C{r},"0.000"),"")'.replace('{r}', str(r))
    B['E%d' % r].font = F(10, True, RUST); B['E%d' % r].alignment = C('right')

# ملخّص ورقة القروض
for i, (nm, tr) in enumerate([('عمر المصري', OTR), ('عبدالعزيز', ZTR)]):
    r = 4 + i
    L['A%d' % r] = nm
    L['B%d' % r] = "=%s+SUMIFS(LG_DEBT,LG_E,A%d)" % (OP('C', nm), r)
    L['C%d' % r] = "='التسوية'!E%d" % tr
    L['D%d' % r] = '=SUMIFS(LG_REP,LG_E,A%d)' % r
    L['E%d' % r] = '=MAX(0,B{r}-C{r}-D{r})'.replace('{r}', str(r))
    L['F%d' % r] = '=IF(E%d>0,%s,0)' % (r, INST(nm, 'MEND', OP('D', nm)))
    L['G%d' % r] = '=IF(E{r}<0.0005,"✔ ما عليه شي",IF(F{r}>0,ROUNDUP(E{r}/F{r},0)&" أسبوع تقريباً","⚠️ حدّد القسط"))'.replace('{r}', str(r))
    for c in 'ABCDEFG': auto(L, '%s%d' % (c, r), MONEY if c in 'BCDEF' else None, True, 11)
    L['E%d' % r].font = F(13, True, RUST); L.row_dimensions[r].height = 24

# ═══════════════════ عمال المهام ═══════════════════
T = wb.create_sheet('عمال المهام')
title(T, 'عمال المهام — بيندفعلهم مباشرة', PLUM, 'E', '="شهر "&TEXT(\'التسوية\'!$C$3,"00")&" / "&\'التسوية\'!$E$3')
for k, w in zip('ABCDE', [13, 20, 40, 16, 32]): T.column_dimensions[k].width = w
note(T, 'A3', 'كل مهمة سطر: التاريخ، العامل، شو عمل، وقديش أخذ. بلا أسعار ثابتة.', 'A3:E3')
header(T, 5, ['التاريخ', 'اسم العامل', 'المهمة', 'المبلغ المدفوع (دينار)', 'ملاحظات'], PLUM, 11, 30)
TE = 6 + TASK_ROWS - 1
for r in range(6, TE + 1):
    edit(T, 'A%d' % r, DIN); edit(T, 'B%d' % r); edit(T, 'C%d' % r); edit(T, 'D%d' % r, MONEY); edit(T, 'E%d' % r)
    for c in 'CE': T['%s%d' % (c, r)].alignment = C('right')
dv_date(T, 'A6:A%d' % TE)
T['C%d' % (TE + 1)] = 'مجموع الشهر'; auto(T, 'C%d' % (TE + 1), None, True, color=SUBT)
T['D%d' % (TE + 1)] = '=SUM(D6:D%d)' % TE; auto(T, 'D%d' % (TE + 1), MONEY, True, 12, color=SUBT)
T.freeze_panes = 'A6'
note(T, 'A%d' % (TE + 3), 'مثال: 07/10/2026 · محمد · تنزيل طبلية جبسمبورد · 10.000', 'A%d:E%d' % (TE + 3, TE + 3))

# عمود مخفي: ترقيم أول ظهور لكل عامل (للملخّص — بلا معادلات مصفوفة)
for r in range(6, TE + 1):
    T['G%d' % r] = '=IF(AND(B{r}<>"",COUNTIF(B$6:B{r},B{r})=1),MAX(G$5:G{p})+1,"")'.replace('{r}', str(r)).replace('{p}', str(r - 1))
T.column_dimensions['G'].hidden = True

# ═══════════════════ ملخص الأيدي العاملة (30/09 بأمره) ═══════════════════
M = wb.create_sheet('ملخص الأيدي العاملة')
title(M, 'ملخص الأيدي العاملة — وضع كل موظف وقديش دفعت هالشهر', NAVY, 'I',
      '="شهر "&TEXT(\'التسوية\'!$C$3,"00")&" / "&\'التسوية\'!$E$3')
for k, w in zip('ABCDEFGHI', [16, 13, 13, 12, 12, 13, 15, 13, 14]): M.column_dimensions[k].width = w
SS = "'التسوية'"
M['A4'] = 'الموظفين الثابتين'; M['A4'].font = F(13, True, NAVY); M.merge_cells('A4:I4'); M['A4'].alignment = C('right')
header(M, 5, ['الموظف', 'ساعات الشهر', 'المستحق', 'إكرامية', 'سلف', 'أقساط دين', 'قبض فعلياً', 'الباقي له', 'الدين المتبقي عليه'], NAVY, 10, 34)
for i, (nm, tr, sh, rows) in enumerate([('عمر المصري', OTR, "'ساعات عمر'", OMR), ('عبدالعزيز', ZTR, "'ساعات عبدالعزيز'", AZZ)]):
    r = 6 + i; M['A%d' % r] = nm
    M['B%d' % r] = '=' + '+'.join('N(%s!E%d)' % (sh, s1) for s1, _ in rows)
    M['C%d' % r] = '=%s!B%d' % (SS, tr); M['D%d' % r] = '=%s!C%d' % (SS, tr)
    M['E%d' % r] = '=%s!D%d' % (SS, tr); M['F%d' % r] = '=%s!E%d' % (SS, tr)
    daily = ('+' + '+'.join('N(%s!G%d)' % (sh, s1) for s1, _ in rows)) if nm == 'عبدالعزيز' else ''
    M['G%d' % r] = '=%s!G%d%s' % (SS, tr, daily)          # عبدالعزيز: الأسبوعي + اللي قبضه يومياً
    M['H%d' % r] = '=%s!H%d' % (SS, tr); M['I%d' % r] = '=%s!M%d' % (SS, tr)
    auto(M, 'A%d' % r, None, True, 12); auto(M, 'B%d' % r, DUR, False, 11)
    for c in 'CDEFGHI': auto(M, '%s%d' % (c, r), MONEY, False, 11)
    M['H%d' % r].font = F(12, True, FINAL_C); M['I%d' % r].font = F(11, True, RUST)
    M.row_dimensions[r].height = 26
# الوضع بكلمات
for i, nm in enumerate(['عمر المصري', 'عبدالعزيز']):
    r = 8 + i
    M['A%d' % r] = '="• "&A%d&": "&IF(H{r0}>0.0005,"المحل لسا مدين إله بـ "&TEXT(H{r0},"0.000"),IF(H{r0}<-0.0005,"قبض زيادة "&TEXT(-H{r0},"0.000"),"مخالص — ما إله شي"))&IF(I{r0}>0.0005,"  ·  وعليه دين "&TEXT(I{r0},"0.000"),"")'.replace('{r0}', str(6 + i)) % (6 + i)
    M.merge_cells('A%d:I%d' % (r, r)); M['A%d' % r].font = F(11, True, RUST); M['A%d' % r].alignment = C('right')

M['A11'] = 'عمال المهام'; M['A11'].font = F(13, True, PLUM); M.merge_cells('A11:I11'); M['A11'].alignment = C('right')
header(M, 12, ['#', 'العامل', 'عدد المهام', 'قبض (دينار)'], PLUM, 10, 28)
WK_ROWS = 15; TW0 = 13
TB, TD, TG = "'عمال المهام'!$B$6:$B$%d" % TE, "'عمال المهام'!$D$6:$D$%d" % TE, "'عمال المهام'!$G$6:$G$%d" % TE
for i in range(WK_ROWS):
    r = TW0 + i
    M['A%d' % r] = '=IF(B%d="","",%d)' % (r, i + 1)
    M['B%d' % r] = '=IFERROR(INDEX(%s,MATCH(%d,%s,0)),"")' % (TB, i + 1, TG)
    M['C%d' % r] = '=IF(B{r}="","",COUNTIF({tb},B{r}))'.replace('{r}', str(r)).replace('{tb}', TB)
    M['D%d' % r] = '=IF(B{r}="","",SUMIFS({td},{tb},B{r}))'.replace('{r}', str(r)).replace('{td}', TD).replace('{tb}', TB)
    for c in 'ABCD': auto(M, '%s%d' % (c, r), MONEY if c == 'D' else None)
TWT = TW0 + WK_ROWS
M['B%d' % TWT] = 'مجموع عمال المهام'; M['C%d' % TWT] = '=COUNTA(%s)' % TB; M['D%d' % TWT] = "='عمال المهام'!D%d" % (TE + 1)
for c in 'ABCD': auto(M, '%s%d' % (c, TWT), MONEY if c == 'D' else None, True, 11, color=SUBT)

P0 = TWT + 2
M['A%d' % P0] = 'قديش دفعت أيدي عاملة هالشهر (كاش طالع فعلياً)'; M['A%d' % P0].font = F(13, True, NAVY)
M.merge_cells('A%d:I%d' % (P0, P0)); M['A%d' % P0].alignment = C('right')
pay = [
    ('رواتب عمر المصري (المدفوع آخر الأسابيع)', "=%s!G%d" % (SS, OTR)),
    ('عبدالعزيز (يومي + آخر الأسابيع)', '=G7'),
    ('سلف كاش للموظفين', '=E6+E7'),
    ('عمال المهام', '=D%d' % TWT),
]
for i, (lb, fm) in enumerate(pay):
    r = P0 + 1 + i; M['A%d' % r] = lb; M.merge_cells('A%d:E%d' % (r, r)); M['F%d' % r] = fm
    auto(M, 'A%d' % r, None, False, 11); M['A%d' % r].alignment = C('right'); auto(M, 'F%d' % r, MONEY, True, 11)
TOT = P0 + 1 + len(pay)
M['A%d' % TOT] = 'المجموع اللي دفعته للأيدي العاملة'; M.merge_cells('A%d:E%d' % (TOT, TOT))
auto(M, 'A%d' % TOT, None, True, 13, color=SUBT, fc=FINAL_C); M['A%d' % TOT].alignment = C('right')
M['F%d' % TOT] = '=SUM(F%d:F%d)' % (P0 + 1, TOT - 1)
auto(M, 'F%d' % TOT, MONEY, True, 16, color=FINAL_BG, fc='FFFFFF')
M['F%d' % TOT].border = Border(left=THICK, right=THICK, top=THICK, bottom=THICK); M.row_dimensions[TOT].height = 32
info = [
    ('كلفة الشغل هالشهر (المستحق + الإكرامية + المهام)', '=C6+C7+D6+D7+D%d' % TWT),
    ('لسا عليك للموظفين (الباقي لهم)', '=MAX(0,H6)+MAX(0,H7)'),
    ('قروض وديون انعطت هالشهر (مش أجور — بترجع أقساط)', '=SUMIFS(LG_DEBT,LG_E,"عمر المصري")+SUMIFS(LG_DEBT,LG_E,"عبدالعزيز")'),
]
for i, (lb, fm) in enumerate(info):
    r = TOT + 2 + i; M['A%d' % r] = lb; M.merge_cells('A%d:E%d' % (r, r)); M['F%d' % r] = fm
    auto(M, 'A%d' % r, None, False, 10); M['A%d' % r].alignment = C('right'); auto(M, 'F%d' % r, MONEY, False, 11)
note(M, 'A%d' % (TOT + 6), 'كل الأرقام بتطلع لحالها من باقي الأوراق — ما في إشي بينكتب هون. «قبض فعلياً» = المدفوع بالتسوية (+ اليومي لعبدالعزيز). السلف كاش طالع، فبتنعدّ مع المدفوع.', 'A%d:I%d' % (TOT + 6, TOT + 6), 9)
M.freeze_panes = 'A4'

# ═══════════════════ التعليمات ═══════════════════
I = wb.create_sheet('التعليمات')
I.sheet_view.rightToLeft = True; I.sheet_view.showGridLines = False
I.column_dimensions['A'].width = 5; I.column_dimensions['B'].width = 118
I['B1'] = 'كيف يشتغل الملف'; I['B1'].font = F(16, True, NAVY)
steps = [
    ('أول كل شهر', None),
    ('1', 'خذ نسخة من الملف وسمّيها برقم الشهر. أول شي فوق بورقة «التسوية»: اختر الشهر والسنة — كل الملف (الأسابيع وتواريخ الأيام) بيتبرمج عليهم.'),
    ('2', 'بورقة «بداية ونهاية الشهر»: انسخ أرقام «نهاية الشهر» من ملف الشهر الماضي لـ«بداية الشهر»: مستحق له · دين متبقي عليه · القسط الأسبوعي.'),
    ('الأسبوع اللي بيقطع بين شهرين', None),
    ('3', 'الملف فيه أيام الشهر بس. مثلاً شهر 10/2026 بيبدأ الخميس 01/10: «الأسبوع الأول» = الخميس والجمعة بس، والسبت 26/09 لـالأربعاء 30/09 كانوا بملف شهر 9.'),
    ('4', 'وآخر الشهر نفس الشي: السبت 31/10 بس بملف شهر 10، والباقي بملف شهر 11. اللي اشتغله وما انقبض بيطلع بـ«نهاية الشهر» ← «مستحق له» ببداية الشهر الجاي، وبينضاف لصافي أول أسبوع.'),
    ('كل يوم', None),
    ('5', 'بورقة الساعات: اختر الدخول والخروج من القائمة (كل 5 دقائق، AM/PM) — أو اكتبه مثل 7:30 AM. عدد الساعات (مثل 11 س 30 د) والعادي والإضافي والمستحق بيطلعوا لحالهم. عبدالعزيز: آخر اليوم «قبض؟ نعم» بس — المبلغ هو مستحق اليوم.'),
    ('6', 'ساعات تحميل الجبسمبورد: مرة وحدة تحت كل أسبوع (مجموع ساعات التحميل بالأسبوع).'),
    ('7', 'سلفة الأسبوع (كاش وسط الأسبوع): سطر بورقة «السلف» — بتنخصم كاملة من نفس الأسبوع. قرض أو دين أو بضاعة: سطر بورقة «القروض والديون» فيه المبلغ والقسط الأسبوعي — القسط بينخصم لحاله كل أسبوع، وفوق الورقة ملخّص: قديش عليه، قديش ضل، وكم أسبوع.'),
    ('آخر الأسبوع', None),
    ('8', 'بورقة «التسوية» اختر الأسبوع — فوق بيطلعلك لكل موظف «الصافي للدفع» (المربّع البرتقالي) = المستحق + الإكرامية − السلف − قسط الدين. (عبدالعزيز: وكمان ناقص اللي قبضه يومياً.)'),
    ('9', 'الإكرامية (Tip) والمدفوع: بجدول الشهر تحت، بسطر الأسبوع (الخانات الصفراء).'),
    ('الحماية', None),
    ('10', 'الخلايا الصفراء بس بتنكتب. الباقي محمي عشان المعادلات ما تنخرب (الحماية بلا كلمة سر — من «مراجعة ← إلغاء حماية الورقة» إذا لزم).'),
    ('التاريخ', None),
    ('11', 'كل التواريخ بالملف يوم/شهر/سنة (05/10/2026). بخانة التاريخ: اختار من القائمة، أو اكتب رقم اليوم لحاله (5 = يوم 5 من الشهر المختار)، أو يوم وشهر (0510 = 05/10)، أو 5/10.'),
    ('آخر الشهر', None),
    ('12', 'ورقة «ملخص الأيدي العاملة» (آخر ورقة): وضع كل موظف (قبض، الباقي له، دينه)، عمال المهام كل واحد قديش قبض، وقديش دفعت أيدي عاملة بالشهر كامل.'),
]
r = 3
for a, b in steps:
    if b is None: I['B%d' % r] = a; I['B%d' % r].font = F(12, True, TEAL); r += 1; continue
    I['A%d' % r] = a; I['A%d' % r].font = F(11, True); I['A%d' % r].alignment = C()
    I['B%d' % r] = b; I['B%d' % r].font = F(11); I['B%d' % r].alignment = C('right'); I.row_dimensions[r].height = 34; r += 1

# ═══════════════════ الترتيب والحماية ═══════════════════
order = ['التسوية', 'ساعات عمر', 'ساعات عبدالعزيز', 'السلف', 'القروض والديون', 'بداية ونهاية الشهر', 'عمال المهام', 'القواعد', 'التعليمات', 'ملخص الأيدي العاملة', 'القوائم']
wb._sheets = [wb[n] for n in order]
wb.active = 0
for ws in wb.worksheets:
    for ref in INPUTS.get(ws.title, []):
        ws[ref].protection = Protection(locked=False)
    ws.protection.sheet = True
    ws.protection.formatColumns = False; ws.protection.formatRows = False
wb.save(OUT); print('saved', OUT)
