# -*- coding: utf-8 -*-
"""نظام الموظفين الشهري — محلات العون لمواد البناء
ملف لكل شهر: تختار الشهر والسنة، والأسابيع وتواريخها بتطلع لحالها.

القواعد (26/09/2026 من صاحب المحل):
  • المدير: 125 دينار كل 10 أيام، بلا ساعات.
  • عمر المصري: السبت–الخميس 7:00 AM–6:00 PM @1.250/ساعة · بعد 6:00 PM @1.500 · الجمعة 2:00 PM–11:00 PM حسب الساعات @1.250.
  • عبدالعزيز: 1.000/ساعة بغض النظر عن اليوم · يقبض نهاية كل يوم.
  • تحميل الجبسمبورد: أي موظف يتحاسب على ساعاته فيه @2.500.
  • عمال المهام: مقاولة لكل مهمة (بلا أسعار ثابتة).
  • المستحق لعمر وعبدالعزيز يُجبر لأقرب ربع دينار (عمر: مستحق الأسبوع · عبدالعزيز: مستحق اليوم لأنه بيقبض يومياً).

الأسبوع: السبت → الجمعة. **الأسبوع بيتبع الشهر اللي فيه جمعته (يوم القبض)** — فكل يوم بيدخل بملف شهر واحد بس.
الدخول والخروج: قائمة منسدلة كل ربع ساعة بنظام 12 ساعة (AM/PM).
"""
import openpyxl, datetime as dt
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName

OUT = '/home/user/Fam-Bam/العون/نظام_الموظفين.xlsx'
NAVY, TEAL, GOLD, PLUM, RUST = '1F4E6B', '0F6E6E', '8A6A00', '5B3A6B', '8B3A2F'
EDIT, AUTO, GREY, WEEKBAR, SUBT = 'FFF6CC', 'EDEDED', 'F5F5F5', 'DCE6EE', 'E3EFE3'
LINE = Side(style='thin', color='9AA5AD'); THICK = Side(style='medium', color=NAVY)
BOX = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
NWMAX = 5                      # أقصى عدد جُمَع بالشهر
BLOCK = 9                      # سطر عنوان الأسبوع + 7 أيام + سطر المجموع
B0 = 6                         # أول سطر بأول بلوك
LOG_ROWS, TASK_ROWS = 150, 60
MONTHS = ['كانون الثاني', 'شباط', 'آذار', 'نيسان', 'أيار', 'حزيران',
          'تموز', 'آب', 'أيلول', 'تشرين الأول', 'تشرين الثاني', 'كانون الأول']
MLABEL = ['%02d - %s' % (i + 1, m) for i, m in enumerate(MONTHS)]
YEARS = list(range(2026, 2036))

def F(sz=11, b=False, c='000000'): return Font(name='Arial', size=sz, bold=b, color=c)
def C(h='center'): return Alignment(horizontal=h, vertical='center', wrap_text=True)
def fill(c): return PatternFill('solid', fgColor=c)
def col(i): return openpyxl.utils.get_column_letter(i)

wb = openpyxl.Workbook()
def name(nm, ref): wb.defined_names[nm] = DefinedName(nm, attr_text=ref)
def dv(ws, formula, rng):
    d = DataValidation(type='list', formula1=formula, allow_blank=True); ws.add_data_validation(d); d.add(rng)
def cellfmt(c, color='FFFFFF', bold=False, fmt=None, sz=11, h='center', fc='000000'):
    c.border = BOX; c.fill = fill(color); c.font = F(sz, bold, fc); c.alignment = C(h)
    if fmt: c.number_format = fmt

DAYNAME = 'CHOOSE(WEEKDAY({d},1),"الأحد","الاثنين","الثلاثاء","الأربعاء","الخميس","الجمعة","السبت")'
def DMY(x): return 'DAY(%s)&"/"&MONTH(%s)' % (x, x)
def RND(x):  # الجبر لربع دينار حسب القاعدة
    return 'IF(ROUND_MODE="للأعلى",CEILING(ROUND((%s),3),0.25),ROUND((%s)*4,0)/4)' % (x, x)

# ═══════════════════ ورقة مخفية: القوائم ═══════════════════
H = wb.create_sheet('القوائم')
def label(m):
    h, mi = divmod(m, 60)
    return '%d:%02d %s' % ((h % 12) or 12, mi, 'AM' if h < 12 else 'PM')
for i, m in enumerate([(6 * 60 + 15 * k) % 1440 for k in range(96)], start=1):   # تبدأ 6:00 AM
    H.cell(i, 1, label(m)); H.cell(i, 2, dt.time(m // 60, m % 60)).number_format = 'h:mm AM/PM'
for i, m in enumerate(MLABEL, start=1): H.cell(i, 4, m)
for i, y in enumerate(YEARS, start=1): H.cell(i, 5, y)
H.sheet_state = 'hidden'
TIME_LIST = "='القوائم'!$A$1:$A$96"
def TV(ref): return "INDEX('القوائم'!$B$1:$B$96,MATCH(%s,'القوائم'!$A$1:$A$96,0))" % ref

# ═══════════════════ التسوية (الواجهة) ═══════════════════
S = wb.create_sheet('التسوية')
S.sheet_view.rightToLeft = True; S.sheet_view.showGridLines = False
for k, w in zip('ABCDEFGHIJKL', [4, 16, 13, 13, 11, 12, 11, 11, 12, 12, 12, 26]): S.column_dimensions[k].width = w
S.merge_cells('A1:L1'); S['A1'] = 'نظام الموظفين — التسوية'
cellfmt(S['A1'], NAVY, True, sz=16, fc='FFFFFF'); S.row_dimensions[1].height = 32
# اختيار الشهر والسنة والأسبوع
for c, txt in (('B3', 'الشهر'), ('E3', 'السنة'), ('H3', 'الأسبوع')):
    S[c] = txt; S[c].font = F(12, True, NAVY); S[c].alignment = C('left')
S.merge_cells('C3:D3'); S['C3'] = MLABEL[9]; cellfmt(S['C3'], EDIT, True, sz=12)
S['F3'] = 2026; cellfmt(S['F3'], EDIT, True, sz=12)
S.merge_cells('I3:L3'); cellfmt(S['I3'], EDIT, True, sz=12)
for cc in 'DJKL': S['%s3' % cc].border = BOX
dv(S, "='القوائم'!$D$1:$D$12", 'C3'); dv(S, "='القوائم'!$E$1:$E$10", 'F3')
S.row_dimensions[3].height = 28
name('SEL_M', "'التسوية'!$C$3"); name('SEL_Y', "'التسوية'!$F$3"); name('SEL_W', "'التسوية'!$I$3")
# حسابات مساعدة (صف 4 مخفي الملامح)
S['B4'] = 'أول الشهر'; S['C4'] = "=DATE(SEL_Y,MATCH(SEL_M,'القوائم'!$D$1:$D$12,0),1)"
S['D4'] = 'آخره'; S['E4'] = '=EOMONTH(C4,0)'
S['F4'] = 'أول جمعة'; S['G4'] = '=C4+MOD(6-WEEKDAY(C4,1),7)'
S['H4'] = 'عدد الأسابيع'; S['I4'] = '=INT((E4-G4)/7)+1'
S['J4'] = 'أول سبت'; S['K4'] = '=G4-6'
for c in ('C4', 'E4', 'G4', 'K4'): S[c].number_format = 'DD/MM/YYYY'
for c in 'BCDEFGHIJK': S['%s4' % c].font = F(9, False, '888888'); S['%s4' % c].alignment = C()
name('M1', "'التسوية'!$C$4"); name('MEND', "'التسوية'!$E$4"); name('NW', "'التسوية'!$I$4"); name('W1', "'التسوية'!$K$4")
# قائمة الأسابيع (للقائمة المنسدلة) بالقوائم المخفية عمود G
for k in range(NWMAX):
    st = 'W1+%d' % (7 * k); en = 'W1+%d' % (7 * k + 6)
    H.cell(k + 1, 7, '=IF(%d<=NW,"الأسبوع %d:  "&%s&" ← "&%s,"")' % (k + 1, k + 1, DMY(st), DMY(en)))
dv(S, "='القوائم'!$G$1:$G$%d" % NWMAX, 'I3')
S['I3'] = "=''"  # placeholder, replaced below
S['I3'] = None
S['B5'] = '=IF(SEL_W="","← اختر الأسبوع من القائمة (بيتغيّر حسب الشهر)","")'
S['B5'].font = F(10, True, 'C00000'); S.merge_cells('B5:L5'); S['B5'].alignment = C('right')
S['C6'] = '=IFERROR(MATCH(SEL_W,\'القوائم\'!$G$1:$G$%d,0),1)' % NWMAX
S['B6'] = 'رقم الأسبوع'; S['D6'] = 'من'; S['E6'] = '=W1+7*(C6-1)'; S['F6'] = 'إلى'; S['G6'] = '=E6+6'
for c in ('E6', 'G6'): S[c].number_format = 'DD/MM/YYYY'
for c in 'BCDEFG': S['%s6' % c].font = F(10, True, NAVY); S['%s6' % c].alignment = C()
name('WK', "'التسوية'!$C$6"); name('WS', "'التسوية'!$E$6"); name('WE', "'التسوية'!$G$6")

# ═══════════════════ القواعد ═══════════════════
K = wb.create_sheet('القواعد')
K.sheet_view.rightToLeft = True; K.sheet_view.showGridLines = False
K.merge_cells('A1:D1'); K['A1'] = 'القواعد والأسعار — كل المعادلات بتقرأ من هون'
cellfmt(K['A1'], GOLD, True, sz=15, fc='FFFFFF'); K.row_dimensions[1].height = 30
for k, w in zip('ABCD', [4, 46, 16, 56]): K.column_dimensions[k].width = w
for i, h in enumerate(['#', 'القاعدة', 'القيمة', 'الشرح'], start=1): cellfmt(K.cell(3, i, h), GOLD, True, fc='FFFFFF')
rules = [
    ('MGR_SAL', 'المدير — راتب كل 10 أيام', 125, '#,##0.000', 'بلا ساعات عمل.'),
    ('MGR_START', 'المدير — بداية أول فترة 10 أيام', dt.date(2026, 9, 26), 'DD/MM/YYYY', 'الفترات بتتسلسل من هالتاريخ كل 10 أيام. ثابت لكل الشهور.'),
    ('OMR_RATE', 'عمر المصري — أجر الساعة العادية', 1.25, '#,##0.000', 'السبت–الخميس لحد بداية الإضافي، والجمعة كل ساعاته.'),
    ('OMR_OT', 'عمر المصري — أجر الساعة بعد الدوام', 1.5, '#,##0.000', 'السبت–الخميس بعد الوقت المكتوب تحت.'),
    ('OMR_OT_AT', 'عمر المصري — بداية الإضافي', dt.time(18, 0), 'h:mm AM/PM', 'الجمعة ما إلها إضافي.'),
    ('OMR_FRI_IN', 'عمر المصري — الجمعة من', dt.time(14, 0), 'h:mm AM/PM', 'للتنبيه بس — الحساب حسب الساعات الفعلية.'),
    ('OMR_FRI_OUT', 'عمر المصري — الجمعة إلى', dt.time(23, 0), 'h:mm AM/PM', ''),
    ('AZZ_RATE', 'عبدالعزيز — أجر الساعة', 1.0, '#,##0.000', 'بغض النظر عن اليوم. بيقبض نهاية كل يوم.'),
    ('GYP_RATE', 'تحميل جبسمبورد — أجر الساعة', 2.5, '#,##0.000', 'ساعات التحميل بتنحسب @2.5 بدل الأجر العادي.'),
    ('ROUND_MODE', 'جبر المستحق (عمر وعبدالعزيز)', 'لأقرب ربع', '@', 'لأقرب ربع: 18.30←18.25 و18.40←18.50 · للأعلى: أي كسر بيطلع للربع اللي فوقه.'),
    ('OMR_INST', 'عمر المصري — قسط أسبوعي ثابت', 0, '#,##0.000', 'دين أو بضاعة. 0 = ما عليه.'),
    ('AZZ_INST', 'عبدالعزيز — قسط أسبوعي ثابت', 0, '#,##0.000', ''),
    ('MGR_INST', 'المدير — قسط ثابت كل 10 أيام', 0, '#,##0.000', ''),
    ('ADV_CAP', 'سقف السلف — نسبة من مستحق الأسبوع', 0.5, '0%', 'تنبيه أحمر لو تجاوزت. 0 = بلا سقف.'),
]
for i, (nm, rule, val, fmt, desc) in enumerate(rules, start=1):
    r = 3 + i
    K.cell(r, 1, i); K.cell(r, 2, rule); K.cell(r, 3, val); K.cell(r, 4, desc)
    for cc in range(1, 5): cellfmt(K.cell(r, cc), h='right' if cc in (2, 4) else 'center')
    cellfmt(K.cell(r, 3), EDIT, True, fmt); K.cell(r, 4).font = F(10, False, '555555'); K.row_dimensions[r].height = 30
    name(nm, "'القواعد'!$C$%d" % r)
    if nm == 'ROUND_MODE': dv(K, '"لأقرب ربع,للأعلى"', 'C%d' % r)
K['B19'] = 'عمال المهام: بلا أسعار ثابتة — كل مهمة مقاولة بسعرها بورقة «المهام».'
K['B19'].font = F(10, True, GOLD); K['B19'].alignment = C('right')

# ═══════════════════ أوراق الساعات (بلوك لكل أسبوع) ═══════════════════
def hours_sheet(title, banner, color, headers, widths, note):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = True; ws.sheet_view.showGridLines = False
    n = len(headers); last = col(n)
    ws.merge_cells('A1:%s1' % last); ws['A1'] = banner; cellfmt(ws['A1'], color, True, sz=15, fc='FFFFFF')
    ws.row_dimensions[1].height = 30
    ws.merge_cells('A2:%s2' % last)
    ws['A2'] = '=SEL_M&"  "&SEL_Y&"   ·   "&NW&" أسابيع   ·   من "&%s&" لحد "&%s' % (DMY('W1'), DMY('W1+7*NW-1'))
    ws['A2'].font = F(12, True, color); ws['A2'].alignment = C()
    ws.merge_cells('A3:%s3' % last); ws['A3'] = note
    ws['A3'].font = F(10, False, '444444'); ws['A3'].alignment = C('right'); ws.row_dimensions[3].height = 34
    for k, w in enumerate(widths, start=1): ws.column_dimensions[col(k)].width = w
    for i, h in enumerate(headers, start=1): cellfmt(ws.cell(5, i, h), color, True, fc='FFFFFF')
    ws.row_dimensions[5].height = 32; ws.freeze_panes = 'A6'
    return ws

def week_header(ws, k, ncols, color):
    r = B0 + BLOCK * k
    ws.merge_cells('A%d:%s%d' % (r, col(ncols), r))
    ws['A%d' % r] = '=IF(%d<=NW,"الأسبوع %d   ·   السبت "&%s&"  ←  الجمعة "&%s,"")' % (
        k + 1, k + 1, DMY('W1+%d' % (7 * k)), DMY('W1+%d' % (7 * k + 6)))
    cellfmt(ws['A%d' % r], WEEKBAR, True, fc=color); ws.row_dimensions[r].height = 22
    return r

# ---- عمر ----
O = hours_sheet('ساعات عمر', 'ساعات عمر المصري', NAVY,
    ['التاريخ', 'اليوم', 'دخول', 'خروج', 'منها جبسمبورد (ساعات)', 'الساعات', 'عادي', 'إضافي', 'المستحق', 'ملاحظات'],
    [12, 10, 12, 12, 13, 9, 9, 9, 11, 28],
    'اختر الدخول والخروج من القائمة. السبت–الخميس: بعد 6:00 PM إضافي @1.5 · الجمعة كل الساعات @1.25 · ساعات الجبسمبورد @2.5 بدلها. مستحق الأسبوع بيتجبر لربع دينار بسطر المجموع.')
O.column_dimensions['K'].hidden = True; O.column_dimensions['L'].hidden = True
OMR_SUB = []
for k in range(NWMAX):
    hr = week_header(O, k, 10, NAVY)
    for d in range(7):
        r = hr + 1 + d
        A = 'A%d' % r
        O[A] = '=IF(%d<=NW,W1+%d,"")' % (k + 1, 7 * k + d)
        O['B%d' % r] = '=IF(%s="","",%s)' % (A, DAYNAME.format(d=A))
        O['K%d' % r] = '=IF(C%d="","",%s)' % (r, TV('C%d' % r))
        O['L%d' % r] = '=IF(D%d="","",%s)' % (r, TV('D%d' % r))
        O['F%d' % r] = '=IF(OR(K{r}="",L{r}=""),"",ROUND(MOD(L{r}-K{r},1)*24,2))'.replace('{r}', str(r))
        O['G%d' % r] = ('=IF(F{r}="","",IF(WEEKDAY(A{r},1)=6,F{r},'
                        'ROUND(MAX(0,MIN(IF(L{r}<K{r},L{r}+1,L{r}),OMR_OT_AT)-K{r})*24,2)))').replace('{r}', str(r))
        O['H%d' % r] = '=IF(F{r}="","",ROUND(F{r}-G{r},2))'.replace('{r}', str(r))
        O['I%d' % r] = ('=IF(F{r}="","",ROUND(G{r}*OMR_RATE+H{r}*OMR_OT'
                        '+MIN(N(E{r}),G{r})*(GYP_RATE-OMR_RATE)+MAX(0,MIN(N(E{r}),F{r})-G{r})*(GYP_RATE-OMR_OT),3))').replace('{r}', str(r))
        for cc in 'ABFGHI': cellfmt(O['%s%d' % (cc, r)], AUTO)
        for cc in 'CDEJ': cellfmt(O['%s%d' % (cc, r)], EDIT)
        O[A].number_format = 'DD/MM/YYYY'; O['I%d' % r].number_format = '#,##0.000'
        for cc in 'EFGH': O['%s%d' % (cc, r)].number_format = '0.00'
        for cc in 'KL': O['%s%d' % (cc, r)].number_format = 'h:mm AM/PM'
        if d == 6:  # الجمعة
            for cc in 'AB': O['%s%d' % (cc, r)].font = F(11, True, NAVY)
    s = hr + 8; OMR_SUB.append(s); lo, hi = hr + 1, hr + 7
    O['B%d' % s] = '=IF(%d<=NW,"المجموع","")' % (k + 1)
    for cc in 'FGH': O['%s%d' % (cc, s)] = '=IF(%d<=NW,SUM(%s%d:%s%d),"")' % (k + 1, cc, lo, cc, hi)
    O['I%d' % s] = '=IF(%d<=NW,%s,"")' % (k + 1, RND('SUM(I%d:I%d)' % (lo, hi)))
    O['J%d' % s] = '=IF(%d<=NW,"قبل الجبر: "&TEXT(SUM(I%d:I%d),"0.000"),"")' % (k + 1, lo, hi)
    for cc in 'ABCDEFGHIJ': cellfmt(O['%s%d' % (cc, s)], SUBT, True)
    O['I%d' % s].number_format = '#,##0.000'; O['J%d' % s].font = F(9, False, '555555')
    for cc in 'FGH': O['%s%d' % (cc, s)].number_format = '0.00'
OEND = B0 + BLOCK * NWMAX - 1
dv(O, TIME_LIST, 'C%d:D%d' % (B0, OEND))
O.conditional_formatting.add('C%d:D%d' % (B0, OEND), FormulaRule(
    formula=['AND(ISNUMBER($A%d),$K%d<>"",$L%d<>"",WEEKDAY($A%d,1)=6,OR($K%d<OMR_FRI_IN,$L%d>OMR_FRI_OUT))' % ((B0,) * 6)], fill=fill('FDE2C8')))
O.conditional_formatting.add('E%d:E%d' % (B0, OEND), FormulaRule(formula=['AND($E%d<>"",ISNUMBER($F%d),$E%d>$F%d)' % ((B0,) * 4)], fill=fill('FBD5D5')))

# ---- عبدالعزيز ----
Z = hours_sheet('ساعات عبدالعزيز', 'ساعات عبدالعزيز — بيقبض نهاية كل يوم', TEAL,
    ['التاريخ', 'اليوم', 'دخول', 'خروج', 'منها جبسمبورد (ساعات)', 'الساعات', 'المستحق (مجبور)', 'قبض؟', 'المقبوض', 'ملاحظات'],
    [12, 10, 12, 12, 13, 9, 12, 8, 11, 28],
    'دينار للساعة أي يوم · ساعات الجبسمبورد @2.5 · مستحق كل يوم بيتجبر لربع دينار لأنه بيقبض يومياً. لما يقبض: «نعم» + المبلغ.')
Z.column_dimensions['K'].hidden = True; Z.column_dimensions['L'].hidden = True
AZZ_SUB = []
for k in range(NWMAX):
    hr = week_header(Z, k, 10, TEAL)
    for d in range(7):
        r = hr + 1 + d; A = 'A%d' % r
        Z[A] = '=IF(%d<=NW,W1+%d,"")' % (k + 1, 7 * k + d)
        Z['B%d' % r] = '=IF(%s="","",%s)' % (A, DAYNAME.format(d=A))
        Z['K%d' % r] = '=IF(C%d="","",%s)' % (r, TV('C%d' % r))
        Z['L%d' % r] = '=IF(D%d="","",%s)' % (r, TV('D%d' % r))
        Z['F%d' % r] = '=IF(OR(K{r}="",L{r}=""),"",ROUND(MOD(L{r}-K{r},1)*24,2))'.replace('{r}', str(r))
        Z['G%d' % r] = '=IF(F%d="","",%s)' % (r, RND('F{r}*AZZ_RATE+MIN(N(E{r}),F{r})*(GYP_RATE-AZZ_RATE)'.replace('{r}', str(r))))
        for cc in 'ABFG': cellfmt(Z['%s%d' % (cc, r)], AUTO)
        for cc in 'CDEHIJ': cellfmt(Z['%s%d' % (cc, r)], EDIT)
        Z[A].number_format = 'DD/MM/YYYY'
        for cc in 'GI': Z['%s%d' % (cc, r)].number_format = '#,##0.000'
        for cc in 'EF': Z['%s%d' % (cc, r)].number_format = '0.00'
        for cc in 'KL': Z['%s%d' % (cc, r)].number_format = 'h:mm AM/PM'
        if d == 6:
            for cc in 'AB': Z['%s%d' % (cc, r)].font = F(11, True, TEAL)
    s = hr + 8; AZZ_SUB.append(s); lo, hi = hr + 1, hr + 7
    Z['B%d' % s] = '=IF(%d<=NW,"المجموع","")' % (k + 1)
    for cc in 'FGI': Z['%s%d' % (cc, s)] = '=IF(%d<=NW,SUM(%s%d:%s%d),"")' % (k + 1, cc, lo, cc, hi)
    Z['J%d' % s] = '=IF(%d<=NW,IF(ROUND(G%d-I%d,3)=0,"✔ مقبوض كامل","باقي له: "&TEXT(G%d-I%d,"0.000")),"")' % (k + 1, s, s, s, s)
    for cc in 'ABCDEFGHIJ': cellfmt(Z['%s%d' % (cc, s)], SUBT, True)
    for cc in 'GI': Z['%s%d' % (cc, s)].number_format = '#,##0.000'
    Z['F%d' % s].number_format = '0.00'; Z['J%d' % s].font = F(10, True, TEAL)
ZEND = B0 + BLOCK * NWMAX - 1
dv(Z, TIME_LIST, 'C%d:D%d' % (B0, ZEND)); dv(Z, '"نعم,لا"', 'H%d:H%d' % (B0, ZEND))
Z.conditional_formatting.add('I%d:I%d' % (B0, ZEND), FormulaRule(formula=['AND($H%d="نعم",ISNUMBER($G%d),$I%d<>$G%d)' % ((B0,) * 4)], fill=fill('FDE2C8')))

# ═══════════════════ السلف والديون ═══════════════════
def plain_sheet(title, banner, color, headers, widths, nrows, note):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = True; ws.sheet_view.showGridLines = False
    n = len(headers); last = col(n)
    ws.merge_cells('A1:%s1' % last); ws['A1'] = banner; cellfmt(ws['A1'], color, True, sz=15, fc='FFFFFF'); ws.row_dimensions[1].height = 30
    ws.merge_cells('A2:%s2' % last); ws['A2'] = '=SEL_M&"  "&SEL_Y'; ws['A2'].font = F(12, True, color); ws['A2'].alignment = C()
    ws.merge_cells('A3:%s3' % last); ws['A3'] = note; ws['A3'].font = F(10, False, '444444'); ws['A3'].alignment = C('right'); ws.row_dimensions[3].height = 34
    for k, w in enumerate(widths, start=1): ws.column_dimensions[col(k)].width = w
    for i, h in enumerate(headers, start=1): cellfmt(ws.cell(5, i, h), color, True, fc='FFFFFF')
    ws.row_dimensions[5].height = 30
    for r in range(6, 6 + nrows):
        for cc in range(1, n + 1): cellfmt(ws.cell(r, cc), EDIT)
    ws.freeze_panes = 'A6'
    return ws
L = plain_sheet('السلف والديون', 'السلف والديون والبضاعة والدفعات', RUST,
    ['التاريخ', 'الموظف', 'النوع', 'المبلغ (دينار)', 'البيان'], [13, 18, 20, 14, 50], LOG_ROWS,
    'سلفة = كاش بينخصم من صافي الأسبوع · دين / بضاعة = عليه وبيتسدّد بأقساط · دفع = ما قبضه · تسديد = دفع هو للمحل · رصيد مدوّر = رصيد آخر الشهر الماضي (له أو عليه).')
LE = 6 + LOG_ROWS - 1
for r in range(6, LE + 1): L['A%d' % r].number_format = 'DD/MM/YYYY'; L['D%d' % r].number_format = '#,##0.000'
dv(L, '"عمر المصري,عبدالعزيز,المدير"', 'B6:B%d' % LE)
dv(L, '"سلفة,دين,بضاعة,دفع,تسديد,رصيد مدوّر له,رصيد مدوّر عليه"', 'C6:C%d' % LE)
name('LOG_D', "'السلف والديون'!$A$6:$A$%d" % LE); name('LOG_E', "'السلف والديون'!$B$6:$B$%d" % LE)
name('LOG_T', "'السلف والديون'!$C$6:$C$%d" % LE); name('LOG_A', "'السلف والديون'!$D$6:$D$%d" % LE)

# ═══════════════════ المهام ═══════════════════
T = plain_sheet('المهام', 'عمال المهام — مقاولة لكل مهمة', PLUM,
    ['التاريخ', 'اسم العامل', 'المهمة', 'السعر المتفق', 'دُفع؟', 'تاريخ الدفع', 'ملاحظات'], [13, 18, 38, 13, 8, 13, 28], TASK_ROWS,
    'بلا أسعار ثابتة: كل مهمة بسعرها المتفق مع العامل.')
TE = 6 + TASK_ROWS - 1
for r in range(6, TE + 1):
    for cc in 'AF': T['%s%d' % (cc, r)].number_format = 'DD/MM/YYYY'
    T['D%d' % r].number_format = '#,##0.000'
dv(T, '"نعم,لا"', 'E6:E%d' % TE)
T['C%d' % (TE + 1)] = 'غير مدفوع'; T['C%d' % (TE + 1)].font = F(11, True); T['C%d' % (TE + 1)].alignment = C('left')
T['D%d' % (TE + 1)] = '=SUMIFS(D6:D%d,E6:E%d,"لا")' % (TE, TE); cellfmt(T['D%d' % (TE + 1)], AUTO, True, '#,##0.000')
T['C%d' % (TE + 2)] = 'مجموع الشهر'; T['C%d' % (TE + 2)].font = F(11, True); T['C%d' % (TE + 2)].alignment = C('left')
T['D%d' % (TE + 2)] = '=SUM(D6:D%d)' % TE; cellfmt(T['D%d' % (TE + 2)], AUTO, True, '#,##0.000')

# ═══════════════════ المدير ═══════════════════
M = wb.create_sheet('المدير')
M.sheet_view.rightToLeft = True; M.sheet_view.showGridLines = False
M.merge_cells('A1:J1'); M['A1'] = 'المدير — 125 دينار كل 10 أيام'; cellfmt(M['A1'], GOLD, True, sz=15, fc='FFFFFF'); M.row_dimensions[1].height = 30
M.merge_cells('A2:J2'); M['A2'] = '=SEL_M&"  "&SEL_Y&"   ·   الفترات اللي بتخلص بهالشهر"'; M['A2'].font = F(12, True, GOLD); M['A2'].alignment = C()
M.merge_cells('A3:J3'); M['A3'] = 'الفترات بتطلع لحالها من «بداية أول فترة» بورقة القواعد. الفترة بتتبع الشهر اللي بتخلص فيه.'
M['A3'].font = F(10, False, '444444'); M['A3'].alignment = C('right')
for k, w in zip('ABCDEFGHIJ', [5, 12, 12, 11, 11, 10, 11, 8, 11, 26]): M.column_dimensions[k].width = w
for i, h in enumerate(['#', 'من', 'إلى', 'الراتب', 'سلف الفترة', 'قسط', 'الصافي', 'دُفع؟', 'المدفوع', 'ملاحظات'], start=1):
    cellfmt(M.cell(5, i, h), GOLD, True, fc='FFFFFF')
M['L4'] = 'أول فترة بالشهر'; M['M4'] = '=MAX(0,ROUNDUP((M1-MGR_START-9)/10,0))'; M['L4'].font = F(8, False, '999999'); M['M4'].font = F(8, False, '999999')
for i in range(4):
    r = 6 + i
    n = '($M$4+%d)' % i
    M['B%d' % r] = '=IF(MGR_START+10*%s+9<=MEND,MGR_START+10*%s,"")' % (n, n)
    M['A%d' % r] = '=IF(B%d="","",%d)' % (r, i + 1)
    M['C%d' % r] = '=IF(B%d="","",B%d+9)' % (r, r)
    M['D%d' % r] = '=IF(B%d="","",MGR_SAL)' % r
    M['E%d' % r] = '=IF(B%d="","",SUMIFS(LOG_A,LOG_E,"المدير",LOG_T,"سلفة",LOG_D,">="&B%d,LOG_D,"<="&C%d))' % (r, r, r)
    M['F%d' % r] = '=IF(B%d="","",MGR_INST)' % r
    M['G%d' % r] = '=IF(B%d="","",D%d-E%d-F%d)' % (r, r, r, r)
    for cc in 'ABCDEFG': cellfmt(M['%s%d' % (cc, r)], AUTO)
    for cc in 'HIJ': cellfmt(M['%s%d' % (cc, r)], EDIT)
    for cc in 'BC': M['%s%d' % (cc, r)].number_format = 'DD/MM/YYYY'
    for cc in 'DEFGI': M['%s%d' % (cc, r)].number_format = '#,##0.000'
dv(M, '"نعم,لا"', 'H6:H9')
M['C10'] = 'مجموع الشهر'; M['C10'].font = F(11, True)
for cc in 'DEFGI': M['%s10' % cc] = '=SUM(%s6:%s9)' % (cc, cc); cellfmt(M['%s10' % cc], AUTO, True, '#,##0.000')

# ═══════════════════ التسوية: الأسبوع المختار + الشهر كامل + الرصيد ═══════════════════
OSUB = "'ساعات عمر'!$I$1:$I$%d" % OEND; OSUBH = "'ساعات عمر'!$F$1:$F$%d" % OEND
ZSUBG = "'ساعات عبدالعزيز'!$G$1:$G$%d" % ZEND; ZSUBI = "'ساعات عبدالعزيز'!$I$1:$I$%d" % ZEND; ZSUBH = "'ساعات عبدالعزيز'!$F$1:$F$%d" % ZEND
def subrow(k): return '(%d+%d*(%s-1))' % (B0 + 8, BLOCK, k)

S['B8'] = '=IF(SEL_W="","","تسوية "&SEL_W)'; S['B8'].font = F(13, True, NAVY); S.merge_cells('B8:L8'); S['B8'].alignment = C('right')
hdr = ['#', 'الموظف', 'الساعات', 'المستحق (مجبور)', 'السلف', 'القسط', 'مقبوض يومياً', 'الصافي للدفع', 'ملاحظة']
for i, h in enumerate(hdr, start=1): cellfmt(S.cell(9, i, h), NAVY, True, fc='FFFFFF')
S.merge_cells('I9:L9')
S.row_dimensions[9].height = 30
emp = [('عمر المصري', OSUBH, OSUB, None, 'OMR_INST'), ('عبدالعزيز', ZSUBH, ZSUBG, ZSUBI, 'AZZ_INST')]
for i, (nm, hrs, due, paid, inst) in enumerate(emp):
    r = 10 + i
    S['A%d' % r] = i + 1; S['B%d' % r] = nm
    S['C%d' % r] = '=N(INDEX(%s,%s))' % (hrs, subrow('WK'))
    S['D%d' % r] = '=N(INDEX(%s,%s))' % (due, subrow('WK'))
    S['E%d' % r] = '=SUMIFS(LOG_A,LOG_E,B%d,LOG_T,"سلفة",LOG_D,">="&WS,LOG_D,"<="&WE)' % r
    S['F%d' % r] = '=IF(D%d=0,0,%s)' % (r, inst)
    S['G%d' % r] = '=N(INDEX(%s,%s))' % (paid, subrow('WK')) if paid else 0
    S['H%d' % r] = '=D%d-E%d-F%d-G%d' % (r, r, r, r)
    S['I%d' % r] = '=IF(AND(ADV_CAP>0,D%d>0,E%d>ADV_CAP*D%d),"⚠️ السلف تجاوزت السقف",IF(H%d<0,"⚠️ السلف أكثر من المستحق",""))' % (r, r, r, r)
    S.merge_cells('I%d:L%d' % (r, r))
    for cc in 'ABCDEFGHI': cellfmt(S['%s%d' % (cc, r)], AUTO)
    S['C%d' % r].number_format = '0.00'
    for cc in 'DEFGH': S['%s%d' % (cc, r)].number_format = '#,##0.000;[Red]-#,##0.000;-'
    S['H%d' % r].font = F(13, True); S['I%d' % r].font = F(10, True, 'C00000')
S['B12'] = 'المدير كل 10 أيام بورقة «المدير» · عمال المهام بورقة «المهام».'; S['B12'].font = F(9, False, '666666'); S.merge_cells('B12:L12'); S['B12'].alignment = C('right')

# الشهر كامل
S['B14'] = '=SEL_M&"  "&SEL_Y&"  —  كل الأسابيع"'; S['B14'].font = F(13, True, TEAL); S.merge_cells('B14:L14'); S['B14'].alignment = C('right')
h2 = ['#', 'من', 'إلى', 'عمر: ساعات', 'عمر: مستحق', 'عمر: سلف', 'عمر: صافي', 'عبدالعزيز: ساعات', 'عبدالعزيز: مستحق', 'عبدالعزيز: مقبوض', 'عبدالعزيز: سلف', 'عبدالعزيز: باقي']
for i, h in enumerate(h2, start=1): cellfmt(S.cell(15, i, h), TEAL, True, sz=10, fc='FFFFFF')
S.row_dimensions[15].height = 34
for k in range(NWMAX):
    r = 16 + k; sr = B0 + 8 + BLOCK * k; on = '%d<=NW' % (k + 1)
    S['A%d' % r] = '=IF(%s,%d,"")' % (on, k + 1)
    S['B%d' % r] = '=IF(%s,W1+%d,"")' % (on, 7 * k); S['C%d' % r] = '=IF(%s,W1+%d,"")' % (on, 7 * k + 6)
    S['D%d' % r] = "=IF(%s,N('ساعات عمر'!F%d),\"\")" % (on, sr)
    S['E%d' % r] = "=IF(%s,N('ساعات عمر'!I%d),\"\")" % (on, sr)
    S['F%d' % r] = '=IF(%s,SUMIFS(LOG_A,LOG_E,"عمر المصري",LOG_T,"سلفة",LOG_D,">="&B%d,LOG_D,"<="&C%d),"")' % (on, r, r)
    S['G%d' % r] = '=IF(%s,E%d-F%d-IF(E%d=0,0,OMR_INST),"")' % (on, r, r, r)
    S['H%d' % r] = "=IF(%s,N('ساعات عبدالعزيز'!F%d),\"\")" % (on, sr)
    S['I%d' % r] = "=IF(%s,N('ساعات عبدالعزيز'!G%d),\"\")" % (on, sr)
    S['J%d' % r] = "=IF(%s,N('ساعات عبدالعزيز'!I%d),\"\")" % (on, sr)
    S['K%d' % r] = '=IF(%s,SUMIFS(LOG_A,LOG_E,"عبدالعزيز",LOG_T,"سلفة",LOG_D,">="&B%d,LOG_D,"<="&C%d),"")' % (on, r, r)
    S['L%d' % r] = '=IF(%s,I%d-J%d-K%d-IF(I%d=0,0,AZZ_INST),"")' % (on, r, r, r, r)
    for cc in 'ABCDEFGHIJKL': cellfmt(S['%s%d' % (cc, r)], AUTO)
    for cc in 'BC': S['%s%d' % (cc, r)].number_format = 'DD/MM'
    for cc in 'DH': S['%s%d' % (cc, r)].number_format = '0.00'
    for cc in 'EFGIJKL': S['%s%d' % (cc, r)].number_format = '#,##0.000;[Red]-#,##0.000;-'
S.conditional_formatting.add('A16:L20', FormulaRule(formula=['AND($A16<>"",$A16=WK)'], fill=fill('FFF2B3')))
r = 21; S['B%d' % r] = 'مجموع الشهر'
for cc in 'DEFGHIJKL': S['%s%d' % (cc, r)] = '=SUM(%s16:%s20)' % (cc, cc)
for cc in 'ABCDEFGHIJKL': cellfmt(S['%s%d' % (cc, r)], SUBT, True)
for cc in 'DH': S['%s%d' % (cc, r)].number_format = '0.00'
for cc in 'EFGIJKL': S['%s%d' % (cc, r)].number_format = '#,##0.000;[Red]-#,##0.000;-'

# رصيد آخر الشهر لكل موظف
S['B23'] = 'رصيد كل موظف آخر الشهر (بيتدوّر للشهر الجاي كـ«رصيد مدوّر»)'; S['B23'].font = F(13, True, RUST); S.merge_cells('B23:L23'); S['B23'].alignment = C('right')
h3 = ['#', 'الموظف', 'رصيد مدوّر', 'مستحق الشهر', 'تسديدات', 'سلف', 'دين + بضاعة', 'مدفوع له', 'الرصيد', 'معناه']
for i, h in enumerate(h3, start=1): cellfmt(S.cell(24, i, h), RUST, True, sz=10, fc='FFFFFF')
S.merge_cells('J24:L24')
earned = {'عمر المصري': '=E21', 'عبدالعزيز': '=I21', 'المدير': "='المدير'!D10"}
paid_extra = {'عمر المصري': '', 'عبدالعزيز': '+J21', 'المدير': ''}
for i, nm in enumerate(['عمر المصري', 'عبدالعزيز', 'المدير']):
    r = 25 + i
    S['A%d' % r] = i + 1; S['B%d' % r] = nm
    S['C%d' % r] = '=SUMIFS(LOG_A,LOG_E,B{r},LOG_T,"رصيد مدوّر له")-SUMIFS(LOG_A,LOG_E,B{r},LOG_T,"رصيد مدوّر عليه")'.replace('{r}', str(r))
    S['D%d' % r] = earned[nm]
    S['E%d' % r] = '=SUMIFS(LOG_A,LOG_E,B%d,LOG_T,"تسديد")' % r
    S['F%d' % r] = '=SUMIFS(LOG_A,LOG_E,B%d,LOG_T,"سلفة")' % r
    S['G%d' % r] = '=SUMIFS(LOG_A,LOG_E,B{r},LOG_T,"دين")+SUMIFS(LOG_A,LOG_E,B{r},LOG_T,"بضاعة")'.replace('{r}', str(r))
    S['H%d' % r] = '=SUMIFS(LOG_A,LOG_E,B%d,LOG_T,"دفع")%s' % (r, paid_extra[nm])
    S['I%d' % r] = '=C%d+D%d+E%d-F%d-G%d-H%d' % ((r,) * 6)
    S['J%d' % r] = '=IF(I%d>0.0005,"المحل مدين إله",IF(I%d<-0.0005,"هو مدين للمحل","متساوي"))' % (r, r)
    S.merge_cells('J%d:L%d' % (r, r))
    for cc in 'ABCDEFGHIJ': cellfmt(S['%s%d' % (cc, r)], AUTO)
    for cc in 'CDEFGHI': S['%s%d' % (cc, r)].number_format = '#,##0.000;[Red]-#,##0.000;-'
    S['I%d' % r].font = F(12, True)
S['B29'] = 'الرصيد = المدوّر + المستحق + التسديدات − السلف − الديون والبضاعة − المدفوع. مدفوع عبدالعزيز = مقبوضه اليومي + أي «دفع» بالسجل.'
S['B29'].font = F(9, False, '555555'); S.merge_cells('B29:L29'); S['B29'].alignment = C('right')

# ═══════════════════ التعليمات ═══════════════════
I = wb.create_sheet('التعليمات')
I.sheet_view.rightToLeft = True; I.sheet_view.showGridLines = False
I.column_dimensions['A'].width = 5; I.column_dimensions['B'].width = 115
I['B1'] = 'كيف يشتغل الملف'; I['B1'].font = F(15, True, NAVY)
steps = [
    ('أول كل شهر', None),
    ('1', 'خذ نسخة جديدة من الملف وسمّيها باسم الشهر. بورقة «التسوية» اختر الشهر والسنة — كل الأسابيع وتواريخها بتطلع لحالها بأوراق الساعات.'),
    ('2', 'الأسبوع من السبت للجمعة، وبيتبع الشهر اللي فيه **جمعته** (يوم القبض). مثلاً أسبوع 26/09 ← 02/10 تبع تشرين الأول. هيك ولا يوم بيتكرر بين ملفين.'),
    ('3', 'رصيد آخر الشهر الماضي لكل موظف (من أسفل ورقة التسوية بملف الشهر الماضي): سطر بورقة «السلف والديون» نوعه «رصيد مدوّر له» أو «رصيد مدوّر عليه».'),
    ('كل يوم', None),
    ('4', 'بورقة الساعات: اختر الدخول والخروج من القائمة (كل ربع ساعة، AM/PM). وساعات تحميل الجبسمبورد بعمودها — هي جزء من ساعات اليوم مش زيادة.'),
    ('5', 'عبدالعزيز: آخر اليوم «قبض؟ نعم» + المقبوض. المستحق اليومي مجبور لربع دينار.'),
    ('6', 'أي سلفة أو دين أو بضاعة أو دفعة: سطر بورقة «السلف والديون».'),
    ('آخر الأسبوع', None),
    ('7', 'بورقة «التسوية» اختر الأسبوع من القائمة — بيطلعلك لكل موظف الساعات والمستحق (مجبور لربع دينار) والسلف والقسط والصافي للدفع. وتحتها جدول الشهر كامل والأسبوع المختار ملوّن.'),
    ('8', 'لما تدفع: سجّل «دفع» بورقة السلف والديون.'),
    ('القواعد', None),
    ('9', 'كل الأسعار والأوقات وطريقة الجبر (لأقرب ربع / للأعلى) بورقة «القواعد». لا تكتب أسعار بأي مكان ثاني.'),
]
r = 3
for a, b in steps:
    if b is None: I['B%d' % r] = a; I['B%d' % r].font = F(12, True, TEAL); r += 1; continue
    I['A%d' % r] = a; I['A%d' % r].font = F(11, True); I['A%d' % r].alignment = C()
    I['B%d' % r] = b.replace('**', ''); I['B%d' % r].font = F(11); I['B%d' % r].alignment = C('right'); I.row_dimensions[r].height = 34; r += 1

# الترتيب: التسوية أولاً، والقوائم مخفية
del wb['Sheet']
wb.move_sheet('التسوية', offset=-wb.sheetnames.index('التسوية'))
wb.active = 0
wb.save(OUT); print('saved', OUT)
