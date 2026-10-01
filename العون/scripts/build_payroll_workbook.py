# -*- coding: utf-8 -*-
"""نظام الموظفين السنوي الأسبوعي — محلات العون لمواد البناء (النسخة 12 · بأمره 01/10/2026)
ملف واحد للسنة: 53 كتلة أسبوعية (السبت → الجمعة). تختار السنة بورقة «التسوية»، والأسابيع وتواريخها بتطلع لحالها.

القاعدة الذهبية: **الأسبوع كامل بينسب للشهر والسنة اللي فيهم جمعته (يوم القبض)** — ما بينقسم أبداً.
  • الأسبوع 1 = اللي جمعته أول جمعة بيناير · آخر أسبوع = اللي فيه آخر جمعة بديسمبر (52 أو 53).
  • الملخص الشهري = كل الأسابيع اللي جمعتها بهالشهر (+ السلف والمهام والبضاعة بشهر جمعة أسبوعها).

القواعد (26/09 + 01/10 من صاحب المحل):
  • عمر المصري: السبت–الخميس @1.250/ساعة · بعد 6:00 PM @1.500 · الجمعة كل ساعاته @1.250 — الأسعار بجدول «ساري من» (تغيير السعر ما بيغيّر الماضي).
  • عبدالعزيز: 1.000/ساعة أي يوم · بيقبض نهاية كل يوم (دفعة مقدّمة) وبيتسوّى الجمعة.
  • تحميل الجبسمبورد @2.500 — مرة وحدة لكل أسبوع. المستحق بيتجبر لربع دينار.
  • الإكرامية والمدفوع يوم الجمعة: بسطر الأسبوع بورقة الساعات (مكان واحد لكل شغل الأسبوع).
  • السقوف: السلفة ≤ 60% من اللي اشتغله لحد يومها (تنبيه) · القسط ≤ 25% من مستحق الأسبوع · الصافي ما ينزل عن 65% من المستحق.
  • «الباقي له» تراكمي: بينتقل لحاله للأسبوع الجاي — نقل يدوي مرة بالسنة بس (ورقة «بداية السنة»).
  • بلا ماكرو وبلا دوال _xlfn (MAXIFS/IFS/XLOOKUP/LET/ISOWEEKNUM ممنوعة — #NAME? عند صاحب المحل).
"""
import openpyxl, datetime as dt
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName

OUT = '/home/user/Fam-Bam/العون/نظام_الموظفين.xlsx'
VERSION = 12
NAVY, TEAL, GOLD, PLUM, RUST = '1F4E6B', '0F6E6E', '8A6A00', '5B3A6B', '8B3A2F'
EDIT, AUTO, WEEKBAR, SUBT, OUTM = 'FFF6CC', 'EDEDED', 'DCE6EE', 'E3EFE3', 'D9D9D9'
LOCKED, CURW = 'E8F3E8', 'FFF2B3'          # أسبوع مقبوض (مقفول) · الأسبوع الحالي
LINE = Side(style='thin', color='9AA5AD'); THICK = Side(style='medium', color=NAVY)
BOX = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
DUR = '[h]" س "mm" د"'
DIN = '[<=3112]00\\/00;dd/mm/yyyy'            # خانة تاريخ بتقبل: 0510 (يوم وشهر) · 05/10/2026
DATE = 'dd/mm/yyyy'; MONEY = '#,##0.000;[Red]-#,##0.000;-'; HRS = '0.00'; PCT = '0%'
NWMAX = 53                     # أقصى عدد أسابيع بالسنة
BLOCK = 11                     # عنوان + 7 أيام + مجموع + جبسمبورد/المستحق + إكرامية/المدفوع
B0 = 6
ADV_ROWS, LOAN_ROWS, TASK_ROWS, RATE_ROWS = 400, 200, 600, 10
EMPS = ['عمر المصري', 'عبدالعزيز']
YEARS = list(range(2026, 2036))
MONTHS = ['كانون الثاني (1)', 'شباط (2)', 'آذار (3)', 'نيسان (4)', 'أيار (5)', 'حزيران (6)',
          'تموز (7)', 'آب (8)', 'أيلول (9)', 'تشرين الأول (10)', 'تشرين الثاني (11)', 'كانون الأول (12)']

def F(sz=11, b=False, c='000000'): return Font(name='Arial', size=sz, bold=b, color=c)
def C(h='center'): return Alignment(horizontal=h, vertical='center', wrap_text=True)
def fill(c): return PatternFill('solid', fgColor=c)
def col(i): return openpyxl.utils.get_column_letter(i)

wb = openpyxl.Workbook()
INPUTS = {}
def name(nm, ref): wb.defined_names[nm] = DefinedName(nm, attr_text=ref)
def dv(ws, formula, rng, prompt=None, title=None):
    d = DataValidation(type='list', formula1=formula, allow_blank=True, showErrorMessage=False)
    if prompt: d.prompt = prompt; d.showInputMessage = True; d.promptTitle = title or ''
    ws.add_data_validation(d); d.add(rng)
def dv_date(ws, rng):
    dv(ws, "='القوائم'!$N$1:$N$14", rng, 'اختار من آخر 14 يوم، أو اكتب يوم وشهر (0510 = 05/10)، أو التاريخ كامل', 'التاريخ')
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
def TXT(x): return 'TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00")' % (x, x)
YEAR_SUB = '="سنة "&YR&"   (السنة بتنختار من ورقة التسوية)"'
def DVAL(x):   # يوم وشهر (0510) ← تاريخ بالسنة (ولو طلع بعد آخر جمعة ← السنة الماضية) · أو تاريخ كامل · رقم لحاله = غلط
    d = 'DATE(YR,MOD({x},100),INT({x}/100))'.format(x=x)
    return ('IF({x}="","",IF(ISNUMBER({x}),IF({x}<=31,"",IF({x}<=3112,IF({d}>SAT1+7*NW-1,DATE(YR-1,MOD({x},100),INT({x}/100)),{d}),{x})),'
            'IFERROR(DATEVALUE({x}),"")))').format(x=x, d=d)
def WKOF(d): return 'IF({d}="","",INT(({d}-SAT1)/7)+1)'.format(d=d)     # رقم أسبوع التاريخ (ممكن يطلع برّا 1..NW)
def MONOF(w): return 'IF({w}="","",IF(OR({w}<1,{w}>NW),"",INDEX(WMON,{w})))'.format(w=w)   # شهر جمعة الأسبوع
def RND(x): return 'IF(ROUND_MODE="للأعلى",CEILING(ROUND((%s),3),0.25),ROUND((%s)*4,0)/4)' % (x, x)

# ═══════════════════ القوائم (مخفية): الأوقات · السنة · جدول الأسابيع ═══════════════════
H = wb.create_sheet('القوائم')
def label(m):
    h, mi = divmod(m, 60)
    return '%d:%02d %s' % ((h % 12) or 12, mi, 'AM' if h < 12 else 'PM')
NT = 288                                                          # كل 5 دقائق · 24 ساعة من 6:00 AM
for i, m in enumerate([(6 * 60 + 5 * k) % 1440 for k in range(NT)], start=1):
    H.cell(i, 1, label(m)); H.cell(i, 2, dt.time(m // 60, m % 60)).number_format = 'h:mm AM/PM'
for i in range(12): H.cell(i + 1, 4, i + 1)
for i, y in enumerate(YEARS, start=1): H.cell(i, 5, y)
H['G4'] = "='التسوية'!$E$3"                                        # السنة
H['G1'] = '=DATE(G4,1,1)+MOD(6-WEEKDAY(DATE(G4,1,1),1),7)-6'       # سبت الأسبوع 1 (جمعته أول جمعة بيناير)
H['G2'] = '=INT((DATE(G4,12,31)-(G1+6))/7)+1'                       # عدد الأسابيع 52/53
H['G3'] = '=IF(AND(TODAY()>=G1,TODAY()<=G1+7*G2-1),INT((TODAY()-G1)/7)+1,1)'   # الأسبوع الحالي
H['G1'].number_format = DATE
name('YR', "'القوائم'!$G$4"); name('SAT1', "'القوائم'!$G$1"); name('NW', "'القوائم'!$G$2"); name('CURWK', "'القوائم'!$G$3")
WT0 = 11
for k in range(NWMAX):     # جدول الأسابيع: السبت · الجمعة · العنوان · شهر الجمعة · الرقم
    r = WT0 + k
    H['H%d' % r] = '=IF(%d<=NW,SAT1+%d,"")' % (k + 1, 7 * k)
    H['I%d' % r] = '=IF(%d<=NW,SAT1+%d,"")' % (k + 1, 7 * k + 6)
    H['J%d' % r] = '=IF(%d<=NW,"الأسبوع %d · "&%s&" ← "&%s,"")' % (k + 1, k + 1, TXT('H%d' % r), TXT('I%d' % r))
    H['K%d' % r] = '=IF(%d<=NW,MONTH(I%d),"")' % (k + 1, r)
    H['L%d' % r] = k + 1
    for c in 'HI': H['%s%d' % (c, r)].number_format = DATE
WTE = WT0 + NWMAX - 1
name('WSAT', "'القوائم'!$H$%d:$H$%d" % (WT0, WTE)); name('WFRI', "'القوائم'!$I$%d:$I$%d" % (WT0, WTE))
name('WTITLE', "'القوائم'!$J$%d:$J$%d" % (WT0, WTE)); name('WMON', "'القوائم'!$K$%d:$K$%d" % (WT0, WTE))
for i in range(14):                                               # آخر 14 يوم (لقائمة التاريخ)
    H['N%d' % (i + 1)] = '=TODAY()-%d' % (13 - i); H['N%d' % (i + 1)].number_format = DATE
H['P1'] = 'ALAWN_PAYROLL'; H['P2'] = VERSION; H['P3'] = 'سنوي'   # علامة النموذج — القارئ بيتأكد منها
H.sheet_state = 'hidden'
TIME_LIST = "='القوائم'!$A$1:$A$%d" % NT
def TV(ref):
    return ("IF(ISNUMBER({r}),MOD({r},1),IFERROR(INDEX('القوائم'!$B$1:$B${n},MATCH(TRIM({r}),'القوائم'!$A$1:$A${n},0)),"
            "IFERROR(MOD(TIMEVALUE({r}),1),\"\")))").format(r=ref, n=NT)

# ═══════════════════ القواعد: ثوابت + جدول أسعار بتاريخ «ساري من» ═══════════════════
K = wb.create_sheet('القواعد')
title(K, 'القواعد والأسعار — كل الحسابات بتقرأ من هون', GOLD, 'E')
for k, w in zip('ABCDE', [4, 40, 16, 16, 44]): K.column_dimensions[k].width = w
header(K, 3, ['#', 'القاعدة', 'القيمة', '', 'الشرح'], GOLD, 11, 26); K.merge_cells('C3:D3')
rules = [
    ('OMR_OT_AT', 'عمر المصري — بداية الإضافي', dt.time(18, 0), 'h:mm AM/PM', 'السبت–الخميس بعد هالوقت إضافي. الجمعة ما إلها إضافي.'),
    ('OMR_FRI_IN', 'عمر المصري — دوام الجمعة من', dt.time(14, 0), 'h:mm AM/PM', 'للتنبيه بس — الحساب حسب الساعات الفعلية.'),
    ('OMR_FRI_OUT', 'عمر المصري — دوام الجمعة إلى', dt.time(23, 0), 'h:mm AM/PM', ''),
    ('ROUND_MODE', 'جبر المستحق', 'لأقرب ربع', '@', 'لأقرب ربع: 18.300←18.250 و18.400←18.500 · للأعلى: أي كسر بيطلع للربع اللي فوقه.'),
    ('ADV_CAP', 'سقف السلفة (من اللي اشتغله لحد يومها)', 0.60, PCT, 'سلفة أكبر من هيك بتتلوّن برتقالي بورقة السلف — تنبيه بس، بتنخصم كاملة.'),
    ('INST_CAP', 'سقف القسط الأسبوعي (من المستحق + الإكرامية)', 0.25, PCT, 'القسط اللي بينخصم بالأسبوع ما بيزيد عن هالنسبة — الباقي بيتأجل.'),
    ('NET_FLOOR', 'الحد الأدنى للصافي (من المستحق)', 0.65, PCT, 'القسط بيقل أو بيتأجل عشان الصافي ما ينزل تحت هالنسبة. السلف بتنخصم كاملة.'),
]
for i, (nm, rule, val, fmt, desc) in enumerate(rules, start=1):
    r = 3 + i
    K.cell(r, 1, i); K.cell(r, 2, rule); K.cell(r, 3, val); K.cell(r, 5, desc); K.merge_cells('C%d:D%d' % (r, r))
    for cc in (1, 2, 5): cellfmt(K.cell(r, cc), h='right' if cc > 1 else 'center')
    edit(K, 'C%d' % r, fmt, True); K.cell(r, 5).font = F(10, False, '555555'); K.row_dimensions[r].height = 30
    name(nm, "'القواعد'!$C$%d" % r)
    if nm == 'ROUND_MODE': dv(K, '"لأقرب ربع,للأعلى"', 'C%d' % r)
RT0 = 3 + len(rules) + 3
K['A%d' % (RT0 - 1)] = 'أسعار الساعة — حسب التاريخ: كل سطر «ساري من» تاريخه لحد السطر اللي بعده. لما يتغيّر السعر: سطر جديد بتاريخه (الأسابيع القديمة ما بتتأثر). السطور لازم تكون مرتّبة بالتاريخ.'
K['A%d' % (RT0 - 1)].font = F(11, True, GOLD); K.merge_cells('A%d:E%d' % (RT0 - 1, RT0 - 1)); K['A%d' % (RT0 - 1)].alignment = C('right'); K.row_dimensions[RT0 - 1].height = 34
header(K, RT0, ['ساري من', 'عمر — الساعة العادية', 'عمر — ساعة الإضافي', 'عبدالعزيز — الساعة', 'تحميل الجبسمبورد — الساعة'], GOLD, 10, 34)
for i in range(RATE_ROWS):
    r = RT0 + 1 + i
    edit(K, 'A%d' % r, DATE, True)
    for c in 'BCDE': edit(K, '%s%d' % (c, r), '0.000')
    K['G%d' % r] = '=IF(A%d="",99999999,A%d)' % (r, r)      # للمطابقة التقريبية (الفاضي آخر الجدول)
K['A%d' % (RT0 + 1)] = dt.date(2026, 1, 1); K['B%d' % (RT0 + 1)] = 1.25; K['C%d' % (RT0 + 1)] = 1.5; K['D%d' % (RT0 + 1)] = 1.0; K['E%d' % (RT0 + 1)] = 2.5
K.column_dimensions['G'].hidden = True
RTE = RT0 + RATE_ROWS
for nm, c in (('RT_K', 'G'), ('RT_OMR', 'B'), ('RT_OT', 'C'), ('RT_AZZ', 'D'), ('RT_GYP', 'E')):
    name(nm, "'القواعد'!$%s$%d:$%s$%d" % (c, RT0 + 1, c, RTE))
def RATE(colname, fri): return 'IFERROR(INDEX(%s,MATCH(%s,RT_K,1)),INDEX(%s,1))' % (colname, fri, colname)
note(K, 'A%d' % (RTE + 2), 'مثال: صار سعر عمر 1.400 من 01/03/2027 ← سطر: 01/03/2027 · 1.400 · 1.500 · 1.000 · 2.500. الأسابيع اللي جمعتها قبل 01/03 بتضل على السعر القديم.', 'A%d:E%d' % (RTE + 2, RTE + 2))

# ═══════════════════ بداية السنة (الافتتاح) ═══════════════════
B = wb.create_sheet('بداية السنة'); BS = "'بداية السنة'"
title(B, 'بداية ونهاية السنة — النقل الوحيد اليدوي: مرة بالسنة', RUST, 'G', YEAR_SUB)
for k, w in zip('ABCDEFG', [16, 14, 18, 16, 16, 14, 30]): B.column_dimensions[k].width = w
B['A4'] = 'بداية السنة — من «نهاية السنة» بملف السنة الماضية (أو الوضع يوم بدأت تستعمل الملف)'; B['A4'].font = F(13, True, RUST); B.merge_cells('A4:G4'); B['A4'].alignment = C('right')
header(B, 5, ['الموظف', 'أسبوع البداية', 'مستحق له من قبل', 'دين متبقي عليه', 'القسط الأسبوعي'], RUST, 10, 34)
OPEN = {}
for i, nm in enumerate(EMPS):
    r = 6 + i; B['A%d' % r] = nm; auto(B, 'A%d' % r, None, True, color='FFFFFF')
    B['B%d' % r] = 40; edit(B, 'B%d' % r, '0', True)
    for c in 'CDE': B['%s%d' % (c, r)] = 0; edit(B, '%s%d' % (c, r), MONEY, True)
    OPEN[nm] = r
note(B, 'A8', '«أسبوع البداية»: أول أسبوع بتستعمل فيه هالملف (أسبوع 1 لسنة كاملة). الأسابيع اللي قبله رمادية وما بتنحسب. «مستحق له» بينضاف لصافي أسبوع البداية · «دين متبقي» والقسط بيكمّلوا ينخصموا كل أسبوع.', 'A8:G8', 9, h=30)
def OP(c, nm): return '%s!$%s$%d' % (BS, c, OPEN[nm])
name('START_O', OP('B', EMPS[0])); name('START_Z', OP('B', EMPS[1]))

# ═══════════════════ أوراق الساعات ═══════════════════
FINAL_C, FINAL_BG = 'C55A11', 'ED7D31'
DUE_BG, DUE_FC, DUE_HEAD = 'E2F0D9', '1B5E20', '2E7D32'
def hours_sheet(ttl, banner, color, heads, widths, text, last):
    ws = wb.create_sheet(ttl)
    title(ws, banner, color, last, YEAR_SUB)
    note(ws, 'A3', text, 'A3:%s3' % last, h=40)
    for k, w in enumerate(widths, start=1): ws.column_dimensions[col(k)].width = w
    header(ws, 5, heads, color, 11, 32); ws.freeze_panes = 'A6'
    for c in 'KLMNOPQ': ws.column_dimensions[c].hidden = True
    return ws

def week_block_head(ws, k, color, last, start_name):
    """عنوان الأسبوع + الخلايا المخفية بسطره: M = فعّال · L = الجمعة · N/O/P = أسعار الأسبوع"""
    r = B0 + BLOCK * k; w = k + 1
    sat = 'SAT1+%d' % (7 * k); fri = 'SAT1+%d' % (7 * k + 6)
    ws.merge_cells('A%d:%s%d' % (r, last, r))
    ws['A%d' % r] = ('=IF(%d<=NW,"الأسبوع %d · السبت "&%s&" ← الجمعة "&%s&" (يوم القبض)"&IF(%d<%s,"   — قبل أسبوع البداية (ما بينحسب)",""),'
                     '"ما في أسبوع %d هالسنة")') % (w, w, TXT(sat), TXT(fri), w, start_name, w)
    cellfmt(ws['A%d' % r], WEEKBAR, True, sz=13, fc=color, h='right'); ws.row_dimensions[r].height = 26
    ws['M%d' % r] = '=IF(AND(%d<=NW,%d>=%s),1,0)' % (w, w, start_name)
    ws['L%d' % r] = '=IF(%d<=NW,%s,"")' % (w, fri); ws['L%d' % r].number_format = DATE
    return r

def day_rows(ws, hr, k, paid_ref):
    for d in range(7):
        r = hr + 1 + d
        ws['A%d' % r] = '=IF(%d<=NW,SAT1+%d,"")' % (k + 1, 7 * k + d)
        ws['B%d' % r] = '=IF(A{r}="","",%s)'.replace('{r}', str(r)) % DAYNAME.format(d='A%d' % r)
        ws['M%d' % r] = '=M%d' % hr
        ws['K%d' % r] = '=IF(OR(M{r}=0,C{r}=""),"",%s)'.replace('{r}', str(r)) % TV('C%d' % r)
        ws['L%d' % r] = '=IF(OR(M{r}=0,D{r}=""),"",%s)'.replace('{r}', str(r)) % TV('D%d' % r)
        ws['Q%d' % r] = '=IF(N(%s)>0,1,0)' % paid_ref                          # 1 = الأسبوع مقبوض (مقفول)
        auto(ws, 'A%d' % r, DATE); auto(ws, 'B%d' % r)
        for c in 'KL': ws['%s%d' % (c, r)].number_format = 'h:mm AM/PM'
        if d == 6:
            for c in 'AB': ws['%s%d' % (c, r)].font = F(11, True, NAVY)

# ---- عمر ----
O = hours_sheet('ساعات عمر', 'ساعات عمر المصري', NAVY,
    ['التاريخ', 'اليوم', 'دخول', 'خروج', 'عدد الساعات', 'منها عادي', 'منها إضافي', 'المستحق (دينار)', 'ملاحظات'],
    [13, 10, 12, 12, 10, 11, 11, 14, 30],
    'اختر الدخول والخروج (كل 5 دقائق، AM/PM) — عدد الساعات والعادي والإضافي والمستحق بيطلعوا لحالهم. '
    'تحت كل أسبوع: ساعات الجبسمبورد · المستحق النهائي (مجبور لربع دينار) · الإكرامية · المدفوع يوم الجمعة. '
    'أول ما تكتب «المدفوع» الأسبوع بيتلوّن أخضر = مقبوض ومقفول: ما تعدّله، أي تصحيح بسطر الأسبوع الجاي.', 'I')
OMR = []   # (سطر المجموع, سطر المستحق النهائي, سطر الإكرامية/المدفوع, سطر العنوان)
for k in range(NWMAX):
    hr = week_block_head(O, k, NAVY, 'I', 'START_O')
    s1, s2, s3 = hr + 8, hr + 9, hr + 10; lo, hi = hr + 1, hr + 7; on = 'M%d=1' % hr
    O['N%d' % hr] = '=' + RATE('RT_OMR', 'L%d' % hr); O['O%d' % hr] = '=' + RATE('RT_OT', 'L%d' % hr); O['P%d' % hr] = '=' + RATE('RT_GYP', 'L%d' % hr)
    day_rows(O, hr, k, 'G%d' % s3)
    for d in range(7):
        r = hr + 1 + d
        O['E%d' % r] = '=IF(OR(M{r}=0,K{r}="",L{r}=""),"",MOD(L{r}-K{r},1))'.replace('{r}', str(r))
        O['F%d' % r] = ('=IF(OR(E{r}="",M{r}=0),"",IF(WEEKDAY(A{r},1)=6,E{r},'
                        'MIN(E{r},MAX(0,MIN(IF(L{r}<K{r},L{r}+1,L{r}),OMR_OT_AT)-K{r}))))').replace('{r}', str(r))
        O['G%d' % r] = '=IF(OR(E{r}="",M{r}=0),"",MAX(0,E{r}-F{r}))'.replace('{r}', str(r))
        O['H%d' % r] = '=IF(OR(E{r}="",M{r}=0),"",ROUND((F{r}*$N${h}+G{r}*$O${h})*24,3))'.replace('{r}', str(r)).replace('{h}', str(hr))
        for c in 'CD': edit(O, '%s%d' % (c, r))
        edit(O, 'I%d' % r)
        for c, fm in zip('EFG', [DUR, DUR, DUR]): auto(O, '%s%d' % (c, r), fm)
        auto(O, 'H%d' % r, MONEY, True, 12, color=DUE_BG, fc=DUE_FC)
    O.merge_cells('A%d:D%d' % (s1, s1)); O['A%d' % s1] = '=IF(%s,"مجموع أيام الأسبوع %d","")' % (on, k + 1)
    for c in 'EFGH': O['%s%d' % (c, s1)] = '=IF(%s,SUM(%s%d:%s%d),"")' % (on, c, lo, c, hi)
    for c in 'ABCDEFGHI': auto(O, '%s%d' % (c, s1), None, True, color=SUBT)
    for c, fm in zip('EFGH', [DUR, DUR, DUR, MONEY]): O['%s%d' % (c, s1)].number_format = fm
    O['H%d' % s1].font = F(12, True, DUE_FC)
    O.merge_cells('A%d:D%d' % (s2, s2)); O['A%d' % s2] = '=IF(%s,"ساعات تحميل الجبسمبورد بهالأسبوع (رقم: 2 أو 1.5) ←","")' % on
    auto(O, 'A%d' % s2, None, True, color=SUBT); edit(O, 'E%d' % s2, HRS, True)
    O.merge_cells('F%d:G%d' % (s2, s2)); O['F%d' % s2] = '=IF(%s,"المستحق النهائي للأسبوع (دينار) ←","")' % on
    auto(O, 'F%d' % s2, None, True, 12, color=SUBT, fc=FINAL_C)
    O['H%d' % s2] = '=IF(%s,%s,"")' % (on, RND('N(H%d)+N(E%d)*($P$%d-$N$%d)' % (s1, s2, hr, hr)))
    auto(O, 'H%d' % s2, MONEY, True, 16, color=SUBT); O.row_dimensions[s2].height = 32
    O['H%d' % s2].border = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)
    O['I%d' % s2] = '=IF(%s,"قبل الجبر: "&TEXT(N(H%d)+N(E%d)*($P$%d-$N$%d),"0.000"),"")' % (on, s1, s2, hr, hr)
    auto(O, 'I%d' % s2, None, color=SUBT); O['I%d' % s2].font = F(9, False, '555555')
    O.merge_cells('A%d:C%d' % (s3, s3)); O['A%d' % s3] = '=IF(%s,"إكرامية الأسبوع (دينار) ←","")' % on
    auto(O, 'A%d' % s3, None, True, color=SUBT); edit(O, 'D%d' % s3, MONEY, True)
    O.merge_cells('E%d:F%d' % (s3, s3)); O['E%d' % s3] = '=IF(%s,"المدفوع يوم الجمعة (دينار) ←","")' % on
    auto(O, 'E%d' % s3, None, True, color=SUBT); edit(O, 'G%d' % s3, MONEY, True)
    O.merge_cells('H%d:I%d' % (s3, s3))
    O['H%d' % s3] = '=IF(%s,IF(N(G%d)>0,"✔ انقبض — الأسبوع مقفول","⏳ لسا ما انقبض"),"")' % (on, s3)
    auto(O, 'H%d' % s3, None, True, 11, color=SUBT, fc='1B5E20'); O.row_dimensions[s3].height = 26
    OMR.append((s1, s2, s3, hr))
OEND = B0 + BLOCK * NWMAX - 1
dv(O, TIME_LIST, 'C%d:D%d' % (B0, OEND))
O.conditional_formatting.add('C%d:D%d' % (B0, OEND), FormulaRule(
    formula=['AND($M%d=1,$K%d<>"",$L%d<>"",WEEKDAY($A%d,1)=6,OR($K%d<OMR_FRI_IN,$L%d>OMR_FRI_OUT))' % ((B0,) * 6)], fill=fill('FDE2C8')))

# ---- عبدالعزيز ----
Z = hours_sheet('ساعات عبدالعزيز', 'ساعات عبدالعزيز — بيقبض نهاية كل يوم وبيتسوّى الجمعة', TEAL,
    ['التاريخ', 'اليوم', 'دخول', 'خروج', 'عدد الساعات', 'مستحق اليوم (دينار)', 'قبض؟', 'ملاحظات'],
    [13, 10, 12, 12, 10, 14, 9, 34],
    'دينار للساعة أي يوم. مستحق كل يوم مجبور لربع دينار. لما يقبض آخر اليوم: «قبض؟ نعم» — المبلغ = مستحق اليوم. '
    'اللي قبضه يومياً دفعة مقدّمة بتنخصم من تسوية الجمعة. تحت كل أسبوع: الجبسمبورد · المستحق النهائي · الإكرامية · المدفوع يوم الجمعة (التسوية).', 'H')
AZZ = []
for k in range(NWMAX):
    hr = week_block_head(Z, k, TEAL, 'H', 'START_Z')
    s1, s2, s3 = hr + 8, hr + 9, hr + 10; lo, hi = hr + 1, hr + 7; on = 'M%d=1' % hr
    Z['N%d' % hr] = '=' + RATE('RT_AZZ', 'L%d' % hr); Z['P%d' % hr] = '=' + RATE('RT_GYP', 'L%d' % hr)
    day_rows(Z, hr, k, 'F%d' % s3)
    for d in range(7):
        r = hr + 1 + d
        Z['E%d' % r] = '=IF(OR(M{r}=0,K{r}="",L{r}=""),"",MOD(L{r}-K{r},1))'.replace('{r}', str(r))
        Z['F%d' % r] = '=IF(OR(E%d="",M%d=0),"",%s)' % (r, r, RND('E%d*24*$N$%d' % (r, hr)))
        for c in 'CDGH': edit(Z, '%s%d' % (c, r))
        auto(Z, 'E%d' % r, DUR); auto(Z, 'F%d' % r, MONEY, True, 12, color=DUE_BG, fc=DUE_FC)
    Z.merge_cells('A%d:D%d' % (s1, s1)); Z['A%d' % s1] = '=IF(%s,"مجموع أيام الأسبوع %d","")' % (on, k + 1)
    for c in 'EF': Z['%s%d' % (c, s1)] = '=IF(%s,SUM(%s%d:%s%d),"")' % (on, c, lo, c, hi)
    Z['G%d' % s1] = '=IF(%s,SUMIFS(F%d:F%d,G%d:G%d,"نعم"),"")' % (on, lo, hi, lo, hi)      # اللي قبضه يومياً
    for c in 'ABCDEFGH': auto(Z, '%s%d' % (c, s1), None, True, color=SUBT)
    Z['E%d' % s1].number_format = DUR; Z['F%d' % s1].number_format = MONEY; Z['F%d' % s1].font = F(12, True, DUE_FC)
    Z['G%d' % s1].number_format = '"قبض "#,##0.000'
    Z.merge_cells('A%d:D%d' % (s2, s2)); Z['A%d' % s2] = '=IF(%s,"ساعات تحميل الجبسمبورد بهالأسبوع (رقم: 2 أو 1.5) ←","")' % on
    auto(Z, 'A%d' % s2, None, True, color=SUBT); edit(Z, 'E%d' % s2, HRS, True)
    Z['F%d' % s2] = '=IF(%s,N(F%d)+%s,"")' % (on, s1, RND('N(E%d)*($P$%d-$N$%d)' % (s2, hr, hr)))
    auto(Z, 'F%d' % s2, MONEY, True, 16, color=SUBT); Z.row_dimensions[s2].height = 32
    Z['F%d' % s2].border = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)
    Z.merge_cells('G%d:H%d' % (s2, s2)); Z['G%d' % s2] = '=IF(%s,"← المستحق النهائي للأسبوع (دينار) مع فرق الجبسمبورد","")' % on
    auto(Z, 'G%d' % s2, None, True, 12, color=SUBT, fc=FINAL_C)
    Z.merge_cells('A%d:B%d' % (s3, s3)); Z['A%d' % s3] = '=IF(%s,"إكرامية الأسبوع ←","")' % on
    auto(Z, 'A%d' % s3, None, True, color=SUBT); edit(Z, 'C%d' % s3, MONEY, True)
    Z.merge_cells('D%d:E%d' % (s3, s3)); Z['D%d' % s3] = '=IF(%s,"المدفوع يوم الجمعة (التسوية) ←","")' % on
    auto(Z, 'D%d' % s3, None, True, color=SUBT); edit(Z, 'F%d' % s3, MONEY, True)
    Z.merge_cells('G%d:H%d' % (s3, s3))
    Z['G%d' % s3] = '=IF(%s,IF(N(F%d)>0,"✔ انقبض — الأسبوع مقفول","⏳ لسا ما انقبض"),"")' % (on, s3)
    auto(Z, 'G%d' % s3, None, True, 11, color=SUBT, fc='1B5E20'); Z.row_dimensions[s3].height = 26
    AZZ.append((s1, s2, s3, hr))
ZEND = B0 + BLOCK * NWMAX - 1
dv(Z, TIME_LIST, 'C%d:D%d' % (B0, ZEND)); dv(Z, '"نعم,لا"', 'G%d:G%d' % (B0, ZEND))

for ws, end, last in ((O, OEND, 'I'), (Z, ZEND, 'H')):
    rng = 'A%d:%s%d' % (B0, last, end)
    ws.conditional_formatting.add(rng, FormulaRule(formula=['$M%d=0' % B0], fill=fill(OUTM)))                      # برّا الفترة: رمادي
    ws.conditional_formatting.add(rng, FormulaRule(formula=['AND($M%d=1,$Q%d=1)' % (B0, B0)], fill=fill(LOCKED)))   # مقبوض: أخضر فاتح (مقفول)
    ws.conditional_formatting.add('A%d:A%d' % (B0, end), FormulaRule(
        formula=['AND(ISNUMBER($A%d),$A%d>=SAT1+7*(CURWK-1),$A%d<=SAT1+7*(CURWK-1)+6)' % ((B0,) * 3)], fill=fill(CURW)))   # الأسبوع الحالي
for ws, rows, c in ((O, OMR, 'H'), (Z, AZZ, 'F')):
    for s1, s2, s3, hr in rows:
        ws.conditional_formatting.add('%s%d' % (c, s2), FormulaRule(formula=['%s%d<>""' % (c, s2)], fill=fill(FINAL_BG), font=Font(color='FFFFFF', bold=True)))

# ═══════════════════ السلف (سجل السنة) ═══════════════════
A = wb.create_sheet('السلف')
title(A, 'السلف الأسبوعية — بتنخصم كاملة من راتب نفس الأسبوع', RUST, 'E', YEAR_SUB)
for k, w in zip('ABCDE', [15, 18, 16, 36, 22]): A.column_dimensions[k].width = w
header(A, 3, ['الموظف', 'سلف هالسنة (دينار)', 'سلف الأسبوع المختار (دينار)', '', ''], RUST, 10, 32)
note(A, 'A7', '• كاش بياخذه وسط الأسبوع ← بينخصم كامل من صافي نفس الأسبوع (الأسبوع اللي فيه تاريخ السلفة). · لو السلفة أكثر من 60% من اللي اشتغله لحد يومها بتتلوّن برتقالي (تنبيه). · القروض والبضاعة بورقة «القروض والديون».', 'A7:E7', 10, '333333', 30)
header(A, 9, ['التاريخ', 'الموظف', 'المبلغ (دينار)', 'البيان', 'التنبيه'], RUST, 11, 30)
AR0, AE = 10, 10 + ADV_ROWS - 1
for r in range(AR0, AE + 1):
    edit(A, 'A%d' % r, DIN); edit(A, 'B%d' % r); edit(A, 'C%d' % r, MONEY); edit(A, 'D%d' % r); A['D%d' % r].alignment = C('right')
    A['G%d' % r] = '=' + DVAL('A%d' % r)                                   # التاريخ الفعلي
    A['H%d' % r] = '=' + WKOF('G%d' % r)                                   # رقم الأسبوع
    A['I%d' % r] = '=' + MONOF('H%d' % r)                                  # شهر جمعة الأسبوع
    # اللي اشتغله لحد يوم السلفة (من أول الأسبوع): مستحق الأيام ≤ التاريخ
    A['J%d' % r] = ('=IF(OR(G{r}="",H{r}="",B{r}=""),"",IF(B{r}="عمر المصري",'
                    "SUMIFS('ساعات عمر'!$H$%d:$H$%d,'ساعات عمر'!$A$%d:$A$%d,\">=\"&(SAT1+7*(H{r}-1)),'ساعات عمر'!$A$%d:$A$%d,\"<=\"&G{r}),"
                    "SUMIFS('ساعات عبدالعزيز'!$F$%d:$F$%d,'ساعات عبدالعزيز'!$A$%d:$A$%d,\">=\"&(SAT1+7*(H{r}-1)),'ساعات عبدالعزيز'!$A$%d:$A$%d,\"<=\"&G{r})))"
                    ).replace('{r}', str(r)) % ((B0, OEND) * 3 + (B0, ZEND) * 3)
    A['E%d' % r] = ('=IF(G{r}="","",IF(OR(H{r}<1,H{r}>NW),"⛔ التاريخ برّا السنة",IF(AND(N(C{r})>0,N(J{r})>0,C{r}>ADV_CAP*J{r}),'
                    '"⚠️ أكثر من "&TEXT(ADV_CAP,"0%")&" من شغله ("&TEXT(J{r},"0.000")&")","")))').replace('{r}', str(r))
    auto(A, 'E%d' % r); A['E%d' % r].font = F(9, True, 'C55A11')
dv(A, '"%s"' % ','.join(EMPS), 'B%d:B%d' % (AR0, AE)); dv_date(A, 'A%d:A%d' % (AR0, AE))
for c in 'GHIJ': A.column_dimensions[c].hidden = True
A['G%d' % AR0].number_format = DATE
A.freeze_panes = 'A10'
for nm, c in (('ADV_D', 'G'), ('ADV_E', 'B'), ('ADV_AMT', 'C'), ('ADV_W', 'H'), ('ADV_M', 'I')):
    name(nm, "'السلف'!$%s$%d:$%s$%d" % (c, AR0, c, AE))
for i, nm in enumerate(EMPS):
    r = 4 + i; A['A%d' % r] = nm
    A['B%d' % r] = '=SUMIFS(ADV_AMT,ADV_E,A%d)' % r
    A['C%d' % r] = '=SUMIFS(ADV_AMT,ADV_E,A%d,ADV_W,WK)' % r
    for c in 'ABC': auto(A, '%s%d' % (c, r), MONEY if c != 'A' else None, True, 12)
A.conditional_formatting.add('A%d:E%d' % (AR0, AE), FormulaRule(formula=['AND($A%d<>"",$G%d="")' % (AR0, AR0)], fill=fill('F8CBAD')))
A.conditional_formatting.add('A%d:E%d' % (AR0, AE), FormulaRule(formula=['LEFT($E%d,1)="⚠"' % AR0], fill=fill('FDE2C8')))
A['A%d' % (AE + 2)] = '="ضل "&COUNTBLANK(A%d:A%d)&" سطر فاضي بالسجل"' % (AR0, AE); A['A%d' % (AE + 2)].font = F(9, False, '555555')
note(A, 'B%d' % (AE + 2), 'مثال: 0510 · عمر المصري · 20.000 · «كاش من الصندوق»', 'B%d:E%d' % (AE + 2, AE + 2))

# ═══════════════════ القروض والديون (سجل السنة) ═══════════════════
LTYPES = ['بضاعة', 'قرض كاش', 'دين قديم']
L = wb.create_sheet('القروض والديون')
title(L, 'القروض والديون — بتنخصم أقساط كل أسبوع', RUST, 'G', YEAR_SUB)
for k, w in zip('ABCDEFG', [14, 16, 18, 16, 15, 16, 40]): L.column_dimensions[k].width = w
header(L, 3, ['الموظف', 'القروض والديون (مع الماضي)', 'انخصم أقساط لحد الأسبوع المختار', 'دفع كاش من جيبته', 'الدين المتبقي', 'القسط الأسبوعي الحالي', 'قديش ضل'], RUST, 10, 34)
expl = [
    '• كل سطر: «النوع»: بضاعة (أخذ من المحل) · قرض كاش (مصاري) · دين قديم. والمبلغ والقسط الأسبوعي ← القسط بينخصم كل أسبوع لحاله لحد ما يخلص.',
    '• البضاعة: اكتب بالبيان الصنف والكمية (مثل: درم ديلوكس ابيض ×1) — بتنزل فاتورة بيع على أودو باسم الموظف. · تغيير القسط: سطر جديد فيه القسط بس. · «دفع كاش من جيبته» بينقص الدين فوراً.',
]
for i, tx in enumerate(expl): note(L, 'A%d' % (7 + i), tx, 'A%d:G%d' % (7 + i, 7 + i), 10, '333333', 26)
header(L, 10, ['التاريخ', 'الموظف', 'النوع', 'المبلغ (دينار)', 'القسط الأسبوعي (دينار)', 'دفع كاش من جيبته (دينار)', 'البيان — للبضاعة: الصنف والكمية'], RUST, 11, 34)
LR0, LE = 11, 11 + LOAN_ROWS - 1
for r in range(LR0, LE + 1):
    edit(L, 'A%d' % r, DIN); edit(L, 'B%d' % r); edit(L, 'C%d' % r)
    for c in 'DEF': edit(L, '%s%d' % (c, r), MONEY)
    edit(L, 'G%d' % r); L['G%d' % r].alignment = C('right')
    L['J%d' % r] = r - LR0 + 1
    L['K%d' % r] = '=' + DVAL('A%d' % r); L['L%d' % r] = '=' + WKOF('K%d' % r); L['M%d' % r] = '=' + MONOF('L%d' % r)
dv(L, '"%s"' % ','.join(EMPS), 'B%d:B%d' % (LR0, LE)); dv_date(L, 'A%d:A%d' % (LR0, LE))
dv(L, '"%s"' % ','.join(LTYPES), 'C%d:C%d' % (LR0, LE), 'بضاعة · قرض كاش · دين قديم')
L.conditional_formatting.add('C%d:C%d' % (LR0, LE), FormulaRule(formula=['AND($D%d<>"",$C%d="")' % (LR0, LR0)], fill=fill('FDE2C8')))
L.conditional_formatting.add('G%d:G%d' % (LR0, LE), FormulaRule(formula=['AND($C%d="بضاعة",$G%d="")' % (LR0, LR0)], fill=fill('FDE2C8')))
L.conditional_formatting.add('A%d:G%d' % (LR0, LE), FormulaRule(formula=['AND($A%d<>"",$K%d="")' % (LR0, LR0)], fill=fill('F8CBAD')))
for c in 'JKLM': L.column_dimensions[c].hidden = True
L.freeze_panes = 'A11'
for nm, c in (('LG_D', 'K'), ('LG_E', 'B'), ('LG_TYPE', 'C'), ('LG_DEBT', 'D'), ('LG_INST', 'E'), ('LG_REP', 'F'), ('LG_IDX', 'J'), ('LG_W', 'L'), ('LG_M', 'M')):
    name(nm, "'القروض والديون'!$%s$%d:$%s$%d" % (c, LR0, c, LE))
L['A%d' % (LE + 2)] = '="ضل "&COUNTBLANK(A%d:A%d)&" سطر فاضي"' % (LR0, LE); L['A%d' % (LE + 2)].font = F(9, False, '555555')
note(L, 'B%d' % (LE + 2), 'أمثلة: 0710 · عمر المصري · بضاعة · 45.000 · القسط 15.000 · «درم ديلوكس ابيض ×2»   |   0910 · عمر المصري · قرض كاش · 497.000 · القسط 30.000   |   2010 · عمر المصري · دفع كاش من جيبته 50.000', 'B%d:G%d' % (LE + 2, LE + 2))

def INST(emp, d, opening):
    """القسط الساري بتاريخ معيّن = آخر قسط انكتب لهالموظف لحد هالتاريخ، وإلا قسط الافتتاح (بلا MAXIFS)."""
    m = 'SUMPRODUCT(MAX((LG_E="%s")*(LG_INST<>"")*(LG_D<=%s)*LG_IDX))' % (emp, d)
    return 'IF(%s=0,%s,INDEX(LG_INST,%s))' % (m, opening, m)

# ═══════════════════ التسوية (شاشة الأسبوع المختار) ═══════════════════
S = wb['Sheet']; S.title = 'التسوية'
title(S, 'التسوية — قديش بدفع لكل موظف يوم الجمعة', NAVY, 'I')
for k, w in zip('ABCDEFGHI', [16, 14, 13, 12, 13, 17, 13, 13, 14]): S.column_dimensions[k].width = w
S['B3'] = 'السنة'; S['B3'].font = F(13, True, NAVY); S['B3'].alignment = C('left')
S['E3'] = 2026; edit(S, 'E3', '0', True, 14); dv(S, "='القوائم'!$E$1:$E$10", 'E3')
S.merge_cells('C3:D3'); S['C3'] = '="سنة "&E3&" · "&NW&" أسبوع"'; auto(S, 'C3', None, True, 12)
S.merge_cells('F3:I3'); note(S, 'F3', '← الملف للسنة كاملة: الأسبوع من السبت للجمعة، وبينسب للشهر اللي فيه جمعته (يوم القبض)', sz=10, color=RUST)
S.row_dimensions[3].height = 30
S['B5'] = 'الأسبوع'; S['B5'].font = F(13, True, NAVY); S['B5'].alignment = C('left')
S.merge_cells('C5:E5'); S['C5'] = None; edit(S, 'C5', None, True, 12)
dv(S, '=WTITLE', 'C5', 'اتركها فاضية = الأسبوع الحالي لحاله · أو اختار أي أسبوع', 'الأسبوع')
S.merge_cells('F5:I5'); S['F5'] = '=IF(C5="","← فاضية: بيعرض الأسبوع الحالي لحاله ("&INDEX(WTITLE,CURWK)&")","← بتعرض أسبوع مختار. امسحها عشان ترجع للأسبوع الحالي")'
S['F5'].font = F(10, False, RUST); S['F5'].alignment = C('right'); S.row_dimensions[5].height = 30
S['N5'] = '=IF(C5="",CURWK,IFERROR(MATCH(C5,WTITLE,0),CURWK))'; S['N5'].font = F(8, False, 'FFFFFF')
name('WK', "'التسوية'!$N$5")
S['A7'] = '="تسوية "&INDEX(WTITLE,WK)&"   (يوم القبض: الجمعة "&%s&")"' % TXT('INDEX(WFRI,WK)')
S['A7'].font = F(14, True, NAVY); S.merge_cells('A7:I7'); S['A7'].alignment = C('right'); S.row_dimensions[7].height = 26
COLS = ['المستحق', 'إكرامية', 'سلف', 'قسط الدين', 'الصافي للدفع', 'المدفوع', 'الباقي له', 'الدين المتبقي']
header(S, 8, ['الموظف'] + COLS, NAVY, 11, 30)

# ═══════════════════ دفتر الأسابيع (كله محسوب) ═══════════════════
W = wb.create_sheet('دفتر الأسابيع'); WS_ = "'دفتر الأسابيع'"
title(W, 'دفتر الأسابيع — كل أسبوع بالسنة لكل موظف (محسوب، ما في كتابة هون)', NAVY, 'L', YEAR_SUB)
for k, w in zip('ABCDEFGHIJKL', [9, 12, 12, 11, 11, 11, 13, 11, 12, 13, 12, 8]): W.column_dimensions[k].width = w
note(W, 'A3', '«الصافي للدفع» = الباقي من الأسبوع الماضي + المستحق + الإكرامية − السلف − القسط (− اللي قبضه يومياً لعبدالعزيز). «الباقي له» = الصافي − المدفوع، وبينتقل لحاله للأسبوع الجاي.', 'A3:L3', 9, h=28)
LEDGER = {}
def week_ledger(top, emp, color, rows, is_azz):
    hs = "'ساعات عبدالعزيز'" if is_azz else "'ساعات عمر'"
    start = 'START_Z' if is_azz else 'START_O'
    W['A%d' % top] = emp; W['A%d' % top].font = F(13, True, color); W.merge_cells('A%d:L%d' % (top, top)); W['A%d' % top].alignment = C('right')
    heads = ['الأسبوع', 'الجمعة'] + COLS + (['قبض يومياً'] if is_azz else ['']) + ['شهر']
    header(W, top + 1, heads, color, 10, 30)
    r0 = top + 2
    for k in range(NWMAX):
        r = r0 + k; w = k + 1; s1, s2, s3, hr = rows[k]
        act = '%s!$M$%d=1' % (hs, hr)
        fri = 'B%d' % r
        W['A%d' % r] = '=IF(%d<=NW,%d,"")' % (w, w)
        W['B%d' % r] = '=IF(%d<=NW,SAT1+%d,"")' % (w, 7 * k + 6)
        W['C%d' % r] = '=IF(%s,N(%s!%s%d),"")' % (act, hs, 'F' if is_azz else 'H', s2)
        W['D%d' % r] = '=IF(%s,N(%s!%s%d),"")' % (act, hs, 'C' if is_azz else 'D', s3)
        W['E%d' % r] = '=IF(%s,SUMIFS(ADV_AMT,ADV_E,"%s",ADV_W,%d),"")' % (act, emp, w)
        W['K%d' % r] = ('=IF(%s,N(%s!G%d),"")' % (act, hs, s1)) if is_azz else '=0'
        owed = '{o}+SUMIFS(LG_DEBT,LG_E,"{e}",LG_D,"<="&{f})-SUMIFS(LG_REP,LG_E,"{e}",LG_D,"<="&{f})'.format(o=OP('D', emp), e=emp, f=fri)
        prev_inst = '0' if k == 0 else 'SUM(F%d:F%d)' % (r0, r - 1)
        prevH = 'IF(%d=%s,%s,N(I%d))' % (w, start, OP('C', emp), r - 1) if k > 0 else 'IF(%d=%s,%s,0)' % (w, start, OP('C', emp))
        room = 'N(C{r})+N(D{r})-N(E{r})-N(K{r})-NET_FLOOR*N(C{r})'.replace('{r}', str(r))
        # القسط = الأصغر من: القسط الساري · باقي الدين · سقف 25% · حد الصافي 65% — ومجبور لتحت لربع دينار (كاش نظيف)
        W['F%d' % r] = '=IF(%s,IF(N(C%d)=0,0,MAX(0,FLOOR(MIN(%s,%s-%s,INST_CAP*(N(C%d)+N(D%d)),%s),0.25))),"")' % (
            act, r, INST(emp, fri, OP('E', emp)), owed, prev_inst, r, r, room)
        W['G%d' % r] = '=IF(%s,%s+N(C{r})+N(D{r})-N(E{r})-N(F{r})-N(K{r}),"")'.replace('{r}', str(r)) % (act, prevH)
        W['H%d' % r] = '=IF(%s,N(%s!%s%d),"")' % (act, hs, 'F' if is_azz else 'G', s3)
        W['I%d' % r] = '=IF(%s,G%d-H%d,"")' % (act, r, r)
        W['J%d' % r] = '=IF(%s,MAX(0,%s-SUM(F%d:F%d)),"")' % (act, owed, r0, r)
        W['L%d' % r] = '=IF(B%d="","",MONTH(B%d))' % (r, r)
        for c in 'ABCDEFGHIJKL': auto(W, '%s%d' % (c, r), MONEY)
        W['A%d' % r].number_format = '0'; W['B%d' % r].number_format = DATE; W['L%d' % r].number_format = '0'
        W['G%d' % r].font = F(11, True, FINAL_C); W['J%d' % r].font = F(10, True, RUST)
        if not is_azz: W['K%d' % r].font = F(8, False, AUTO)
    tr = r0 + NWMAX
    W['A%d' % tr] = 'مجموع السنة'; W.merge_cells('A%d:B%d' % (tr, tr))
    for c in 'CDEFGHK': W['%s%d' % (c, tr)] = '=SUM(%s%d:%s%d)' % (c, r0, c, tr - 1)
    for c in 'ABCDEFGHIJKL': auto(W, '%s%d' % (c, tr), MONEY, True, color=SUBT)
    W.conditional_formatting.add('A%d:L%d' % (r0, tr - 1), FormulaRule(formula=['$A%d=WK' % r0], fill=fill(CURW)))
    W.conditional_formatting.add('A%d:L%d' % (r0, tr - 1), FormulaRule(formula=['AND($A%d<>"",N($H%d)>0)' % (r0, r0)], fill=fill(LOCKED)))
    for c in 'ABCDEFGHIJKL':
        name('%s_%s' % ('LZ' if is_azz else 'LO', c), "%s!$%s$%d:$%s$%d" % (WS_, c, r0, c, tr - 1))
    LEDGER[emp] = (r0, tr)
    return r0, tr
OR0, OTR = week_ledger(5, EMPS[0], NAVY, OMR, False)
ZR0, ZTR = week_ledger(OTR + 3, EMPS[1], TEAL, AZZ, True)
W.freeze_panes = 'A5'

# صفوف الأسبوع المختار بالتسوية
for i, (nm, pre) in enumerate([(EMPS[0], 'LO'), (EMPS[1], 'LZ')]):
    r = 9 + i; S['A%d' % r] = nm; auto(S, 'A%d' % r, None, True, 12)
    for c, src in zip('BCDEFGHI', 'CDEFGHIJ'):
        S['%s%d' % (c, r)] = '=N(INDEX(%s_%s,WK))' % (pre, src)
        auto(S, '%s%d' % (c, r), MONEY, False, 12)
    auto(S, 'F%d' % r, MONEY, True, 16, color=FINAL_BG, fc='FFFFFF')
    S['F%d' % r].border = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)
    S['I%d' % r].font = F(11, True, RUST); S.row_dimensions[r].height = 32
S['A11'] = 'الكاش اللي بتطلعه الجمعة (الاثنين)'; S.merge_cells('A11:E11'); auto(S, 'A11', None, True, 12, color=SUBT); S['A11'].alignment = C('right')
S['F11'] = '=F9+F10'; auto(S, 'F11', MONEY, True, 16, color=FINAL_BG, fc='FFFFFF'); S['F11'].border = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)
S.row_dimensions[11].height = 30
S['A13'] = '=HYPERLINK("#\'ساعات عمر\'!A"&(%d+%d*(WK-1)),"← افتح ساعات عمر لهالأسبوع")' % (B0, BLOCK)
S['A14'] = '=HYPERLINK("#\'ساعات عبدالعزيز\'!A"&(%d+%d*(WK-1)),"← افتح ساعات عبدالعزيز لهالأسبوع")' % (B0, BLOCK)
S['A15'] = '=HYPERLINK("#\'دفتر الأسابيع\'!A"&(%d+WK-1),"← دفتر الأسابيع (كل أسابيع السنة)")' % OR0
for r in (13, 14, 15): S.merge_cells('A%d:E%d' % (r, r)); S['A%d' % r].font = F(12, True, '1F4E79'); S['A%d' % r].alignment = C('right'); S.row_dimensions[r].height = 24
note(S, 'A17', 'الإكرامية والمدفوع يوم الجمعة بتنكتبوا تحت كل أسبوع بورقة الساعات (الخانات الصفراء). · عبدالعزيز: الصافي بعد ما انخصم اللي قبضه يومياً. · «الباقي له» بينتقل لحاله للأسبوع الجاي.', 'A17:I17', 9, h=28)
S.freeze_panes = 'A6'

# نهاية السنة (بورقة بداية السنة)
B['A11'] = 'نهاية السنة — هالأرقام بتنكتب بـ«بداية السنة» بملف السنة الجاية (مع أسبوع بداية = 1)'; B['A11'].font = F(13, True, RUST); B.merge_cells('A11:G11'); B['A11'].alignment = C('right')
header(B, 12, ['الموظف', 'آخر أسبوع', 'مستحق له لسا ما قبضه', 'دين متبقي عليه', 'القسط الأسبوعي'], RUST, 10, 34)
for i, (nm, pre) in enumerate([(EMPS[0], 'LO'), (EMPS[1], 'LZ')]):
    r = 13 + i
    B['A%d' % r] = nm; B['B%d' % r] = '=NW'
    B['C%d' % r] = '=N(INDEX(%s_I,NW))' % pre
    B['D%d' % r] = '=N(INDEX(%s_J,NW))' % pre
    B['E%d' % r] = '=IF(D%d>0,%s,0)' % (r, INST(nm, 'INDEX(WFRI,NW)', '$E$%d' % OPEN[nm]))
    for c in 'ABCDE': auto(B, '%s%d' % (c, r), MONEY if c in 'CDE' else None, True, 12)
    B['B%d' % r].number_format = '0'
    B.merge_cells('F%d:G%d' % (r, r))
    B['F%d' % r] = '=IF(C{r}>0.0005,"المحل لسا مدين إله بـ "&TEXT(C{r},"0.000"),IF(C{r}<-0.0005,"قبض زيادة "&TEXT(-C{r},"0.000"),"ما إله شي"))&IF(D{r}>0.0005,"  ·  وعليه دين "&TEXT(D{r},"0.000"),"")'.replace('{r}', str(r))
    B['F%d' % r].font = F(10, True, RUST); B['F%d' % r].alignment = C('right')

# ملخّص ورقة القروض (لحد الأسبوع المختار)
for i, (nm, pre) in enumerate([(EMPS[0], 'LO'), (EMPS[1], 'LZ')]):
    r = 4 + i
    L['A%d' % r] = nm
    L['B%d' % r] = '=%s+SUMIFS(LG_DEBT,LG_E,A%d,LG_W,"<="&WK)' % (OP('D', nm), r)
    L['C%d' % r] = '=SUMIFS(%s_F,%s_A,"<="&WK)' % (pre, pre)
    L['D%d' % r] = '=SUMIFS(LG_REP,LG_E,A%d,LG_W,"<="&WK)' % r
    L['E%d' % r] = '=MAX(0,B{r}-C{r}-D{r})'.replace('{r}', str(r))
    L['F%d' % r] = '=IF(E%d>0,%s,0)' % (r, INST(nm, 'INDEX(WFRI,WK)', OP('E', nm)))
    L['G%d' % r] = '=IF(E{r}<0.0005,"✔ ما عليه شي",IF(F{r}>0,ROUNDUP(E{r}/F{r},0)&" أسبوع تقريباً","⚠️ حدّد القسط"))'.replace('{r}', str(r))
    for c in 'ABCDEFG': auto(L, '%s%d' % (c, r), MONEY if c in 'BCDEF' else None, True, 11)
    L['E%d' % r].font = F(13, True, RUST); L.row_dimensions[r].height = 24

# ═══════════════════ عمال المهام (سجل السنة) ═══════════════════
T = wb.create_sheet('عمال المهام')
title(T, 'عمال المهام — بيندفعلهم مباشرة', PLUM, 'E', YEAR_SUB)
for k, w in zip('ABCDE', [13, 20, 40, 16, 32]): T.column_dimensions[k].width = w
note(T, 'A3', 'كل مهمة سطر: التاريخ، العامل، شو عمل، وقديش أخذ. بلا أسعار ثابتة. بتنحسب على شهر جمعة أسبوعها.', 'A3:E3')
header(T, 5, ['التاريخ', 'اسم العامل', 'المهمة', 'المبلغ المدفوع (دينار)', 'ملاحظات'], PLUM, 11, 30)
TE = 6 + TASK_ROWS - 1
for r in range(6, TE + 1):
    edit(T, 'A%d' % r, DIN); edit(T, 'B%d' % r); edit(T, 'C%d' % r); edit(T, 'D%d' % r, MONEY); edit(T, 'E%d' % r)
    for c in 'CE': T['%s%d' % (c, r)].alignment = C('right')
    T['G%d' % r] = '=IF(AND(B{r}<>"",COUNTIF(B$6:B{r},B{r})=1),MAX(G$5:G{p})+1,"")'.replace('{r}', str(r)).replace('{p}', str(r - 1))
    T['H%d' % r] = '=' + DVAL('A%d' % r); T['I%d' % r] = '=' + WKOF('H%d' % r); T['J%d' % r] = '=' + MONOF('I%d' % r)
dv_date(T, 'A6:A%d' % TE)
for c in 'GHIJ': T.column_dimensions[c].hidden = True
T.conditional_formatting.add('A6:E%d' % TE, FormulaRule(formula=['AND($A6<>"",$H6="")'], fill=fill('F8CBAD')))
T['C%d' % (TE + 1)] = 'مجموع السنة'; auto(T, 'C%d' % (TE + 1), None, True, color=SUBT)
T['D%d' % (TE + 1)] = '=SUM(D6:D%d)' % TE; auto(T, 'D%d' % (TE + 1), MONEY, True, 12, color=SUBT)
T.freeze_panes = 'A6'
T['A%d' % (TE + 3)] = '="ضل "&COUNTBLANK(A6:A%d)&" سطر فاضي"' % TE; T['A%d' % (TE + 3)].font = F(9, False, '555555')
note(T, 'B%d' % (TE + 3), 'مثال: 0710 · محمد · تنزيل طبلية جبسمبورد · 10.000', 'B%d:E%d' % (TE + 3, TE + 3))
for nm, c in (('TASK_AMT', 'D'), ('TASK_B', 'B'), ('TASK_W', 'I'), ('TASK_M', 'J'), ('TASK_G', 'G')):
    name(nm, "'عمال المهام'!$%s$6:$%s$%d" % (c, c, TE))

# ═══════════════════ ملخص الأيدي العاملة ═══════════════════
M = wb.create_sheet('ملخص الأيدي العاملة')
title(M, 'ملخص الأيدي العاملة — شهر بشهر (حسب جمعة الأسبوع) ووضع كل موظف', NAVY, 'J', YEAR_SUB)
for k, w in zip('ABCDEFGHIJ', [18, 12, 12, 11, 11, 14, 13, 11, 12, 13]): M.column_dimensions[k].width = w
M['A4'] = 'قديش دفعت أيدي عاملة كل شهر (كاش طالع فعلياً) — الشهر = الأشهر اللي فيها جمعة القبض'; M['A4'].font = F(13, True, NAVY); M.merge_cells('A4:J4'); M['A4'].alignment = C('right')
header(M, 5, ['الشهر', 'عدد الجمع', 'رواتب عمر', 'عبدالعزيز (يومي + جمعة)', 'سلف كاش', 'عمال المهام', 'المجموع كاش', 'كلفة الشغل', 'بضاعة أخذوها', 'قروض كاش انعطت'], NAVY, 9, 40)
for m in range(1, 13):
    r = 5 + m
    M['A%d' % r] = MONTHS[m - 1]
    M['B%d' % r] = '=COUNTIF(WMON,%d)' % m
    M['C%d' % r] = '=SUMIFS(LO_H,LO_L,%d)' % m
    M['D%d' % r] = '=SUMIFS(LZ_H,LZ_L,%d)+SUMIFS(LZ_K,LZ_L,%d)' % (m, m)
    M['E%d' % r] = '=SUMIFS(ADV_AMT,ADV_M,%d)' % m
    M['F%d' % r] = '=SUMIFS(TASK_AMT,TASK_M,%d)' % m
    M['G%d' % r] = '=C{r}+D{r}+E{r}+F{r}'.replace('{r}', str(r))
    M['H%d' % r] = '=SUMIFS(LO_C,LO_L,%d)+SUMIFS(LO_D,LO_L,%d)+SUMIFS(LZ_C,LZ_L,%d)+SUMIFS(LZ_D,LZ_L,%d)+F%d' % (m, m, m, m, r)
    M['I%d' % r] = '=SUMIFS(LG_DEBT,LG_TYPE,"بضاعة",LG_M,%d)' % m
    M['J%d' % r] = '=SUMIFS(LG_DEBT,LG_TYPE,"قرض كاش",LG_M,%d)' % m
    for c in 'ABCDEFGHIJ': auto(M, '%s%d' % (c, r), MONEY if c not in 'AB' else None)
    M['B%d' % r].number_format = '0'; M['A%d' % r].alignment = C('right'); M['G%d' % r].font = F(11, True, FINAL_C)
MT = 18
M['A%d' % MT] = 'مجموع السنة'
for c in 'BCDEFGHIJ': M['%s%d' % (c, MT)] = '=SUM(%s6:%s17)' % (c, c)
for c in 'ABCDEFGHIJ': auto(M, '%s%d' % (c, MT), MONEY if c not in 'AB' else None, True, color=SUBT)
M['B%d' % MT].number_format = '0'; M['G%d' % MT].font = F(13, True, 'FFFFFF'); M['G%d' % MT].fill = fill(FINAL_BG)
M.conditional_formatting.add('A6:J17', FormulaRule(formula=['INDEX(WMON,WK)=ROW()-5'], fill=fill(CURW)))
note(M, 'A%d' % (MT + 1), '«كلفة الشغل» = المستحق + الإكرامية + المهام (الأجر كامل، سواء انقبض كاش أو بضاعة). «المجموع كاش» = اللي طلع من الدرج فعلياً. السلف كاش فبتنعدّ · البضاعة والأقساط لا. الأشهر اللي فيها 5 جمع بتطلع أعلى — قارن بمعدل الأسبوع.', 'A%d:J%d' % (MT + 1, MT + 1), 9, h=30)

E0 = MT + 3
M['A%d' % E0] = '="وضع كل موظف لحد "&INDEX(WTITLE,WK)'; M['A%d' % E0].font = F(13, True, NAVY); M.merge_cells('A%d:J%d' % (E0, E0)); M['A%d' % E0].alignment = C('right')
header(M, E0 + 1, ['الموظف', 'ساعات', 'مستحق + إكرامية', 'قبض كاش', '+ سلف', '+ أقساط', '+ الباقي له', 'المجموع', 'الدين المتبقي', 'النتيجة'], NAVY, 9, 36)
for i, (nm, pre, hs, rows) in enumerate([(EMPS[0], 'LO', "'ساعات عمر'", OMR), (EMPS[1], 'LZ', "'ساعات عبدالعزيز'", AZZ)]):
    r = E0 + 2 + i
    M['A%d' % r] = nm
    M['B%d' % r] = '=' + '+'.join('IF(%d<=WK,N(%s!E%d),0)' % (k + 1, hs, rows[k][0]) for k in range(NWMAX))
    M['C%d' % r] = '=%s+SUMIFS(%s_C,%s_A,"<="&WK)+SUMIFS(%s_D,%s_A,"<="&WK)' % (OP('C', nm), pre, pre, pre, pre)
    M['D%d' % r] = '=SUMIFS(%s_H,%s_A,"<="&WK)+SUMIFS(%s_K,%s_A,"<="&WK)' % (pre, pre, pre, pre)
    M['E%d' % r] = '=SUMIFS(%s_E,%s_A,"<="&WK)' % (pre, pre)
    M['F%d' % r] = '=SUMIFS(%s_F,%s_A,"<="&WK)' % (pre, pre)
    M['G%d' % r] = '=N(INDEX(%s_I,WK))' % pre
    M['H%d' % r] = '=D{r}+E{r}+F{r}+G{r}'.replace('{r}', str(r))
    M['I%d' % r] = '=N(INDEX(%s_J,WK))' % pre
    M['J%d' % r] = '=IF(ABS(H{r}-C{r})<0.0005,"✔ مطابق","⚠️ فرق "&TEXT(H{r}-C{r},"0.000"))'.replace('{r}', str(r))
    auto(M, 'A%d' % r, None, True, 11); auto(M, 'B%d' % r, DUR)
    for c in 'CDEFGHI': auto(M, '%s%d' % (c, r), MONEY, c in 'CH', 11)
    M['I%d' % r].font = F(11, True, RUST); auto(M, 'J%d' % r, None, True, 10, fc='1B5E20')
for i, nm in enumerate(EMPS):
    r = E0 + 4 + i; src = E0 + 2 + i
    M['A%d' % r] = '="• "&A{s}&": "&IF(G{s}>0.0005,"المحل لسا مدين إله بـ "&TEXT(G{s},"0.000"),IF(G{s}<-0.0005,"قبض زيادة "&TEXT(-G{s},"0.000"),"مخالص — ما إله شي"))&IF(I{s}>0.0005,"  ·  وعليه دين "&TEXT(I{s},"0.000"),"")'.replace('{s}', str(src))
    M.merge_cells('A%d:J%d' % (r, r)); M['A%d' % r].font = F(11, True, RUST); M['A%d' % r].alignment = C('right')
note(M, 'A%d' % (E0 + 6), 'المطابقة: (مستحق له من قبل + كل المستحق + الإكرامية) = قبض كاش + سلف + أقساط + الباقي له — ولا دينار مكرّر.', 'A%d:J%d' % (E0 + 6, E0 + 6), 9)

T0 = E0 + 8
M['A%d' % T0] = 'عمال المهام — السنة، وشهر بتختاره'; M['A%d' % T0].font = F(13, True, PLUM); M.merge_cells('A%d:F%d' % (T0, T0)); M['A%d' % T0].alignment = C('right')
M['H%d' % T0] = 'الشهر:'; M['H%d' % T0].font = F(11, True, PLUM); M['H%d' % T0].alignment = C('left')
M['I%d' % T0] = '=INDEX(WMON,WK)'; edit(M, 'I%d' % T0, '0', True, 12); dv(M, "='القوائم'!$D$1:$D$12", 'I%d' % T0, 'رقم الشهر 1–12 (بيطلع لحاله شهر الأسبوع المختار)')
header(M, T0 + 1, ['#', 'العامل', 'مهام السنة', 'قبض بالسنة', 'مهام الشهر', 'قبض بالشهر'], PLUM, 10, 28)
TW0, WK_ROWS = T0 + 2, 20
for i in range(WK_ROWS):
    r = TW0 + i
    M['A%d' % r] = '=IF(B%d="","",%d)' % (r, i + 1)
    M['B%d' % r] = '=IFERROR(INDEX(TASK_B,MATCH(%d,TASK_G,0)),"")' % (i + 1)
    M['C%d' % r] = '=IF(B{r}="","",COUNTIF(TASK_B,B{r}))'.replace('{r}', str(r))
    M['D%d' % r] = '=IF(B{r}="","",SUMIFS(TASK_AMT,TASK_B,B{r}))'.replace('{r}', str(r))
    M['E%d' % r] = '=IF(B{r}="","",COUNTIFS(TASK_B,B{r},TASK_M,$I$%d))'.replace('{r}', str(r)) % T0
    M['F%d' % r] = '=IF(B{r}="","",SUMIFS(TASK_AMT,TASK_B,B{r},TASK_M,$I$%d))'.replace('{r}', str(r)) % T0
    for c in 'ABCDEF': auto(M, '%s%d' % (c, r), MONEY if c in 'DF' else None)
TWT = TW0 + WK_ROWS
M['B%d' % TWT] = 'المجموع'; M['C%d' % TWT] = '=COUNTA(TASK_B)'; M['D%d' % TWT] = "='عمال المهام'!D%d" % (TE + 1)
M['E%d' % TWT] = '=COUNTIFS(TASK_M,$I$%d)' % T0; M['F%d' % TWT] = '=SUMIFS(TASK_AMT,TASK_M,$I$%d)' % T0
for c in 'ABCDEF': auto(M, '%s%d' % (c, TWT), MONEY if c in 'DF' else None, True, 11, color=SUBT)
M.freeze_panes = 'A5'

# ═══════════════════ التعليمات ═══════════════════
I = wb.create_sheet('التعليمات')
I.sheet_view.rightToLeft = True; I.sheet_view.showGridLines = False
I.column_dimensions['A'].width = 5; I.column_dimensions['B'].width = 118
I['B1'] = 'كيف يشتغل الملف (النظام الأسبوعي السنوي)'; I['B1'].font = F(16, True, NAVY)
steps = [
    ('الفكرة', None),
    ('1', 'ملف واحد للسنة كاملة. الأسبوع من السبت لمساء الجمعة، ومساء الجمعة الموظف بياخذ أجرته. الأسبوع كامل بينسب للشهر اللي فيه جمعته (يوم القبض) — ما بينقسم أبداً بين شهرين.'),
    ('2', 'الأسبوع 1 = اللي جمعته أول جمعة بيناير. السنة فيها 52 أو 53 أسبوع (2027 فيها 53). الأيام اللي آخر ديسمبر وجمعتها بيناير بتروح لملف السنة الجاية.'),
    ('أول السنة (أو أول ما تبلّش)', None),
    ('3', 'بورقة «التسوية» اختر السنة. بورقة «بداية السنة»: أسبوع البداية (أسبوع 1 لسنة كاملة — أو رقم الأسبوع اللي بلّشت فيه)، ومستحق له من قبل، ودين عليه، والقسط — من «نهاية السنة» بملف السنة الماضية.'),
    ('كل يوم', None),
    ('4', 'بورقة الساعات: روح على الأسبوع الحالي (ملوّن أصفر، أو من رابط «افتح ساعات…» بالتسوية). اختر الدخول والخروج من القائمة (كل 5 دقائق، AM/PM) أو اكتبه مثل 7:30 AM. عبدالعزيز: آخر اليوم «قبض؟ نعم».'),
    ('5', 'سلفة كاش: سطر بورقة «السلف» بتاريخها (اختار من آخر 14 يوم، أو اكتب 0510 = 05/10). قرض أو بضاعة: سطر بورقة «القروض والديون» فيه النوع والمبلغ والقسط — والبضاعة بالبيان: الصنف والكمية.'),
    ('6', 'عامل مهمة: سطر بورقة «عمال المهام» نفس اليوم.'),
    ('آخر الأسبوع (الجمعة)', None),
    ('7', 'ورقة «التسوية» بتفتح على الأسبوع الحالي لحالها: لكل موظف «الصافي للدفع» (البرتقالي) = الباقي من الأسبوع الماضي + المستحق + الإكرامية − السلف − القسط (عبدالعزيز: − اللي قبضه يومياً). وتحتهم مجموع الكاش اللي بتطلعه.'),
    ('8', 'تحت الأسبوع بورقة الساعات: اكتب ساعات الجبسمبورد، والإكرامية، و«المدفوع يوم الجمعة». أول ما تكتب المدفوع الأسبوع بيتلوّن أخضر = مقبوض ومقفول. ما تعدّل أسبوع مقبوض — أي تصحيح بسطر الأسبوع الجاي.'),
    ('9', '«الباقي له» (لو دفعت أقل أو أكثر) بينتقل لحاله للأسبوع الجاي — ما في نقل يدوي.'),
    ('السقوف (بتشتغل لحالها)', None),
    ('10', 'السلفة أكثر من 60% من اللي اشتغله لحد يومها بتتلوّن برتقالي (تنبيه بس). القسط ما بيزيد عن 25% من المستحق + الإكرامية، والصافي ما بينزل عن 65% من المستحق — الباقي من القسط بيتأجل لحاله. النسب بورقة «القواعد».'),
    ('الأسعار', None),
    ('11', 'جدول الأسعار بورقة «القواعد» فيه «ساري من»: لما يتغيّر سعر الساعة اكتب سطر جديد بتاريخه. الأسابيع القديمة بتضل على سعرها — ما بيتغيّر الماضي.'),
    ('آخر الشهر', None),
    ('12', 'ورقة «ملخص الأيدي العاملة»: جدول 12 شهر — قديش دفعت كاش، كلفة الشغل، البضاعة، القروض — حسب جمعة كل أسبوع. وتحته وضع كل موظف والمطابقة وعمال المهام.'),
    ('13', 'ارفع الملف لكلود وقلّه «هاد ملف الموظفين لشهر كذا»: البضاعة بتنزل فواتير بيع على أودو باسم الموظف، والأقساط تحصيل. كل شي مسودة أول، وما بينرحّل إلا بـ«رحّل». والملف ما بينرحّل مرتين.'),
    ('الحماية والنسخ', None),
    ('14', 'الخلايا الصفراء بس بتنكتب. الباقي محمي بلا كلمة سر (مراجعة ← إلغاء حماية الورقة إذا لزم). الملف فيه سنة كاملة — خلّيه على الدرايف عشان النسخ تنحفظ لحالها.'),
    ('15', 'كل التواريخ يوم/شهر/سنة. بخانة التاريخ: اختار من القائمة (آخر 14 يوم)، أو اكتب يوم وشهر (0510 = 05/10)، أو التاريخ كامل. رقم اليوم لحاله ما بينفع بالملف السنوي (بيتلوّن أحمر).'),
]
r = 3
for a, b in steps:
    if b is None: I['B%d' % r] = a; I['B%d' % r].font = F(12, True, TEAL); r += 1; continue
    I['A%d' % r] = a; I['A%d' % r].font = F(11, True); I['A%d' % r].alignment = C()
    I['B%d' % r] = b; I['B%d' % r].font = F(11); I['B%d' % r].alignment = C('right'); I.row_dimensions[r].height = 36; r += 1

# ═══════════════════ الترتيب والحماية ═══════════════════
order = ['التسوية', 'ساعات عمر', 'ساعات عبدالعزيز', 'السلف', 'القروض والديون', 'عمال المهام', 'دفتر الأسابيع',
         'بداية السنة', 'القواعد', 'ملخص الأيدي العاملة', 'التعليمات', 'القوائم']
wb._sheets = [wb[n] for n in order]
wb.active = 0
for ws in wb.worksheets:
    for ref in INPUTS.get(ws.title, []):
        ws[ref].protection = Protection(locked=False)
    ws.protection.sheet = True
    ws.protection.formatColumns = False; ws.protection.formatRows = False
wb.save(OUT); print('saved', OUT)
