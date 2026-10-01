# -*- coding: utf-8 -*-
"""نظام الموظفين السنوي الأسبوعي — محلات العون لمواد البناء (النسخة 13 · بأمره 01/10/2026)
ملف واحد للسنة: 53 كتلة أسبوعية (السبت → الجمعة). السنة بتنختار من ورقة «بداية السنة».

القاعدة الذهبية: **الأسبوع كامل بينسب للشهر والسنة اللي فيهم جمعته (يوم القبض)** — ما بينقسم أبداً.
  • الأسبوع 1 = اللي جمعته أول جمعة بيناير · آخر أسبوع = اللي فيه آخر جمعة بديسمبر (52 أو 53).
  • الملخص الشهري = كل الأسابيع اللي جمعتها بهالشهر (+ السلف والمهام والبضاعة بشهر جمعة أسبوعها).

القواعد (26/09 + 01/10 من صاحب المحل):
  • عمر المصري: السبت–الخميس @1.250/ساعة · بعد 6:00 PM @1.500 · الجمعة كل ساعاته @1.250 — الأسعار بجدول «ساري من».
  • عبدالعزيز: 1.000/ساعة أي يوم · بيقبض نهاية كل يوم (دفعة مقدّمة) وبيتسوّى الجمعة.
  • تحميل الجبسمبورد @2.500 — مرة وحدة لكل أسبوع. المستحق بيتجبر لربع دينار.
  • الإكرامية والمدفوع يوم الجمعة: بسطر الأسبوع بورقة الساعات.
  • السقوف: السلفة ≤ 60% من اللي اشتغله لحد يومها (تنبيه) · القسط ≤ 25% · الصافي ≥ 65%.
  • «الباقي له» تراكمي — نقل يدوي مرة بالسنة بس (ورقة «بداية السنة»).
  • 01/10 (النسخة 13): التواريخ كاملة dd/mm/yyyy بكل مكان · عناوين الأعمدة فوق كل أسبوع · التواريخ قبل رقم الأسبوع ·
    لونين واضحين (أصفر = بتعبّيه إنت · رمادي = محسوب) + مفتاح ألوان · دفتر الأسابيع من/إلى · السنة من قائمة بورقة بداية السنة.
  • بلا ماكرو وبلا دوال _xlfn.
"""
import openpyxl, datetime as dt
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName

OUT = '/home/user/Fam-Bam/العون/نظام_الموظفين.xlsx'
VERSION = 13
NAVY, TEAL, GOLD, PLUM, RUST = '1F4E6B', '0F6E6E', '8A6A00', '5B3A6B', '8B3A2F'
EDIT, AUTO, WEEKBAR, SUBT, OUTM = 'FFF6CC', 'EDEDED', 'DCE6EE', 'E3EFE3', 'D9D9D9'
LOCKED, CURW = 'C6E0B4', 'BDD7EE'          # أسبوع مقبوض (مقفول) أخضر · الأسبوع الحالي أزرق
LINE = Side(style='thin', color='9AA5AD'); THICK = Side(style='medium', color=NAVY)
BOX = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
DUR = '[h]" س "mm" د"'
DIN = '[<=3112]00\\/00;dd/mm/yyyy'            # خانة الإدخال بتقبل 0510 — والتاريخ الكامل بيطلع بالخانة اللي جنبها
DATE = 'dd/mm/yyyy'; MONEY = '#,##0.000;[Red]-#,##0.000;-'; HRS = '0.00'; PCT = '0%'
NWMAX = 53
BLOCK = 12                     # عنوان الأسبوع + عناوين الأعمدة + 7 أيام + مجموع + جبسمبورد/المستحق + إكرامية/المدفوع
B0 = 7
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
    dv(ws, "='القوائم'!$N$1:$N$14", rng, 'اختار من آخر 14 يوم، أو اكتب يوم وشهر (0510 = 05/10)، أو التاريخ كامل — التاريخ الكامل بيطلع بالخانة اللي جنبها', 'التاريخ')
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
def header(ws, row, heads, color, sz=10, height=34, start=1):
    for i, h in enumerate(heads, start=start): cellfmt(ws.cell(row, i, h), color, True, sz=sz, fc='FFFFFF')
    ws.row_dimensions[row].height = height
def legend(ws, row, items, last):
    """مفتاح الألوان: [(نص, لون)]"""
    for i, (tx, color) in enumerate(items):
        c = ws.cell(row, 1 + i, tx); cellfmt(c, color, True, sz=9, h='center')
    ws.row_dimensions[row].height = 22
LEGEND = [('🟨 بتعبّيها إنت', EDIT), ('محسوبة — ما بتنكتب', AUTO), ('أسبوع مقبوض = مقفول', LOCKED), ('الأسبوع الحالي', CURW), ('برّا الفترة', OUTM)]

DAYNAME = 'CHOOSE(WEEKDAY({d},1),"الأحد","الاثنين","الثلاثاء","الأربعاء","الخميس","الجمعة","السبت")'
def TXT(x): return 'TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00")&"/"&YEAR(%s)' % (x, x, x)      # تاريخ كامل dd/mm/yyyy
def WTEXT(sat, fri, w): return '"من السبت "&%s&" إلى الجمعة "&%s&" · الأسبوع %d"' % (TXT(sat), TXT(fri), w)
YEAR_SUB = '="سنة "&YR&"   (السنة بتنختار من ورقة «بداية السنة»)"'
def DVAL(x):   # يوم وشهر (0510) ← تاريخ بالسنة (ولو طلع بعد آخر جمعة ← السنة الماضية) · أو تاريخ كامل · رقم لحاله = غلط
    d = 'DATE(YR,MOD({x},100),INT({x}/100))'.format(x=x)
    return ('IF({x}="","",IF(ISNUMBER({x}),IF({x}<=31,"",IF({x}<=3112,IF({d}>SAT1+7*NW-1,DATE(YR-1,MOD({x},100),INT({x}/100)),{d}),{x})),'
            'IFERROR(DATEVALUE({x}),"")))').format(x=x, d=d)
def WKOF(d): return 'IF({d}="","",INT(({d}-SAT1)/7)+1)'.format(d=d)
def MONOF(w): return 'IF({w}="","",IF(OR({w}<1,{w}>NW),"",INDEX(WMON,{w})))'.format(w=w)
def RND(x): return 'IF(ROUND_MODE="للأعلى",CEILING(ROUND((%s),3),0.25),ROUND((%s)*4,0)/4)' % (x, x)

# ═══════════════════ القوائم (مخفية) ═══════════════════
H = wb.create_sheet('القوائم')
def label(m):
    h, mi = divmod(m, 60)
    return '%d:%02d %s' % ((h % 12) or 12, mi, 'AM' if h < 12 else 'PM')
NT = 288
for i, m in enumerate([(6 * 60 + 5 * k) % 1440 for k in range(NT)], start=1):
    H.cell(i, 1, label(m)); H.cell(i, 2, dt.time(m // 60, m % 60)).number_format = 'h:mm AM/PM'
for i in range(12): H.cell(i + 1, 4, i + 1)
for i, y in enumerate(YEARS, start=1): H.cell(i, 5, y)
H['G4'] = "='بداية السنة'!$C$3"                                      # السنة (من قائمة ورقة بداية السنة)
H['G1'] = '=DATE(G4,1,1)+MOD(6-WEEKDAY(DATE(G4,1,1),1),7)-6'
H['G2'] = '=INT((DATE(G4,12,31)-(G1+6))/7)+1'
H['G3'] = '=IF(AND(TODAY()>=G1,TODAY()<=G1+7*G2-1),INT((TODAY()-G1)/7)+1,1)'
H['G1'].number_format = DATE
name('YR', "'القوائم'!$G$4"); name('SAT1', "'القوائم'!$G$1"); name('NW', "'القوائم'!$G$2"); name('CURWK', "'القوائم'!$G$3")
WT0 = 11
for k in range(NWMAX):
    r = WT0 + k
    H['H%d' % r] = '=IF(%d<=NW,SAT1+%d,"")' % (k + 1, 7 * k)
    H['I%d' % r] = '=IF(%d<=NW,SAT1+%d,"")' % (k + 1, 7 * k + 6)
    H['J%d' % r] = '=IF(%d<=NW,%s,"")' % (k + 1, WTEXT('H%d' % r, 'I%d' % r, k + 1))
    H['K%d' % r] = '=IF(%d<=NW,MONTH(I%d),"")' % (k + 1, r)
    H['L%d' % r] = k + 1
    for c in 'HI': H['%s%d' % (c, r)].number_format = DATE
WTE = WT0 + NWMAX - 1
name('WSAT', "'القوائم'!$H$%d:$H$%d" % (WT0, WTE)); name('WFRI', "'القوائم'!$I$%d:$I$%d" % (WT0, WTE))
name('WTITLE', "'القوائم'!$J$%d:$J$%d" % (WT0, WTE)); name('WMON', "'القوائم'!$K$%d:$K$%d" % (WT0, WTE))
for i in range(14):
    H['N%d' % (i + 1)] = '=TODAY()-%d' % (13 - i); H['N%d' % (i + 1)].number_format = DATE
H['P1'] = 'ALAWN_PAYROLL'; H['P2'] = VERSION; H['P3'] = 'سنوي'
H.sheet_state = 'hidden'
TIME_LIST = "='القوائم'!$A$1:$A$%d" % NT
def TV(ref):
    return ("IF(ISNUMBER({r}),MOD({r},1),IFERROR(INDEX('القوائم'!$B$1:$B${n},MATCH(TRIM({r}),'القوائم'!$A$1:$A${n},0)),"
            "IFERROR(MOD(TIMEVALUE({r}),1),\"\")))").format(r=ref, n=NT)

# ═══════════════════ القواعد ═══════════════════
K = wb.create_sheet('القواعد')
title(K, 'القواعد والأسعار — كل الحسابات بتقرأ من هون', GOLD, 'E')
for k, w in zip('ABCDE', [4, 40, 16, 16, 44]): K.column_dimensions[k].width = w
legend(K, 2, LEGEND[:2], 'E')
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
K['A%d' % (RT0 - 1)] = 'أسعار الساعة — حسب التاريخ: كل سطر «ساري من» تاريخه لحد السطر اللي بعده. لما يتغيّر السعر: سطر جديد بتاريخه الكامل (الأسابيع القديمة ما بتتأثر). السطور مرتّبة بالتاريخ.'
K['A%d' % (RT0 - 1)].font = F(11, True, GOLD); K.merge_cells('A%d:E%d' % (RT0 - 1, RT0 - 1)); K['A%d' % (RT0 - 1)].alignment = C('right'); K.row_dimensions[RT0 - 1].height = 34
header(K, RT0, ['ساري من', 'عمر — الساعة العادية', 'عمر — ساعة الإضافي', 'عبدالعزيز — الساعة', 'تحميل الجبسمبورد — الساعة'], GOLD, 10, 34)
for i in range(RATE_ROWS):
    r = RT0 + 1 + i
    edit(K, 'A%d' % r, DATE, True)
    for c in 'BCDE': edit(K, '%s%d' % (c, r), '0.000')
    K['G%d' % r] = '=IF(A%d="",99999999,A%d)' % (r, r)
K['A%d' % (RT0 + 1)] = dt.date(2026, 1, 1); K['B%d' % (RT0 + 1)] = 1.25; K['C%d' % (RT0 + 1)] = 1.5; K['D%d' % (RT0 + 1)] = 1.0; K['E%d' % (RT0 + 1)] = 2.5
K.column_dimensions['G'].hidden = True
RTE = RT0 + RATE_ROWS
for nm, c in (('RT_K', 'G'), ('RT_OMR', 'B'), ('RT_OT', 'C'), ('RT_AZZ', 'D'), ('RT_GYP', 'E')):
    name(nm, "'القواعد'!$%s$%d:$%s$%d" % (c, RT0 + 1, c, RTE))
def RATE(colname, fri): return 'IFERROR(INDEX(%s,MATCH(%s,RT_K,1)),INDEX(%s,1))' % (colname, fri, colname)
note(K, 'A%d' % (RTE + 2), 'مثال: صار سعر عمر 1.400 من 01/03/2027 ← سطر: 01/03/2027 · 1.400 · 1.500 · 1.000 · 2.500. الأسابيع اللي جمعتها قبل 01/03/2027 بتضل على السعر القديم.', 'A%d:E%d' % (RTE + 2, RTE + 2))

# ═══════════════════ بداية السنة (السنة + الافتتاح) ═══════════════════
B = wb.create_sheet('بداية السنة'); BS = "'بداية السنة'"
title(B, 'بداية ونهاية السنة — السنة، والنقل الوحيد اليدوي: مرة بالسنة', RUST, 'G')
for k, w in zip('ABCDEFG', [16, 14, 18, 16, 16, 14, 30]): B.column_dimensions[k].width = w
B['A3'] = 'السنة ←'; B['A3'].font = F(13, True, RUST); B['A3'].alignment = C('left')
B['C3'] = 2026; edit(B, 'C3', '0', True, 14); dv(B, "='القوائم'!$E$1:$E$10", 'C3', 'اختار السنة — كل الملف بيتبرمج عليها', 'السنة')
B.merge_cells('D3:G3'); B['D3'] = '="السنة "&C3&" فيها "&NW&" أسبوع · الأسبوع 1 من السبت "&%s&" · آخر أسبوع بيخلص الجمعة "&%s' % (TXT('SAT1'), TXT('SAT1+7*NW-1'))
B['D3'].font = F(10, True, RUST); B['D3'].alignment = C('right'); B.row_dimensions[3].height = 30
B['A4'] = 'بداية السنة — من «نهاية السنة» بملف السنة الماضية (أو الوضع يوم بدأت تستعمل الملف)'; B['A4'].font = F(13, True, RUST); B.merge_cells('A4:G4'); B['A4'].alignment = C('right')
header(B, 5, ['الموظف', 'أسبوع البداية', 'مستحق له من قبل', 'دين متبقي عليه', 'القسط الأسبوعي', 'بداية الأسبوع', ''], RUST, 10, 34)
B.merge_cells('F5:G5')
OPEN = {}
for i, nm in enumerate(EMPS):
    r = 6 + i; B['A%d' % r] = nm; auto(B, 'A%d' % r, None, True, color='FFFFFF')
    B['B%d' % r] = 40; edit(B, 'B%d' % r, '0', True)
    for c in 'CDE': B['%s%d' % (c, r)] = 0; edit(B, '%s%d' % (c, r), MONEY, True)
    B.merge_cells('F%d:G%d' % (r, r)); B['F%d' % r] = '=IFERROR("من السبت "&%s&" إلى الجمعة "&%s,"")' % (TXT('INDEX(WSAT,B%d)' % r), TXT('INDEX(WFRI,B%d)' % r))
    auto(B, 'F%d' % r, None, False, 10)
    OPEN[nm] = r
note(B, 'A8', '«أسبوع البداية»: أول أسبوع بتستعمل فيه هالملف (أسبوع 1 لسنة كاملة). الأسابيع اللي قبله رمادية وما بتنحسب. «مستحق له» بينضاف لصافي أسبوع البداية · «دين متبقي» والقسط بيكمّلوا ينخصموا كل أسبوع.', 'A8:G8', 9, h=30)
legend(B, 9, LEGEND[:2], 'G')
def OP(c, nm): return '%s!$%s$%d' % (BS, c, OPEN[nm])
name('START_O', OP('B', EMPS[0])); name('START_Z', OP('B', EMPS[1]))

# ═══════════════════ أوراق الساعات ═══════════════════
FINAL_C, FINAL_BG = 'C55A11', 'ED7D31'
DUE_BG, DUE_FC = 'E2F0D9', '1B5E20'
def hours_sheet(ttl, banner, color, widths, text, last):
    ws = wb.create_sheet(ttl)
    title(ws, banner, color, last, YEAR_SUB)
    note(ws, 'A3', text, 'A3:%s3' % last, h=40)
    for k, w in enumerate(widths, start=1): ws.column_dimensions[col(k)].width = w
    legend(ws, 5, LEGEND, last); ws.freeze_panes = 'A6'
    for c in 'KLMNOPQR': ws.column_dimensions[c].hidden = True
    return ws

def week_block_head(ws, k, color, last, start_name, heads):
    """عنوان الأسبوع (التواريخ أولاً ثم رقم الأسبوع) + عناوين الأعمدة + الخلايا المخفية"""
    r = B0 + BLOCK * k; w = k + 1
    sat = 'SAT1+%d' % (7 * k); fri = 'SAT1+%d' % (7 * k + 6)
    ws.merge_cells('A%d:%s%d' % (r, last, r))
    ws['A%d' % r] = ('=IF(%d<=NW,%s&"  (يوم القبض: الجمعة)"&IF(%d<%s,"   — قبل أسبوع البداية (ما بينحسب)",""),"ما في أسبوع %d هالسنة")'
                     ) % (w, WTEXT(sat, fri, w), w, start_name, w)
    cellfmt(ws['A%d' % r], WEEKBAR, True, sz=13, fc=color, h='right'); ws.row_dimensions[r].height = 28
    header(ws, r + 1, heads, color, 10, 26)
    ws['M%d' % r] = '=IF(AND(%d<=NW,%d>=%s),1,0)' % (w, w, start_name)
    ws['L%d' % r] = '=IF(%d<=NW,%s,"")' % (w, fri); ws['L%d' % r].number_format = DATE
    ws['R%d' % r] = w                                                   # رقم الأسبوع (لتلوين الأسبوع الحالي)
    return r

def day_rows(ws, hr, k, paid_ref):
    for d in range(7):
        r = hr + 2 + d
        ws['A%d' % r] = '=IF(%d<=NW,SAT1+%d,"")' % (k + 1, 7 * k + d)
        ws['B%d' % r] = '=IF(A{r}="","",%s)'.replace('{r}', str(r)) % DAYNAME.format(d='A%d' % r)
        ws['M%d' % r] = '=M%d' % hr
        ws['K%d' % r] = '=IF(OR(M{r}=0,C{r}=""),"",%s)'.replace('{r}', str(r)) % TV('C%d' % r)
        ws['L%d' % r] = '=IF(OR(M{r}=0,D{r}=""),"",%s)'.replace('{r}', str(r)) % TV('D%d' % r)
        ws['Q%d' % r] = '=IF(N(%s)>0,1,0)' % paid_ref
        auto(ws, 'A%d' % r, DATE); auto(ws, 'B%d' % r)
        for c in 'KL': ws['%s%d' % (c, r)].number_format = 'h:mm AM/PM'
        if d == 6:
            for c in 'AB': ws['%s%d' % (c, r)].font = F(11, True, NAVY)

# ---- عمر ----
O_HEADS = ['التاريخ', 'اليوم', 'دخول', 'خروج', 'عدد الساعات', 'منها عادي', 'منها إضافي', 'المستحق (دينار)', 'ملاحظات']
O = hours_sheet('ساعات عمر', 'ساعات عمر المصري', NAVY, [14, 10, 12, 12, 10, 11, 11, 14, 30],
    'اختر الدخول والخروج (كل 5 دقائق، AM/PM) — عدد الساعات والعادي والإضافي والمستحق بيطلعوا لحالهم. '
    'تحت كل أسبوع: ساعات الجبسمبورد · المستحق النهائي · الإكرامية · المدفوع يوم الجمعة. '
    'أول ما تكتب «المدفوع» الأسبوع كله بيتلوّن أخضر = مقبوض ومقفول: ما تعدّله، أي تصحيح بسطر الأسبوع الجاي.', 'I')
OMR = []   # (سطر المجموع, سطر المستحق النهائي, سطر الإكرامية/المدفوع, سطر العنوان)
for k in range(NWMAX):
    hr = week_block_head(O, k, NAVY, 'I', 'START_O', O_HEADS)
    s1, s2, s3 = hr + 9, hr + 10, hr + 11; lo, hi = hr + 2, hr + 8; on = 'M%d=1' % hr
    O['N%d' % hr] = '=' + RATE('RT_OMR', 'L%d' % hr); O['O%d' % hr] = '=' + RATE('RT_OT', 'L%d' % hr); O['P%d' % hr] = '=' + RATE('RT_GYP', 'L%d' % hr)
    day_rows(O, hr, k, 'G%d' % s3)
    for d in range(7):
        r = hr + 2 + d
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
Z_HEADS = ['التاريخ', 'اليوم', 'دخول', 'خروج', 'عدد الساعات', 'مستحق اليوم (دينار)', 'قبض؟', 'ملاحظات']
Z = hours_sheet('ساعات عبدالعزيز', 'ساعات عبدالعزيز — بيقبض نهاية كل يوم وبيتسوّى الجمعة', TEAL, [14, 10, 12, 12, 10, 14, 9, 34],
    'دينار للساعة أي يوم. مستحق كل يوم مجبور لربع دينار. لما يقبض آخر اليوم: «قبض؟ نعم» — المبلغ = مستحق اليوم. '
    'اللي قبضه يومياً دفعة مقدّمة بتنخصم من تسوية الجمعة. تحت كل أسبوع: الجبسمبورد · المستحق النهائي · الإكرامية · المدفوع يوم الجمعة (التسوية).', 'H')
AZZ = []
for k in range(NWMAX):
    hr = week_block_head(Z, k, TEAL, 'H', 'START_Z', Z_HEADS)
    s1, s2, s3 = hr + 9, hr + 10, hr + 11; lo, hi = hr + 2, hr + 8; on = 'M%d=1' % hr
    Z['N%d' % hr] = '=' + RATE('RT_AZZ', 'L%d' % hr); Z['P%d' % hr] = '=' + RATE('RT_GYP', 'L%d' % hr)
    day_rows(Z, hr, k, 'F%d' % s3)
    for d in range(7):
        r = hr + 2 + d
        Z['E%d' % r] = '=IF(OR(M{r}=0,K{r}="",L{r}=""),"",MOD(L{r}-K{r},1))'.replace('{r}', str(r))
        Z['F%d' % r] = '=IF(OR(E%d="",M%d=0),"",%s)' % (r, r, RND('E%d*24*$N$%d' % (r, hr)))
        for c in 'CDGH': edit(Z, '%s%d' % (c, r))
        auto(Z, 'E%d' % r, DUR); auto(Z, 'F%d' % r, MONEY, True, 12, color=DUE_BG, fc=DUE_FC)
    Z.merge_cells('A%d:D%d' % (s1, s1)); Z['A%d' % s1] = '=IF(%s,"مجموع أيام الأسبوع %d","")' % (on, k + 1)
    for c in 'EF': Z['%s%d' % (c, s1)] = '=IF(%s,SUM(%s%d:%s%d),"")' % (on, c, lo, c, hi)
    Z['G%d' % s1] = '=IF(%s,SUMIFS(F%d:F%d,G%d:G%d,"نعم"),"")' % (on, lo, hi, lo, hi)
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
    ws.conditional_formatting.add(rng, FormulaRule(formula=['$M%d=0' % B0], fill=fill(OUTM)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['AND($M%d=1,$Q%d=1)' % (B0, B0)], fill=fill(LOCKED)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['$R%d=CURWK' % B0], fill=fill(CURW), font=Font(bold=True, color=NAVY)))   # سطر عنوان الأسبوع الحالي
for ws, rows, c in ((O, OMR, 'H'), (Z, AZZ, 'F')):
    for s1, s2, s3, hr in rows:
        ws.conditional_formatting.add('%s%d' % (c, s2), FormulaRule(formula=['%s%d<>""' % (c, s2)], fill=fill(FINAL_BG), font=Font(color='FFFFFF', bold=True)))

# ═══════════════════ السلف ═══════════════════
A = wb.create_sheet('السلف')
title(A, 'السلف الأسبوعية — بتنخصم كاملة من راتب نفس الأسبوع', RUST, 'F', YEAR_SUB)
for k, w in zip('ABCDEF', [13, 15, 16, 15, 34, 24]): A.column_dimensions[k].width = w
header(A, 3, ['الموظف', 'سلف هالسنة (دينار)', 'سلف الأسبوع المختار (دينار)', '', '', ''], RUST, 10, 32)
note(A, 'A7', '• كاش بياخذه وسط الأسبوع ← بينخصم كامل من صافي نفس الأسبوع. · لو السلفة أكثر من 60% من اللي اشتغله لحد يومها بتتلوّن برتقالي (تنبيه). · القروض والبضاعة بورقة «القروض والديون».', 'A7:F7', 10, '333333', 30)
legend(A, 8, LEGEND[:2], 'F')
header(A, 9, ['التاريخ (إدخال)', 'التاريخ الكامل', 'الموظف', 'المبلغ (دينار)', 'البيان', 'التنبيه'], RUST, 11, 30)
AR0, AE = 10, 10 + ADV_ROWS - 1
for r in range(AR0, AE + 1):
    edit(A, 'A%d' % r, DIN); edit(A, 'C%d' % r); edit(A, 'D%d' % r, MONEY); edit(A, 'E%d' % r); A['E%d' % r].alignment = C('right')
    A['B%d' % r] = '=' + DVAL('A%d' % r); auto(A, 'B%d' % r, DATE)
    A['H%d' % r] = '=' + WKOF('B%d' % r)
    A['I%d' % r] = '=' + MONOF('H%d' % r)
    A['J%d' % r] = ('=IF(OR(B{r}="",H{r}="",C{r}=""),"",IF(C{r}="عمر المصري",'
                    "SUMIFS('ساعات عمر'!$H$%d:$H$%d,'ساعات عمر'!$A$%d:$A$%d,\">=\"&(SAT1+7*(H{r}-1)),'ساعات عمر'!$A$%d:$A$%d,\"<=\"&B{r}),"
                    "SUMIFS('ساعات عبدالعزيز'!$F$%d:$F$%d,'ساعات عبدالعزيز'!$A$%d:$A$%d,\">=\"&(SAT1+7*(H{r}-1)),'ساعات عبدالعزيز'!$A$%d:$A$%d,\"<=\"&B{r})))"
                    ).replace('{r}', str(r)) % ((B0, OEND) * 3 + (B0, ZEND) * 3)
    A['F%d' % r] = ('=IF(A{r}="","",IF(B{r}="","⛔ التاريخ غلط (اكتب يوم وشهر مثل 0510 أو التاريخ كامل)",IF(OR(H{r}<1,H{r}>NW),"⛔ التاريخ برّا السنة",IF(AND(N(D{r})>0,N(J{r})>0,D{r}>ADV_CAP*J{r}),'
                    '"⚠️ أكثر من "&TEXT(ADV_CAP,"0%")&" من شغله ("&TEXT(J{r},"0.000")&")",""))))').replace('{r}', str(r))
    auto(A, 'F%d' % r); A['F%d' % r].font = F(9, True, 'C55A11')
dv(A, '"%s"' % ','.join(EMPS), 'C%d:C%d' % (AR0, AE)); dv_date(A, 'A%d:A%d' % (AR0, AE))
for c in 'HIJ': A.column_dimensions[c].hidden = True
A.freeze_panes = 'A10'
for nm, c in (('ADV_D', 'B'), ('ADV_E', 'C'), ('ADV_AMT', 'D'), ('ADV_W', 'H'), ('ADV_M', 'I')):
    name(nm, "'السلف'!$%s$%d:$%s$%d" % (c, AR0, c, AE))
for i, nm in enumerate(EMPS):
    r = 4 + i; A['A%d' % r] = nm
    A['B%d' % r] = '=SUMIFS(ADV_AMT,ADV_E,A%d)' % r
    A['C%d' % r] = '=SUMIFS(ADV_AMT,ADV_E,A%d,ADV_W,WK)' % r
    for c in 'ABC': auto(A, '%s%d' % (c, r), MONEY if c != 'A' else None, True, 12)
A.conditional_formatting.add('A%d:F%d' % (AR0, AE), FormulaRule(formula=['LEFT($F%d,1)="⛔"' % AR0], fill=fill('F8CBAD')))
A.conditional_formatting.add('A%d:F%d' % (AR0, AE), FormulaRule(formula=['LEFT($F%d,1)="⚠"' % AR0], fill=fill('FDE2C8')))
A['A%d' % (AE + 2)] = '="ضل "&COUNTBLANK(A%d:A%d)&" سطر فاضي بالسجل"' % (AR0, AE); A['A%d' % (AE + 2)].font = F(9, False, '555555')
note(A, 'B%d' % (AE + 2), 'مثال: 0510 (بيطلع 05/10/2026) · عمر المصري · 20.000 · «كاش من الصندوق»', 'B%d:F%d' % (AE + 2, AE + 2))

# ═══════════════════ القروض والديون ═══════════════════
LTYPES = ['بضاعة', 'قرض كاش', 'دين قديم']
L = wb.create_sheet('القروض والديون')
title(L, 'القروض والديون — بتنخصم أقساط كل أسبوع', RUST, 'H', YEAR_SUB)
for k, w in zip('ABCDEFGH', [13, 15, 15, 12, 15, 15, 16, 40]): L.column_dimensions[k].width = w
header(L, 3, ['الموظف', 'القروض والديون (مع الماضي)', 'انخصم أقساط لحد الأسبوع المختار', 'دفع كاش من جيبته', 'الدين المتبقي', 'القسط الأسبوعي الحالي', 'قديش ضل', ''], RUST, 10, 34)
expl = [
    '• كل سطر: «النوع»: بضاعة (أخذ من المحل) · قرض كاش (مصاري) · دين قديم. والمبلغ والقسط الأسبوعي ← القسط بينخصم كل أسبوع لحاله لحد ما يخلص.',
    '• البضاعة: اكتب بالبيان الصنف والكمية (مثل: درم ديلوكس ابيض ×1) — بتنزل فاتورة بيع على أودو باسم الموظف. · تغيير القسط: سطر جديد فيه القسط بس. · «دفع كاش من جيبته» بينقص الدين فوراً.',
]
for i, tx in enumerate(expl): note(L, 'A%d' % (7 + i), tx, 'A%d:H%d' % (7 + i, 7 + i), 10, '333333', 26)
legend(L, 9, LEGEND[:2], 'H')
header(L, 10, ['التاريخ (إدخال)', 'التاريخ الكامل', 'الموظف', 'النوع', 'المبلغ (دينار)', 'القسط الأسبوعي (دينار)', 'دفع كاش من جيبته (دينار)', 'البيان — للبضاعة: الصنف والكمية'], RUST, 11, 34)
LR0, LE = 11, 11 + LOAN_ROWS - 1
for r in range(LR0, LE + 1):
    edit(L, 'A%d' % r, DIN); edit(L, 'C%d' % r); edit(L, 'D%d' % r)
    for c in 'EFG': edit(L, '%s%d' % (c, r), MONEY)
    edit(L, 'H%d' % r); L['H%d' % r].alignment = C('right')
    L['B%d' % r] = '=' + DVAL('A%d' % r); auto(L, 'B%d' % r, DATE)
    L['J%d' % r] = r - LR0 + 1
    L['L%d' % r] = '=' + WKOF('B%d' % r); L['M%d' % r] = '=' + MONOF('L%d' % r)
dv(L, '"%s"' % ','.join(EMPS), 'C%d:C%d' % (LR0, LE)); dv_date(L, 'A%d:A%d' % (LR0, LE))
dv(L, '"%s"' % ','.join(LTYPES), 'D%d:D%d' % (LR0, LE), 'بضاعة · قرض كاش · دين قديم')
L.conditional_formatting.add('D%d:D%d' % (LR0, LE), FormulaRule(formula=['AND($E%d<>"",$D%d="")' % (LR0, LR0)], fill=fill('FDE2C8')))
L.conditional_formatting.add('H%d:H%d' % (LR0, LE), FormulaRule(formula=['AND($D%d="بضاعة",$H%d="")' % (LR0, LR0)], fill=fill('FDE2C8')))
L.conditional_formatting.add('A%d:H%d' % (LR0, LE), FormulaRule(formula=['AND($A%d<>"",$B%d="")' % (LR0, LR0)], fill=fill('F8CBAD')))
for c in 'JLM': L.column_dimensions[c].hidden = True
L.freeze_panes = 'A11'
for nm, c in (('LG_D', 'B'), ('LG_E', 'C'), ('LG_TYPE', 'D'), ('LG_DEBT', 'E'), ('LG_INST', 'F'), ('LG_REP', 'G'), ('LG_IDX', 'J'), ('LG_W', 'L'), ('LG_M', 'M')):
    name(nm, "'القروض والديون'!$%s$%d:$%s$%d" % (c, LR0, c, LE))
L['A%d' % (LE + 2)] = '="ضل "&COUNTBLANK(A%d:A%d)&" سطر فاضي"' % (LR0, LE); L['A%d' % (LE + 2)].font = F(9, False, '555555')
note(L, 'B%d' % (LE + 2), 'أمثلة: 0710 · عمر المصري · بضاعة · 45.000 · القسط 15.000 · «درم ديلوكس ابيض ×2»   |   0910 · عمر المصري · قرض كاش · 497.000 · القسط 30.000   |   2010 · عمر المصري · دفع كاش من جيبته 50.000', 'B%d:H%d' % (LE + 2, LE + 2))

def INST(emp, d, opening):
    m = 'SUMPRODUCT(MAX((LG_E="%s")*(LG_INST<>"")*(LG_D<=%s)*LG_IDX))' % (emp, d)
    return 'IF(%s=0,%s,INDEX(LG_INST,%s))' % (m, opening, m)

# ═══════════════════ التسوية ═══════════════════
S = wb['Sheet']; S.title = 'التسوية'
title(S, 'التسوية — قديش بدفع لكل موظف يوم الجمعة', NAVY, 'I')
for k, w in zip('ABCDEFGHI', [16, 14, 13, 12, 13, 17, 13, 13, 14]): S.column_dimensions[k].width = w
S['B3'] = 'السنة'; S['B3'].font = F(13, True, NAVY); S['B3'].alignment = C('left')
S['C3'] = '=YR'; auto(S, 'C3', '0', True, 14)
S.merge_cells('D3:I3'); S['D3'] = '="فيها "&NW&" أسبوع · السنة بتنختار من ورقة «بداية السنة»"'; S['D3'].font = F(10, False, RUST); S['D3'].alignment = C('right')
S.row_dimensions[3].height = 30
S['B5'] = 'الأسبوع'; S['B5'].font = F(13, True, NAVY); S['B5'].alignment = C('left')
S.merge_cells('C5:F5'); S['C5'] = None; edit(S, 'C5', None, True, 12)
dv(S, '=WTITLE', 'C5', 'اتركها فاضية = الأسبوع الحالي لحاله · أو اختار أي أسبوع (من السبت … إلى الجمعة … · رقم الأسبوع)', 'الأسبوع')
S.merge_cells('G5:I5'); S['G5'] = '=IF(C5="","← فاضية = الأسبوع الحالي لحاله","← أسبوع مختار. امسحها عشان ترجع للحالي")'
S['G5'].font = F(10, False, RUST); S['G5'].alignment = C('right'); S.row_dimensions[5].height = 30
S['N5'] = '=IF(C5="",CURWK,IFERROR(MATCH(C5,WTITLE,0),CURWK))'; S['N5'].font = F(8, False, 'FFFFFF')
name('WK', "'التسوية'!$N$5")
legend(S, 6, LEGEND[:4], 'I')
S['A7'] = '="تسوية: "&INDEX(WTITLE,WK)'
S['A7'].font = F(14, True, NAVY); S.merge_cells('A7:I7'); S['A7'].alignment = C('right'); S.row_dimensions[7].height = 26
COLS = ['المستحق', 'إكرامية', 'سلف', 'قسط الدين', 'الصافي للدفع', 'المدفوع', 'الباقي له', 'الدين المتبقي']
header(S, 8, ['الموظف'] + COLS, NAVY, 11, 30)

# ═══════════════════ دفتر الأسابيع ═══════════════════
W = wb.create_sheet('دفتر الأسابيع'); WS_ = "'دفتر الأسابيع'"
title(W, 'دفتر الأسابيع — كل أسبوع بالسنة لكل موظف (محسوب، ما في كتابة هون)', NAVY, 'M', YEAR_SUB)
for k, w in zip('ABCDEFGHIJKLM', [9, 13, 13, 12, 11, 11, 11, 13, 11, 12, 13, 12, 7]): W.column_dimensions[k].width = w
note(W, 'A3', '«الصافي للدفع» = الباقي من الأسبوع الماضي + المستحق + الإكرامية − السلف − القسط (− اللي قبضه يومياً لعبدالعزيز). «الباقي له» = الصافي − المدفوع، وبينتقل لحاله للأسبوع الجاي.', 'A3:M3', 9, h=28)
legend(W, 4, [LEGEND[1], LEGEND[2], LEGEND[3]], 'M')
LKEYS = ['WK', 'SAT', 'FRI', 'DUE', 'TIP', 'ADV', 'INST', 'NET', 'PAID', 'REM', 'DEBT', 'DAILY', 'MON']   # A..M
def LN(pre, key): return '%s_%s' % (pre, key)
def week_ledger(top, emp, color, rows, is_azz):
    hs = "'ساعات عبدالعزيز'" if is_azz else "'ساعات عمر'"
    start = 'START_Z' if is_azz else 'START_O'; pre = 'LZ' if is_azz else 'LO'
    W['A%d' % top] = emp; W['A%d' % top].font = F(13, True, color); W.merge_cells('A%d:M%d' % (top, top)); W['A%d' % top].alignment = C('right')
    header(W, top + 1, ['الأسبوع', 'من السبت', 'إلى الجمعة'] + COLS + (['قبض يومياً'] if is_azz else ['']) + ['شهر'], color, 10, 30)
    r0 = top + 2
    for k in range(NWMAX):
        r = r0 + k; w = k + 1; s1, s2, s3, hr = rows[k]
        act = '%s!$M$%d=1' % (hs, hr); fri = 'C%d' % r
        W['A%d' % r] = '=IF(%d<=NW,%d,"")' % (w, w)
        W['B%d' % r] = '=IF(%d<=NW,SAT1+%d,"")' % (w, 7 * k)
        W['C%d' % r] = '=IF(%d<=NW,SAT1+%d,"")' % (w, 7 * k + 6)
        W['D%d' % r] = '=IF(%s,N(%s!%s%d),"")' % (act, hs, 'F' if is_azz else 'H', s2)
        W['E%d' % r] = '=IF(%s,N(%s!%s%d),"")' % (act, hs, 'C' if is_azz else 'D', s3)
        W['F%d' % r] = '=IF(%s,SUMIFS(ADV_AMT,ADV_E,"%s",ADV_W,%d),"")' % (act, emp, w)
        W['L%d' % r] = ('=IF(%s,N(%s!G%d),"")' % (act, hs, s1)) if is_azz else '=0'
        owed = '{o}+SUMIFS(LG_DEBT,LG_E,"{e}",LG_D,"<="&{f})-SUMIFS(LG_REP,LG_E,"{e}",LG_D,"<="&{f})'.format(o=OP('D', emp), e=emp, f=fri)
        prev_inst = '0' if k == 0 else 'SUM(G%d:G%d)' % (r0, r - 1)
        prevH = 'IF(%d=%s,%s,N(J%d))' % (w, start, OP('C', emp), r - 1) if k > 0 else 'IF(%d=%s,%s,0)' % (w, start, OP('C', emp))
        room = 'N(D{r})+N(E{r})-N(F{r})-N(L{r})-NET_FLOOR*N(D{r})'.replace('{r}', str(r))
        # القسط = الأصغر من: الساري · باقي الدين · سقف 25% · حد الصافي 65% — مجبور لتحت لربع دينار
        W['G%d' % r] = '=IF(%s,IF(N(D%d)=0,0,MAX(0,FLOOR(MIN(%s,%s-%s,INST_CAP*(N(D%d)+N(E%d)),%s),0.25))),"")' % (
            act, r, INST(emp, fri, OP('E', emp)), owed, prev_inst, r, r, room)
        W['H%d' % r] = '=IF(%s,%s+N(D{r})+N(E{r})-N(F{r})-N(G{r})-N(L{r}),"")'.replace('{r}', str(r)) % (act, prevH)
        W['I%d' % r] = '=IF(%s,N(%s!%s%d),"")' % (act, hs, 'F' if is_azz else 'G', s3)
        W['J%d' % r] = '=IF(%s,H%d-I%d,"")' % (act, r, r)
        W['K%d' % r] = '=IF(%s,MAX(0,%s-SUM(G%d:G%d)),"")' % (act, owed, r0, r)
        W['M%d' % r] = '=IF(C%d="","",MONTH(C%d))' % (r, r)
        for c in 'ABCDEFGHIJKLM': auto(W, '%s%d' % (c, r), MONEY)
        W['A%d' % r].number_format = '0'; W['B%d' % r].number_format = DATE; W['C%d' % r].number_format = DATE; W['M%d' % r].number_format = '0'
        W['H%d' % r].font = F(11, True, FINAL_C); W['K%d' % r].font = F(10, True, RUST)
        if not is_azz: W['L%d' % r].font = F(8, False, AUTO)
    tr = r0 + NWMAX
    W['A%d' % tr] = 'مجموع السنة'; W.merge_cells('A%d:C%d' % (tr, tr))
    for c in 'DEFGHIL': W['%s%d' % (c, tr)] = '=SUM(%s%d:%s%d)' % (c, r0, c, tr - 1)
    for c in 'ABCDEFGHIJKLM': auto(W, '%s%d' % (c, tr), MONEY, True, color=SUBT)
    W.conditional_formatting.add('A%d:M%d' % (r0, tr - 1), FormulaRule(formula=['AND($A%d<>"",N($I%d)>0)' % (r0, r0)], fill=fill(LOCKED)))
    W.conditional_formatting.add('A%d:M%d' % (r0, tr - 1), FormulaRule(formula=['$A%d=WK' % r0], fill=fill(CURW), font=Font(bold=True)))
    for c, key in zip('ABCDEFGHIJKLM', LKEYS):
        name(LN(pre, key), "%s!$%s$%d:$%s$%d" % (WS_, c, r0, c, tr - 1))
    return r0, tr
OR0, OTR = week_ledger(6, EMPS[0], NAVY, OMR, False)
ZR0, ZTR = week_ledger(OTR + 3, EMPS[1], TEAL, AZZ, True)
W.freeze_panes = 'A6'

# صفوف الأسبوع المختار بالتسوية
for i, (nm, pre) in enumerate([(EMPS[0], 'LO'), (EMPS[1], 'LZ')]):
    r = 9 + i; S['A%d' % r] = nm; auto(S, 'A%d' % r, None, True, 12)
    for c, key in zip('BCDEFGHI', LKEYS[3:11]):
        S['%s%d' % (c, r)] = '=N(INDEX(%s,WK))' % LN(pre, key)
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
S.freeze_panes = 'A7'

# نهاية السنة
B['A11'] = 'نهاية السنة — هالأرقام بتنكتب بـ«بداية السنة» بملف السنة الجاية (مع أسبوع بداية = 1)'; B['A11'].font = F(13, True, RUST); B.merge_cells('A11:G11'); B['A11'].alignment = C('right')
header(B, 12, ['الموظف', 'آخر أسبوع', 'مستحق له لسا ما قبضه', 'دين متبقي عليه', 'القسط الأسبوعي', 'الوضع', ''], RUST, 10, 34); B.merge_cells('F12:G12')
for i, (nm, pre) in enumerate([(EMPS[0], 'LO'), (EMPS[1], 'LZ')]):
    r = 13 + i
    B['A%d' % r] = nm; B['B%d' % r] = '="الأسبوع "&NW&" (الجمعة "&%s&")"' % TXT('INDEX(WFRI,NW)')
    B['C%d' % r] = '=N(INDEX(%s,NW))' % LN(pre, 'REM')
    B['D%d' % r] = '=N(INDEX(%s,NW))' % LN(pre, 'DEBT')
    B['E%d' % r] = '=IF(D%d>0,%s,0)' % (r, INST(nm, 'INDEX(WFRI,NW)', '$E$%d' % OPEN[nm]))
    for c in 'ABCDE': auto(B, '%s%d' % (c, r), MONEY if c in 'CDE' else None, True, 12)
    B['B%d' % r].font = F(9, False, '333333')
    B.merge_cells('F%d:G%d' % (r, r))
    B['F%d' % r] = '=IF(C{r}>0.0005,"المحل لسا مدين إله بـ "&TEXT(C{r},"0.000"),IF(C{r}<-0.0005,"قبض زيادة "&TEXT(-C{r},"0.000"),"ما إله شي"))&IF(D{r}>0.0005,"  ·  وعليه دين "&TEXT(D{r},"0.000"),"")'.replace('{r}', str(r))
    B['F%d' % r].font = F(10, True, RUST); B['F%d' % r].alignment = C('right')

# ملخّص ورقة القروض
for i, (nm, pre) in enumerate([(EMPS[0], 'LO'), (EMPS[1], 'LZ')]):
    r = 4 + i
    L['A%d' % r] = nm
    L['B%d' % r] = '=%s+SUMIFS(LG_DEBT,LG_E,A%d,LG_W,"<="&WK)' % (OP('D', nm), r)
    L['C%d' % r] = '=SUMIFS(%s,%s,"<="&WK)' % (LN(pre, 'INST'), LN(pre, 'WK'))
    L['D%d' % r] = '=SUMIFS(LG_REP,LG_E,A%d,LG_W,"<="&WK)' % r
    L['E%d' % r] = '=MAX(0,B{r}-C{r}-D{r})'.replace('{r}', str(r))
    L['F%d' % r] = '=IF(E%d>0,%s,0)' % (r, INST(nm, 'INDEX(WFRI,WK)', OP('E', nm)))
    L['G%d' % r] = '=IF(E{r}<0.0005,"✔ ما عليه شي",IF(F{r}>0,ROUNDUP(E{r}/F{r},0)&" أسبوع تقريباً","⚠️ حدّد القسط"))'.replace('{r}', str(r))
    for c in 'ABCDEFG': auto(L, '%s%d' % (c, r), MONEY if c in 'BCDEF' else None, True, 11)
    L['E%d' % r].font = F(13, True, RUST); L.row_dimensions[r].height = 24

# ═══════════════════ عمال المهام ═══════════════════
T = wb.create_sheet('عمال المهام')
title(T, 'عمال المهام — بيندفعلهم مباشرة', PLUM, 'F', YEAR_SUB)
for k, w in zip('ABCDEF', [13, 15, 20, 40, 16, 30]): T.column_dimensions[k].width = w
note(T, 'A3', 'كل مهمة سطر: التاريخ، العامل، شو عمل، وقديش أخذ. بلا أسعار ثابتة. بتنحسب على شهر جمعة أسبوعها.', 'A3:F3')
legend(T, 4, LEGEND[:2], 'F')
header(T, 5, ['التاريخ (إدخال)', 'التاريخ الكامل', 'اسم العامل', 'المهمة', 'المبلغ المدفوع (دينار)', 'ملاحظات'], PLUM, 11, 30)
TE = 6 + TASK_ROWS - 1
for r in range(6, TE + 1):
    edit(T, 'A%d' % r, DIN); edit(T, 'C%d' % r); edit(T, 'D%d' % r); edit(T, 'E%d' % r, MONEY); edit(T, 'F%d' % r)
    for c in 'DF': T['%s%d' % (c, r)].alignment = C('right')
    T['B%d' % r] = '=' + DVAL('A%d' % r); auto(T, 'B%d' % r, DATE)
    T['G%d' % r] = '=IF(AND(C{r}<>"",COUNTIF(C$6:C{r},C{r})=1),MAX(G$5:G{p})+1,"")'.replace('{r}', str(r)).replace('{p}', str(r - 1))
    T['I%d' % r] = '=' + WKOF('B%d' % r); T['J%d' % r] = '=' + MONOF('I%d' % r)
dv_date(T, 'A6:A%d' % TE)
for c in 'GIJ': T.column_dimensions[c].hidden = True
T.conditional_formatting.add('A6:F%d' % TE, FormulaRule(formula=['AND($A6<>"",$B6="")'], fill=fill('F8CBAD')))
T['D%d' % (TE + 1)] = 'مجموع السنة'; auto(T, 'D%d' % (TE + 1), None, True, color=SUBT)
T['E%d' % (TE + 1)] = '=SUM(E6:E%d)' % TE; auto(T, 'E%d' % (TE + 1), MONEY, True, 12, color=SUBT)
T.freeze_panes = 'A6'
T['A%d' % (TE + 3)] = '="ضل "&COUNTBLANK(A6:A%d)&" سطر فاضي"' % TE; T['A%d' % (TE + 3)].font = F(9, False, '555555')
note(T, 'B%d' % (TE + 3), 'مثال: 0710 (بيطلع 07/10/2026) · محمد · تنزيل طبلية جبسمبورد · 10.000', 'B%d:F%d' % (TE + 3, TE + 3))
for nm, c in (('TASK_AMT', 'E'), ('TASK_B', 'C'), ('TASK_W', 'I'), ('TASK_M', 'J'), ('TASK_G', 'G')):
    name(nm, "'عمال المهام'!$%s$6:$%s$%d" % (c, c, TE))

# ═══════════════════ ملخص الأيدي العاملة ═══════════════════
M = wb.create_sheet('ملخص الأيدي العاملة')
title(M, 'ملخص الأيدي العاملة — شهر بشهر (حسب جمعة الأسبوع) ووضع كل موظف', NAVY, 'L', YEAR_SUB)
for k, w in zip('ABCDEFGHIJKL', [17, 12, 12, 8, 11, 12, 10, 11, 13, 12, 11, 12]): M.column_dimensions[k].width = w
M['A4'] = 'قديش دفعت أيدي عاملة كل شهر (كاش طالع فعلياً) — الشهر = الأسابيع اللي جمعتها (يوم القبض) فيه'; M['A4'].font = F(13, True, NAVY); M.merge_cells('A4:L4'); M['A4'].alignment = C('right')
header(M, 5, ['الشهر', 'من السبت', 'إلى الجمعة', 'عدد الجمع', 'رواتب عمر', 'عبدالعزيز (يومي + جمعة)', 'سلف كاش', 'عمال المهام', 'المجموع كاش', 'كلفة الشغل', 'بضاعة أخذوها', 'قروض كاش انعطت'], NAVY, 9, 40)
for m in range(1, 13):
    r = 5 + m
    M['A%d' % r] = MONTHS[m - 1]
    M['B%d' % r] = '=IFERROR(INDEX(WSAT,MATCH(%d,WMON,0)),"")' % m
    M['C%d' % r] = '=IFERROR(INDEX(WFRI,MATCH(%d,WMON,0)+COUNTIF(WMON,%d)-1),"")' % (m, m)
    M['D%d' % r] = '=COUNTIF(WMON,%d)' % m
    M['E%d' % r] = '=SUMIFS(LO_PAID,LO_MON,%d)' % m
    M['F%d' % r] = '=SUMIFS(LZ_PAID,LZ_MON,%d)+SUMIFS(LZ_DAILY,LZ_MON,%d)' % (m, m)
    M['G%d' % r] = '=SUMIFS(ADV_AMT,ADV_M,%d)' % m
    M['H%d' % r] = '=SUMIFS(TASK_AMT,TASK_M,%d)' % m
    M['I%d' % r] = '=E{r}+F{r}+G{r}+H{r}'.replace('{r}', str(r))
    M['J%d' % r] = '=SUMIFS(LO_DUE,LO_MON,%d)+SUMIFS(LO_TIP,LO_MON,%d)+SUMIFS(LZ_DUE,LZ_MON,%d)+SUMIFS(LZ_TIP,LZ_MON,%d)+H%d' % (m, m, m, m, r)
    M['K%d' % r] = '=SUMIFS(LG_DEBT,LG_TYPE,"بضاعة",LG_M,%d)' % m
    M['L%d' % r] = '=SUMIFS(LG_DEBT,LG_TYPE,"قرض كاش",LG_M,%d)' % m
    for c in 'ABCDEFGHIJKL': auto(M, '%s%d' % (c, r), MONEY if c not in 'ABCD' else None)
    M['B%d' % r].number_format = DATE; M['C%d' % r].number_format = DATE; M['D%d' % r].number_format = '0'
    M['A%d' % r].alignment = C('right'); M['I%d' % r].font = F(11, True, FINAL_C)
MT = 18
M['A%d' % MT] = 'مجموع السنة'
M['B%d' % MT] = '=SAT1'; M['C%d' % MT] = '=SAT1+7*NW-1'
for c in 'DEFGHIJKL': M['%s%d' % (c, MT)] = '=SUM(%s6:%s17)' % (c, c)
for c in 'ABCDEFGHIJKL': auto(M, '%s%d' % (c, MT), MONEY if c not in 'ABCD' else None, True, color=SUBT)
M['B%d' % MT].number_format = DATE; M['C%d' % MT].number_format = DATE; M['D%d' % MT].number_format = '0'
M['I%d' % MT].font = F(13, True, 'FFFFFF'); M['I%d' % MT].fill = fill(FINAL_BG)
M.conditional_formatting.add('A6:L17', FormulaRule(formula=['INDEX(WMON,WK)=ROW()-5'], fill=fill(CURW)))
note(M, 'A%d' % (MT + 1), '«كلفة الشغل» = المستحق + الإكرامية + المهام (الأجر كامل، سواء انقبض كاش أو بضاعة). «المجموع كاش» = اللي طلع من الدرج فعلياً. السلف كاش فبتنعدّ · البضاعة والأقساط لا. الأشهر اللي فيها 5 جمع بتطلع أعلى — قارن بمعدل الأسبوع.', 'A%d:L%d' % (MT + 1, MT + 1), 9, h=30)

E0 = MT + 3
M['A%d' % E0] = '="وضع كل موظف لحد: "&INDEX(WTITLE,WK)'; M['A%d' % E0].font = F(13, True, NAVY); M.merge_cells('A%d:L%d' % (E0, E0)); M['A%d' % E0].alignment = C('right')
header(M, E0 + 1, ['الموظف', 'ساعات', 'مستحق + إكرامية', 'قبض كاش', '+ سلف', '+ أقساط', '+ الباقي له', 'المجموع', 'الدين المتبقي', 'النتيجة'], NAVY, 9, 36)
for i, (nm, pre, hs, rows) in enumerate([(EMPS[0], 'LO', "'ساعات عمر'", OMR), (EMPS[1], 'LZ', "'ساعات عبدالعزيز'", AZZ)]):
    r = E0 + 2 + i; wk = LN(pre, 'WK')
    M['A%d' % r] = nm
    M['B%d' % r] = '=' + '+'.join('IF(%d<=WK,N(%s!E%d),0)' % (k + 1, hs, rows[k][0]) for k in range(NWMAX))
    M['C%d' % r] = '=%s+SUMIFS(%s,%s,"<="&WK)+SUMIFS(%s,%s,"<="&WK)' % (OP('C', nm), LN(pre, 'DUE'), wk, LN(pre, 'TIP'), wk)
    M['D%d' % r] = '=SUMIFS(%s,%s,"<="&WK)+SUMIFS(%s,%s,"<="&WK)' % (LN(pre, 'PAID'), wk, LN(pre, 'DAILY'), wk)
    M['E%d' % r] = '=SUMIFS(%s,%s,"<="&WK)' % (LN(pre, 'ADV'), wk)
    M['F%d' % r] = '=SUMIFS(%s,%s,"<="&WK)' % (LN(pre, 'INST'), wk)
    M['G%d' % r] = '=N(INDEX(%s,WK))' % LN(pre, 'REM')
    M['H%d' % r] = '=D{r}+E{r}+F{r}+G{r}'.replace('{r}', str(r))
    M['I%d' % r] = '=N(INDEX(%s,WK))' % LN(pre, 'DEBT')
    M['J%d' % r] = '=IF(ABS(H{r}-C{r})<0.0005,"✔ مطابق","⚠️ فرق "&TEXT(H{r}-C{r},"0.000"))'.replace('{r}', str(r))
    auto(M, 'A%d' % r, None, True, 11); auto(M, 'B%d' % r, DUR)
    for c in 'CDEFGHI': auto(M, '%s%d' % (c, r), MONEY, c in 'CH', 11)
    M['I%d' % r].font = F(11, True, RUST); auto(M, 'J%d' % r, None, True, 10, fc='1B5E20')
for i, nm in enumerate(EMPS):
    r = E0 + 4 + i; src = E0 + 2 + i
    M['A%d' % r] = '="• "&A{s}&": "&IF(G{s}>0.0005,"المحل لسا مدين إله بـ "&TEXT(G{s},"0.000"),IF(G{s}<-0.0005,"قبض زيادة "&TEXT(-G{s},"0.000"),"مخالص — ما إله شي"))&IF(I{s}>0.0005,"  ·  وعليه دين "&TEXT(I{s},"0.000"),"")'.replace('{s}', str(src))
    M.merge_cells('A%d:L%d' % (r, r)); M['A%d' % r].font = F(11, True, RUST); M['A%d' % r].alignment = C('right')
note(M, 'A%d' % (E0 + 6), 'المطابقة: (مستحق له من قبل + كل المستحق + الإكرامية) = قبض كاش + سلف + أقساط + الباقي له — ولا دينار مكرّر.', 'A%d:L%d' % (E0 + 6, E0 + 6), 9)

T0 = E0 + 8
M['A%d' % T0] = 'عمال المهام — السنة، وشهر بتختاره'; M['A%d' % T0].font = F(13, True, PLUM); M.merge_cells('A%d:F%d' % (T0, T0)); M['A%d' % T0].alignment = C('right')
M['H%d' % T0] = 'الشهر:'; M['H%d' % T0].font = F(11, True, PLUM); M['H%d' % T0].alignment = C('left')
M['I%d' % T0] = '=INDEX(WMON,WK)'; edit(M, 'I%d' % T0, '0', True, 12); dv(M, "='القوائم'!$D$1:$D$12", 'I%d' % T0, 'رقم الشهر 1–12 (بيطلع لحاله شهر الأسبوع المختار)')
M.merge_cells('J%d:L%d' % (T0, T0)); M['J%d' % T0] = '=IFERROR("من "&%s&" إلى "&%s,"")' % (TXT('INDEX(WSAT,MATCH(I%d,WMON,0))' % T0), TXT('INDEX(WFRI,MATCH(I%d,WMON,0)+COUNTIF(WMON,I%d)-1)' % (T0, T0)))
M['J%d' % T0].font = F(9, False, '555555'); M['J%d' % T0].alignment = C('right')
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
M['B%d' % TWT] = 'المجموع'; M['C%d' % TWT] = '=COUNTA(TASK_B)'; M['D%d' % TWT] = "='عمال المهام'!E%d" % (TE + 1)
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
    ('الألوان', None),
    ('3', 'أصفر = خانة بتعبّيها إنت · رمادي = محسوبة (ما بتنكتب) · أخضر = أسبوع مقبوض ومقفول · أزرق = الأسبوع الحالي · رمادي غامق = برّا الفترة. المفتاح موجود فوق كل ورقة.'),
    ('أول السنة (أو أول ما تبلّش)', None),
    ('4', 'بورقة «بداية السنة»: اختر السنة من القائمة (كل الملف بيتبرمج عليها)، وأسبوع البداية (أسبوع 1 لسنة كاملة — أو رقم الأسبوع اللي بلّشت فيه)، ومستحق له من قبل، ودين عليه، والقسط — من «نهاية السنة» بملف السنة الماضية.'),
    ('كل يوم', None),
    ('5', 'بورقة الساعات: كل أسبوع عنوانه فيه تواريخه كاملة ورقمه، وتحته عناوين الأعمدة. الأسبوع الحالي عنوانه أزرق (أو من رابط «افتح ساعات…» بالتسوية). اختر الدخول والخروج من القائمة (كل 5 دقائق، AM/PM) أو اكتبه مثل 7:30 AM. عبدالعزيز: آخر اليوم «قبض؟ نعم».'),
    ('6', 'سلفة كاش: سطر بورقة «السلف» — التاريخ: اختار من آخر 14 يوم، أو اكتب يوم وشهر (0510)، أو التاريخ كامل؛ التاريخ الكامل بيطلع بالخانة اللي جنبها. قرض أو بضاعة: سطر بورقة «القروض والديون» فيه النوع والمبلغ والقسط — والبضاعة بالبيان: الصنف والكمية.'),
    ('7', 'عامل مهمة: سطر بورقة «عمال المهام» نفس اليوم.'),
    ('آخر الأسبوع (الجمعة)', None),
    ('8', 'ورقة «التسوية» بتفتح على الأسبوع الحالي لحالها (أو تختار أسبوع من القائمة: من السبت … إلى الجمعة … · رقم الأسبوع): لكل موظف «الصافي للدفع» (البرتقالي) = الباقي من الأسبوع الماضي + المستحق + الإكرامية − السلف − القسط (عبدالعزيز: − اللي قبضه يومياً). وتحتهم مجموع الكاش.'),
    ('9', 'تحت الأسبوع بورقة الساعات: اكتب ساعات الجبسمبورد، والإكرامية، و«المدفوع يوم الجمعة». أول ما تكتب المدفوع الأسبوع كله بيتلوّن أخضر = مقبوض ومقفول. ما تعدّل أسبوع مقبوض — أي تصحيح بسطر الأسبوع الجاي.'),
    ('10', '«الباقي له» (لو دفعت أقل أو أكثر) بينتقل لحاله للأسبوع الجاي — ما في نقل يدوي. كل أسابيع السنة بورقة «دفتر الأسابيع» (من السبت … إلى الجمعة …).'),
    ('السقوف (بتشتغل لحالها)', None),
    ('11', 'السلفة أكثر من 60% من اللي اشتغله لحد يومها بتتلوّن برتقالي (تنبيه بس). القسط ما بيزيد عن 25% من المستحق + الإكرامية، والصافي ما بينزل عن 65% من المستحق — الباقي من القسط بيتأجل لحاله. النسب بورقة «القواعد».'),
    ('الأسعار', None),
    ('12', 'جدول الأسعار بورقة «القواعد» فيه «ساري من»: لما يتغيّر سعر الساعة اكتب سطر جديد بتاريخه. الأسابيع القديمة بتضل على سعرها — ما بيتغيّر الماضي.'),
    ('آخر الشهر', None),
    ('13', 'ورقة «ملخص الأيدي العاملة»: جدول 12 شهر (كل شهر من السبت … إلى الجمعة … بالتواريخ الكاملة) — قديش دفعت كاش، كلفة الشغل، البضاعة، القروض. وتحته وضع كل موظف والمطابقة وعمال المهام.'),
    ('14', 'ارفع الملف لكلود وقلّه «هاد ملف الموظفين لشهر كذا»: البضاعة بتنزل فواتير بيع على أودو باسم الموظف، والأقساط تحصيل. كل شي مسودة أول، وما بينرحّل إلا بـ«رحّل». والملف ما بينرحّل مرتين.'),
    ('الحماية والنسخ', None),
    ('15', 'الخلايا الصفراء بس بتنكتب. الباقي محمي بلا كلمة سر (مراجعة ← إلغاء حماية الورقة إذا لزم). الملف فيه سنة كاملة — خلّيه على الدرايف عشان النسخ تنحفظ لحالها.'),
    ('16', 'كل التواريخ بالملف كاملة: يوم/شهر/سنة (05/10/2026).'),
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
