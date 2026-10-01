# -*- coding: utf-8 -*-
"""نظام الموظفين السنوي الأسبوعي — محلات العون لمواد البناء (النسخة 16 · بأمره 01/10/2026)
ملف واحد للسنة: 53 كتلة أسبوعية (السبت → الجمعة). السنة بتنختار من ورقة «السنة».

القاعدة الذهبية: **الأسبوع كامل بينسب للشهر والسنة اللي فيهم جمعته (يوم القبض)** — ما بينقسم أبداً.
  • الأسبوع 1 = اللي جمعته أول جمعة بيناير · آخر أسبوع = اللي فيه آخر جمعة بديسمبر (52 أو 53).

القواعد (26/09 + 01/10 من صاحب المحل):
  • عمر المصري: السبت–الخميس @1.250/ساعة · بعد 6:00 PM @1.500 · الجمعة كل ساعاته @1.250 — الأسعار بجدول «ساري من».
  • عبدالعزيز: 1.000/ساعة أي يوم · بيقبض نهاية كل يوم (دفعة مقدّمة) وبيتسوّى الجمعة.
  • تحميل الجبسمبورد: كل ساعة تحميل = 2.500 دينار (ساعات التحميل جزء من ساعات الدوام: سعرها العادي محسوب
    بساعات اليوم، وبينضاف الفرق لـ2.500).
  • الإكرامية خانة إدخال كل أسبوع وبتنجمع داخل المستحق النهائي — بلا جبر تلقائي (الإكرامية هي اللي بتجبره) — النسخة 17.
  • القسط بينخصم زي ما انكتب (النسخة 17)؛ السقوف تنبيه بالتسوية، أو «تخفيض تلقائي» من القواعد.
  • السقوف: السلفة ≤ 60% من اللي اشتغله لحد يومها (تنبيه) · القسط ≤ 25% · الصافي ≥ 65%.
  • «الباقي له» تراكمي — نقل يدوي مرة بالسنة بس (ورقة «السنة»).

النسخة 16 (01/10):
  • 🔴 إصلاح سنة 1900: الاسم SAT1 كان عنوان خلية حقيقي بإكسل (عمود SAT سطر 1) فإكسل قرأه خلية فاضية = 0.
    صار WEEK_ONE، وفي فحص إجباري بآخر السكربت إنه ولا اسم معرّف بيشبه عنوان خلية.
  • البرتقالي = الخلايا اللي بتعبّيها إنت · عناوين الأعمدة أزرق فاتح بخط أسود · بلا مفاتيح ألوان مكتوبة.
  • عنوان الأسبوع: التاريخين ثم رقم الأسبوع (بلا «من/إلى» وبلا «يوم القبض») · بلا سطر «سنة …» فوق.
  • الساعات بصيغة ساعة:دقائق (11:30) بلا أحرف · قائمة الأوقات بس على خانات الدخول والخروج.
  • أزرار واضحة بالتسوية للانتقال لأسبوع عمر وعبدالعزيز ودفتر الأسابيع.
  • بلا ماكرو وبلا دوال _xlfn.
"""
import re, openpyxl, datetime as dt
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.workbook.defined_name import DefinedName

OUT = '/home/user/Fam-Bam/العون/نظام_الموظفين.xlsx'
VERSION = 17
NAVY, TEAL, GOLD, PLUM, RUST = '1F4E79', '0F6E6E', '7F6000', '5B3A6B', '833C0B'   # ألوان عناوين الأوراق (خط أبيض)
# ── ترميز الألوان (بأمره 01/10) ──
INPUT = 'F4B183'                 # برتقالي: بتعبّيها إنت
AUTO = 'F2F2F2'                  # رمادي فاتح: محسوبة
HEAD_BG, HEAD_FC = 'BDD7EE', '000000'      # عناوين الأعمدة: أزرق فاتح، خط أسود
WEEK_BG, WEEK_FC = '9BC2E6', '000000'      # عنوان الأسبوع: أزرق أغمق، خط أسود
SUBT = 'E7E6E6'                  # سطور المجموع والتسميات
FINAL_BG, FINAL_FC = '262626', 'FFFFFF'    # المستحق النهائي والصافي للدفع: فحمي بخط أبيض (الشغلة الغامقة الوحيدة)
LOCKED = 'C6E0B4'                # أسبوع مقبوض (مقفول): أخضر
CURW, CURW_FC, CURW_ROWS = '2F5597', 'FFFFFF', 'DDEBF7'   # الأسبوع الحالي: عنوانه أزرق غامق، وتاريخ أيامه أزرق فاتح
OUTM, OUTM_FC = 'D9D9D9', '595959'          # برّا الفترة
ERR_BG, ERR_FC = 'FFC7CE', '800000'         # خطأ
WARN_BG, WARN_FC = 'FFEB9C', '6B3E00'       # تحذير
BTN_BG, BTN_FC = '2F75B5', 'FFFFFF'         # أزرار الانتقال
LINE = Side(style='thin', color='8EA9C1'); THICK = Side(style='thick', color='000000')
BOX = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)
TBOX = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)
DUR = '[h]:mm'                   # ساعة:دقائق بلا أحرف (11:30) — وبالمجاميع كمان (69:30)
DIN = 'dd/mm/yyyy'
DATE = 'dd/mm/yyyy'; MONEY = '#,##0.000;[Red]-#,##0.000;-'; HRS = '0.00'; PCT = '0%'
NWMAX = 53
BLOCK = 14                       # عنوان + عناوين الأعمدة + 7 أيام + جبسمبورد + إكرامية + المستحق النهائي + المدفوع + سطر فاصل
B0 = 5
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
    ws.add_data_validation(d)
    for part in rng.split(): d.add(part)
def dv_date(ws, rng):
    dv(ws, "='القوائم'!$N$1:$N$14", rng, 'اختار اليوم من القائمة، أو اكتب التاريخ كامل 05/10/2026، أو اسم اليوم بس', 'التاريخ')
def cellfmt(c, color='FFFFFF', bold=False, fmt=None, sz=11, h='center', fc='000000'):
    c.border = BOX; c.fill = fill(color); c.font = F(sz, bold, fc); c.alignment = C(h)
    if fmt: c.number_format = fmt
def edit(ws, ref, fmt=None, bold=False, sz=11):
    c = ws[ref]; cellfmt(c, INPUT, bold, fmt, sz); INPUTS.setdefault(ws.title, []).append(ref)
def auto(ws, ref, fmt=None, bold=False, sz=11, color=AUTO, fc='000000'):
    cellfmt(ws[ref], color, bold, fmt, sz, fc=fc)
def final_box(ws, ref, sz=16):
    cellfmt(ws[ref], FINAL_BG, True, MONEY, sz, fc=FINAL_FC); ws[ref].border = TBOX
def title(ws, text, color, last):
    ws.sheet_view.rightToLeft = True; ws.sheet_view.showGridLines = False
    ws.merge_cells('A1:%s1' % last); ws['A1'] = text; cellfmt(ws['A1'], color, True, sz=18, fc='FFFFFF')
    ws.row_dimensions[1].height = 40
def note(ws, ref, text, rng=None, sz=10, color='404040', h=None):
    if rng: ws.merge_cells(rng)
    ws[ref] = text; ws[ref].font = F(sz, False, color); ws[ref].alignment = C('right')
    if h: ws.row_dimensions[ws[ref].row].height = h
def header(ws, row, heads, sz=10, height=32, start=1):
    for i, h in enumerate(heads, start=start): cellfmt(ws.cell(row, i, h), HEAD_BG, True, sz=sz + 3, fc=HEAD_FC)   # عناوين أكبر (بأمره 01/10)
    ws.row_dimensions[row].height = height + 14
def label(ws, ref, text, rng=None, sz=11, h='right', fc='000000', color=SUBT):
    if rng: ws.merge_cells(rng)
    ws[ref] = text; cellfmt(ws[ref], color, True, sz=sz, h=h, fc=fc)

DAYNAME = 'CHOOSE(WEEKDAY({d},1),"الأحد","الاثنين","الثلاثاء","الأربعاء","الخميس","الجمعة","السبت")'
def TXT(x): return 'TEXT(DAY(%s),"00")&"/"&TEXT(MONTH(%s),"00")&"/"&YEAR(%s)' % (x, x, x)      # تاريخ كامل dd/mm/yyyy
def WTEXT(sat, fri, w): return '%s&"  -  "&%s&"   ·   الأسبوع %d"' % (TXT(sat), TXT(fri), w)       # التاريخين ثم رقم الأسبوع
DAYS_ARR = '{"الأحد","الاثنين","الثلاثاء","الأربعاء","الخميس","الجمعة","السبت"}'
def DVAL(x):
    """تاريخ من: رقم تاريخ · نص «الثلاثاء 29/09/2026» أو «29/09/2026» · اسم اليوم بس (آخر مرة إجا لحد اليوم)"""
    t = 'RIGHT(TRIM({x}),10)'.format(x=x)
    return ('IF({x}="","",IF(ISNUMBER({x}),IF({x}>=36526,{x},""),'
            'IFERROR(DATE(VALUE(RIGHT({t},4)),VALUE(MID({t},4,2)),VALUE(LEFT({t},2))),'
            'IFERROR(TODAY()-MOD(WEEKDAY(TODAY(),1)-MATCH(TRIM({x}),{d},0),7),IFERROR(DATEVALUE({x}),"")))))').format(x=x, t=t, d=DAYS_ARR)
def WKOF(d): return 'IF({d}="","",INT(({d}-WEEK_ONE)/7)+1)'.format(d=d)
def MONOF(w): return 'IF({w}="","",IF(OR({w}<1,{w}>NW),"",INDEX(WMON,{w})))'.format(w=w)
def RND(x): return 'IF(ROUND_MODE="للأعلى",CEILING(ROUND((%s),3),0.25),ROUND((%s)*4,0)/4)' % (x, x)

# ═══════════════════ القوائم (مخفية) ═══════════════════
H = wb.create_sheet('القوائم')
def tlabel(m):
    h, mi = divmod(m, 60)
    return '%d:%02d %s' % ((h % 12) or 12, mi, 'AM' if h < 12 else 'PM')
NT = 288
for i, m in enumerate([(6 * 60 + 5 * k) % 1440 for k in range(NT)], start=1):
    H.cell(i, 1, tlabel(m)); H.cell(i, 2, dt.time(m // 60, m % 60)).number_format = 'h:mm AM/PM'
for i in range(12): H.cell(i + 1, 4, i + 1)
for i, y in enumerate(YEARS, start=1): H.cell(i, 5, y)
H['G4'] = "=IF(AND(ISNUMBER(السنة!$C$3),السنة!$C$3>=2000,السنة!$C$3<=2100),السنة!$C$3,2026)"   # السنة — احتياط 2026
H['G1'] = '=DATE(G4,1,1)+MOD(6-WEEKDAY(DATE(G4,1,1),1),7)-6'       # سبت الأسبوع 1
H['G2'] = '=INT((DATE(G4,12,31)-(G1+6))/7)+1'
H['G3'] = '=IF(AND(TODAY()>=G1,TODAY()<=G1+7*G2-1),INT((TODAY()-G1)/7)+1,1)'
H['G1'].number_format = DATE
# ⚠️ أسماء ما بتشبه عنوان خلية (SAT1 كان عنوان خلية بإكسل ← سنة 1900)
name('YR', "'القوائم'!$G$4"); name('WEEK_ONE', "'القوائم'!$G$1"); name('NW', "'القوائم'!$G$2"); name('CURWK', "'القوائم'!$G$3")
WT0 = 11
for k in range(NWMAX):
    r = WT0 + k
    H['H%d' % r] = '=IF(%d<=NW,WEEK_ONE+%d,"")' % (k + 1, 7 * k)
    H['I%d' % r] = '=IF(%d<=NW,WEEK_ONE+%d,"")' % (k + 1, 7 * k + 6)
    H['J%d' % r] = '=IF(%d<=NW,%s,"")' % (k + 1, WTEXT('H%d' % r, 'I%d' % r, k + 1))
    H['K%d' % r] = '=IF(%d<=NW,MONTH(I%d),"")' % (k + 1, r)
    H['L%d' % r] = k + 1
    for c in 'HI': H['%s%d' % (c, r)].number_format = DATE
WTE = WT0 + NWMAX - 1
name('WSAT', "'القوائم'!$H$%d:$H$%d" % (WT0, WTE)); name('WFRI', "'القوائم'!$I$%d:$I$%d" % (WT0, WTE))
name('WTITLE', "'القوائم'!$J$%d:$J$%d" % (WT0, WTE)); name('WMON', "'القوائم'!$K$%d:$K$%d" % (WT0, WTE))
for i in range(14):   # قائمة التاريخ: «اليوم dd/mm/yyyy» — من اليوم ولورا 13 يوم
    H['N%d' % (i + 1)] = '=%s&" "&%s' % (DAYNAME.format(d='(TODAY()-%d)' % i), TXT('(TODAY()-%d)' % i))
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
header(K, 3, ['#', 'القاعدة', 'القيمة', '', 'الشرح'], 11, 26); K.merge_cells('C3:D3')
rules = [
    ('OMR_OT_AT', 'عمر المصري — بداية الإضافي', dt.time(18, 0), 'h:mm AM/PM', 'السبت–الخميس بعد هالوقت إضافي. الجمعة ما إلها إضافي.'),
    ('OMR_FRI_IN', 'عمر المصري — دوام الجمعة من', dt.time(14, 0), 'h:mm AM/PM', 'للتنبيه بس — الحساب حسب الساعات الفعلية.'),
    ('OMR_FRI_OUT', 'عمر المصري — دوام الجمعة إلى', dt.time(23, 0), 'h:mm AM/PM', ''),
    ('ROUND_MODE', 'جبر مستحق عبدالعزيز اليومي', 'لأقرب ربع', '@', 'بس لقبض عبدالعزيز اليومي. المستحق النهائي للأسبوع ما بينجبر — الإكرامية هي اللي بتجبره.'),
    ('ADV_CAP', 'سقف السلفة (من اللي اشتغله لحد يومها)', 0.60, PCT, 'سلفة أكبر من هيك بتطلع تنبيه بورقة السلف — بتنخصم كاملة.'),
    ('INST_CAP', 'سقف القسط الأسبوعي (من المستحق)', 0.25, PCT, 'قسط أكبر من هالنسبة بيطلع تنبيه بالتسوية (أو بيقل لحاله إذا اخترت «تخفيض تلقائي» تحت).'),
    ('NET_FLOOR', 'الحد الأدنى للصافي (من المستحق)', 0.65, PCT, 'صافي أقل من هالنسبة بيطلع تنبيه بالتسوية (القبض اليومي لعبدالعزيز محسوب قبض مش خصم).'),
    ('INST_MODE', 'القسط اللي كتبته', 'ينخصم كامل', '@', 'ينخصم كامل: بينخصم القسط زي ما كتبته، والسقفين فوق تنبيه بس · تخفيض تلقائي: القسط بيقل لحاله لحد السقفين والباقي بيتأجل.'),
]
for i, (nm, rule, val, fmt, desc) in enumerate(rules, start=1):
    r = 3 + i
    K.cell(r, 1, i); K.cell(r, 2, rule); K.cell(r, 3, val); K.cell(r, 5, desc); K.merge_cells('C%d:D%d' % (r, r))
    for cc in (1, 2, 5): cellfmt(K.cell(r, cc), h='right' if cc > 1 else 'center')
    edit(K, 'C%d' % r, fmt, True); K.cell(r, 5).font = F(10, False, '404040'); K.row_dimensions[r].height = 30
    name(nm, "'القواعد'!$C$%d" % r)
    if nm == 'ROUND_MODE': dv(K, '"لأقرب ربع,للأعلى"', 'C%d' % r)
    if nm == 'INST_MODE': dv(K, '"ينخصم كامل,تخفيض تلقائي"', 'C%d' % r)
RT0 = 3 + len(rules) + 3
K['A%d' % (RT0 - 1)] = 'أسعار الساعة حسب التاريخ: كل سطر «ساري من» تاريخه لحد السطر اللي بعده. لما يتغيّر السعر: سطر جديد بتاريخه (الأسابيع القديمة ما بتتأثر).'
K['A%d' % (RT0 - 1)].font = F(11, True, GOLD); K.merge_cells('A%d:E%d' % (RT0 - 1, RT0 - 1)); K['A%d' % (RT0 - 1)].alignment = C('right'); K.row_dimensions[RT0 - 1].height = 34
header(K, RT0, ['ساري من', 'عمر — الساعة العادية', 'عمر — ساعة الإضافي', 'عبدالعزيز — الساعة', 'تحميل الجبسمبورد — الساعة'], 10, 34)
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

# ═══════════════════ السنة (السنة + الافتتاح) ═══════════════════
B = wb.create_sheet('السنة')
title(B, 'السنة — والنقل الوحيد اليدوي: مرة بالسنة', RUST, 'G')
for k, w in zip('ABCDEFG', [16, 14, 18, 16, 16, 14, 30]): B.column_dimensions[k].width = w
B['A3'] = 'السنة ←'; B['A3'].font = F(13, True, RUST); B['A3'].alignment = C('left')
B['C3'] = 2026; edit(B, 'C3', '0', True, 14); dv(B, "='القوائم'!$E$1:$E$10", 'C3', 'اختار السنة — كل الملف بيتبرمج عليها', 'السنة')
B.merge_cells('D3:G3'); B['D3'] = '="فيها "&NW&" أسبوع · الأسبوع 1: "&%s&"  -  "&%s' % (TXT('WEEK_ONE'), TXT('WEEK_ONE+6'))
B['D3'].font = F(10, True, RUST); B['D3'].alignment = C('right'); B.row_dimensions[3].height = 30
B['A4'] = 'بداية السنة — من «نهاية السنة» بملف السنة الماضية (أو الوضع يوم بدأت تستعمل الملف)'; B['A4'].font = F(13, True, RUST); B.merge_cells('A4:G4'); B['A4'].alignment = C('right')
header(B, 5, ['الموظف', 'أسبوع البداية', 'مستحق له من قبل', 'دين متبقي عليه', 'القسط الأسبوعي', 'الأسبوع', ''], 10, 34)
B.merge_cells('F5:G5')
OPEN = {}
for i, nm in enumerate(EMPS):
    r = 6 + i; B['A%d' % r] = nm; auto(B, 'A%d' % r, None, True, color='FFFFFF')
    B['B%d' % r] = 40; edit(B, 'B%d' % r, '0', True)
    for c in 'CDE': B['%s%d' % (c, r)] = 0; edit(B, '%s%d' % (c, r), MONEY, True)
    B.merge_cells('F%d:G%d' % (r, r)); B['F%d' % r] = '=IFERROR(%s&"  -  "&%s,"")' % (TXT('INDEX(WSAT,B%d)' % r), TXT('INDEX(WFRI,B%d)' % r))
    auto(B, 'F%d' % r, None, False, 10)
    OPEN[nm] = r
note(B, 'A8', '«أسبوع البداية»: أول أسبوع بتستعمل فيه هالملف (أسبوع 1 لسنة كاملة). الأسابيع اللي قبله رمادية وما بتنحسب. «مستحق له» بينضاف لصافي أسبوع البداية · «دين متبقي» والقسط بيكمّلوا ينخصموا كل أسبوع.', 'A8:G8', 9, h=30)
def OP(c, nm): return 'السنة!$%s$%d' % (c, OPEN[nm])
name('START_O', 'السنة!$B$%d' % OPEN[EMPS[0]]); name('START_Z', 'السنة!$B$%d' % OPEN[EMPS[1]])

# ═══════════════════ أوراق الساعات ═══════════════════
# أعمدة مخفية: K/L وقت الدخول/الخروج · M الأسبوع فعّال · N/O/P أسعار الأسبوع · Q الأسبوع مقبوض · R رقم الأسبوع (سطر العنوان) · S رقم الأسبوع (كل الكتلة) · T سطر المستحق النهائي
def hours_sheet(ttl, banner, color, widths, text, last):
    ws = wb.create_sheet(ttl)
    title(ws, banner, color, last)
    note(ws, 'A3', text, 'A3:%s3' % last, h=44)
    for k, w in enumerate(widths, start=1): ws.column_dimensions[col(k)].width = w
    ws.freeze_panes = 'A5'
    for c in 'KLMNOPQRSTUVW': ws.column_dimensions[c].hidden = True
    return ws

def block_head(ws, k, last, start_name, heads, paid_ref):
    r = B0 + BLOCK * k; w = k + 1
    sat = 'WEEK_ONE+%d' % (7 * k); fri = 'WEEK_ONE+%d' % (7 * k + 6)
    ws.merge_cells('A%d:%s%d' % (r, last, r))
    ws['A%d' % r] = '=IF(%d<=NW,%s&IF(%d<%s,"   (قبل أسبوع البداية — ما بينحسب)",""),"ما في أسبوع %d هالسنة")' % (w, WTEXT(sat, fri, w), w, start_name, w)
    cellfmt(ws['A%d' % r], WEEK_BG, True, sz=16, fc=WEEK_FC, h='right'); ws.row_dimensions[r].height = 34
    header(ws, r + 1, heads, 10, 26)
    ws['L%d' % r] = '=IF(%d<=NW,%s,"")' % (w, fri); ws['L%d' % r].number_format = DATE
    ws['R%d' % r] = w
    for rr in range(r, r + BLOCK):
        ws['M%d' % rr] = '=IF(AND(%d<=NW,%d>=%s),1,0)' % (w, w, start_name) if rr == r else '=$M$%d' % r
        ws['Q%d' % rr] = '=IF(N(%s)>0,1,0)' % paid_ref
        ws['S%d' % rr] = w
    ws['T%d' % (r + 11)] = 1
    return r

def day_rows(ws, hr, k):
    for d in range(7):
        r = hr + 2 + d
        ws['A%d' % r] = '=IF(%d<=NW,WEEK_ONE+%d,"")' % (k + 1, 7 * k + d)
        ws['B%d' % r] = '=IF(A{r}="","",%s)'.replace('{r}', str(r)) % DAYNAME.format(d='A%d' % r)
        ws['K%d' % r] = '=IF(OR(M{r}=0,C{r}=""),"",%s)'.replace('{r}', str(r)) % TV('C%d' % r)
        ws['L%d' % r] = '=IF(OR(M{r}=0,D{r}=""),"",%s)'.replace('{r}', str(r)) % TV('D%d' % r)
        auto(ws, 'A%d' % r, DATE); auto(ws, 'B%d' % r)
        for c in 'KL': ws['%s%d' % (c, r)].number_format = 'h:mm AM/PM'
        if d == 6:
            for c in 'AB': ws['%s%d' % (c, r)].font = F(11, True, NAVY)

def tail_rows(ws, hr, last, final_col):
    """سطور آخر الأسبوع (بلا سطر مجموع): جبسمبورد · إكرامية · المستحق النهائي · المدفوع · سطر فاصل"""
    s2, s3, s4, s5, sp = hr + 9, hr + 10, hr + 11, hr + 12, hr + 13; on = 'M%d=1' % hr
    # جبسمبورد
    label(ws, 'A%d' % s2, '=IF(%s,"ساعات تحميل الجبسمبورد بهالأسبوع ←","")' % on, 'A%d:D%d' % (s2, s2))
    edit(ws, 'E%d' % s2, HRS, True)
    ws.merge_cells('F%d:%s%d' % (s2, last, s2))
    ws['F%d' % s2] = '=IF(%s,"كل ساعة تحميل = "&TEXT($P$%d,"0.000")&" دينار (منها "&TEXT($N$%d,"0.000")&" محسوبة أصلاً بساعات اليوم)","")' % (on, hr, hr)
    auto(ws, 'F%d' % s2, None, False, 10, color=SUBT, fc='404040'); ws['F%d' % s2].alignment = C('right')
    # إكرامية (يدوي — هي اللي بتجبر المستحق)
    label(ws, 'A%d' % s3, '=IF(%s,"إكرامية الأسبوع (دينار) ←","")' % on, 'A%d:D%d' % (s3, s3))
    edit(ws, 'E%d' % s3, MONEY, True)
    ws.merge_cells('F%d:%s%d' % (s3, last, s3)); auto(ws, 'F%d' % s3, None, color=SUBT)
    # المستحق النهائي = أيام الأسبوع + فرق الجبسمبورد + الإكرامية (بلا جبر تلقائي)
    prev = col(openpyxl.utils.column_index_from_string(final_col) - 1)
    label(ws, 'A%d' % s4, '=IF(%s,"المستحق النهائي للأسبوع","")' % on, 'A%d:%s%d' % (s4, prev, s4), sz=13)
    ws['%s%d' % (final_col, s4)] = '=IF({on},ROUND(N($V${h})+N(E{g})*($P${h}-$N${h})+N(E{t}),3),"")'.format(on=on, h=hr, g=s2, t=s3)
    final_box(ws, '%s%d' % (final_col, s4))
    if final_col != last:
        ws.merge_cells('%s%d:%s%d' % (col(openpyxl.utils.column_index_from_string(final_col) + 1), s4, last, s4))
        auto(ws, '%s%d' % (col(openpyxl.utils.column_index_from_string(final_col) + 1), s4), None, color=SUBT)
    ws.row_dimensions[s4].height = 36
    # المدفوع
    label(ws, 'A%d' % s5, '=IF(%s,"المدفوع (دينار) ←","")' % on, 'A%d:D%d' % (s5, s5))
    edit(ws, 'E%d' % s5, MONEY, True)
    ws.merge_cells('F%d:%s%d' % (s5, last, s5))
    ws['F%d' % s5] = '=IF(%s,IF(N(E%d)>0,"✔ انقبض — الأسبوع مقفول","⏳ لسا ما انقبض"),"")' % (on, s5)
    auto(ws, 'F%d' % s5, None, True, 11, color=SUBT, fc='1E5631'); ws.row_dimensions[s5].height = 26
    ws.row_dimensions[sp].height = 10; ws['T%d' % sp] = 2
    # مخفي على سطر العنوان: U مجموع الساعات · V مجموع مستحق الأيام
    ws['U%d' % hr] = '=IF(%s,SUM(E%d:E%d),0)' % (on, hr + 2, hr + 8)
    return s2, s3, s4, s5

# ---- عمر ----
O_HEADS = ['التاريخ', 'اليوم', 'دخول', 'خروج', 'عدد الساعات', 'منها عادي', 'منها إضافي', 'المستحق (دينار)', 'ملاحظات']
O = hours_sheet('ساعات عمر', 'ساعات عمر المصري', NAVY, [14, 11, 13, 13, 12, 12, 12, 16, 30],
    'الدخول والخروج من القائمة (كل 5 دقائق، AM/PM). تحت كل أسبوع: ساعات التحميل والإكرامية والمدفوع (الخانات البرتقالية). '
    'أول ما تكتب «المدفوع» الأسبوع كله بيتلوّن أخضر = مقبوض ومقفول — أي تصحيح بسطر الأسبوع الجاي.', 'I')
OMR = []; TIME_RNG = []
for k in range(NWMAX):
    hr = B0 + BLOCK * k; s5 = hr + 12
    block_head(O, k, 'I', 'START_O', O_HEADS, 'E%d' % s5)
    O['N%d' % hr] = '=' + RATE('RT_OMR', 'L%d' % hr); O['O%d' % hr] = '=' + RATE('RT_OT', 'L%d' % hr); O['P%d' % hr] = '=' + RATE('RT_GYP', 'L%d' % hr)
    day_rows(O, hr, k); TIME_RNG.append('C%d:D%d' % (hr + 2, hr + 8))
    for d in range(7):
        r = hr + 2 + d
        O['E%d' % r] = '=IF(OR(M{r}=0,K{r}="",L{r}=""),"",MOD(L{r}-K{r},1))'.replace('{r}', str(r))
        O['F%d' % r] = ('=IF(OR(E{r}="",M{r}=0),"",IF(WEEKDAY(A{r},1)=6,E{r},'
                        'MIN(E{r},MAX(0,MIN(IF(L{r}<K{r},L{r}+1,L{r}),OMR_OT_AT)-K{r}))))').replace('{r}', str(r))
        O['G%d' % r] = '=IF(OR(E{r}="",M{r}=0),"",MAX(0,E{r}-F{r}))'.replace('{r}', str(r))
        O['H%d' % r] = '=IF(OR(E{r}="",M{r}=0),"",ROUND((F{r}*$N${h}+G{r}*$O${h})*24,3))'.replace('{r}', str(r)).replace('{h}', str(hr))
        for c in 'CD': edit(O, '%s%d' % (c, r))
        edit(O, 'I%d' % r); O['I%d' % r].alignment = C('right')
        for c in 'EFG': auto(O, '%s%d' % (c, r), DUR)
        auto(O, 'H%d' % r, MONEY, True, 12, fc=NAVY)
    O['V%d' % hr] = '=SUM(H%d:H%d)' % (hr + 2, hr + 8)
    s2, s3, s4, s5 = tail_rows(O, hr, 'I', 'H')
    OMR.append(dict(s2=s2, s3=s3, s4=s4, s5=s5, hr=hr, due='H%d' % s4, paid='E%d' % s5))
OEND = B0 + BLOCK * NWMAX - 1
dv(O, TIME_LIST, ' '.join(TIME_RNG))
O.conditional_formatting.add('C%d:D%d' % (B0, OEND), FormulaRule(
    formula=['AND($M%d=1,$K%d<>"",$L%d<>"",WEEKDAY($A%d,1)=6,OR($K%d<OMR_FRI_IN,$L%d>OMR_FRI_OUT))' % ((B0,) * 6)], fill=fill(WARN_BG)))

# ---- عبدالعزيز ----
Z_HEADS = ['التاريخ', 'اليوم', 'دخول', 'خروج', 'عدد الساعات', 'مستحق اليوم (دينار)', 'قبض؟', 'ملاحظات']
Z = hours_sheet('ساعات عبدالعزيز', 'ساعات عبدالعزيز — بيقبض نهاية كل يوم وبيتسوّى الجمعة', TEAL, [14, 11, 13, 13, 12, 16, 11, 34],
    'دينار للساعة أي يوم. لما يقبض آخر اليوم: «قبض؟ نعم» — المبلغ = مستحق اليوم، وبينخصم من تسوية الجمعة. '
    'تحت كل أسبوع: ساعات التحميل والإكرامية والمدفوع (الخانات البرتقالية).', 'H')
AZZ = []; TIME_RNG = []; YN_RNG = []
for k in range(NWMAX):
    hr = B0 + BLOCK * k; s5 = hr + 12
    block_head(Z, k, 'H', 'START_Z', Z_HEADS, 'E%d' % s5)
    Z['N%d' % hr] = '=' + RATE('RT_AZZ', 'L%d' % hr); Z['P%d' % hr] = '=' + RATE('RT_GYP', 'L%d' % hr)
    day_rows(Z, hr, k); TIME_RNG.append('C%d:D%d' % (hr + 2, hr + 8)); YN_RNG.append('G%d:G%d' % (hr + 2, hr + 8))
    for d in range(7):
        r = hr + 2 + d
        Z['E%d' % r] = '=IF(OR(M{r}=0,K{r}="",L{r}=""),"",MOD(L{r}-K{r},1))'.replace('{r}', str(r))
        Z['F%d' % r] = '=IF(OR(E%d="",M%d=0),"",%s)' % (r, r, RND('E%d*24*$N$%d' % (r, hr)))
        for c in 'CDGH': edit(Z, '%s%d' % (c, r))
        Z['H%d' % r].alignment = C('right')
        auto(Z, 'E%d' % r, DUR); auto(Z, 'F%d' % r, MONEY, True, 12, fc=NAVY)
    Z['V%d' % hr] = '=SUM(F%d:F%d)' % (hr + 2, hr + 8)
    Z['W%d' % hr] = '=IF(M%d=1,SUMIFS(F%d:F%d,G%d:G%d,"نعم"),0)' % (hr, hr + 2, hr + 8, hr + 2, hr + 8)   # قبض يومياً
    s2, s3, s4, s5 = tail_rows(Z, hr, 'H', 'F')
    AZZ.append(dict(s2=s2, s3=s3, s4=s4, s5=s5, hr=hr, due='F%d' % s4, paid='E%d' % s5))
ZEND = B0 + BLOCK * NWMAX - 1
dv(Z, TIME_LIST, ' '.join(TIME_RNG)); dv(Z, '"نعم,لا"', ' '.join(YN_RNG))

for ws, end, last in ((O, OEND, 'I'), (Z, ZEND, 'H')):
    rng = 'A%d:%s%d' % (B0, last, end)
    ws.conditional_formatting.add(rng, FormulaRule(formula=['AND($M%d=0,N($T%d)<>2)' % (B0, B0)], fill=fill(OUTM), font=Font(color=OUTM_FC), stopIfTrue=True))   # برّا الفترة
    ws.conditional_formatting.add(rng, FormulaRule(formula=['AND($R%d=CURWK,$Q%d=0)' % (B0, B0)], fill=fill(CURW), font=Font(bold=True, color=CURW_FC)))   # عنوان الأسبوع الحالي
    ws.conditional_formatting.add(rng, FormulaRule(formula=['AND($Q%d=1,N($T%d)=0)' % (B0, B0)], fill=fill(LOCKED)))   # مقبوض (المستحق النهائي بيضل غامق)
    ws.conditional_formatting.add('A%d:B%d' % (B0, end), FormulaRule(formula=['AND($S%d=CURWK,$R%d<>CURWK,ISNUMBER($A%d))' % (B0, B0, B0)], fill=fill(CURW_ROWS)))   # تاريخ أيام الأسبوع الحالي

# ═══════════════════ السلف ═══════════════════
# التاريخ (برتقالي): اختار من القائمة (اليوم مع تاريخه — آخر 14 يوم) أو اكتب التاريخ كامل أو اسم اليوم بس ← «اليوم» بيطلع لحاله
def DAYCOL(a, g):
    return '=IF({g}="","",{dn}&IF(OR(ISNUMBER({a}),LEN(TRIM({a}))>=10),"","  "&{t}))'.format(a=a, g=g, dn=DAYNAME.format(d=g), t=TXT(g))
A = wb.create_sheet('السلف')
title(A, 'السلف الأسبوعية — بتنخصم كاملة من راتب نفس الأسبوع', RUST, 'F')
for k, w in zip('ABCDEF', [22, 13, 16, 15, 34, 30]): A.column_dimensions[k].width = w
header(A, 3, ['الموظف', 'سلف هالسنة (دينار)', 'سلف الأسبوع المختار (دينار)', '', '', ''], 10, 32)
note(A, 'A7', 'التاريخ: اختار اليوم من القائمة (فيها اليوم وتاريخه)، أو اكتب التاريخ كامل، أو اسم اليوم بس — والثاني بيطلع لحاله. السلفة بتنخصم كاملة من صافي نفس الأسبوع، ولو أكثر من 60% من اللي اشتغله لحد يومها بيطلع تنبيه.', 'A7:F7', 10, '404040', 34)
header(A, 9, ['التاريخ', 'اليوم', 'الموظف', 'المبلغ (دينار)', 'البيان', 'التنبيه'], 11, 30)
AR0, AE = 10, 10 + ADV_ROWS - 1
for r in range(AR0, AE + 1):
    edit(A, 'A%d' % r, DIN); edit(A, 'C%d' % r); edit(A, 'D%d' % r, MONEY); edit(A, 'E%d' % r); A['E%d' % r].alignment = C('right')
    A['G%d' % r] = '=' + DVAL('A%d' % r)
    A['B%d' % r] = DAYCOL('A%d' % r, 'G%d' % r); auto(A, 'B%d' % r)
    A['H%d' % r] = '=' + WKOF('G%d' % r)
    A['I%d' % r] = '=' + MONOF('H%d' % r)
    A['J%d' % r] = ('=IF(OR(G{r}="",H{r}="",C{r}=""),"",IF(C{r}="عمر المصري",'
                    "SUMIFS('ساعات عمر'!$H$%d:$H$%d,'ساعات عمر'!$A$%d:$A$%d,\">=\"&(WEEK_ONE+7*(H{r}-1)),'ساعات عمر'!$A$%d:$A$%d,\"<=\"&G{r}),"
                    "SUMIFS('ساعات عبدالعزيز'!$F$%d:$F$%d,'ساعات عبدالعزيز'!$A$%d:$A$%d,\">=\"&(WEEK_ONE+7*(H{r}-1)),'ساعات عبدالعزيز'!$A$%d:$A$%d,\"<=\"&G{r})))"
                    ).replace('{r}', str(r)) % ((B0, OEND) * 3 + (B0, ZEND) * 3)
    A['F%d' % r] = ('=IF(A{r}="","",IF(G{r}="","⛔ التاريخ غلط — اختاره من القائمة أو اكتبه كامل 05/10/2026",IF(OR(H{r}<1,H{r}>NW),"⛔ التاريخ برّا السنة",IF(AND(N(D{r})>0,N(J{r})>0,D{r}>ADV_CAP*J{r}),'
                    '"⚠️ أكثر من "&TEXT(ADV_CAP,"0%")&" من شغله ("&TEXT(J{r},"0.000")&")",""))))').replace('{r}', str(r))
    auto(A, 'F%d' % r, None, True, 9)
dv(A, '"%s"' % ','.join(EMPS), 'C%d:C%d' % (AR0, AE)); dv_date(A, 'A%d:A%d' % (AR0, AE))
for c in 'GHIJ': A.column_dimensions[c].hidden = True
A.freeze_panes = 'A10'
for nm, c in (('ADV_D', 'G'), ('ADV_E', 'C'), ('ADV_AMT', 'D'), ('ADV_W', 'H'), ('ADV_M', 'I')):
    name(nm, "'السلف'!$%s$%d:$%s$%d" % (c, AR0, c, AE))
for i, nm in enumerate(EMPS):
    r = 4 + i; A['A%d' % r] = nm
    A['B%d' % r] = '=SUMIFS(ADV_AMT,ADV_E,A%d)' % r
    A['C%d' % r] = '=SUMIFS(ADV_AMT,ADV_E,A%d,ADV_W,WK)' % r
    for c in 'ABC': auto(A, '%s%d' % (c, r), MONEY if c != 'A' else None, True, 12)
A.conditional_formatting.add('A%d:F%d' % (AR0, AE), FormulaRule(formula=['LEFT($F%d,1)="⛔"' % AR0], fill=fill(ERR_BG), font=Font(color=ERR_FC, bold=True)))
A.conditional_formatting.add('A%d:F%d' % (AR0, AE), FormulaRule(formula=['LEFT($F%d,1)="⚠"' % AR0], fill=fill(WARN_BG), font=Font(color=WARN_FC, bold=True)))
A['A%d' % (AE + 2)] = '="ضل "&COUNTBLANK(A%d:A%d)&" سطر فاضي بالسجل"' % (AR0, AE); A['A%d' % (AE + 2)].font = F(9, False, '404040')

# ═══════════════════ القروض والديون (نفس شكل السلف) ═══════════════════
LTYPES = ['بضاعة', 'قرض كاش', 'دين قديم']
L = wb.create_sheet('القروض والديون')
title(L, 'القروض والديون — بتنخصم أقساط كل أسبوع', RUST, 'H')
for k, w in zip('ABCDEFGH', [22, 13, 16, 13, 15, 15, 16, 38]): L.column_dimensions[k].width = w
header(L, 3, ['الموظف', 'القروض والديون (مع الماضي)', 'انخصم أقساط لحد الأسبوع المختار', 'دفعات كاش سداد', 'الدين المتبقي', 'القسط الأسبوعي الحالي', 'قديش ضل', ''], 10, 34)
note(L, 'A7', 'التاريخ زي ورقة السلف (اليوم بيطلع لحاله). النوع: بضاعة · قرض كاش · دين قديم — والبضاعة بالبيان: الصنف والكمية. القسط بينخصم كل أسبوع لحاله لحد ما يخلص · تغيير القسط: سطر جديد فيه القسط بس · «دفعات كاش سداد» = اللي دفعه من جيبته.', 'A7:H7', 10, '404040', 34)
header(L, 9, ['التاريخ', 'اليوم', 'الموظف', 'النوع', 'المبلغ (دينار)', 'القسط الأسبوعي (دينار)', 'دفعات كاش سداد (دينار)', 'البيان — للبضاعة: الصنف والكمية'], 11, 34)
LR0, LE = 10, 10 + LOAN_ROWS - 1
for r in range(LR0, LE + 1):
    edit(L, 'A%d' % r, DIN); edit(L, 'C%d' % r); edit(L, 'D%d' % r)
    for c in 'EFG': edit(L, '%s%d' % (c, r), MONEY)
    edit(L, 'H%d' % r); L['H%d' % r].alignment = C('right')
    L['K%d' % r] = '=' + DVAL('A%d' % r)
    L['B%d' % r] = DAYCOL('A%d' % r, 'K%d' % r); auto(L, 'B%d' % r)
    L['J%d' % r] = r - LR0 + 1
    L['L%d' % r] = '=' + WKOF('K%d' % r); L['M%d' % r] = '=' + MONOF('L%d' % r)
dv(L, '"%s"' % ','.join(EMPS), 'C%d:C%d' % (LR0, LE)); dv_date(L, 'A%d:A%d' % (LR0, LE))
dv(L, '"%s"' % ','.join(LTYPES), 'D%d:D%d' % (LR0, LE), 'بضاعة · قرض كاش · دين قديم')
L.conditional_formatting.add('A%d:H%d' % (LR0, LE), FormulaRule(formula=['AND($A%d<>"",$K%d="")' % (LR0, LR0)], fill=fill(ERR_BG), font=Font(color=ERR_FC, bold=True)))
L.conditional_formatting.add('D%d:D%d' % (LR0, LE), FormulaRule(formula=['AND($E%d<>"",$D%d="")' % (LR0, LR0)], fill=fill(WARN_BG)))
L.conditional_formatting.add('H%d:H%d' % (LR0, LE), FormulaRule(formula=['AND($D%d="بضاعة",$H%d="")' % (LR0, LR0)], fill=fill(WARN_BG)))
for c in 'JKLM': L.column_dimensions[c].hidden = True
L.freeze_panes = 'A10'
for nm, c in (('LG_D', 'K'), ('LG_E', 'C'), ('LG_TYPE', 'D'), ('LG_DEBT', 'E'), ('LG_INST', 'F'), ('LG_REP', 'G'), ('LG_IDX', 'J'), ('LG_W', 'L'), ('LG_M', 'M')):
    name(nm, "'القروض والديون'!$%s$%d:$%s$%d" % (c, LR0, c, LE))
L['A%d' % (LE + 2)] = '="ضل "&COUNTBLANK(A%d:A%d)&" سطر فاضي بالسجل"' % (LR0, LE); L['A%d' % (LE + 2)].font = F(9, False, '404040')

def INST(emp, d, opening):
    m = 'SUMPRODUCT(MAX((LG_E="%s")*(LG_INST<>"")*(LG_D<=%s)*LG_IDX))' % (emp, d)
    return 'IF(%s=0,%s,INDEX(LG_INST,%s))' % (m, opening, m)

# ═══════════════════ التسوية ═══════════════════
S = wb['Sheet']; S.title = 'التسوية'
title(S, 'التسوية — قديش بدفع لكل موظف يوم الجمعة', NAVY, 'I')
for k, w in zip('ABCDEFGHI', [17, 14, 12, 12, 16, 13, 13, 14, 34]): S.column_dimensions[k].width = w
S['B3'] = 'السنة'; S['B3'].font = F(13, True, NAVY); S['B3'].alignment = C('left')
S['C3'] = '=YR'; auto(S, 'C3', '0', True, 14)
S.merge_cells('D3:H3'); S['D3'] = '="فيها "&NW&" أسبوع"'; S['D3'].font = F(10, False, '404040'); S['D3'].alignment = C('right')
S.row_dimensions[3].height = 30
S['B5'] = 'الأسبوع'; S['B5'].font = F(13, True, NAVY); S['B5'].alignment = C('left')
S.merge_cells('C5:F5'); S['C5'] = None; edit(S, 'C5', None, True, 12)
dv(S, '=WTITLE', 'C5', 'اتركها فاضية = الأسبوع الحالي لحاله · أو اختار أي أسبوع', 'الأسبوع')
S.merge_cells('G5:H5'); S['G5'] = '=IF(C5="","← فاضية = الأسبوع الحالي","← امسحها ترجع للحالي")'
S['G5'].font = F(10, False, '404040'); S['G5'].alignment = C('right'); S.row_dimensions[5].height = 30
S['N5'] = '=IF(C5="",CURWK,IFERROR(MATCH(C5,WTITLE,0),CURWK))'; S['N5'].font = F(8, False, 'FFFFFF')
name('WK', "'التسوية'!$N$5")
S['A7'] = '=INDEX(WTITLE,WK)'
S['A7'].font = F(14, True, NAVY); S.merge_cells('A7:H7'); S['A7'].alignment = C('right'); S.row_dimensions[7].height = 26
COLS = ['المستحق', 'سلف', 'قسط الدين', 'الصافي للدفع', 'المدفوع', 'الباقي له', 'الدين المتبقي']
header(S, 8, ['الموظف'] + COLS + ['تنبيه'], 11, 30)

# ═══════════════════ دفتر الأسابيع ═══════════════════
W = wb.create_sheet('دفتر الأسابيع'); WS_ = "'دفتر الأسابيع'"
title(W, 'دفتر الأسابيع — كل أسبوع بالسنة لكل موظف (محسوب، ما في كتابة هون)', NAVY, 'L')
for k, w in zip('ABCDEFGHIJKL', [9, 13, 13, 13, 11, 11, 13, 11, 12, 13, 12, 7]): W.column_dimensions[k].width = w
note(W, 'A3', '«الصافي للدفع» = الباقي من الأسبوع الماضي + المستحق − السلف − القسط (− اللي قبضه يومياً لعبدالعزيز). «الباقي له» = الصافي − المدفوع، وبينتقل لحاله للأسبوع الجاي.', 'A3:L3', 9, h=28)
LKEYS = ['WK', 'SAT', 'FRI', 'DUE', 'ADV', 'INST', 'NET', 'PAID', 'REM', 'DEBT', 'DAILY', 'MON']   # A..L
def LN(pre, key): return '%s_%s' % (pre, key)
def week_ledger(top, emp, color, rows, is_azz):
    hs = "'ساعات عبدالعزيز'" if is_azz else "'ساعات عمر'"
    start = 'START_Z' if is_azz else 'START_O'; pre = 'LZ' if is_azz else 'LO'
    W['A%d' % top] = emp; W['A%d' % top].font = F(13, True, color); W.merge_cells('A%d:L%d' % (top, top)); W['A%d' % top].alignment = C('right')
    header(W, top + 1, ['الأسبوع', 'السبت', 'الجمعة'] + COLS + (['قبض يومياً'] if is_azz else ['']) + ['شهر'], 10, 30)
    r0 = top + 2
    for k in range(NWMAX):
        r = r0 + k; w = k + 1; b = rows[k]
        act = '%s!$M$%d=1' % (hs, b['hr']); fri = 'C%d' % r
        W['A%d' % r] = '=IF(%d<=NW,%d,"")' % (w, w)
        W['B%d' % r] = '=IF(%d<=NW,WEEK_ONE+%d,"")' % (w, 7 * k)
        W['C%d' % r] = '=IF(%d<=NW,WEEK_ONE+%d,"")' % (w, 7 * k + 6)
        W['D%d' % r] = '=IF(%s,N(%s!%s),"")' % (act, hs, b['due'])
        W['E%d' % r] = '=IF(%s,SUMIFS(ADV_AMT,ADV_E,"%s",ADV_W,%d),"")' % (act, emp, w)
        W['K%d' % r] = ('=IF(%s,N(%s!W%d),"")' % (act, hs, b['hr'])) if is_azz else '=0'
        owed = '{o}+SUMIFS(LG_DEBT,LG_E,"{e}",LG_D,"<="&{f})-SUMIFS(LG_REP,LG_E,"{e}",LG_D,"<="&{f})'.format(o=OP('D', emp), e=emp, f=fri)
        prev_inst = '0' if k == 0 else 'SUM(F%d:F%d)' % (r0, r - 1)
        prevH = 'IF(%d=%s,%s,N(I%d))' % (w, start, OP('C', emp), r - 1) if k > 0 else 'IF(%d=%s,%s,0)' % (w, start, OP('C', emp))
        room = 'N(D{r})-N(E{r})-NET_FLOOR*N(D{r})'.replace('{r}', str(r))   # القبض اليومي مش خصم
        inst_w, left = INST(emp, fri, OP('E', emp)), '%s-%s' % (owed, prev_inst)
        W['F%d' % r] = ('=IF(%s,IF(N(D%d)=0,0,MAX(0,IF(INST_MODE="تخفيض تلقائي",FLOOR(MIN(%s,%s,INST_CAP*N(D%d),%s),0.25),'
                        'MIN(%s,%s)))),"")') % (act, r, inst_w, left, r, room, inst_w, left)
        W['G%d' % r] = '=IF(%s,%s+N(D{r})-N(E{r})-N(F{r})-N(K{r}),"")'.replace('{r}', str(r)) % (act, prevH)
        W['H%d' % r] = '=IF(%s,N(%s!%s),"")' % (act, hs, b['paid'])
        W['I%d' % r] = '=IF(%s,G%d-H%d,"")' % (act, r, r)
        W['J%d' % r] = '=IF(%s,MAX(0,%s-SUM(F%d:F%d)),"")' % (act, owed, r0, r)
        W['L%d' % r] = '=IF(C%d="","",MONTH(C%d))' % (r, r)
        for c in 'ABCDEFGHIJKL': auto(W, '%s%d' % (c, r), MONEY)
        W['A%d' % r].number_format = '0'; W['B%d' % r].number_format = DATE; W['C%d' % r].number_format = DATE; W['L%d' % r].number_format = '0'
        W['G%d' % r].font = F(11, True, NAVY); W['J%d' % r].font = F(10, True, RUST)
        if not is_azz: W['K%d' % r].font = F(8, False, AUTO)
    tr = r0 + NWMAX
    W['A%d' % tr] = 'مجموع السنة'; W.merge_cells('A%d:C%d' % (tr, tr))
    for c in 'DEFGHK': W['%s%d' % (c, tr)] = '=SUM(%s%d:%s%d)' % (c, r0, c, tr - 1)
    for c in 'ABCDEFGHIJKL': auto(W, '%s%d' % (c, tr), MONEY, True, color=SUBT)
    W.conditional_formatting.add('A%d:L%d' % (r0, tr - 1), FormulaRule(formula=['$A%d=WK' % r0], fill=fill(CURW_ROWS), font=Font(bold=True), stopIfTrue=True))
    W.conditional_formatting.add('A%d:L%d' % (r0, tr - 1), FormulaRule(formula=['AND($A%d<>"",N($H%d)>0)' % (r0, r0)], fill=fill(LOCKED)))
    for c, key in zip('ABCDEFGHIJKL', LKEYS):
        name(LN(pre, key), "%s!$%s$%d:$%s$%d" % (WS_, c, r0, c, tr - 1))
    return r0, tr
OR0, OTR = week_ledger(5, EMPS[0], NAVY, OMR, False)
ZR0, ZTR = week_ledger(OTR + 3, EMPS[1], TEAL, AZZ, True)
W.freeze_panes = 'A5'

# صفوف الأسبوع المختار بالتسوية
for i, (nm, pre) in enumerate([(EMPS[0], 'LO'), (EMPS[1], 'LZ')]):
    r = 9 + i; S['A%d' % r] = nm; auto(S, 'A%d' % r, None, True, 12)
    for c, key in zip('BCDEFGH', LKEYS[3:10]):
        S['%s%d' % (c, r)] = '=N(INDEX(%s,WK))' % LN(pre, key)
        auto(S, '%s%d' % (c, r), MONEY, False, 12)
    final_box(S, 'E%d' % r)
    S['H%d' % r].font = F(11, True, RUST); S.row_dimensions[r].height = 40
    S['I%d' % r] = ('=IF(AND(D{r}>0.0005,D{r}>INST_CAP*B{r}+0.0005),"⚠️ القسط "&TEXT(D{r},"0.000")&" أكثر من "&TEXT(INST_CAP,"0%")&" من المستحق","")'
                    '&IF(AND(B{r}>0,B{r}-C{r}-D{r}<NET_FLOOR*B{r}-0.0005),"  ⚠️ الصافي "&TEXT((B{r}-C{r}-D{r})/B{r},"0%")&" من المستحق (أقل من "&TEXT(NET_FLOOR,"0%")&")","")'
                    '&IF(AND(INST_MODE="تخفيض تلقائي",H{r}>0.0005,D{r}<\'القروض والديون\'!F{lr}-0.0005),"  ↓ القسط انخصم "&TEXT(D{r},"0.000")&" بدل "&TEXT(\'القروض والديون\'!F{lr},"0.000"),"")'
                    ).replace('{r}', str(r)).replace('{lr}', str(4 + i))
    auto(S, 'I%d' % r, None, True, 10, fc=WARN_FC); S['I%d' % r].alignment = C('right')
label(S, 'A11', 'الكاش اللي بتطلعه الجمعة (الاثنين)', 'A11:D11', 12)
S['E11'] = '=E9+E10'; final_box(S, 'E11'); S.row_dimensions[11].height = 34
# أزرار الانتقال (سطور 13–15): زر أزرق — اضغط عليه بيوديك عالأسبوع المختار
BTNS = [
    (13, '"ساعات عمر"', "'ساعات عمر'", '%d+%d*(WK-1)' % (B0, BLOCK)),
    (14, '"ساعات عبدالعزيز"', "'ساعات عبدالعزيز'", '%d+%d*(WK-1)' % (B0, BLOCK)),
    (15, '"دفتر الأسابيع"', "'دفتر الأسابيع'", '%d+WK-1' % OR0),
]
for r, nm_txt, sh, row_expr in BTNS:
    S.merge_cells('A%d:E%d' % (r, r))
    S['A%d' % r] = '=HYPERLINK("#%s!A"&(%s),"▶  اضغط هون: افتح "&%s&" — الأسبوع "&WK)' % (sh.replace('"', '""'), row_expr, nm_txt)
    cellfmt(S['A%d' % r], BTN_BG, True, sz=13, fc=BTN_FC, h='center'); S['A%d' % r].border = TBOX
    for c in 'BCDE': S['%s%d' % (c, r)].border = TBOX
    S.row_dimensions[r].height = 32
note(S, 'A17', 'الأزرار الزرقا فوق (سطور 13–15) بتنقلك عالأسبوع المختار مباشرة. الإكرامية والمدفوع بتنكتبوا تحت كل أسبوع بورقة الساعات. «الباقي له» بينتقل لحاله للأسبوع الجاي.', 'A17:H17', 9, h=28)
S.freeze_panes = 'A7'

# نهاية السنة
B['A11'] = 'نهاية السنة — هالأرقام بتنكتب بورقة «السنة» بملف السنة الجاية (مع أسبوع بداية = 1)'; B['A11'].font = F(13, True, RUST); B.merge_cells('A11:G11'); B['A11'].alignment = C('right')
header(B, 12, ['الموظف', 'آخر أسبوع', 'مستحق له لسا ما قبضه', 'دين متبقي عليه', 'القسط الأسبوعي', 'الوضع', ''], 10, 34); B.merge_cells('F12:G12')
for i, (nm, pre) in enumerate([(EMPS[0], 'LO'), (EMPS[1], 'LZ')]):
    r = 13 + i
    B['A%d' % r] = nm; B['B%d' % r] = '="الأسبوع "&NW&" ("&%s&")"' % TXT('INDEX(WFRI,NW)')
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
title(T, 'عمال المهام — بيندفعلهم مباشرة', PLUM, 'F')
for k, w in zip('ABCDEF', [22, 13, 20, 38, 16, 28]): T.column_dimensions[k].width = w
note(T, 'A3', 'كل مهمة سطر: التاريخ (اليوم بيطلع لحاله — أو اختار اليوم من القائمة)، العامل، شو عمل، وقديش أخذ. بتنحسب على شهر جمعة أسبوعها.', 'A3:F3', h=30)
header(T, 5, ['التاريخ', 'اليوم', 'اسم العامل', 'المهمة', 'المبلغ المدفوع (دينار)', 'ملاحظات'], 11, 30)
TE = 6 + TASK_ROWS - 1
for r in range(6, TE + 1):
    edit(T, 'A%d' % r, DIN); edit(T, 'C%d' % r); edit(T, 'D%d' % r); edit(T, 'E%d' % r, MONEY); edit(T, 'F%d' % r)
    for c in 'DF': T['%s%d' % (c, r)].alignment = C('right')
    T['H%d' % r] = '=' + DVAL('A%d' % r)
    T['B%d' % r] = DAYCOL('A%d' % r, 'H%d' % r); auto(T, 'B%d' % r)
    T['G%d' % r] = '=IF(AND(C{r}<>"",COUNTIF(C$6:C{r},C{r})=1),MAX(G$5:G{p})+1,"")'.replace('{r}', str(r)).replace('{p}', str(r - 1))
    T['I%d' % r] = '=' + WKOF('H%d' % r); T['J%d' % r] = '=' + MONOF('I%d' % r)
dv_date(T, 'A6:A%d' % TE)
for c in 'GHIJ': T.column_dimensions[c].hidden = True
T.conditional_formatting.add('A6:F%d' % TE, FormulaRule(formula=['AND($A6<>"",$H6="")'], fill=fill(ERR_BG), font=Font(color=ERR_FC, bold=True)))
label(T, 'D%d' % (TE + 1), 'مجموع السنة')
T['E%d' % (TE + 1)] = '=SUM(E6:E%d)' % TE; auto(T, 'E%d' % (TE + 1), MONEY, True, 12, color=SUBT)
T.freeze_panes = 'A6'
T['A%d' % (TE + 3)] = '="ضل "&COUNTBLANK(A6:A%d)&" سطر فاضي بالسجل"' % TE; T['A%d' % (TE + 3)].font = F(9, False, '404040')
for nm, c in (('TASK_AMT', 'E'), ('TASK_B', 'C'), ('TASK_W', 'I'), ('TASK_M', 'J'), ('TASK_G', 'G')):
    name(nm, "'عمال المهام'!$%s$6:$%s$%d" % (c, c, TE))

# ═══════════════════ ملخص الأيدي العاملة ═══════════════════
M = wb.create_sheet('ملخص الأيدي العاملة')
title(M, 'ملخص الأيدي العاملة — شهر بشهر (حسب جمعة الأسبوع) ووضع كل موظف', NAVY, 'L')
for k, w in zip('ABCDEFGHIJKL', [17, 12, 12, 8, 11, 12, 10, 11, 13, 12, 11, 12]): M.column_dimensions[k].width = w
M['A3'] = 'قديش دفعت أيدي عاملة كل شهر (كاش طالع فعلياً) — الشهر = الأسابيع اللي جمعتها فيه'; M['A3'].font = F(13, True, NAVY); M.merge_cells('A3:L3'); M['A3'].alignment = C('right')
header(M, 5, ['الشهر', 'أول سبت', 'آخر جمعة', 'عدد الجمع', 'رواتب عمر', 'عبدالعزيز (يومي + جمعة)', 'سلف كاش', 'عمال المهام', 'المجموع كاش', 'كلفة الشغل', 'بضاعة أخذوها', 'قروض كاش انعطت'], 9, 40)
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
    M['J%d' % r] = '=SUMIFS(LO_DUE,LO_MON,%d)+SUMIFS(LZ_DUE,LZ_MON,%d)+H%d' % (m, m, r)
    M['K%d' % r] = '=SUMIFS(LG_DEBT,LG_TYPE,"بضاعة",LG_M,%d)' % m
    M['L%d' % r] = '=SUMIFS(LG_DEBT,LG_TYPE,"قرض كاش",LG_M,%d)' % m
    for c in 'ABCDEFGHIJKL': auto(M, '%s%d' % (c, r), MONEY if c not in 'ABCD' else None)
    M['B%d' % r].number_format = DATE; M['C%d' % r].number_format = DATE; M['D%d' % r].number_format = '0'
    M['A%d' % r].alignment = C('right'); M['I%d' % r].font = F(11, True, NAVY)
MT = 18
M['A%d' % MT] = 'مجموع السنة'
M['B%d' % MT] = '=WEEK_ONE'; M['C%d' % MT] = '=WEEK_ONE+7*NW-1'
for c in 'DEFGHIJKL': M['%s%d' % (c, MT)] = '=SUM(%s6:%s17)' % (c, c)
for c in 'ABCDEFGHIJKL': auto(M, '%s%d' % (c, MT), MONEY if c not in 'ABCD' else None, True, color=SUBT)
M['B%d' % MT].number_format = DATE; M['C%d' % MT].number_format = DATE; M['D%d' % MT].number_format = '0'
final_box(M, 'I%d' % MT, 13)
M.conditional_formatting.add('A6:L17', FormulaRule(formula=['INDEX(WMON,WK)=ROW()-5'], fill=fill(CURW_ROWS)))
note(M, 'A%d' % (MT + 1), '«كلفة الشغل» = المستحق + المهام (الأجر كامل، سواء انقبض كاش أو بضاعة). «المجموع كاش» = اللي طلع من الدرج فعلياً: السلف بتنعدّ · البضاعة والأقساط لا. الأشهر اللي فيها 5 جمع بتطلع أعلى.', 'A%d:L%d' % (MT + 1, MT + 1), 9, h=30)

E0 = MT + 3
M['A%d' % E0] = '="وضع كل موظف لحد الأسبوع "&WK&":  "&INDEX(WTITLE,WK)'; M['A%d' % E0].font = F(13, True, NAVY); M.merge_cells('A%d:L%d' % (E0, E0)); M['A%d' % E0].alignment = C('right')
header(M, E0 + 1, ['الموظف', 'ساعات', 'المستحق', 'قبض كاش', '+ سلف', '+ أقساط', '+ الباقي له', 'المجموع', 'الدين المتبقي', 'النتيجة'], 9, 36)
for i, (nm, pre, hs, rows) in enumerate([(EMPS[0], 'LO', "'ساعات عمر'", OMR), (EMPS[1], 'LZ', "'ساعات عبدالعزيز'", AZZ)]):
    r = E0 + 2 + i; wk = LN(pre, 'WK')
    M['A%d' % r] = nm
    M['B%d' % r] = '=' + '+'.join('IF(%d<=WK,N(%s!U%d),0)' % (k + 1, hs, rows[k]['hr']) for k in range(NWMAX))
    M['C%d' % r] = '=%s+SUMIFS(%s,%s,"<="&WK)' % (OP('C', nm), LN(pre, 'DUE'), wk)
    M['D%d' % r] = '=SUMIFS(%s,%s,"<="&WK)+SUMIFS(%s,%s,"<="&WK)' % (LN(pre, 'PAID'), wk, LN(pre, 'DAILY'), wk)
    M['E%d' % r] = '=SUMIFS(%s,%s,"<="&WK)' % (LN(pre, 'ADV'), wk)
    M['F%d' % r] = '=SUMIFS(%s,%s,"<="&WK)' % (LN(pre, 'INST'), wk)
    M['G%d' % r] = '=N(INDEX(%s,WK))' % LN(pre, 'REM')
    M['H%d' % r] = '=D{r}+E{r}+F{r}+G{r}'.replace('{r}', str(r))
    M['I%d' % r] = '=N(INDEX(%s,WK))' % LN(pre, 'DEBT')
    M['J%d' % r] = '=IF(ABS(H{r}-C{r})<0.0005,"✔ مطابق","⚠️ فرق "&TEXT(H{r}-C{r},"0.000"))'.replace('{r}', str(r))
    auto(M, 'A%d' % r, None, True, 11); auto(M, 'B%d' % r, DUR)
    for c in 'CDEFGHI': auto(M, '%s%d' % (c, r), MONEY, c in 'CH', 11)
    M['I%d' % r].font = F(11, True, RUST); auto(M, 'J%d' % r, None, True, 10, fc='1E5631')
for i, nm in enumerate(EMPS):
    r = E0 + 4 + i; src = E0 + 2 + i
    M['A%d' % r] = '="• "&A{s}&": "&IF(G{s}>0.0005,"المحل لسا مدين إله بـ "&TEXT(G{s},"0.000"),IF(G{s}<-0.0005,"قبض زيادة "&TEXT(-G{s},"0.000"),"مخالص — ما إله شي"))&IF(I{s}>0.0005,"  ·  وعليه دين "&TEXT(I{s},"0.000"),"")'.replace('{s}', str(src))
    M.merge_cells('A%d:L%d' % (r, r)); M['A%d' % r].font = F(11, True, RUST); M['A%d' % r].alignment = C('right')
note(M, 'A%d' % (E0 + 6), 'المطابقة: (مستحق له من قبل + كل المستحق) = قبض كاش + سلف + أقساط + الباقي له — ولا دينار مكرّر.', 'A%d:L%d' % (E0 + 6, E0 + 6), 9)

T0 = E0 + 8
M['A%d' % T0] = 'عمال المهام — السنة، وشهر بتختاره'; M['A%d' % T0].font = F(13, True, PLUM); M.merge_cells('A%d:F%d' % (T0, T0)); M['A%d' % T0].alignment = C('right')
M['H%d' % T0] = 'الشهر:'; M['H%d' % T0].font = F(11, True, PLUM); M['H%d' % T0].alignment = C('left')
M['I%d' % T0] = '=INDEX(WMON,WK)'; edit(M, 'I%d' % T0, '0', True, 12); dv(M, "='القوائم'!$D$1:$D$12", 'I%d' % T0, 'رقم الشهر 1–12')
M.merge_cells('J%d:L%d' % (T0, T0)); M['J%d' % T0] = '=IFERROR(%s&"  -  "&%s,"")' % (TXT('INDEX(WSAT,MATCH(I%d,WMON,0))' % T0), TXT('INDEX(WFRI,MATCH(I%d,WMON,0)+COUNTIF(WMON,I%d)-1)' % (T0, T0)))
M['J%d' % T0].font = F(9, False, '404040'); M['J%d' % T0].alignment = C('right')
header(M, T0 + 1, ['#', 'العامل', 'مهام السنة', 'قبض بالسنة', 'مهام الشهر', 'قبض بالشهر'], 10, 28)
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
    ('1', 'ملف واحد للسنة كاملة. الأسبوع من السبت لمساء الجمعة، ومساء الجمعة الموظف بياخذ أجرته. الأسبوع كامل بينسب للشهر اللي فيه جمعته — ما بينقسم أبداً بين شهرين.'),
    ('2', 'الأسبوع 1 = اللي جمعته أول جمعة بيناير. السنة فيها 52 أو 53 أسبوع (2027 فيها 53).'),
    ('الألوان', None),
    ('3', 'برتقالي = خانة بتعبّيها إنت · رمادي = محسوبة · أسود = المستحق النهائي والصافي للدفع · أخضر = أسبوع مقبوض ومقفول · أزرق = الأسبوع الحالي · أحمر = خطأ · أصفر = تحذير.'),
    ('أول السنة', None),
    ('4', 'بورقة «السنة»: اختر السنة من القائمة (كل الملف بيتبرمج عليها)، وأسبوع البداية، ومستحق له من قبل، ودين عليه، والقسط — من «نهاية السنة» بملف السنة الماضية.'),
    ('كل يوم', None),
    ('5', 'بورقة الساعات: الأسبوع الحالي عنوانه أزرق غامق (أو من الأزرار بالتسوية). اختر الدخول والخروج من القائمة (كل 5 دقائق) أو اكتبه مثل 7:30 AM. عدد الساعات بيطلع ساعة:دقائق (11:30). عبدالعزيز: آخر اليوم «قبض؟ نعم».'),
    ('6', 'سلفة كاش: سطر بورقة «السلف» — التاريخ: اختار اليوم من القائمة أو اكتب التاريخ كامل (05/10/2026) أو اسم اليوم بس، و«اليوم» بيطلع لحاله. قرض أو بضاعة: سطر بورقة «القروض والديون». عامل مهمة: سطر بورقة «عمال المهام».'),
    ('آخر الأسبوع (الجمعة)', None),
    ('7', 'تحت الأسبوع بورقة الساعات: ساعات تحميل الجبسمبورد، والإكرامية (بتنجمع مع المستحق)، و«المستحق النهائي» = الأيام + فرق التحميل + الإكرامية (بلا جبر تلقائي — الإكرامية هي اللي بتجبره). ولما تدفع اكتب «المدفوع» — الأسبوع كله بيتلوّن أخضر = مقفول.'),
    ('8', 'ورقة «التسوية» بتفتح على الأسبوع الحالي: «الصافي للدفع» لكل موظف = الباقي من الأسبوع الماضي + المستحق − السلف − القسط (عبدالعزيز: − اللي قبضه يومياً). وتحتهم الكاش الكلي، وتحته 3 أزرار زرقا بتوديك عالأسبوع بورقة كل موظف وبدفتر الأسابيع.'),
    ('السقوف والأسعار', None),
    ('9', 'السلفة فوق 60% من اللي اشتغله لحد يومها = تنبيه أصفر. القسط بينخصم زي ما كتبته؛ لو أكثر من 25% من المستحق أو الصافي أقل من 65% بيطلع تنبيه بالتسوية (أو اختار «تخفيض تلقائي» بالقواعد). جدول الأسعار «ساري من» بورقة «القواعد» — تغيير السعر ما بيغيّر الأسابيع القديمة.'),
    ('آخر الشهر', None),
    ('10', 'ورقة «ملخص الأيدي العاملة»: 12 شهر — الكاش، كلفة الشغل، البضاعة، القروض. ارفع الملف لكلود: البضاعة بتنزل فواتير بيع على أودو باسم الموظف، والأقساط تحصيل — مسودة أول وما بينرحّل إلا بـ«رحّل».'),
    ('الحماية', None),
    ('11', 'الخلايا البرتقالية بس بتنكتب. الباقي محمي بلا كلمة سر (مراجعة ← إلغاء حماية الورقة إذا لزم). خلّي الملف على الدرايف عشان النسخ تنحفظ لحالها.'),
]
r = 3
for a, b in steps:
    if b is None: I['B%d' % r] = a; I['B%d' % r].font = F(12, True, TEAL); r += 1; continue
    I['A%d' % r] = a; I['A%d' % r].font = F(11, True); I['A%d' % r].alignment = C()
    I['B%d' % r] = b; I['B%d' % r].font = F(11); I['B%d' % r].alignment = C('right'); I.row_dimensions[r].height = 36; r += 1

# ═══════════════════ الترتيب والحماية ═══════════════════
order = ['التسوية', 'ساعات عمر', 'ساعات عبدالعزيز', 'السلف', 'القروض والديون', 'عمال المهام', 'دفتر الأسابيع',
         'السنة', 'القواعد', 'ملخص الأيدي العاملة', 'التعليمات', 'القوائم']
wb._sheets = [wb[n] for n in order]
wb.active = 0
for ws in wb.worksheets:
    for ref in INPUTS.get(ws.title, []):
        ws[ref].protection = Protection(locked=False)
    ws.protection.sheet = True
    ws.protection.formatColumns = False; ws.protection.formatRows = False

# 🔴 فحص إجباري: ولا اسم معرّف بيشبه عنوان خلية بإكسل (A1 · SAT1 · XFD100 · R/C) — هاد اللي عمل سنة 1900
def col_num(s):
    n = 0
    for ch in s.upper(): n = n * 26 + (ord(ch) - 64)
    return n
for nm in list(wb.defined_names.keys()):
    m = re.fullmatch(r'([A-Za-z]{1,3})([0-9]+)', nm)
    assert not (m and col_num(m.group(1)) <= 16384), 'اسم معرّف بيتلخبط مع عنوان خلية: %s' % nm
    assert nm.upper() not in ('R', 'C') and not re.fullmatch(r'[RrCc][0-9]*([RrCc][0-9]*)?', nm), 'اسم R1C1: %s' % nm
wb.save(OUT); print('saved', OUT)
