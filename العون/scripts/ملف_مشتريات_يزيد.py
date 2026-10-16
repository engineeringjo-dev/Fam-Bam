# -*- coding: utf-8 -*-
"""ملف عمل مشتريات يزيد العمري — بيتعبّى عند صاحب المحل وبيرجع (صاحب المحل 08/10)
- 18 فاتورة شراء كلها بتاريخ 28/03/2024 (فاتورة 18 = «الورق الفرط — دفتر سافيتو الصغير»)
- الخانات الصفرا بيعبّيها صاحب المحل: الإفرادي الناقص · الصنف المعتمد · ملاحظاته · الإجماليات معادلات
- الدفعات بتتشيّك بعدين (قراره)"""
import json, datetime, openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
d = json.load(open('drive/مسودات/مشتريات يزيد.json', encoding='utf-8'))
INV, JARD = d['inv'], d['jard']
SALES, PAID = 2778.750, 850.000
DATE = datetime.date(2024, 3, 28)
OUT = 'drive/مشتريات_يزيد_للتعبئة.xlsx'
wb = openpyxl.Workbook()
thin = Side(style='thin', color='999999'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
FILL = lambda c: PatternFill('solid', fgColor=c)
YEL, RED, GRN, ORG, BLU, PUR = 'FFF2CC', 'FFC7CE', 'E2EFDA', 'FCE4D6', 'DDEBF7', 'E4DFEC'
def head(ws, row, cols, color='1F4E6B'):
    for i, (h, w) in enumerate(cols, 1):
        c = ws.cell(row, i, h); c.fill = FILL(color); c.font = Font(b=True, color='FFFFFF')
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True); c.border = B
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[row].height = 32
def title(ws, t, sub, n):
    ws.sheet_view.rightToLeft = True
    ws.cell(1, 1, t).font = Font(b=True, size=14, color='1F4E6B'); ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n)
    ws.cell(2, 1, sub).font = Font(size=10, color='555555'); ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n)
def cell(ws, r, c, v, fmt=None, fill=None, bold=False, wrap=False):
    x = ws.cell(r, c, v); x.border = B
    x.alignment = Alignment(horizontal='right' if isinstance(v, str) and not str(v).startswith('=') else 'center', vertical='center', wrap_text=wrap)
    if fmt: x.number_format = fmt
    if fill: x.fill = FILL(fill)
    if bold: x.font = Font(b=True)
    return x
GRADE = lambda s: ('✓ قوي', GRN) if s >= 0.85 else ('؟ ضعيف', YEL) if s >= 0.6 else ('🆕 مش معرّف', ORG)

# ── البنود (كل الفواتير 1–18 بشيت واحد)
rows = []
for v in INV:
    for x in v['rows']:
        tags = [w for w in ('؟؟', 'قديم', 'سيء', 'تالف') if w in x['b']] + ([x['note']] if x.get('note') else [])
        rows.append((v['n'], x['b'], x['q'], x['p'] if x['ok'] else None, x['m'], ' · '.join(tags)))
for x in JARD:
    rows.append((18, x['b'], x['q'], x.get('p'), x['m'], ' · '.join(filter(None, [x['sec'], x['note']]))))
ws = wb.active; ws.title = 'البنود'
title(ws, 'مشتريات من يزيد العمري — 18 فاتورة بتاريخ 28/03/2024', 'عبّي الخانات الصفرا بس: الإفرادي الناقص · الصنف المعتمد (الكود) · ملاحظاتك — والإجمالي بيتحسب لحاله', 13)
cols = [('الفاتورة', 9), ('#', 5), ('البند (زي الورقة)', 34), ('الكمية', 9), ('الإفرادي', 11), ('الإجمالي', 12), ('الكود المقترح', 11),
        ('الصنف المقترح بالكتالوج', 36), ('سعر البيع بالكتالوج', 11), ('التطابق', 13), ('✍️ الصنف المعتمد (كود)', 15), ('✍️ ملاحظاتك', 30), ('ملاحظات الورقة', 34)]
head(ws, 4, cols); ws.freeze_panes = 'D5'
r = 5; n_prev = None; k = 0
for n, b, q, p, m, note in rows:
    k = k + 1 if n == n_prev else 1; n_prev = n
    s, c = m if m and m[1] else (0, None)
    g, gc = GRADE(s) if c else ('🆕 مش معرّف', ORG)
    sug = c if c and s >= 0.6 else None
    cell(ws, r, 1, n); cell(ws, r, 2, k); cell(ws, r, 3, b); cell(ws, r, 4, q, '0.###', None if q is not None else RED)
    cell(ws, r, 5, p, '0.000', None if p else YEL)
    cell(ws, r, 6, '=IF(OR(D{0}="",E{0}=""),"",ROUND(D{0}*E{0},3))'.format(r), '0.000', ORG)
    cell(ws, r, 7, sug['default_code'] if sug else None); cell(ws, r, 8, sug['name'] if sug else None)
    cell(ws, r, 9, sug['lst_price'] if sug else None, '0.000'); cell(ws, r, 10, g, fill=gc)
    cell(ws, r, 11, None, fill=YEL); cell(ws, r, 12, None, fill=PUR); cell(ws, r, 13, note, wrap=True)
    r += 1
last = r - 1
ws.auto_filter.ref = 'A4:M%d' % last
dv = DataValidation(type='decimal', operator='greaterThanOrEqual', formula1='0', showErrorMessage=True, errorTitle='رقم بس', error='اكتب السعر رقم')
dv.add('E5:E%d' % last); ws.add_data_validation(dv)
# الإفرادي الناقص أصفر لحد ما يتعبّى
ws.conditional_formatting.add('E5:E%d' % last, FormulaRule(formula=['AND($C5<>"",$E5="")'], fill=FILL(YEL)))

# ── الملخّص (معادلات من شيت البنود)
S = wb.create_sheet('الملخّص', 0)
title(S, 'ملخّص مشتريات يزيد — بيتحدّث لحاله من شيت «البنود»', 'التاريخ لكل الفواتير 28/03/2024 · فاتورة 18 = الورق الفرط (دفتر سافيتو الصغير) · الدفعات بتتشيّك بعدين', 7)
head(S, 4, [('الفاتورة', 12), ('التاريخ', 13), ('عدد البنود', 11), ('بلا سعر', 10), ('المجموع (من البنود)', 17), ('مجموع الورقة', 15), ('الفرق', 12)])
paper = {v['n']: v['paper'] for v in INV}
for i in range(1, 19):
    rr = 4 + i
    cell(S, rr, 1, 'فاتورة %d' % i); cell(S, rr, 2, DATE, 'dd/mm/yyyy')
    cell(S, rr, 3, "=COUNTIF('البنود'!$A$5:$A$%d,%d)" % (last, i))
    cell(S, rr, 4, "=COUNTIFS('البنود'!$A$5:$A$%d,%d,'البنود'!$E$5:$E$%d,\"\")" % (last, i, last))
    cell(S, rr, 5, "=SUMIF('البنود'!$A$5:$A$%d,%d,'البنود'!$F$5:$F$%d)" % (last, i, last), '0.000', BLU)
    cell(S, rr, 6, paper.get(i), '0.000')
    cell(S, rr, 7, '=IF(F{0}="","",E{0}-F{0})'.format(rr), '0.000')
S.conditional_formatting.add('D5:D22', FormulaRule(formula=['D5>0'], fill=FILL(RED)))
S.conditional_formatting.add('G5:G22', FormulaRule(formula=['AND(G5<>"",ABS(G5)>0.001)'], fill=FILL(RED)))
cell(S, 23, 1, 'المجموع', bold=True); cell(S, 23, 3, '=SUM(C5:C22)', bold=True); cell(S, 23, 4, '=SUM(D5:D22)', bold=True)
cell(S, 23, 5, '=SUM(E5:E22)', '#,##0.000', BLU, True); cell(S, 23, 6, '=SUM(F5:F22)', '#,##0.000', None, True)
S.cell(25, 1, 'حساب يزيد المتوقع').font = Font(b=True, size=12, color='1F4E6B')
for i, (k_, v_, note) in enumerate([('بضاعة بعناها إله (على أودو)', SALES, 'عليه'), ('دفعات أعطيناه (على أودو — بتتشيّك بعدين)', PAID, 'عليه'),
                                    ('بضاعة اشتريناها منه (هالملف)', '=-E23', 'إله'), ('الصافي اللي عليه', '=E26+E27+E28', '')]):
    rr = 26 + i
    cell(S, rr, 1, k_); S.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=4)
    cell(S, rr, 5, v_, '#,##0.000', BLU if 'الصافي' in k_ else None, 'الصافي' in k_); cell(S, rr, 6, note)

# ── التعليمات
I = wb.create_sheet('طريقة التعبئة')
title(I, 'طريقة التعبئة', '', 2)
head(I, 4, [('#', 5), ('شو تعمل', 105)])
for i, t in enumerate([
    'اشتغل على شيت «البنود» بس — الملخّص بيتحدّث لحاله.',
    'الخانة الصفرا بعمود «الإفرادي» = بند بلا سعر: اكتب سعر الحبة. فاتورة 18 (دفتر سافيتو) كلها بلا أسعار.',
    'عمود «✍️ الصنف المعتمد (كود)»: إذا الصنف المقترح صح اكتب كوده (أو انسخه من «الكود المقترح») · إذا غلط اكتب الكود الصح · إذا مش موجود بالكتالوج اتركه فاضي واكتب بملاحظاتك «جديد» ← بنعرضه على مدير الأصناف.',
    'عمود «التطابق»: أخضر = قوي · أصفر = ضعيف (غالباً غلط، دقّقه) · برتقالي = مش موجود بالكتالوج.',
    'عمود «✍️ ملاحظاتك»: أي إشي — «سعر مقطوع» · لون أو نوع … (بتنزل على البند نفسه) · ما في بند بينحذف إلا بكلمتك.',
    'بالملخّص: عمود «بلا سعر» أحمر لحد ما تخلص أسعار الفاتورة · «الفرق» عن مجموع الورقة لازم يضل صفر للفواتير 1–17.',
    'لما تخلص ابعتلي الملف ← بطلّعلك مسودة نهائية وما بنرحّل على أودو إلا بأمرك.'], 1):
    cell(I, 4 + i, 1, i); cell(I, 4 + i, 2, t, wrap=True); I.row_dimensions[4 + i].height = 32
# أسئلة مفتوحة
Q = wb.create_sheet('أسئلة')
title(Q, 'أسئلة مفتوحة', '', 3)
head(Q, 4, [('#', 5), ('السؤال', 85), ('✍️ جوابك', 40)])
QA = [('«خردوات 30 JD» بفاتورة 18: مبلغ مقطوع 30 دينار؟', 'مبلغ مقطوع ✓ (انحط: الكمية 1 × 30)'),
      ('«حبل تالف» بفاتورة 17 كميته 0 وسعره 0: بنشيله؟', 'يضل مبدئياً ✓'),
      ('«اضافة المواسير السعودي» بآخر دفتر سافيتو: عنوان بلا بنود — في بنود ناقصة؟', '📌 فاتورة لاحقة — رح يعطيها (تذكير)'),
      ('«محابس ستيم قديم 32» مكتوب جنبها «غير معلومة الورقة اين» — بتضل بفاتورة 18؟', 'نعم ✓')]
for i, (t, a) in enumerate(QA, 1):
    cell(Q, 4 + i, 1, i); cell(Q, 4 + i, 2, t, wrap=True); cell(Q, 4 + i, 3, a, fill=GRN, wrap=True); Q.row_dimensions[4 + i].height = 30
cell(Q, 10, 2, '🔴 ما في ولا بند بينحذف — كل إشي بيضل موجود (صاحب المحل 08/10)', bold=True)
wb.save(OUT); print('✓', OUT, '· بنود', last - 4)
