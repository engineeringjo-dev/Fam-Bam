# -*- coding: utf-8 -*-
"""نموذج طلبية ورشة — محلات العون لمواد البناء (نسخة مبسّطة)"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

OUT = '/home/user/Fam-Bam/نموذج_طلبية_ورشة.xlsx'
R0, R1 = 9, 43                       # صفوف البيانات (35 سطر)
المدققون = 'عمر المصري,ابو علي'

NAVY, TEAL = '1F4E6B', '0F6E6E'
EDIT, AUTO = 'FFF6CC', 'EDEDED'
GREEN, RED = 'D6F0D6', 'FBD5D5'
LINE  = Side(style='thin', color='9AA5AD')
THICK = Side(style='medium', color=NAVY)
BOX   = Border(left=LINE, right=LINE, top=LINE, bottom=LINE)

def F(sz=11, b=False, c='000000'):
    return Font(name='Arial', size=sz, bold=b, color=c)
def C(h='center'):
    return Alignment(horizontal=h, vertical='center', wrap_text=True)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'طلبية'
ws.sheet_view.rightToLeft = True
ws.sheet_view.showGridLines = False

for k, v in {'A': 5, 'B': 42, 'C': 9, 'D': 14, 'E': 13, 'F': 16, 'G': 24}.items():
    ws.column_dimensions[k].width = v

def cells(rng):
    v = ws[rng]
    if hasattr(v, 'value'):
        return [v]
    out = []
    for row in v:
        out.extend([row] if hasattr(row, 'value') else list(row))
    return out

def lab(cell, txt):
    ws[cell] = txt
    ws[cell].font = F(10, True); ws[cell].alignment = C('right')
    ws[cell].fill = PatternFill('solid', fgColor='E8EDF0'); ws[cell].border = BOX

def inp(rng):
    if ':' in rng:
        ws.merge_cells(rng)
    first = rng.split(':')[0]
    ws[first].fill = PatternFill('solid', fgColor=EDIT)
    ws[first].font = F(11, True); ws[first].alignment = C()
    for c in cells(rng):
        c.border = BOX

def auto(cell):
    ws[cell].fill = PatternFill('solid', fgColor=AUTO)
    ws[cell].font = F(11, True); ws[cell].alignment = C(); ws[cell].border = BOX

# ── العنوان ──
ws.merge_cells('A1:G1')
ws['A1'] = 'طلبية ورشة — محلات العون لمواد البناء'
ws['A1'].font = F(16, True, 'FFFFFF'); ws['A1'].alignment = C()
ws['A1'].fill = PatternFill('solid', fgColor=NAVY)
ws.row_dimensions[1].height = 30

# ── الترويسة ──
for r in (2, 3):
    ws.row_dimensions[r].height = 23
lab('A2', 'الورشة / المشروع');   inp('B2')
lab('C2', 'رقم الطلبية');         inp('D2')
lab('E2', 'التاريخ');             inp('F2:G2')
lab('A3', 'المهندس / المستلِم');  inp('B3')
lab('C3', 'صورة الأصل محفوظة؟');  inp('D3')
lab('E3', 'الورقة');              inp('F3:G3')

# ── شريط التدقيق ──
ws.row_dimensions[5].height = 26
lab('A5', 'البنود المُدخلة')
ws['B5'] = '=COUNTA(B{}:B{})'.format(R0, R1); auto('B5')
lab('C5', 'المدقَّقة')
ws['D5'] = '=COUNTA(F{}:F{})'.format(R0, R1); auto('D5')
ws.merge_cells('E5:G5'); auto('E5')
ws['E5'] = ('=IF(B5=0,"⛔ ما في بنود مُدخلة",'
            'IF(D5<B5,"⛔ باقي "&(B5-D5)&" بند بدون تدقيق",'
            '"✅ كل البنود مدقَّقة"))')
ws['E5'].font = F(12, True)

# ── رأس الجدول ──
HDR = ['#', 'البند', 'الكمية', 'السعر الإفرادي', 'التاريخ', 'دقّقها', 'ملاحظات']
ws.row_dimensions[8].height = 28
for i, h in enumerate(HDR, start=1):
    c = ws.cell(8, i, h)
    c.font = F(11, True, 'FFFFFF'); c.alignment = C()
    c.fill = PatternFill('solid', fgColor=NAVY)
    c.border = Border(left=LINE, right=LINE, top=THICK, bottom=THICK)

# ── الصفوف ──
for r in range(R0, R1 + 1):
    ws.row_dimensions[r].height = 20
    ws.cell(r, 1, '=IF(B{}="","",ROW()-{})'.format(r, R0 - 1))
    for col in range(1, 8):
        c = ws.cell(r, col)
        c.border = BOX; c.font = F(11)
        c.alignment = C('right' if col in (2, 7) else 'center')
        c.fill = PatternFill('solid', fgColor=AUTO if col == 1 else 'FFFFFF')
    ws.cell(r, 3).number_format = '0.###'
    ws.cell(r, 4).number_format = '0.000'
    ws.cell(r, 5).number_format = 'DD/MM/YYYY'

# ── التواقيع ──
S = R1 + 2
ws.merge_cells('A{0}:G{0}'.format(S))
ws['A{}'.format(S)] = 'لا تُرسَل الطلبية إلا وكل بند مؤشَّر عليه مدقِّقه'
ws['A{}'.format(S)].font = F(10, True, 'FFFFFF'); ws['A{}'.format(S)].alignment = C()
ws['A{}'.format(S)].fill = PatternFill('solid', fgColor=TEAL)
ws.row_dimensions[S].height = 22

for j, (who, hint) in enumerate([('أدخلها', 'نقلها من الورقة الأصلية بند بند'),
                                 ('دقّقها', 'قارَنها بالورقة الأصلية: البند صح والعدد صح')]):
    r = S + 1 + j
    ws.row_dimensions[r].height = 25
    lab('A{}'.format(r), who);         inp('B{}'.format(r))
    lab('C{}'.format(r), 'التاريخ');   inp('D{}'.format(r))
    ws.merge_cells('E{0}:G{0}'.format(r))
    ws['E{}'.format(r)] = hint
    ws['E{}'.format(r)].font = F(9, False, '555555'); ws['E{}'.format(r)].alignment = C()
    ws['E{}'.format(r)].border = BOX
AUDITOR_SIG = 'B{}'.format(S + 2)

# ── القوائم المنسدلة ──
def dv(formula, target):
    d = DataValidation(type='list', formula1=formula, allow_blank=True, showDropDown=False)
    ws.add_data_validation(d); d.add(target)

dv('"%s"' % المدققون, 'F{}:F{}'.format(R0, R1))
dv('"%s"' % المدققون, AUDITOR_SIG)
dv('"نعم,لا"', 'D3')

num = DataValidation(type='decimal', operator='greaterThan', formula1='0', allow_blank=True,
                     showErrorMessage=True, errorTitle='كمية غير صحيحة',
                     error='الكمية رقم فقط. «ربطة» و«جوز» و«لفة» تُكتب بخانة الملاحظات.')
ws.add_data_validation(num); num.add('C{}:C{}'.format(R0, R1))

# ── التنسيق الشرطي ──
ws.conditional_formatting.add('E5', FormulaRule(
    formula=['LEFT(E5,1)="✅"'], fill=PatternFill('solid', bgColor=GREEN), font=F(12, True, '17632A')))
ws.conditional_formatting.add('E5', FormulaRule(
    formula=['LEFT(E5,1)="⛔"'], fill=PatternFill('solid', bgColor=RED), font=F(12, True, 'A11212')))
for col in ('C', 'E'):
    ws.conditional_formatting.add('{0}{1}:{0}{2}'.format(col, R0, R1), FormulaRule(
        formula=['AND($B{0}<>"",${1}{0}="")'.format(R0, col)],
        fill=PatternFill('solid', bgColor=RED)))
ws.conditional_formatting.add('A{}:G{}'.format(R0, R1), FormulaRule(
    formula=['$F{}<>""'.format(R0)], fill=PatternFill('solid', bgColor='F0F8F0')))

# ── الطباعة والحماية ──
ws.print_area = 'A1:G{}'.format(S + 2)
ws.page_setup.orientation = 'landscape'
ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_margins.left = ws.page_margins.right = 0.3
ws.page_margins.top = ws.page_margins.bottom = 0.4
ws.print_title_rows = '8:8'
ws.freeze_panes = 'A9'

for row in ws.iter_rows(min_row=1, max_row=S + 2, max_col=7):
    for c in row:
        c.protection = Protection(locked=True)
for rng in ['B2', 'D2', 'F2:G2', 'B3', 'D3', 'F3:G3',
            'B{}:G{}'.format(R0, R1),
            'B{}'.format(S + 1), 'D{}'.format(S + 1), AUDITOR_SIG, 'D{}'.format(S + 2)]:
    for c in cells(rng):
        c.protection = Protection(locked=False)
ws.protection.sheet = True
ws.protection.formatCells = False
ws.protection.selectLockedCells = False

# ══════════ التعليمات ══════════
wi = wb.create_sheet('التعليمات')
wi.sheet_view.rightToLeft = True
wi.sheet_view.showGridLines = False
wi.column_dimensions['A'].width = 4
wi.column_dimensions['B'].width = 92

wi.merge_cells('A1:B1')
wi['A1'] = 'طريقة العمل'
wi['A1'].font = F(16, True, 'FFFFFF'); wi['A1'].alignment = C()
wi['A1'].fill = PatternFill('solid', fgColor=NAVY)
wi.row_dimensions[1].height = 30

STEPS = [
    ('sec', 'الموظف اللي بطلّع البضاعة'),
    ('1', 'ورقة وحدة لكل طلبية، ولها رقم متسلسل. اكتب الرقم على ورقة الطلبية الأصلية كمان.'),
    ('2', 'أدخل البنود بند بند. وكل بند بتدخله — اشطبه بالقلم على الورقة الأصلية. '
          'أي بند بضل بدون شطب = بند نسيته. هاي أهم خطوة بكل النموذج.'),
    ('3', 'الكمية رقم فقط. «ربطة» و«جوز» و«لفة» تُكتب بخانة الملاحظات.'),
    ('4', 'التاريخ إلزامي لكل بند — كل تاريخ بصير فاتورة لحاله بالنظام.'),
    ('5', 'السعر الإفرادي: إذا المهندس اتفق مع المقاول على سعر، اكتبه. وإلا اتركه فاضي.'),
    ('6', 'صوّر الورقة الأصلية واحفظها بنفس رقم الطلبية.'),
    ('sec', 'المدقِّق — عمر المصري أو أبو علي'),
    ('7', 'التدقيق من الورقة الأصلية، مش من الشاشة. امسك الورقة واقرأ منها.'),
    ('8', 'كل بند تتأكد منه (البند صح والعدد صح) — اختر اسمك بخانة «دقّقها» لهذا البند.'),
    ('9', 'العدّاد فوق ما بيصير أخضر إلا لما كل بند مُدخل يكون عليه اسم مدقِّق.'),
    ('10', 'وقّع بأسفل الورقة وابعث الملف + صورة الورقة الأصلية للأستاذ.'),
    ('sec', 'الفكرة باختصار'),
    ('◆', 'البند المنسي ما بيترك أثر على الإكسل — عشان هيك التدقيق لازم يكون مقابل الورقة '
          'الأصلية لا مقابل الشاشة، ومعه شطب بالقلم على الورقة وقت الإدخال.'),
]
r = 3
for kind, txt in STEPS:
    if kind == 'sec':
        wi.merge_cells('A{0}:B{0}'.format(r))
        wi['A{}'.format(r)] = txt
        wi['A{}'.format(r)].font = F(12, True, 'FFFFFF')
        wi['A{}'.format(r)].fill = PatternFill('solid', fgColor=TEAL)
        wi['A{}'.format(r)].alignment = C()
        wi.row_dimensions[r].height = 24
    else:
        wi['A{}'.format(r)] = kind
        wi['A{}'.format(r)].font = F(11, True, NAVY); wi['A{}'.format(r)].alignment = C()
        wi['B{}'.format(r)] = txt
        wi['B{}'.format(r)].font = F(11); wi['B{}'.format(r)].alignment = C('right')
        wi.row_dimensions[r].height = 30 if len(txt) < 100 else 44
    r += 1

wi.print_area = 'A1:B{}'.format(r - 1)
wi.page_setup.orientation = 'portrait'
wi.page_setup.paperSize = wi.PAPERSIZE_A4
wi.page_setup.fitToWidth = wi.page_setup.fitToHeight = 1
wi.sheet_properties.pageSetUpPr.fitToPage = True

wb.active = 0
wb.save(OUT)
print('تم:', OUT)
