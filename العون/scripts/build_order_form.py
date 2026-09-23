# -*- coding: utf-8 -*-
"""نموذج ورشة — محلات العون لمواد البناء
ملف واحد لكل ورشة: ورقة «الورشة» (الصادر) + ورقة «المرتجعات» + التعليمات.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

OUT = '/home/user/Fam-Bam/العون/نموذج_طلبية_ورشة.xlsx'
المدققون = 'عمر المصري,ابو علي'

NAVY, TEAL, RUST = '1F4E6B', '0F6E6E', '9A4A00'
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


def build(ws, *, title, banner, rows, who_col, count_label='البنود المُدخلة',
          name_formula=None, foot=None):
    """يبني ورقة جدول: ترويسة اسم الورشة + عدّاد التدقيق + جدول ٧ أعمدة."""
    R0, R1 = 7, 6 + rows
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False
    for k, v in {'A': 5, 'B': 52, 'C': 11, 'D': 12, 'E': 13, 'F': 18, 'G': 18}.items():
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

    def auto(cell):
        ws[cell].fill = PatternFill('solid', fgColor=AUTO)
        ws[cell].font = F(11, True); ws[cell].alignment = C(); ws[cell].border = BOX

    ws.merge_cells('A1:G1')
    ws['A1'] = banner
    ws['A1'].font = F(16, True, 'FFFFFF'); ws['A1'].alignment = C()
    ws['A1'].fill = PatternFill('solid', fgColor=title)
    ws.row_dimensions[1].height = 30

    ws.row_dimensions[2].height = 26
    lab('A2', 'الورشة')
    ws.merge_cells('B2:G2')
    if name_formula:
        ws['B2'] = name_formula; auto('B2')
    else:
        ws['B2'].fill = PatternFill('solid', fgColor=EDIT)
        for c in cells('B2:G2'):
            c.border = BOX
    ws['B2'].font = F(13, True); ws['B2'].alignment = C()

    ws.row_dimensions[4].height = 26
    lab('A4', count_label)
    ws['B4'] = '=COUNTA(B{}:B{})'.format(R0, R1); auto('B4')
    lab('C4', 'المدقَّقة')
    ws['D4'] = '=COUNTA(G{}:G{})'.format(R0, R1); auto('D4')
    ws.merge_cells('E4:G4'); auto('E4')
    ws['E4'] = ('=IF(B4=0,"⛔ ما في بنود مُدخلة",'
                'IF(D4<B4,"⛔ باقي "&(B4-D4)&" بند بدون تدقيق",'
                '"✅ كل البنود مدقَّقة"))')
    ws['E4'].font = F(12, True)

    for i, h in enumerate(['#', 'البند', 'الكمية', 'الإفرادي', 'التاريخ', who_col, 'المدقق'], start=1):
        c = ws.cell(6, i, h)
        c.font = F(11, True, 'FFFFFF'); c.alignment = C()
        c.fill = PatternFill('solid', fgColor=title)
        c.border = Border(left=LINE, right=LINE, top=THICK, bottom=THICK)
    ws.row_dimensions[6].height = 28

    for r in range(R0, R1 + 1):
        ws.row_dimensions[r].height = 21
        ws.cell(r, 1, '=IF(B{}="","",ROW()-{})'.format(r, R0 - 1))
        for col in range(1, 8):
            c = ws.cell(r, col)
            c.border = BOX; c.font = F(11)
            c.alignment = C('right' if col == 2 else 'center')
            c.fill = PatternFill('solid', fgColor=AUTO if col == 1 else 'FFFFFF')
        ws.cell(r, 4).number_format = '0.000'
        ws.cell(r, 5).number_format = 'DD/MM/YYYY'

    W = R1 + 2
    ws.merge_cells('A{0}:G{0}'.format(W))
    ws['A{}'.format(W)] = foot or 'لا يُرسَل الملف إلا وكل بند مكتوب عليه اسم مدقِّقه'
    ws['A{}'.format(W)].font = F(10, True, 'FFFFFF'); ws['A{}'.format(W)].alignment = C()
    ws['A{}'.format(W)].fill = PatternFill('solid', fgColor=TEAL)
    ws.row_dimensions[W].height = 22

    d = DataValidation(type='list', formula1='"%s"' % المدققون, allow_blank=True, showDropDown=False)
    ws.add_data_validation(d); d.add('G{}:G{}'.format(R0, R1))

    ws.conditional_formatting.add('E4', FormulaRule(
        formula=['LEFT(E4,1)="✅"'], fill=PatternFill('solid', bgColor=GREEN), font=F(12, True, '17632A')))
    ws.conditional_formatting.add('E4', FormulaRule(
        formula=['LEFT(E4,1)="⛔"'], fill=PatternFill('solid', bgColor=RED), font=F(12, True, 'A11212')))
    for col in ('C', 'E'):
        ws.conditional_formatting.add('{0}{1}:{0}{2}'.format(col, R0, R1), FormulaRule(
            formula=['AND($B{0}<>"",${1}{0}="")'.format(R0, col)],
            fill=PatternFill('solid', bgColor=RED)))
    ws.conditional_formatting.add('A{}:G{}'.format(R0, R1), FormulaRule(
        formula=['$G{}<>""'.format(R0)], fill=PatternFill('solid', bgColor='F0F8F0')))

    ws.print_area = 'A1:G{}'.format(W)
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top = ws.page_margins.bottom = 0.4
    ws.print_title_rows = '1:6'
    ws.freeze_panes = 'A7'

    for row in ws.iter_rows(min_row=1, max_row=W, max_col=7):
        for c in row:
            c.protection = Protection(locked=True)
    unlock = ['B{}:G{}'.format(R0, R1)] + ([] if name_formula else ['B2'])
    for rng in unlock:
        for c in cells(rng):
            c.protection = Protection(locked=False)
    ws.protection.sheet = True
    ws.protection.formatCells = False
    ws.protection.selectLockedCells = False


# ── ورقة الصادر ──
out = wb.active
out.title = 'الورشة'
build(out, title=NAVY, banner='ورشة — محلات العون لمواد البناء',
      rows=200, who_col='المستلم')

# ── ورقة المرتجعات ──
ret = wb.create_sheet('المرتجعات')
build(ret, title=RUST, banner='مرتجعات الورشة — محلات العون لمواد البناء',
      rows=60, who_col='المُرجِع', count_label='البنود المرتجعة',
      name_formula="=IF('الورشة'!B2=\"\",\"\",'الورشة'!B2)",
      foot='الإفرادي هنا = السعر اللي انباع فيه للمشروع، مش سعر الكتالوج')

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
    ('sec', 'ورقة «الورشة» — البضاعة اللي بتطلع'),
    ('1', 'ورقة وحدة لكل ورشة، من أول بند لآخر بند. اسم الورشة بيتكتب مرة وحدة فوق وبس.'),
    ('2', 'أدخل البنود بند بند. وكل بند بتدخله — اشطبه بالقلم على الورقة الأصلية. '
          'أي بند بضل بدون شطب = بند نسيته. هاي أهم خطوة بكل النموذج.'),
    ('3', 'التاريخ إلزامي لكل بند — كل تاريخ بصير فاتورة لحاله بالنظام.'),
    ('4', 'الإفرادي: إذا المهندس اتفق مع المقاول على سعر، اكتبه. وإلا اتركه فاضي.'),
    ('5', 'المستلم: اكتب اسم اللي استلم البضاعة على كل بند.'),
    ('sec', 'ورقة «المرتجعات» — البضاعة اللي بترجع'),
    ('6', 'المرتجعات لها ورقة مستقلة (التبويب البرتقالي). المرتجع ما بينكتب أبداً بورقة الورشة.'),
    ('7', 'اسم الورشة فوق بيتعبّى لحاله من ورقة الورشة — لا تكتبه.'),
    ('8', 'الكمية تنكتب عادي بالموجب — الورقة كلها مرتجعات، ما في داعي للسالب.'),
    ('9', 'الإفرادي = السعر اللي انباع فيه للمشروع، مش سعر الكتالوج. '
          'وإذا الراجع مكسور أو مفتوح وبدك تخصم، اكتب سعر أقل — هاد اللي رح يترجع للمقاول. '
          'وإذا ما بتعرف السعر اتركه فاضي وأنا بجيبه من فواتير المشروع.'),
    ('10', 'المُرجِع: اسم اللي رجّع البضاعة.'),
    ('sec', 'المدقِّق — عمر المصري أو أبو علي'),
    ('11', 'التدقيق من الورقة الأصلية، مش من الشاشة. امسك الورقة واقرأ منها.'),
    ('12', 'كل بند تتأكد منه (البند صح والعدد صح) — اختر اسمك بخانة «المدقق».'),
    ('13', 'العدّاد فوق ما بيصير أخضر إلا لما كل بند مُدخل يكون عليه اسم مدقِّق. '
           'وكل ورقة إلها عدّادها — لازم الاثنين يصيروا أخضر.'),
    ('14', 'ابعث الملف + صورة الورقة الأصلية للأستاذ.'),
    ('15', 'وقت الطباعة: حدّد المدى المعبّى فقط، لأنه الورقة فيها سطور فاضية جاهزة.'),
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
        wi['A{}'.format(r)].fill = PatternFill('solid', fgColor=RUST if 'المرتجعات' in txt else TEAL)
        wi['A{}'.format(r)].alignment = C()
        wi.row_dimensions[r].height = 24
    else:
        wi['A{}'.format(r)] = kind
        wi['A{}'.format(r)].font = F(11, True, NAVY); wi['A{}'.format(r)].alignment = C()
        wi['B{}'.format(r)] = txt
        wi['B{}'.format(r)].font = F(11); wi['B{}'.format(r)].alignment = C('right')
        wi.row_dimensions[r].height = 30 if len(txt) < 100 else 46
    r += 1
wi.print_area = 'A1:B{}'.format(r - 1)
wi.page_setup.orientation = 'portrait'
wi.page_setup.paperSize = wi.PAPERSIZE_A4
wi.page_setup.fitToWidth = wi.page_setup.fitToHeight = 1
wi.sheet_properties.pageSetUpPr.fitToPage = True

wb.active = 0
wb.save(OUT)
print('تم:', OUT)
