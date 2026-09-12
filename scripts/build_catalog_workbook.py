# -*- coding: utf-8 -*-
"""Build the Arabic product workbook for محلات العون from the live Odoo catalogue.

Two jobs in one file:
  * «أصناف جديدة» — an Odoo-import-ready entry sheet whose barcode column is a
    formula, so the shop keeps numbering products in the shop's own 9xxxxx block
    without any server-side automation.
  * «الكتالوج» — the current catalogue, so a name can be checked against what
    already exists before it is entered twice.

Re-run any time to pull a fresh copy:  python scripts/build_catalog_workbook.py
"""
import sys, json
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from scripts.check_connection import load_dotenv

load_dotenv(Path(__file__).resolve().parent.parent / ".env")

from odoo_mcp.client import OdooClient                      # noqa: E402
from odoo_mcp.config import OdooConfig                      # noqa: E402
from openpyxl import Workbook                               # noqa: E402
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side  # noqa: E402
from openpyxl.utils import get_column_letter                # noqa: E402
from openpyxl.worksheet.datavalidation import DataValidation  # noqa: E402
from openpyxl.comments import Comment                       # noqa: E402

OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("كتالوج_العون.xlsx")
BLANK_ROWS = 300

# Technical field names in row 1 — Odoo's importer matches on these directly,
# so the sheet needs no column mapping when it is imported.
COLS = [
    ("name",             "اسم الصنف",            34, "input"),
    ("default_code",     "المرجع الداخلي",       14, "formula"),
    ("barcode",          "الباركود",             14, "formula"),
    ("categ_id",         "فئة المنتج",           16, "list"),
    ("pos_categ_ids",    "فئة نقطة البيع",       20, "list"),
    ("list_price",       "سعر البيع",            11, "input"),
    ("uom_id",           "الوحدة",               11, "list"),
    ("taxes_id",         "ضريبة البيع",          13, "list"),
    ("available_in_pos", "يظهر بنقطة البيع",     16, "auto"),
    ("is_storable",      "يُخزَّن",               11, "auto"),
    ("sale_ok",          "قابل للبيع",           12, "auto"),
    ("purchase_ok",      "قابل للشراء",          12, "auto"),
]
CHECK_COL = len(COLS) + 1

INK, PAPER = "1F2A37", "FFFFFF"
HEAD_BG, HEAD_FG = "1F3A5F", "FFFFFF"
INPUT_BG, FORMULA_BG, AUTO_BG = "FFF6D9", "E8EEF6", "F1F3F5"
WARN = "B42318"

thin = Side(style="thin", color="C6CCD4")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def fetch():
    c = OdooClient(OdooConfig.from_env())
    fields = [f for f, *_ in COLS]
    ids = c.execute_kw("product.template", "search", [[]],
                       {"limit": 8000, "order": "default_code"})
    rows = []
    for i in range(0, len(ids), 200):
        rows += c.call_method("product.template", "export_data",
                              [ids[i:i + 200], ["id"] + fields])["datas"]
    seq = c.read("ir.sequence", [29], ["number_next_actual"])[0]
    return {
        "rows": rows,
        "next": seq["number_next_actual"],
        "categories": sorted(x["display_name"] for x in
                             c.search_read("product.category", [], ["display_name"], limit=300)),
        "pos_categories": sorted(
            (f"{x['parent_id'][1]} / {x['name']}" if x["parent_id"] else x["name"])
            for x in c.search_read("pos.category", [], ["name", "parent_id"], limit=300)),
        "uoms": sorted(x["name"] for x in c.search_read("uom.uom", [], ["name"], limit=200)),
        "taxes": sorted(x["name"] for x in c.search_read(
            "account.tax", [("type_tax_use", "=", "sale")], ["name"], limit=100)),
    }


def style_header(ws, labels, row=1):
    for i, text in enumerate(labels, start=1):
        cell = ws.cell(row=row, column=i, value=text)
        cell.font = Font(name="Arial", bold=True, size=10, color=HEAD_FG)
        cell.fill = PatternFill("solid", fgColor=HEAD_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


def sheet_entry(wb, data):
    """The sheet the shop types into and then imports straight back into Odoo."""
    ws = wb.create_sheet("أصناف جديدة")
    last = len(data["rows"]) + 1        # bound the lookup: a whole-column
                                        # COUNTIF scans a million rows per line
    ws.sheet_view.rightToLeft = True

    style_header(ws, [f for f, *_ in COLS] + ["تدقيق — لا تستورد هذا العمود"])
    ws.row_dimensions[1].height = 30
    for i, (_, _, width, _) in enumerate(COLS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.column_dimensions[get_column_letter(CHECK_COL)].width = 26

    # Arabic meaning of each technical header, as a cell comment.
    for i, (field, label, _, kind) in enumerate(COLS, start=1):
        note = {"input": "اكتب هنا", "formula": "معادلة — لا تكتب فوقها",
                "list": "اختر من القائمة", "auto": "يتعبّى تلقائيًا"}[kind]
        ws.cell(row=1, column=i).comment = Comment(f"{label}\n({note})", "العون", height=90, width=190)

    base = data["next"] - 1          # first filled row lands on data["next"]
    for r in range(2, 2 + BLANK_ROWS):
        a = f"$A{r}"
        filled = f'IF({a}="",""'     # every column stays empty until a name is typed

        ws.cell(row=r, column=1).fill = PatternFill("solid", fgColor=INPUT_BG)
        ws.cell(row=r, column=2, value=f'={filled},$C{r})')
        ws.cell(row=r, column=3,
                value=f'={filled},TEXT({base}+COUNTA($A$2:{a}),"000000"))')
        ws.cell(row=r, column=6).fill = PatternFill("solid", fgColor=INPUT_BG)
        ws.cell(row=r, column=7, value=f'={filled},"حبة")')
        ws.cell(row=r, column=8, value=f'={filled},"0%")')
        for col in (9, 10, 11, 12):
            ws.cell(row=r, column=col, value=f'={filled},TRUE)')
        ws.cell(row=r, column=CHECK_COL,
                value=f'={filled},IF(COUNTIF(\'الكتالوج\'!$C$2:$C${last},{a})>0,'
                      f'"⚠ الاسم موجود مسبقًا","جديد ✔"))')

        for col in range(1, CHECK_COL + 1):
            cell = ws.cell(row=r, column=col)
            cell.font = Font(name="Arial", size=10, color=INK)
            cell.border = BOX
            if col in (2, 3):
                cell.fill = PatternFill("solid", fgColor=FORMULA_BG)
                cell.font = Font(name="Arial", size=10, bold=True, color=INK)
            elif col in (7, 8, 9, 10, 11, 12):
                cell.fill = PatternFill("solid", fgColor=AUTO_BG)
            elif col in (4, 5):
                cell.fill = PatternFill("solid", fgColor=INPUT_BG)
        ws.cell(row=r, column=6).number_format = "#,##0.00"

    ranges = {4: ("A", len(data["categories"])), 5: ("B", len(data["pos_categories"])),
              7: ("C", len(data["uoms"])), 8: ("D", len(data["taxes"]))}
    for col, (letter, count) in ranges.items():
        dv = DataValidation(type="list",
                            formula1=f"'القوائم'!${letter}$2:${letter}${count + 1}",
                            allow_blank=True, showDropDown=False)
        ws.add_data_validation(dv)
        dv.add(f"{get_column_letter(col)}2:{get_column_letter(col)}{1 + BLANK_ROWS}")

    ws.freeze_panes = "A2"
    return ws


def sheet_catalog(wb, data):
    ws = wb.create_sheet("الكتالوج")
    ws.sheet_view.rightToLeft = True
    style_header(ws, ["المعرّف الخارجي (id)"] + [f for f, *_ in COLS])
    widths = [40] + [w for _, _, w, _ in COLS]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for r, row in enumerate(data["rows"], start=2):
        for col, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=col, value=val)
            cell.font = Font(name="Arial", size=9, color=INK)
            if col == 7:
                cell.number_format = "#,##0.00"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS) + 1)}{len(data['rows']) + 1}"
    return ws


def sheet_lists(wb, data):
    ws = wb.create_sheet("القوائم")
    ws.sheet_view.rightToLeft = True
    style_header(ws, ["فئات المنتجات", "فئات نقطة البيع", "الوحدات", "ضرائب البيع"])
    for i, key in enumerate(("categories", "pos_categories", "uoms", "taxes"), start=1):
        ws.column_dimensions[get_column_letter(i)].width = 30
        for r, val in enumerate(data[key], start=2):
            ws.cell(row=r, column=i, value=val).font = Font(name="Arial", size=10, color=INK)
    return ws


def sheet_help(wb, data):
    ws = wb.create_sheet("التعليمات")
    ws.sheet_view.rightToLeft = True
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 78

    def line(r, a, b, c, bold=False, color=INK, size=10, fill=None):
        for col, val in ((1, a), (2, b), (3, c)):
            cell = ws.cell(row=r, column=col, value=val)
            cell.font = Font(name="Arial", size=size, bold=bold, color=color)
            cell.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True)
            if fill:
                cell.fill = PatternFill("solid", fgColor=fill)

    line(1, "", "دليل إضافة الأصناف", "محلات العون لمواد البناء — alawn.odoo.com",
         bold=True, size=14)
    line(2, "", "", f"الرقم التالي المتاح: {data['next']}   ·   الأصناف الحالية: {len(data['rows'])}",
         color="5A6672")

    rows = [
        ("", "الخطوات", "", True, HEAD_BG, "E8EEF6"),
        ("١", "افتح «أصناف جديدة»", "اكتب اسم الصنف بالعمود الأصفر الأول. الباركود والمرجع الداخلي بيتعبّوا لحالهم بمجرد ما تكتب الاسم.", False, INK, None),
        ("٢", "املأ الأصفر فقط", "الأعمدة الصفراء هي اللي بتكتب فيها: الاسم، فئة المنتج، فئة نقطة البيع، سعر البيع. الباقي بيتعبّى تلقائيًا.", False, INK, None),
        ("٣", "راقب عمود التدقيق", "آخر عمود بيقارن الاسم مع الكتالوج. إذا طلع «⚠ الاسم موجود مسبقًا» فالصنف معرّف من قبل — لا تكرّره.", False, INK, None),
        ("٤", "احذف الأسطر الفاضية", "قبل الاستيراد، احذف كل سطر ما كتبت فيه اسم. لا تستورد العمود الأخير (تدقيق).", False, WARN, None),
        ("٥", "استورد", "أودو ← المنتجات ← ⚙️ ← Import records ← ارفع الملف ← اختر ورقة «أصناف جديدة» ← Test ثم Import.", False, INK, None),
        ("٦", "بعد الاستيراد", "ارجع شغّل هذا الملف من جديد ليتحدّث الكتالوج ويبدأ الترقيم من الرقم الصحيح.", False, INK, None),
        ("", "", "", False, INK, None),
        ("", "مثال على سطر مكتمل", "", True, HEAD_BG, "E8EEF6"),
    ]
    r = 4
    for a, b, c, bold, color, fill in rows:
        line(r, a, b, c, bold=bold, color=color, fill=fill)
        r += 1

    ex_head = r
    style_header(ws, [f for f, *_ in COLS], row=ex_head)
    example = ["كوع 2 إنش أبيض", str(data["next"]), str(data["next"]), "صحي",
               "صحي", "1.25", "حبة", "0%", "TRUE", "TRUE", "TRUE", "TRUE"]
    for i, val in enumerate(example, start=1):
        cell = ws.cell(row=ex_head + 1, column=i, value=val)
        cell.font = Font(name="Arial", size=10, color=INK)
        cell.fill = PatternFill("solid", fgColor=INPUT_BG)
        cell.border = BOX
    line(ex_head + 2, "", "",
         "المرجع الداخلي والباركود متساويان دائمًا — هذا هو النمط المعتمد بالمحل (4547 صنف عليه).",
         color="5A6672", size=9)
    line(ex_head + 4, "", "ملاحظات",
         "• لا تستعمل رقمًا يبدأ بصفر، ولا رقمًا بطول 8 أو 13 خانة بالضبط — أودو بتفسّرها كباركود عالمي.\n"
         "• الأصناف اللي عليها باركود مصنّع (EAN) اكتبه بيدك بعمود الباركود بدل الرقم التلقائي.\n"
         "• للترقيم داخل أودو مباشرة: المنتجات ← فلتر «أصناف بدون باركود» ← حدّد ← ⚙️ Actions ← «ترقيم الأصناف الجديدة».")
    ws.row_dimensions[ex_head + 4].height = 58
    return ws


def main():
    data = fetch()
    print(f"سُحب {len(data['rows'])} صنف · الرقم التالي {data['next']}")
    wb = Workbook()
    wb.remove(wb.active)
    sheet_help(wb, data)
    sheet_entry(wb, data)
    sheet_catalog(wb, data)
    sheet_lists(wb, data)
    wb.active = 0
    wb.save(OUT)
    print(f"✔ {OUT}")


if __name__ == "__main__":
    main()
