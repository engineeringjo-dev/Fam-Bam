# -*- coding: utf-8 -*-
"""نظام الموظفين النسخة 26 ← 27: إلغاء «عمال المهام» (صاحب المحل 10/10 — «نقلنا عمال المهام على التكاليف التشغيلية»)

  python3 scripts/نظام_الموظفين_بلا_عمال_المهام.py <الملف المشفّر 26.xlsx> <الناتج 27.xlsx>

- شيت «عمال المهام» بينشال + الأسماء المعرّفة TASK_*.
- «ملخص الأيدي العاملة»: عمود «عمال المهام» بالجدول الشهري بينشال (المجموع كاش وكلفة الشغل والبضاعة والقروض بيرجعوا خانة لورا
  ومعادلاتهم بلا المهام) · قسم «عمال المهام — السنة، وشهر بتختاره» (الأسطر 29–51) بينشال.
- التعليمات والملاحظة تحت الجدول بتتعدّل · كل الباقي زي ما هو · الملف بيتشفّر بنفس كلمة السر ****.
"""
import sys, io, re, copy, msoffcrypto, msoffcrypto.format.ooxml, openpyxl

SRC, OUT, PWD = sys.argv[1], sys.argv[2], __import__('os').environ.get('ALAWN_XLSX_PWD', '')
buf = io.BytesIO(); f = msoffcrypto.OfficeFile(open(SRC, 'rb')); f.load_key(PWD); f.decrypt(buf); buf.seek(0)
wb = openpyxl.load_workbook(buf)

# 1) الشيت والأسماء
del wb['عمال المهام']
for n in [n for n in list(wb.defined_names) if n.startswith('TASK_')]:
    del wb.defined_names[n]

# 2) الملخص — الجدول الشهري (أسطر 5–18): I..L ← H..K
Z = wb['ملخص الأيدي العاملة']
for r in range(5, 19):
    for c in range(8, 12):                               # H..K ياخدوا قيمة وستايل I..L
        src, dst = Z.cell(r, c + 1), Z.cell(r, c)
        dst.value = src.value; dst._style = copy.copy(src._style)
    Z.cell(r, 12).value = None; Z.cell(r, 12)._style = copy.copy(Z.cell(4, 12)._style)
for r in range(6, 18):
    Z.cell(r, 8).value = '=E{0}+F{0}+G{0}'.format(r)                                     # المجموع كاش (بلا المهام)
    Z.cell(r, 9).value = re.sub(r'\+H%d$' % r, '', str(Z.cell(r, 9).value))              # كلفة الشغل = المستحق بس
    assert 'H%d' % r not in Z.cell(r, 9).value, Z.cell(r, 9).value
for c, L in ((8, 'H'), (9, 'I'), (10, 'J'), (11, 'K')):
    Z.cell(18, c).value = '=SUM({0}6:{0}17)'.format(L)
for c in 'HIJK':                                          # عرض الأعمدة بيرجع مع خاناته
    Z.column_dimensions[c].width = Z.column_dimensions[chr(ord(c) + 1)].width
Z.column_dimensions['L'].width = 4
cf = Z.conditional_formatting
rules = [(str(rng.sqref), r) for rng in cf for r in rng.rules]
Z.conditional_formatting = type(cf)()
for sq, rule in rules:
    Z.conditional_formatting.add('A6:K17' if sq == 'A6:L17' else sq, rule)
Z['A19'] = Z['A19'].value.replace('«كلفة الشغل» = المستحق + المهام (', '«كلفة الشغل» = المستحق (') + \
    ' · عمال المهام صاروا بنموذج «التكاليف التشغيلية».'

# 3) الملخص — قسم «عمال المهام» (29–51) بينشال
for m in [m for m in Z.merged_cells.ranges if m.min_row >= 28]:
    Z.unmerge_cells(str(m))
Z.data_validations.dataValidation = [d for d in Z.data_validations.dataValidation if 'I29' not in str(d.sqref)]
for r in range(28, Z.max_row + 1):
    for c in range(1, 13):
        Z.cell(r, c).value = None; Z.cell(r, c)._style = copy.copy(Z.cell(28, 13)._style)

# 4) التعليمات
I = wb['التعليمات']
I['B12'] = I['B12'].value.replace('عامل مهمة: سطر بورقة «عمال المهام».', 'عمال المهام: بنموذج «التكاليف التشغيلية» (مش هون).')

# 5) فحص: ولا معادلة بتأشّر على المهام
for ws in wb:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and ('TASK_' in c.value or 'عمال المهام\'!' in c.value):
                sys.exit('⛔ لسا في مرجع للمهام: %s!%s %s' % (ws.title, c.coordinate, c.value[:80]))

plain = io.BytesIO(); wb.save(plain); plain.seek(0)
with open(OUT, 'wb') as o:
    msoffcrypto.format.ooxml.OOXMLFile(plain).encrypt(PWD, o)
print('✓', OUT, wb.sheetnames)
