# -*- coding: utf-8 -*-
"""مراقبة مجلّد «ورشات» على Google Drive — أداة مدير الورشات

الوصول للدرايف بيصير عبر أدوات كلود (MCP) لا عبر بايثون، فهاي الأداة بتاخذ
**ناتج الجرد** وبتقارنه مع آخر لقطة محفوظة وبتطلّع الفرق:
   الجديد 🆕 · المعدَّل ✏️ · المحذوف 🗑️

  python3 scripts/مراقبة_درايف.py --لقطة  < listing.json     # حفظ اللقطة الأولى
  python3 scripts/مراقبة_درايف.py --فرق   < listing.json     # مقارنة + تحديث اللقطة
  python3 scripts/مراقبة_درايف.py --اعرض                      # اللقطة الحالية
"""
import sys, os, json, datetime

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
اللقطة = os.path.join(ROOT, 'drive', 'لقطة_ورشات.json')
BAR = '─' * 78
تجاهل = ('.onetoc2',)          # ملفات OneNote التقنية — مش ورشات


def حمّل():
    if not os.path.exists(اللقطة): return {}
    return json.load(open(اللقطة))


def اقرأ_المدخل():
    d = json.load(sys.stdin)
    files = d.get('files', d if isinstance(d, list) else [])
    out = {}
    for f in files:
        if f['title'].endswith(تجاهل): continue
        out[f['id']] = {'اسم': f['title'], 'عُدّل': f.get('modifiedTime'),
                        'حجم': f.get('fileSize'), 'رابط': f.get('viewUrl'),
                        'مالك': f.get('owner')}
    return out


def احفظ(d):
    d = dict(d); d['_وقت_الفحص'] = datetime.datetime.utcnow().isoformat(timespec='seconds') + 'Z'
    json.dump(d, open(اللقطة, 'w'), ensure_ascii=False, indent=1)


def فرق(قديم, جديد):
    قديم = {k: v for k, v in قديم.items() if not k.startswith('_')}
    ج = [k for k in جديد if k not in قديم]
    ح = [k for k in قديم if k not in جديد]
    ع = [k for k in جديد if k in قديم and (جديد[k]['عُدّل'] != قديم[k]['عُدّل']
                                            or جديد[k]['حجم'] != قديم[k]['حجم'])]
    return ج, ع, ح


def اطبع(قديم, جديد):
    ج, ع, ح = فرق(قديم, جديد)
    print(BAR); print('مراقبة مجلّد «ورشات» — %d ملف' % len(جديد)); print(BAR)
    if not (ج or ع or ح):
        print('✅ ما في أي تحديث منذ آخر فحص (%s).' % قديم.get('_وقت_الفحص', '—')); return False
    for k in ج:
        print('🆕 ملف جديد   : %-34s  %s' % (جديد[k]['اسم'], جديد[k]['عُدّل']))
    for k in ع:
        print('✏️  انعدّل      : %-34s  %s → %s  (حجم %s → %s)'
              % (جديد[k]['اسم'], قديم[k]['عُدّل'], جديد[k]['عُدّل'], قديم[k]['حجم'], جديد[k]['حجم']))
    for k in ح:
        print('🗑️  اختفى      : %-34s  (آخر تعديل %s)' % (قديم[k]['اسم'], قديم[k]['عُدّل']))
    print(BAR)
    print('الحصيلة: %d جديد · %d معدَّل · %d مختفٍ' % (len(ج), len(ع), len(ح)))
    print('🔴 كل ملف بالقائمة يُمرَّر على فحص الازدواج قبل أي إدخال (مدير الورشات).')
    print(BAR)
    return True


def main(argv):
    if not argv: raise SystemExit(__doc__)
    if argv[0] == '--اعرض':
        s = حمّل()
        print('آخر فحص:', s.get('_وقت_الفحص', '—'))
        for k, v in sorted(((k, v) for k, v in s.items() if not k.startswith('_')),
                           key=lambda z: z[1]['عُدّل'] or '', reverse=True):
            print('  %-34s  %s  %s' % (v['اسم'], v['عُدّل'], v['حجم']))
        return
    جديد = اقرأ_المدخل()
    if argv[0] == '--لقطة':
        احفظ(جديد); print('💾 انحفظت اللقطة الأولى: %d ملف' % len(جديد)); return
    if argv[0] == '--فرق':
        اطبع(حمّل(), جديد); احفظ(جديد); return
    raise SystemExit(__doc__)


if __name__ == '__main__':
    main(sys.argv[1:])
