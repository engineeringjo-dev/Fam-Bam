# -*- coding: utf-8 -*-
"""شجرة الدرايف — فحص مجلّد «ورشات» بأقل استهلاك وبدقة كاملة (صاحب المحل 14/10)

الفكرة: ما بنحمّل ولا ملف إلا إذا تغيّر.
  1) جرد خفيف (بيانات الملفات بس: الاسم · المجلّد · وقت التعديل · الحجم — بلا محتوى) لكل المجلّدات
     ← الشجرة `drive/شجرة_الدرايف.json` بتعرف آخر نسخة قرأناها من كل ملف.
  2) أي ملف وقت تعديله أو حجمه تغيّر (أو جديد) ← بس هو بينحمّل (بمساعد عشان المحتوى ما يعبّي المحادثة).
  3) المحتوى بيتقرأ سطر سطر ← لقطة `drive/لقطات/<id>.json` ← **فرق على مستوى السطر** مع اللقطة القبل
     (سطر جديد · سطر تغيّر · سطر انشال) — الدقة كاملة لأن أي تعديل بالملف بيغيّر وقته.

  python3 scripts/شجرة_الدرايف.py --جرد                 # يقرأ آخر نتائج الجرد من سجل الجلسة ويقارن ← قائمة اللي لازم ينحمّل
  python3 scripts/شجرة_الدرايف.py --خزّن <id> [ملف.xlsx]  # يخزّن نسخة محمّلة (أو يدوّر عليها بسجلات الجلسة) ويطلّع الفرق
  python3 scripts/شجرة_الدرايف.py --بذرة <id> <ملف.xlsx> <modifiedTime>   # تخزين نسخة قديمة كنقطة بداية (بلا فرق)
  python3 scripts/شجرة_الدرايف.py --اعرض                # الشجرة الحالية

قواعد: الحذف من الدرايف ممنوع · الفحص بأمر صاحب المحل بس · ما في ولا إشي على أودو هون.
"""
import sys, os, json, glob, re, base64, datetime, hashlib

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
TREE = os.path.join(ROOT, 'drive', 'شجرة_الدرايف.json')
SNAP = os.path.join(ROOT, 'drive', 'لقطات')
CACHE = os.path.join(ROOT, 'drive', 'لقطات', 'ملفات')   # نسخ xlsx (مش بتنرفع على git)
ROOT_FOLDER = '1R3uPVHzsz-C-ebwAaE43FTg_2qUNhk4q'
# مجلّدات مسكّرة: بتنعدّ بالشجرة بس ما بتنحمّل إلا مع --كل (توفير)
مسكّر = ('ورشات مغلقة', 'ديون قديمة', 'مؤسسة ايلا مغلق')
تجاهل = ('الورشات Live',)
PROJ = '/root/.claude/projects'
SKIP_SHEETS = ('التعليمات',)


def حمّل():
    return json.load(open(TREE, encoding='utf-8')) if os.path.exists(TREE) else {'مجلّدات': {}, 'ملفات': {}}


def احفظ(t):
    os.makedirs(os.path.dirname(TREE), exist_ok=True)
    json.dump(t, open(TREE, 'w', encoding='utf-8'), ensure_ascii=False, indent=1, sort_keys=True)


def سجلات():
    return sorted(glob.glob(PROJ + '/**/*.jsonl', recursive=True) + glob.glob('/tmp/claude-0/**/tasks/*.output', recursive=True),
                  key=os.path.getmtime)


def _walk(o, cb):
    if isinstance(o, dict):
        cb(o)
        for v in o.values(): _walk(v, cb)
    elif isinstance(o, list):
        for v in o: _walk(v, cb)


def نتائج_الجرد(نافذة_دقائق=20):
    """نتائج search_files من سجلات الجلسة — **بس جولة الجرد الأخيرة** (خلال 20 دقيقة من آخر نتيجة)
    عشان ملف انحذف أو انتقل من زمان ما يضل طالع ← {id: file} (الأحدث بيغلب)."""
    names, hits = {}, []
    for p in سجلات():
        for line in open(p, errors='ignore'):
            if 'search_files' not in line and '"files"' not in line: continue
            try: d = json.loads(line)
            except Exception: continue
            ts = d.get('timestamp', '')
            def cb(o):
                if o.get('type') == 'tool_use' and 'search_files' in str(o.get('name', '')):
                    names[o.get('id')] = True
                if o.get('type') == 'tool_result' and o.get('tool_use_id') in names and ts:
                    c = o.get('content'); s = c if isinstance(c, str) else json.dumps(c, ensure_ascii=False)
                    for m in re.finditer(r'\{"files":\[.*?\](?:,"nextPageToken":"[^"]*")?\}', s):
                        try: hits.append((ts, json.loads(m.group(0))['files']))
                        except Exception: pass
            _walk(d, cb)
    if not hits: return {}
    last = max(h[0] for h in hits)
    lim = (datetime.datetime.fromisoformat(last.replace('Z', '+00:00')) - datetime.timedelta(minutes=نافذة_دقائق)).isoformat()
    out = {}
    for ts, fs in sorted(hits, key=lambda h: h[0]):
        if datetime.datetime.fromisoformat(ts.replace('Z', '+00:00')).isoformat() >= lim:
            for f in fs: out[f['id']] = f
    return out


def جرد():
    t = حمّل(); files = نتائج_الجرد()
    folders = {f['id']: f for f in files.values() if f['mimeType'].endswith('folder')}
    folders[ROOT_FOLDER] = {'title': 'ورشات', 'parentId': None}
    def path(pid):
        p = []
        while pid and pid in folders and pid != ROOT_FOLDER:
            p.append(folders[pid]['title']); pid = folders[pid].get('parentId')
        return '/'.join(reversed(p)) or '(الرئيسي)'
    inside = {fid for fid in folders if fid == ROOT_FOLDER or path(fid) != '(الرئيسي)' or folders[fid].get('parentId') == ROOT_FOLDER}
    xl = {i: f for i, f in files.items() if not f['mimeType'].endswith('folder') and f.get('parentId') in inside}
    t['مجلّدات'] = {i: {'اسم': folders[i]['title'], 'مسار': path(i)} for i in inside}
    جديد, تغيّر, انتقل, نفسه = [], [], [], []
    for i, f in xl.items():
        old = t['ملفات'].get(i)
        rec = old or {}
        rec.update({'اسم': f['title'], 'مجلّد': path(f.get('parentId')), 'تعديل_درايف': f['modifiedTime'], 'حجم_درايف': f.get('fileSize')})
        skip = f['title'] in تجاهل or (any(m in rec['مجلّد'].split('/') for m in مسكّر) and '--كل' not in sys.argv)
        if skip: نفسه.append(i)
        elif not old or not old.get('تعديل_مقروء'): جديد.append(i)
        elif old.get('تعديل_مقروء') != f['modifiedTime']: تغيّر.append(i)
        else: نفسه.append(i)
        if old and old.get('مجلّد') and old['مجلّد'] != rec['مجلّد']: انتقل.append((i, old['مجلّد'], rec['مجلّد']))
        t['ملفات'][i] = rec
    غاب = [i for i in t['ملفات'] if i not in xl]
    for i in غاب: t['ملفات'][i]['غاب_عن_الجرد'] = datetime.date.today().isoformat()
    t['آخر_جرد'] = datetime.datetime.now().isoformat(timespec='minutes')
    احفظ(t)
    ل = lambda ids: '\n'.join('   %s  %s  [%s]  %s' % (t['ملفات'][i]['تعديل_درايف'][:16].replace('T', ' '), t['ملفات'][i]['اسم'], t['ملفات'][i]['مجلّد'], i) for i in sorted(ids, key=lambda i: t['ملفات'][i]['تعديل_درايف'], reverse=True))
    print('🌳 الجرد: %d ملف · %d مجلّد' % (len(xl), len(inside)))
    if جديد: print('\n🆕 جديد أو ما انقرأ قبل (%d):\n%s' % (len(جديد), ل(جديد)))
    if تغيّر: print('\n✏️ تغيّر من آخر قراءة (%d):\n%s' % (len(تغيّر), ل(تغيّر)))
    for i, a, b in انتقل: print('📦 انتقل: %s  [%s ← %s]' % (t['ملفات'][i]['اسم'], a, b))
    for i in غاب: print('❓ مش بالجرد (انحذف أو انتقل لبرا): %s' % t['ملفات'][i]['اسم'])
    print('\n✅ ما تغيّر أو مسكّر (%d) — ما بينحمّلوا (المسكّر مع --كل)' % len(نفسه))
    print('\nلازم ينحمّل:', ' '.join(sorted(جديد + تغيّر)))


def دوّر_على_محتوى(fid):
    """آخر محتوى محمّل للملف بسجلات الجلسة (نتيجة مباشرة أو ملف tool-results)."""
    ids, best = {}, None
    for p in سجلات():
        for line in open(p, errors='ignore'):
            if fid not in line and 'download_file_content' not in line: continue
            try: d = json.loads(line)
            except Exception: continue
            def cb(o):
                nonlocal best
                if o.get('type') == 'tool_use' and 'download_file_content' in str(o.get('name', '')):
                    ids[o['id']] = (o.get('input') or {}).get('fileId')
                if o.get('type') == 'tool_result' and ids.get(o.get('tool_use_id')) == fid:
                    s = json.dumps(o.get('content'), ensure_ascii=False)
                    m = re.search(r'(/root/\.claude/projects/[^"\\ ]*tool-results/[^"\\ ]+\.txt)', s)
                    if m and os.path.exists(m.group(1)): best = ('file', m.group(1))
                    elif 'UEsDB' in s:
                        i = s.find('UEsDB'); best = ('b64', s[i:s.find('\\"', i) if '\\"' in s[i:] else s.find('"', i)])
            _walk(d, cb)
    if not best: return None
    if best[0] == 'file':
        raw = open(best[1]).read()
        try:
            j = json.loads(raw); j = json.loads(j[0]['text']) if isinstance(j, list) else j; b = j['content']
        except Exception:
            i = raw.find('UEsDB'); b = raw[i:raw.find('"', i)]
    else: b = best[1]
    return base64.b64decode(b)


def صفوف(path):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True); out = []
    for ws in wb:
        if ws.title in SKIP_SHEETS: continue
        for r in ws.iter_rows():
            vals = [c.value for c in r]
            if not any(v not in (None, '') for v in vals): continue
            vals = [v.strftime('%d/%m/%Y') if hasattr(v, 'strftime') else v for v in vals]
            while vals and vals[-1] in (None, ''): vals.pop()
            out.append({'ورقة': ws.title, 'سطر': r[0].row, 'قيم': vals})
    return out


def خزّن(fid, path=None, modified=None, بذرة=False):
    t = حمّل(); rec = t['ملفات'].setdefault(fid, {})
    os.makedirs(CACHE, exist_ok=True)
    dst = os.path.join(CACHE, fid + '.xlsx')
    if path: data = open(path, 'rb').read()
    else:
        data = دوّر_على_محتوى(fid)
        if not data: sys.exit('⛔ ما لقيت نسخة محمّلة للملف %s — حمّله أول' % fid)
    open(dst, 'wb').write(data)
    new = صفوف(dst)
    sp = os.path.join(SNAP, fid + '.json')
    old = json.load(open(sp, encoding='utf-8'))['صفوف'] if os.path.exists(sp) else None
    json.dump({'اسم': rec.get('اسم'), 'تعديل': modified or rec.get('تعديل_درايف'), 'صفوف': new}, open(sp, 'w', encoding='utf-8'), ensure_ascii=False)
    rec['تعديل_مقروء'] = modified or rec.get('تعديل_درايف'); rec['بصمة'] = hashlib.md5(data).hexdigest()[:10]
    rec['قُرئ'] = datetime.datetime.now().isoformat(timespec='minutes')
    احفظ(t)
    if بذرة: print('🌱 بذرة: %s ← %d سطر' % (rec.get('اسم', fid), len(new))); return
    print('══ %s  [%s]' % (rec.get('اسم', fid), rec.get('مجلّد', '')))
    if old is None:
        print('   (أول قراءة — كل الأسطر)'); show = [('🆕', r) for r in new if isinstance((r['قيم'] or [None])[0], (int, float))]
    else:
        key = lambda r: (r['ورقة'], r['سطر'])
        o = {key(r): r for r in old}; n = {key(r): r for r in new}
        show = [('🆕', n[k]) for k in n if k not in o] + [('✏️', n[k]) for k in n if k in o and o[k]['قيم'] != n[k]['قيم']] + [('🗑️', o[k]) for k in o if k not in n]
        if not show: print('   ✅ المحتوى ما تغيّر (بس وقت الحفظ)')
    for s, r in sorted(show, key=lambda x: (x[1]['ورقة'], x[1]['سطر'])):
        print('   %s %s:%d  %s' % (s, r['ورقة'], r['سطر'], ' | '.join('' if v is None else str(v) for v in r['قيم'])[:160]))


if __name__ == '__main__':
    a = sys.argv[1:]
    if not a or a[0] == '--اعرض':
        t = حمّل()
        for i, r in sorted(t['ملفات'].items(), key=lambda x: (x[1].get('مجلّد', ''), x[1].get('اسم', ''))):
            ok = '✅' if r.get('تعديل_مقروء') == r.get('تعديل_درايف') else '✏️'
            print('%s [%s] %s  درايف %s · مقروء %s' % (ok, r.get('مجلّد'), r.get('اسم'), (r.get('تعديل_درايف') or '')[:16], (r.get('تعديل_مقروء') or '—')[:16]))
    elif a[0] == '--جرد': جرد()
    elif a[0] == '--خزّن': خزّن(a[1], a[2] if len(a) > 2 else None)
    elif a[0] == '--بذرة': خزّن(a[1], a[2], a[3], بذرة=True)
