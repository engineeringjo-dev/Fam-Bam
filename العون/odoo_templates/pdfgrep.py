# -*- coding: utf-8 -*-
"""بحث موثوق داخل PDF عربي: بيرجّع الأحرف من صيغ العرض (presentation forms) للأصل."""
import subprocess, unicodedata, sys, re
def text(p):
    t=subprocess.run(['pdftotext','-layout',p,'-'],capture_output=True).stdout.decode('utf-8','replace')
    out=[]
    for ch in t:
        if 0xFB50<=ord(ch)<=0xFEFF:
            d=unicodedata.normalize('NFKD',ch)
            out.append(''.join(c for c in d if not unicodedata.combining(c)))
        else: out.append(ch)
    s=''.join(out)
    s=s.replace('‏','').replace('‎','').replace('‫','').replace('‬','')
    s=re.sub(r'[ً-ْ]','',s)          # شكل
    s=s.replace('ـ','')                         # تطويل
    return s
def has(p,word):
    w=re.sub(r'[ً-ْ]','',word)
    t=text(p)
    return w in t or w[::-1] in t
if __name__=='__main__':
    p=sys.argv[1]
    for w in sys.argv[2:]:
        print(('  ✔ موجود ' if has(p,w) else '  ✘ غير موجود ')+f'«{w}»')
