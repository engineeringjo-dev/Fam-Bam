# -*- coding: utf-8 -*-
"""بحث ذكي بالكاتالوج — بيلقط اختلافات الإملاء والكلفة المطابقة."""
import json, re, sys, unicodedata
from difflib import SequenceMatcher
cat = json.load(open('cat.json'))

def norm(s):
    s = unicodedata.normalize('NFKC', s or '')
    s = re.sub(r'[ً-ْٰـ]', '', s)          # تشكيل وتطويل
    for a, b in [('أ','ا'),('إ','ا'),('آ','ا'),('ى','ي'),('ة','ه'),
                 ('ؤ','و'),('ئ','ي'),('“','"'),('”','"'),('×','*')]:
        s = s.replace(a, b)
    s = re.sub(r'[()\[\]\-_/,.،]', ' ', s)
    return re.sub(r'\s+', ' ', s).strip().lower()

def skel(w):
    """هيكل الكلمة: بشيل حروف المد عشان كمبرسور≈كمبريسه"""
    w = re.sub(r'[اويه]', '', w)
    return re.sub(r'(.)\1+', r'\1', w)

STOP = set('من في على مع الى عالي جوده الجوده حبه قطعه'.split())
def toks(s):
    return [t for t in norm(s).split() if t not in STOP]

def score(q, p):
    qt, pt = toks(q), toks(p['name'])
    if not qt or not pt: return 0.0
    tot = 0.0
    for a in qt:
        best = 0.0
        for b in pt:
            if a == b: best = max(best, 1.0)
            elif skel(a) and skel(a) == skel(b): best = max(best, 0.85)   # إملاء مختلف
            elif a in b or b in a: best = max(best, 0.70)
            else:
                r = SequenceMatcher(None, a, b).ratio()
                if r > 0.75: best = max(best, r * 0.8)
        w = 2.5 if (re.search(r'\d', a) or re.search(r'[a-z]', a)) else 1.0
        tot += best * w
    den = sum(2.5 if (re.search(r'\d', a) or re.search(r'[a-z]', a)) else 1.0 for a in qt)
    return tot / den

def find(desc, cost=None, n=6):
    out = []
    for p in cat:
        s = score(desc, p)
        sp = p['standard_price'] or 0
        tag = ''
        if cost and sp:
            if abs(sp - cost) < 0.0005: s += 1.2; tag = '💰 الكلفة مطابقة تماماً'
            elif abs(sp - cost) / cost < 0.08: s += 0.4; tag = '≈ كلفة قريبة'
        if s > 0.25: out.append((s, p, tag))
    out.sort(key=lambda z: -z[0])
    return out[:n]

if __name__ == '__main__':
    desc = sys.argv[1]; cost = float(sys.argv[2]) if len(sys.argv) > 2 else None
    print('🔎 «%s»%s' % (desc, '  كلفة %.3f' % cost if cost else ''))
    for s, p, tag in find(desc, cost):
        print('   %.2f  %-14s %-46s بيع %7.3f كلفة %6.3f  %s' % (
            s, p['default_code'] or '-', p['name'][:46], p['list_price'], p['standard_price'], tag))
