# -*- coding: utf-8 -*-
"""تعليم الفواتير المدقّقة — بحقل «المرجع» (ref) الأصلي.
ظاهر وقابل للتعديل بنموذج الفاتورة · مخفي عن الزبون:
قالب الفاتورة ما بيطبع ref إطلاقاً · وقالب الكشف بيقص من «✓» فما بتظهر العلامة."""
from jrpc import x
MARK = '✓ مدقّقة'

def mark(names, date=None, note=''):
    """تعليم فواتير كمدقّقة. names: اسم أو قائمة أسماء."""
    import datetime
    date = date or datetime.date.today().isoformat()
    if isinstance(names, str): names = [names]
    out = []
    for n in names:
        ids = x('account.move','search',[('name','=',n)])
        if not ids: print('✘ ما لقيت',n); continue
        m = x('account.move','read',ids,['ref'])[0]
        tag = f"{MARK} {date}" + (f" · {note}" if note else '')
        cur = (m['ref'] or '').split(' | ')
        cur = [c for c in cur if c and MARK not in c]
        x('account.move','write',ids,{'ref':' | '.join(cur+[tag])})
        out.append(n); print(f"  ✔ {n} → {' | '.join(cur+[tag])}")
    return out

def unmark(names):
    if isinstance(names,str): names=[names]
    for n in names:
        ids=x('account.move','search',[('name','=',n)])
        if not ids: continue
        m=x('account.move','read',ids,['ref'])[0]
        cur=[c for c in (m['ref'] or '').split(' | ') if c and MARK not in c]
        x('account.move','write',ids,{'ref':' | '.join(cur) or False})
        print('  ↩',n)

def status(partner=None):
    d=[('move_type','=','out_invoice'),('state','=','posted')]
    if partner: d.append(('partner_id','=',partner))
    ms=x('account.move','search_read',d,['name','partner_id','invoice_date','amount_total','ref'],order='invoice_date')
    ok=[m for m in ms if MARK in (m['ref'] or '')]
    print(f"مدقّقة: {len(ok)} من {len(ms)}")
    return ms, ok
