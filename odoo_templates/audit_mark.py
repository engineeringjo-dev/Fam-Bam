# -*- coding: utf-8 -*-
"""تعليم الفواتير المدقّقة — بحقل «المستند المصدر» (invoice_origin) الأصلي.
مخفي عن الزبون: قالبا الفاتورة وكشف الحساب ما بيطبعوه إطلاقاً.
ممنوع استعمال ref — قالب الكشف بيطبعه لو فيه مسافة."""
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
        m = x('account.move','read',ids,['invoice_origin'])[0]
        tag = f"{MARK} {date}" + (f" · {note}" if note else '')
        cur = (m['invoice_origin'] or '').split(' | ')
        cur = [c for c in cur if c and MARK not in c]
        x('account.move','write',ids,{'invoice_origin':' | '.join(cur+[tag])})
        out.append(n); print(f"  ✔ {n} → {' | '.join(cur+[tag])}")
    return out

def unmark(names):
    if isinstance(names,str): names=[names]
    for n in names:
        ids=x('account.move','search',[('name','=',n)])
        if not ids: continue
        m=x('account.move','read',ids,['invoice_origin'])[0]
        cur=[c for c in (m['invoice_origin'] or '').split(' | ') if c and MARK not in c]
        x('account.move','write',ids,{'invoice_origin':' | '.join(cur) or False})
        print('  ↩',n)

def status(partner=None):
    d=[('move_type','=','out_invoice'),('state','=','posted')]
    if partner: d.append(('partner_id','=',partner))
    ms=x('account.move','search_read',d,['name','partner_id','invoice_date','amount_total','invoice_origin'],order='invoice_date')
    ok=[m for m in ms if MARK in (m['invoice_origin'] or '')]
    print(f"مدقّقة: {len(ok)} من {len(ms)}")
    return ms, ok
