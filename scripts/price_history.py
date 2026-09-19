# -*- coding: utf-8 -*-
"""تاريخ أسعار صنف: بأي فاتورة انباع وبأي سعر ولأي مشروع.
الاستعمال:  python3 price_hist.py 500414            (كل الزبائن)
            python3 price_hist.py 500414 19         (زبون معيّن)
            python3 price_hist.py "كوع فاتح"        (بحث بالاسم)
"""
from jrpc import x
import sys
term=sys.argv[1]; pid=int(sys.argv[2]) if len(sys.argv)>2 else None
pp=x('product.product','search_read',[('default_code','=',term)],['id','default_code','name'])
if not pp:
    pp=x('product.product','search_read',[('name','like',term)],['id','default_code','name'],limit=20)
if not pp: sys.exit('ما لقيت صنف بهالاسم/الكود')
for p in pp:
    dom=[('product_id','=',p['id']),('display_type','=','product'),('parent_state','=','posted'),
         ('move_id.move_type','in',['out_invoice','out_refund'])]
    if pid: dom.append(('partner_id','=',pid))
    ls=x('account.move.line','search_read',dom,['move_id','date','quantity','price_unit','partner_id'])
    if not ls: continue
    ls.sort(key=lambda l:str(l['date']))
    print('\n=== [%s] %s'%(p['default_code'],p['name']))
    for l in ls:
        print('  %-11s %-18s %-24s ×%-6g %.3f'%(l['date'],l['move_id'][1].split(' ')[0],
              (l['partner_id'] or [0,'-'])[1][:24],l['quantity'],l['price_unit']))
    pr=sorted({round(l['price_unit'],3) for l in ls})
    print('  الأسعار المستعملة: %s | آخر سعر: %.3f | الأقل: %.3f'%(pr,ls[-1]['price_unit'],min(pr)))
