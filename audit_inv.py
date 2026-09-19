# -*- coding: utf-8 -*-
from jrpc import x
ms=sorted(x('account.move','search_read',[('move_type','=','out_invoice')],
   ['name','partner_id','invoice_date','amount_total','state','invoice_line_ids','write_date']),
   key=lambda m:m['name'])
for m in ms:
    ls=x('account.move.line','read',m['invoice_line_ids'],['product_id','price_unit','quantity'])
    pids=[l['product_id'][0] for l in ls if l['product_id']]
    pp={p['id']:p for p in x('product.product','read',pids,['default_code','name','list_price'])} if pids else {}
    diff=sum(1 for l in ls if l['product_id'] and abs(pp[l['product_id'][0]]['list_price']-l['price_unit'])>0.0001)
    print('%-18s %-26s %s %9.3f  أسطر=%2d  مختلف عن الكاتالوج=%2d  آخر تعديل %s'%(
      m['name'], (m['partner_id'][1] if m['partner_id'] else '')[:26], m['invoice_date'],
      m['amount_total'], len(ls), diff, m['write_date']))
