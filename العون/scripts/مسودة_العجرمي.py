# -*- coding: utf-8 -*-
"""مسودة مشروع ايمن العجرمي — البنود الجديدة 23/09/2026 (الباقي مرحّل سابقاً)"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'odoo_templates'))
from jrpc import x

PARTNER = 21
البنود = [   # (تاريخ, كود, كمية, سعر, ملاحظة السطر)
 ('2026-09-23','500533',20,1.500,'20 متر'),
 ('2026-09-23','500414', 5,2.000,''),
 ('2026-09-23','500464', 6,3.300,''),
 ('2026-09-23','500370', 4,10.000,''),
 ('2026-09-23','500368', 7,5.000,''),
 ('2026-09-23','500409',13,5.250,'13 متر'),
 ('2026-09-23','500413', 2,4.600,''),
 ('2026-09-23','500378', 3,4.500,''),
]
codes=sorted({b[1] for b in البنود})
prods=x('product.product','search_read',[('default_code','in',codes)],
        ['id','default_code','name','list_price'],context={'active_test':False})
m={p['default_code']:p for p in prods}
assert not [c for c in codes if c not in m], [c for c in codes if c not in m]
مؤرشف=[c for c in codes if not x('product.product','search_count',[('default_code','=',c)])]
assert not مؤرشف, 'أكواد مؤرشفة: %s'%مؤرشف

if __name__ == '__main__':
    print('| # | التاريخ | الكود | الصنف | كمية | إفرادي | إجمالي |')
    print('|--:|---|---|---|--:|--:|--:|')
    ك=0.0
    for i,(d,c,q,pu,n) in enumerate(البنود,1):
        t=q*pu; ك+=t
        اسم=m[c]['name']+((' «%s»'%n) if n else '')
        فرق='' if abs(m[c]['list_price']-pu)<0.0005 else ' ⚠️كتالوج %.3f'%m[c]['list_price']
        print('| %d | %s/%s | `%s` | %s%s | %g | %.3f | **%.3f** |'%(i,d[8:],d[5:7],c,اسم,فرق,q,pu,t))
    print('| | | | **المجموع** | | | **%.3f** |'%ك)
    print('\nرصيد العجرمي الحالي 670.900 → **%.3f**'%(670.9+ك))
