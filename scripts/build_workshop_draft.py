# -*- coding: utf-8 -*-
import sys, json; sys.path.insert(0,'/home/user/Fam-Bam/odoo_templates')
from jrpc import x
SP='/tmp/claude-0/-home-user-Fam-Bam/a76a5bbe-02f9-54b9-9aba-c7781c667b82/scratchpad'
items = json.load(open(SP+'/wt_items.json'))
# (كود, حالة, ملاحظة)  — حالة: ok / q
M = {
 1:('500410','ok',''), 2:('500431','ok',''), 3:('500461','ok','قطعتين'),
 4:('500422','ok',''), 5:('500447','ok','فولتراب 2×4'), 6:('500456','ok',''),
 7:('500414','ok',''), 8:('500449','q','الغطاء: شامل ولا صنف منفصل؟'),
 9:('500412','ok',''), 10:('500420','ok',''), 11:('500416','ok',''), 12:('500425','ok',''),
 13:('700021','ok','14 يارد HIGH POWER'),
 14:('500377','ok','اسود 1 ك شنل امريكي'),
 15:('500145','ok','سكني'),
 16:('200153','ok','سعر خاص 1.000 (الكاتالوج 1.500)'),
 17:('500032','ok','ايطالي 1/2"×16ملم'),
 18:('500385','ok','القصير · 1 جوز'),
 19:('500161','ok','اتسعّر 1.500 بالكاتالوج'),
 20:('500269','ok','دُمج مع البند 21 — صنف واحد ×19'),
 21:(None,'merged','اندمج بالبند 20'),
 22:('500456','ok',''),
 23:('100310','ok','10م'),
 24:('500634','ok','جدلة كتكت'),
 25:('200342','ok','شوال سافيتو'),
 26:('900061','ok','تعبئة 200'),
 27:('100038','ok','قفل 50 ملم XCORT'),
 28:('900920','ok','مع عصا زان'),
 29:('500449','ok',''),
 30:('100581','ok','زوايا مجلفن 2×2 سم'),
 31:('200342','ok',''),
 32:('500041','ok','(كبس)'),
}
OVR = {16: 1.000}          # أسعار خاصة بالفاتورة
QTY = {20: 19, 21: 0}      # دمج لوني جلدة التدكيك
codes=[c for c,_,_ in M.values() if c]
pm={p['default_code']:p for p in x('product.product','search_read',
     [('default_code','in',codes)],['id','default_code','name','list_price'])}
rows=[]
for it in items:
    code,st,note = M[it['n']]
    qty = QTY.get(it['n'], 6 if it['qty']=='6م' else (it['qty'] or 0))
    p = pm.get(code)
    price = OVR.get(it['n'], p['list_price'] if p else None)
    rows.append({'n':it['n'],'desc':it['desc'],'qty':qty,'date':it['date'],
                 'code':code,'pname':p['name'] if p else None,'price':price,
                 'total':(qty*price) if price is not None else None,'st':st,'note':note})
json.dump(rows, open(SP+'/wt_rows.json','w'), ensure_ascii=False, indent=1)
from collections import OrderedDict
byd=OrderedDict()
for r in rows: byd.setdefault(r['date'],[]).append(r)
G=0; GQ=0
for d,rs in byd.items():
    t=sum(r['total'] or 0 for r in rs); q=sum(1 for r in rs if r['st']=='q')
    G+=t; GQ+=q
    print('\n══ %s — %d بند · %.3f %s' % (d,len(rs),t,'· ⏸️ %d بدهم قرار'%q if q else '✅'))
    for r in rs:
        s='✅' if r['st']=='ok' else '⏸️'
        print(' %s %2d %-30s ×%-4g %-8s %-42s %7s %9s  %s' % (
          s, r['n'], r['desc'][:30], r['qty'], r['code'] or '——',
          (r['pname'] or '— ما لقيت مقابل —')[:42],
          '%.3f'%r['price'] if r['price'] is not None else '—',
          '%.3f'%r['total'] if r['total'] is not None else '—', r['note']))
print('\n█ المجموع المحسوب: %.3f  ·  بنود بدها قرار: %d من %d' % (G,GQ,len(rows)))
