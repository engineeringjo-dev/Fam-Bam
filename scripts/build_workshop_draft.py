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
 14:('500377','q','اسود ولا احمر؟ (الاحمر 1ك مش موجود — بس 1/4 و1/8)'),
 15:('500145','q','سكني @0.250 ولا برتقالي @0.350؟'),
 16:(None,'q','ما في «رش اسود» عادي — الموجود رش مقاوم للحرارة @8.000'),
 17:('500032','ok','ايطالي 1/2"×16ملم'),
 18:(None,'q','ما لقيت «مربط كوليكتر» بالكاتالوج'),
 19:('500161','q','السعر بالكاتالوج صفر — بدّه تسعير'),
 20:('500269','q','عندك صنف واحد «كاب 32 احمر وازرق» — نفصله لونين؟'),
 21:('500269','q','نفس السؤال'),
 22:('500456','ok',''),
 23:('100310','q','10م @0.250 ولا 20م @0.500؟'),
 24:('500634','q','«جدلة كتكت» @0.750 — هي المقصودة؟'),
 25:('200342','ok','شوال سافيتو'),
 26:('700012','q','عندك تعبئة 100 @1.500 — الطلب 200 حبة'),
 27:(None,'q','ما في «قفل باب» بالكاتالوج إطلاقاً'),
 28:('900920','ok','مع عصا زان'),
 29:('500449','ok',''),
 30:(None,'q','ما في «زوايا ستانلس 2×2»'),
 31:('200342','ok',''),
 32:('500041','ok','(كبس)'),
}
codes=[c for c,_,_ in M.values() if c]
pm={p['default_code']:p for p in x('product.product','search_read',
     [('default_code','in',codes)],['id','default_code','name','list_price'])}
rows=[]
for it in items:
    code,st,note = M[it['n']]
    qty = 6 if it['qty']=='6م' else (it['qty'] or 0)
    p = pm.get(code)
    price = p['list_price'] if p else None
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
