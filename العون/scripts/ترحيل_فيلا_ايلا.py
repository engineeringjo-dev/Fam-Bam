# -*- coding: utf-8 -*-
import sys, json, importlib.util, io, contextlib
sys.path.insert(0,'odoo_templates')
from jrpc import x
spec=importlib.util.spec_from_file_location('a','/tmp/claude-0/-home-user-Fam-Bam/a76a5bbe-02f9-54b9-9aba-c7781c667b82/scratchpad/ayla.py')
mod=importlib.util.module_from_spec(spec)
with contextlib.redirect_stdout(io.StringIO()): spec.loader.exec_module(mod)
m=mod.m
PARTNER, JOURNAL = 23, 8

# فحص ازدواج إلزامي
موجود={i['invoice_date'] for i in x('account.move','search_read',
        [('partner_id','=',PARTNER),('move_type','=','out_invoice')],['invoice_date'])}

from collections import OrderedDict
فواتير=OrderedDict()
for d,c,q,pu,ن in mod.البنود: فواتير.setdefault(d,[]).append((c,q,pu,ن))
تعارض=[d for d in فواتير if d in موجود]
if تعارض: raise SystemExit('⛔ تواريخ موجودة أصلاً: %s'%تعارض)

انشئت=[]
for d,rows in sorted(فواتير.items()):
    lines=[]
    for c,q,pu,ن in rows:
        p=m[c]
        اسم=p['name']+((' — %s'%ن) if ن else '')
        lines.append([0,0,{'product_id':p['id'],'name':اسم,'quantity':q,
                           'price_unit':pu,'tax_ids':[[6,0,[]]]}])
    mid=x('account.move','create',[{'move_type':'out_invoice','partner_id':PARTNER,
          'journal_id':JOURNAL,'invoice_date':d,'date':d,'invoice_line_ids':lines}])[0]
    متوقع=round(sum(q*pu for _,q,pu,_ in rows),3)
    فعلي=x('account.move','search_read',[('id','=',mid)],['amount_total'])[0]['amount_total']
    assert abs(فعلي-متوقع)<0.001, 'خلل بالمجموع %s: متوقع %s فعلي %s'%(d,متوقع,فعلي)
    انشئت.append({'id':mid,'date':d,'total':متوقع})
    print('مسودة %s = %.3f (id %d)'%(d,متوقع,mid))

json.dump(انشئت,open('odoo_backup/ROLLBACK_ayla_2024_invoices.json','w'),ensure_ascii=False,indent=1)
ids=[a['id'] for a in انشئت]
x('account.move','action_post',[ids])
بعد=x('account.move','search_read',[('id','in',ids)],['name','invoice_date','amount_total','state'],order='invoice_date')
print('\n✅ انرحّلت:')
ك=0
for i in بعد:
    ك+=i['amount_total']
    print('  %s %s %.3f %s'%(i['invoice_date'],i['name'],i['amount_total'],i['state']))
    assert i['state']=='posted'
print('المجموع %.3f'%ك)
assert abs(ك-436.600)<0.001, ك
json.dump(بعد,open('odoo_backup/ROLLBACK_ayla_2024_invoices.json','w'),ensure_ascii=False,indent=1)
