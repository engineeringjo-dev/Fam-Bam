# -*- coding: utf-8 -*-
from jrpc import x
KEY='alawn.view_price_history_list'
arch='''<list string="تاريخ أسعار الأصناف" create="false" edit="false" delete="false" default_order="date desc">
  <field name="date"/>
  <field name="move_id"/>
  <field name="partner_id"/>
  <field name="product_id"/>
  <field name="name"/>
  <field name="quantity" sum="الكمية"/>
  <field name="price_unit"/>
  <field name="price_subtotal" sum="الإجمالي"/>
</list>'''
v=x('ir.ui.view','search',[('key','=',KEY)])
if v: x('ir.ui.view','write',v,{'arch_db':arch}); vid=v[0]
else:
    vid=x('ir.ui.view','create',[{'name':'تاريخ أسعار الأصناف','type':'list','model':'account.move.line','key':KEY,'arch_db':arch,'priority':99}])
    if isinstance(vid,list): vid=vid[0]
print('view',vid)

dom="[('display_type','=','product'),('parent_state','=','posted'),('move_id.move_type','in',['out_invoice','out_refund'])]"
ctx="{'search_default_groupby_product':1}"
avals={'name':'تاريخ أسعار الأصناف','res_model':'account.move.line','view_mode':'list','domain':dom,
       'context':"{}", 'help':'<p>كل سعر انباع فيه أي صنف: لأي زبون · بأي فاتورة · بأي تاريخ.<br/>ابحث باسم الصنف أو كوده، أو جمّع حسب الصنف/الزبون.</p>'}
a=x('ir.actions.act_window','search',[('name','=','تاريخ أسعار الأصناف')])
if a: x('ir.actions.act_window','write',a,avals); aid=a[0]
else:
    aid=x('ir.actions.act_window','create',[avals])
    if isinstance(aid,list): aid=aid[0]
print('action',aid)
# اربط العرض بالإجراء
cur=x('ir.actions.act_window.view','search',[('act_window_id','=',aid)])
if cur: x('ir.actions.act_window.view','unlink',cur)
x('ir.actions.act_window.view','create',[{'act_window_id':aid,'view_id':vid,'view_mode':'list','sequence':1}])

m=x('ir.ui.menu','search',[('name','=','تاريخ أسعار الأصناف')])
mvals={'name':'تاريخ أسعار الأصناف','parent_id':150,'action':'ir.actions.act_window,%d'%aid,'sequence':99}
if m: x('ir.ui.menu','write',m,mvals); mid=m[0]
else:
    mid=x('ir.ui.menu','create',[mvals])
    if isinstance(mid,list): mid=mid[0]
print('menu',mid,'تحت: المحاسبة ← التقارير')
n=x('account.move.line','search_count',[('display_type','=','product'),('parent_state','=','posted'),('move_id.move_type','in',['out_invoice','out_refund'])])
print('عدد الأسطر الي رح تظهر:',n)
