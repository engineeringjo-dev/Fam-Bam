# -*- coding: utf-8 -*-
from jrpc import x
KEY='alawn.report_customer_statement'
arch=open('tpl_stmt.xml',encoding='utf-8').read()
v=x('ir.ui.view','search',[('key','=',KEY)])
if v:
    x('ir.ui.view','write',v,{'arch_db':arch}); vid=v[0]; print('حُدّث القالب',vid)
else:
    vid=x('ir.ui.view','create',[{'name':'كشف حساب العميل - العون','type':'qweb','key':KEY,'arch_db':arch}])
    if isinstance(vid,list): vid=vid[0]
    print('انعمل القالب',vid)
pm=x('ir.model','search',[('model','=','res.partner')])[0]
r=x('ir.actions.report','search',[('report_name','=',KEY)])
vals={'name':'كشف حساب العميل','model':'res.partner','report_type':'qweb-pdf','report_name':KEY,
      'report_file':KEY,'print_report_name':"'كشف حساب - %s' % (object.name)",
      'binding_model_id':pm,'binding_type':'report'}
if r: x('ir.actions.report','write',r,vals); rid=r[0]; print('حُدّث التقرير',rid)
else:
    rid=x('ir.actions.report','create',[vals])
    if isinstance(rid,list): rid=rid[0]
    print('انعمل التقرير',rid)

# تحقّق بالتصيير
code=("p=env['ir.config_parameter'].sudo()\n"
      "try:\n"
      "    h=env['ir.actions.report']._render_qweb_html('%s',[19])[0]\n"
      "    h=h.decode() if isinstance(h,bytes) else str(h)\n"
      "    p.set_param('alawn.rlen',str(len(h)))\n"
      "    i=h.find('كشف حساب العميل')\n"
      "    p.set_param('alawn.rsnip',h[i-200:i+2500] if i>0 else h[:2000])\n"
      "except Exception as e:\n"
      "    p.set_param('alawn.rlen','0')\n"
      "    p.set_param('alawn.rsnip','ERR: '+repr(e)[:900])\n")%KEY
sid=x('ir.actions.server','create',[{'name':'alawn render stmt','model_id':pm,'state':'code','code':code}])
if isinstance(sid,list): sid=sid[0]
x('ir.actions.server','run',[sid],context={'active_model':'res.partner','active_id':19,'active_ids':[19]})
print('طول الناتج:',x('ir.config_parameter','search_read',[('key','=','alawn.rlen')],['value'])[0]['value'])
print(x('ir.config_parameter','search_read',[('key','=','alawn.rsnip')],['value'])[0]['value'][:2500])
x('ir.actions.server','unlink',[sid])
