# -*- coding: utf-8 -*-
from jrpc import x
import base64
KEY='alawn.report_invoice_alawn'
arch=open('tpl_inv.xml',encoding='utf-8').read()
v=x('ir.ui.view','search',[('key','=',KEY)])
if v: x('ir.ui.view','write',v,{'arch_db':arch}); vid=v[0]; print('حُدّث القالب',vid)
else:
    vid=x('ir.ui.view','create',[{'name':'فاتورة العون','type':'qweb','key':KEY,'arch_db':arch}])
    if isinstance(vid,list): vid=vid[0]
    print('انعمل القالب',vid)
mm=x('ir.model','search',[('model','=','account.move')])[0]
pf=x('report.paperformat','search',[('name','=','كشف حساب العون A4')])
M="['يناير','فبراير','مارس','أبريل','مايو','يونيو','يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']"
prn=("'فاتورة %s - %s %d %s %d' % (object.name, object.partner_id.name, object.invoice_date.day, "
     + M + "[object.invoice_date.month-1], object.invoice_date.year)")
vals={'name':'فاتورة (محلات العون)','model':'account.move','report_type':'qweb-pdf','report_name':KEY,
      'report_file':KEY,'print_report_name':prn,'binding_model_id':mm,'binding_type':'report'}
if pf: vals['paperformat_id']=pf[0]
r=x('ir.actions.report','search',[('report_name','=',KEY)])
if r: x('ir.actions.report','write',r,vals); rid=r[0]; print('حُدّث التقرير',rid)
else:
    rid=x('ir.actions.report','create',[vals])
    if isinstance(rid,list): rid=rid[0]
    print('انعمل التقرير',rid)
mid=x('account.move','search',[('name','=','INV/2026/00021')])[0]
code=("p=env['ir.config_parameter'].sudo()\n"
      "try:\n"
      "    pdf=env['ir.actions.report']._render_qweb_pdf(%d,[%d])[0]\n"
      "    p.set_param('alawn.invb64', b64encode(pdf).decode())\n"
      "except Exception as e:\n"
      "    p.set_param('alawn.invb64','')\n"
      "    p.set_param('alawn.err','ERR: '+repr(e)[:400])\n")%(rid,mid)
sid=x('ir.actions.server','create',[{'name':'inv pdf','model_id':mm,'state':'code','code':code}])
if isinstance(sid,list): sid=sid[0]
x('ir.actions.server','run',[sid],context={'active_model':'account.move','active_id':mid,'active_ids':[mid]})
b=x('ir.config_parameter','search_read',[('key','=','alawn.invb64')],['value'])
if b and b[0]['value']:
    open('/home/user/Fam-Bam/فاتورة_من_اودو.pdf','wb').write(base64.b64decode(b[0]['value']))
    print('حجم PDF:',len(base64.b64decode(b[0]['value'])))
else:
    e=x('ir.config_parameter','search_read',[('key','=','alawn.err')],['value'])
    print('✘',e[0]['value'] if e else 'بلا ناتج')
x('ir.actions.server','unlink',[sid])
x('ir.config_parameter','unlink',x('ir.config_parameter','search',[('key','in',['alawn.invb64','alawn.err'])]))
