# -*- coding: utf-8 -*-
from jrpc import x
vals={'name':'كشف حساب العون A4','format':'A4','orientation':'Portrait','margin_top':8,'margin_bottom':8,
      'margin_left':7,'margin_right':7,'header_line':False,'header_spacing':4,'dpi':90}
p=x('report.paperformat','search',[('name','=',vals['name'])])
if p: x('report.paperformat','write',p,vals); pid=p[0]
else:
    pid=x('report.paperformat','create',[vals])
    if isinstance(pid,list): pid=pid[0]
x('ir.actions.report','write',[784],{'paperformat_id':pid})
print('paperformat',pid,'مربوط بالتقرير 784')
