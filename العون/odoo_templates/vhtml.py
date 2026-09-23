# -*- coding: utf-8 -*-
"""فحص قاطع: يطبع التقرير HTML (نص مضبوط، بدون التباس استخراج PDF) ويفتش فيه."""
from jrpc import x
import base64
def html(report_name, res_id):
    rid=x('ir.actions.report','search',[('report_name','=',report_name)])[0]
    pm=x('ir.model','search',[('model','=','res.partner')])[0]
    code=("p=env['ir.config_parameter'].sudo()\n"
          "try:\n"
          "    h=env['ir.actions.report']._render_qweb_html(%d,[%d])[0]\n"
          "    p.set_param('alawn.h','__'+b64encode(h).decode())\n"
          "except Exception as e:\n"
          "    p.set_param('alawn.h','ERR '+repr(e)[:300])\n")%(rid,res_id)
    sid=x('ir.actions.server','create',[{'name':'tmp_html','model_id':pm,'state':'code','code':code}])
    if isinstance(sid,list): sid=sid[0]
    try:
        x('ir.actions.server','run',[sid],context={'active_model':'res.partner','active_id':1,'active_ids':[1]})
        v=x('ir.config_parameter','search_read',[('key','=','alawn.h')],['value'])[0]['value']
        if not v.startswith('__'): raise RuntimeError(v[:300])
        return base64.b64decode(v[2:]).decode('utf-8')
    finally:
        x('ir.actions.server','unlink',[sid])
        assert not x('ir.actions.server','search_count',[('id','=',sid)]), 'الإجراء ما انحذف!'
        p=x('ir.config_parameter','search',[('key','=','alawn.h')])
        if p: x('ir.config_parameter','unlink',p)
