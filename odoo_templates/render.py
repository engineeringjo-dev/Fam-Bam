# -*- coding: utf-8 -*-
"""طباعة تقرير أودو PDF — مع تنظيف مضمون لإجراء الخادم المؤقت.
لا تستعمل أي طريقة ثانية: إجراءات الخادم المتروكة بتنعدّ كـ«كود مخصص» وبتهدد الاشتراك."""
from jrpc import x
import base64

def render(report_name, res_id, out_path):
    rid = x('ir.actions.report','search',[('report_name','=',report_name)])[0]
    pm  = x('ir.model','search',[('model','=','res.partner')])[0]
    code = ("p=env['ir.config_parameter'].sudo()\n"
            "try:\n"
            "    pdf=env['ir.actions.report']._render_qweb_pdf(%d,[%d])[0]\n"
            "    p.set_param('alawn.b','__'+b64encode(pdf).decode())\n"
            "except Exception as e:\n"
            "    p.set_param('alawn.b','ERR '+repr(e)[:400])\n") % (rid, res_id)
    sid = x('ir.actions.server','create',[{'name':'tmp_render','model_id':pm,'state':'code','code':code}])
    if isinstance(sid,list): sid = sid[0]
    try:
        x('ir.actions.server','run',[sid],
          context={'active_model':'res.partner','active_id':res_id,'active_ids':[res_id]})
        v = x('ir.config_parameter','search_read',[('key','=','alawn.b')],['value'])
        val = v[0]['value'] if v else ''
        if not val.startswith('__'):
            raise RuntimeError('فشل الطباعة: '+val[:300])
        open(out_path,'wb').write(base64.b64decode(val[2:]))
    finally:
        # ── التنظيف إجباري ومُتحقَّق منه ──
        x('ir.actions.server','unlink',[sid])
        assert not x('ir.actions.server','search_count',[('id','=',sid)]), \
               'خطر: إجراء الخادم %d ما انحذف — بينعدّ كود مخصص!' % sid
        p = x('ir.config_parameter','search',[('key','=','alawn.b')])
        if p: x('ir.config_parameter','unlink',p)
        audit()
    return out_path

def audit(loud=True):
    """فحص: ما في أي أثر بينعدّ كـ«كود مخصص» عند أودو."""
    sa = x('ir.actions.server','search_read',[('state','=','code')],['id','name'])
    md = {a['res_id'] for a in x('ir.model.data','search_read',
          [('model','=','ir.actions.server'),('res_id','in',[a['id'] for a in sa])],['res_id'])}
    orph = [a for a in sa if a['id'] not in md]
    mf = x('ir.model.fields','search_count',[('state','=','manual'),('compute','!=',False)])
    mm = x('ir.model','search_count',[('state','=','manual')])
    if orph or mf or mm:
        raise RuntimeError('⚠ أثر كود مخصص: إجراءات=%s · حقول compute=%d · نماذج=%d' % (orph, mf, mm))
    if loud: print('  🧹 نظيف — صفر كود مخصص')
    return True
