# -*- coding: utf-8 -*-
"""كشوفات الحساب بثلاث خيارات (صاحب المحل 12/10 — «3 ازرار بس ما بدي يطلب زيادة اشتراك»)

زر الطباعة على الجهة فيه 3 خيارات بس، ونوع الكشف (عميل · مورّد · موحّد) بيختاره لحاله:
   كشف حساب شامل             ← كل التاريخ
   كشف حساب عن فترة          ← من/إلى من مربّع «فترة الكشف» بصفحة الجهة (x_period_from / x_period_to)
   كشف حساب منذ آخر تصفير    ← من اليوم اللي بعد آخر يوم صار فيه الرصيد صفر (بينحسب جوّا القالب)
الأسماء من خبير المحاسبة: «رصيد أول المدة» · «رصيد آخر المدة».
بلا Studio · بلا نافذة/نموذج جديد · بلا أتمتة · بلا حقول محسوبة (باقة Standard).

  python3 scripts/بناء_الكشوفات_الثلاثة.py <مجلد القوالب الأصلية>     # بيبني وبينشر

المجلد فيه orig_customer.xml · orig_vendor.xml · orig_joint.xml (القوالب قبل التعديل — نسخها بـodoo_backup/ROLLBACK_stmt3_*).
"""
import sys, os, json
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'odoo_templates'))
from jrpc import x

SRC = sys.argv[1]
OUT = os.path.join(HERE, '..', 'odoo_templates')
F = "'%.3f' % (round({0}, 3) or 0.0)"


def rep(s, a, b, n=1):
    assert s.count(a) == n, ('%d× بدل %d: %s' % (s.count(a), n, a[:90]))
    return s.replace(a, b)


def prelude(moves):
    p = """
  <t t-set="smode" t-value="stmt_mode or 'full'"/>
  <t t-set="allines" t-value="lines"/>
  <t t-set="zb" t-value="allines.mapped('balance') if smode == 'zero' else []"/>
  <t t-set="zdt" t-value="allines.mapped('date') if smode == 'zero' else []"/>
  <t t-set="zc" t-value="[sum(zb[:i + 1]) for i in range(len(zb))]"/>
  <t t-set="zlast" t-value="([False] + [zdt[i] for i in range(len(zc)) if abs(zc[i]) &lt; 0.001 and (i == len(zdt) - 1 or zdt[i + 1] != zdt[i])])[-1]"/>
  <t t-set="sfrom" t-value="((stmt_from and datetime.date.fromisoformat(stmt_from)) or o.x_period_from) if smode == 'period' else ((zlast + datetime.timedelta(days=1)) if (smode == 'zero' and zlast) else False)"/>
  <t t-set="sto" t-value="((stmt_to and datetime.date.fromisoformat(stmt_to)) or o.x_period_to) if smode == 'period' else False"/>
  <t t-set="pbad" t-value="smode == 'period' and not (sfrom and sto and sfrom &lt;= sto)"/>
  <t t-set="openbal" t-value="0 if pbad else (sum(allines.filtered(lambda z: z.date &lt; sfrom).mapped('balance')) if sfrom else 0)"/>
  <t t-set="lines" t-value="allines[:0] if pbad else allines.filtered(lambda z: (not sfrom or z.date &gt;= sfrom) and (not sto or z.date &lt;= sto))"/>
  <t t-set="mlabel" t-value="{'full': 'شامل', 'period': 'عن فترة', 'zero': 'منذ آخر تصفير'}.get(smode, '')"/>
"""
    if moves:
        p += """  <t t-set="moves" t-value="moves[:0] if pbad else moves.filtered(lambda z: (not sfrom or z.invoice_date &gt;= sfrom) and (not sto or z.invoice_date &lt;= sto))"/>
  <t t-set="refs" t-value="refs[:0] if pbad else refs.filtered(lambda z: (not sfrom or z.invoice_date &gt;= sfrom) and (not sto or z.invoice_date &lt;= sto))"/>
"""
    return p


def header(s, title):
    s = rep(s, """<span style="direction:ltr;display:inline-block" t-out="lines and lines[0].date.strftime('%d/%m/%Y') or '-'"/>""",
            """<span style="direction:ltr;display:inline-block" t-out="sfrom.strftime('%d/%m/%Y') if sfrom else (lines and lines[0].date.strftime('%d/%m/%Y') or '-')"/>""")
    s = rep(s, """<span style="direction:ltr;display:inline-block" t-out="lines and lines[-1].date.strftime('%d/%m/%Y') or '-'"/>""",
            """<span style="direction:ltr;display:inline-block" t-out="sto.strftime('%d/%m/%Y') if sto else (lines and lines[-1].date.strftime('%d/%m/%Y') or (sfrom and datetime.date.today().strftime('%d/%m/%Y')) or '-')"/>""")
    s = rep(s, '>%s</h2>' % title, '>%s</h2>\n  <div style="text-align:center;font-size:12px;color:#305496;font-weight:bold;margin-bottom:2px" t-out="mlabel"/>' % title)
    s = rep(s, '\n  <table style="width:100%;border-collapse:collapse;font-size:11.5px">', """
  <div t-if="pbad" style="margin:10px 0;padding:10px;border:2px solid #b00;background:#fdecea;font-size:13px;text-align:center;font-weight:bold">
    ⚠ عبّي «فترة الكشف» (من تاريخ · إلى تاريخ) بصفحة الحساب واحفظ، وبعدين اطبع «كشف حساب عن فترة» مرة ثانية.
  </div>
  <table style="width:100%;border-collapse:collapse;font-size:11.5px">""")
    return s


def open_row(ncols, val):
    return """
      <tr t-if="sfrom" style="background:#e8e8e8;font-weight:bold">
        <td style="border:1px solid #999;padding:3px;text-align:center;direction:ltr" t-out="sfrom.strftime('%d/%m/%Y')"/>
        <td style="border:1px solid #999;padding:3px"/>
        <td style="border:1px solid #999;padding:3px">رصيد أول المدة</td>""" + \
        '\n        <td style="border:1px solid #999;padding:3px"/>' * (ncols - 4) + """
        <td style="border:1px solid #999;padding:3px;text-align:center;direction:ltr" t-out="%s"/>
      </tr>""" % F.format(val)


def close_row(ncols, val):
    return """
      <tr t-if="sfrom or sto" style="background:#dfe7f3;font-weight:bold">
        <td style="border:1px solid #999;padding:4px;text-align:center;direction:ltr" t-out="(sto or datetime.date.today()).strftime('%d/%m/%Y')"/>
        <td style="border:1px solid #999;padding:4px"/>
        <td colspan="{0}" style="border:1px solid #999;padding:4px">رصيد آخر المدة</td>
        <td style="border:1px solid #999;padding:4px;text-align:center;direction:ltr" t-out="{1}"/>
      </tr>""".format(ncols - 3, F.format(val))


def body(path, key):
    """الأصلي (container > foreach > basic_layout) ← قالب داخلي لجهة وحدة + الأصلي بينادي عليه."""
    L = open(path, encoding='utf-8').read().rstrip().split('\n')
    assert L[1].strip() == '<t t-call="web.html_container">' and L[2].strip() == '<t t-foreach="docs" t-as="o">' and L[3].strip() == '<t t-call="web.basic_layout">'
    assert [l.strip() for l in L[-4:]] == ['</t>'] * 4
    return '\n'.join(L[3:-3])


def customer(s):
    s = rep(s, """order='invoice_date asc, name asc')"/>
  <t t-set="ordF\"""", """order='invoice_date asc, name asc')"/>""" + prelude(True) + """  <t t-set="ordF\"""")
    s = header(s, 'كشف حساب العميل')
    s = rep(s, '      <t t-set="bal" t-value="0"/>\n', '      <t t-set="bal" t-value="openbal"/>' + open_row(7, 'openbal') + '\n')
    s = s.replace("t-out=\"'%.3f' % bal\"", 't-out="%s"' % F.format('bal'))
    tot = "(sum(lines.mapped('debit')) - sum(lines.mapped('credit')))"
    s = rep(s, "'%.3f' % " + tot, F.format('openbal + ' + tot), 2)
    s = rep(s, "      </tr>\n    </tbody>\n  </table>", "      </tr>" + close_row(7, 'openbal + ' + tot) + "\n    </tbody>\n  </table>")
    s = rep(s, """  <p style="page-break-before:always"/>
  <h3 style="font-size:15px;margin:4px 0 8px 0">تفاصيل الفواتير</h3>""", """  <t t-if="moves">
  <p style="page-break-before:always"/>
  <h3 style="font-size:15px;margin:4px 0 8px 0">تفاصيل الفواتير</h3>
  </t>""")
    return s


def vendor(s):
    s = rep(s, """order='invoice_date asc, name asc')"/>
  <t t-set="ordF\"""", """order='invoice_date asc, name asc')"/>""" + prelude(True) + """  <t t-set="ordF\"""")
    s = header(s, 'كشف حساب المورّد')
    s = rep(s, '      <t t-set="bal" t-value="0"/>\n', '      <t t-set="bal" t-value="-openbal"/>' + open_row(7, '-openbal') + '\n')
    s = s.replace("t-out=\"'%.3f' % bal\"", 't-out="%s"' % F.format('bal'))
    tot = "(sum(lines.mapped('credit')) - sum(lines.mapped('debit')))"
    s = rep(s, "'%.3f' % " + tot, F.format('tot_v'), 2)
    s = rep(s, "'%.3f' % -" + tot, F.format('-tot_v'))
    s = rep(s, '<t t-set="vbal" t-value="' + tot + '"/>', '<t t-set="tot_v" t-value="-openbal + ' + tot + '"/><t t-set="vbal" t-value="tot_v"/>')
    # المجموع بسطر الإجمالي بينحسب قبل vbal ← نعرّفه بأول الجدول
    s = rep(s, '      <t t-set="bal" t-value="-openbal"/>', '      <t t-set="tot_v" t-value="-openbal + ' + tot + '"/>\n      <t t-set="bal" t-value="-openbal"/>')
    s = rep(s, "      </tr>\n    </tbody>\n  </table>", "      </tr>" + close_row(7, 'tot_v') + "\n    </tbody>\n  </table>")
    return s


def joint(s):
    s = rep(s, """('parent_state','=','posted')], order='date asc, id asc')"/>
  <t t-set="rl\"""", """('parent_state','=','posted')], order='date asc, id asc')"/>""" + prelude(False) + """  <t t-set="rl\"""")
    s = rep(s, """<t t-set="net" t-value="sum(lines.mapped('debit')) - sum(lines.mapped('credit'))"/>""",
            """<t t-set="net" t-value="openbal + sum(lines.mapped('debit')) - sum(lines.mapped('credit'))"/>""")
    s = header(s, 'كشف حساب موحّد')
    s = rep(s, '      <t t-set="bal" t-value="0"/>\n', '      <t t-set="bal" t-value="openbal"/>' + open_row(6, 'openbal') + '\n')
    s = s.replace("t-out=\"'%.3f' % bal\"", 't-out="%s"' % F.format('bal'))
    s = rep(s, "t-out=\"'%.3f' % net\"", 't-out="%s"' % F.format('net'), 2)
    s = rep(s, "      </tr>\n    </tbody>\n  </table>", "      </tr>" + close_row(6, 'net') + "\n    </tbody>\n  </table>")
    return s


def wrap(name, inner_key):
    return ('<t t-name="%s">\n<t t-call="web.html_container">\n<t t-foreach="docs" t-as="o">\n'
            '<t t-call="%s"/>\n</t>\n</t>\n</t>') % (name, inner_key)


DISPATCH = """<t t-name="alawn.stmt_dispatch">
<t t-call="web.html_container">
<t t-foreach="docs" t-as="o">
  <t t-set="k_s" t-value="o.env['account.move'].search_count([('partner_id','=',o.id),('move_type','in',['out_invoice','out_refund']),('state','=','posted')])"/>
  <t t-set="k_b" t-value="o.env['account.move'].search_count([('partner_id','=',o.id),('move_type','in',['in_invoice','in_refund']),('state','=','posted')])"/>
  <t t-set="k_p" t-value="o.env['account.move.line'].search_count([('partner_id','=',o.id),('parent_state','=','posted'),('account_id.account_type','=','liability_payable')])"/>
  <t t-if="k_s and (k_b or k_p)"><t t-call="alawn.stmt_body_joint"/></t>
  <t t-elif="k_b or k_p"><t t-call="alawn.stmt_body_vendor"/></t>
  <t t-else=""><t t-call="alawn.stmt_body_customer"/></t>
</t>
</t>
</t>"""

MODES = [('alawn.stmt_full', 'full', 'كشف حساب شامل', "'كشف حساب شامل - %s' % (object.name)"),
         ('alawn.stmt_period', 'period', 'كشف حساب عن فترة', "'كشف حساب عن فترة - %s' % (object.name)"),
         ('alawn.stmt_zero', 'zero', 'كشف حساب منذ آخر تصفير', "'كشف حساب - %s' % (object.name)")]

FORM = """<data>
  <xpath expr="//notebook" position="before">
    <group string="فترة الكشف — لطباعة «كشف حساب عن فترة»">
      <field name="x_period_from" string="من تاريخ"/>
      <field name="x_period_to" string="إلى تاريخ"/>
    </group>
  </xpath>
</data>"""


def view(key, arch, name, **kw):
    v = x('ir.ui.view', 'search', [('key', '=', key)])
    if v:
        x('ir.ui.view', 'write', v, dict(arch_db=arch, **kw)); return v[0]
    r = x('ir.ui.view', 'create', [dict(name=name, type=kw.pop('type', 'qweb'), key=key, arch_db=arch, **kw)])
    return r[0] if isinstance(r, list) else r


def main():
    built = {
        'alawn.stmt_body_customer': '<t t-name="alawn.stmt_body_customer">\n' + customer(body(os.path.join(SRC, 'orig_customer.xml'), '')) + '\n</t>',
        'alawn.stmt_body_vendor': '<t t-name="alawn.stmt_body_vendor">\n' + vendor(body(os.path.join(SRC, 'orig_vendor.xml'), '')) + '\n</t>',
        'alawn.stmt_body_joint': '<t t-name="alawn.stmt_body_joint">\n' + joint(body(os.path.join(SRC, 'orig_joint.xml'), '')) + '\n</t>',
        'alawn.report_customer_statement': wrap('alawn.report_customer_statement', 'alawn.stmt_body_customer'),
        'alawn.report_vendor_statement': wrap('alawn.report_vendor_statement', 'alawn.stmt_body_vendor'),
        'alawn.report_joint_statement': wrap('alawn.report_joint_statement', 'alawn.stmt_body_joint'),
        'alawn.stmt_dispatch': DISPATCH,
    }
    for key, mode, _, _ in MODES:
        built[key] = '<t t-name="%s"><t t-set="stmt_mode" t-value="%r"/><t t-call="alawn.stmt_dispatch"/></t>' % (key, mode)
    os.makedirs(os.path.join(OUT, 'كشوفات_3'), exist_ok=True)
    for k, a in built.items():
        open(os.path.join(OUT, 'كشوفات_3', k + '.xml'), 'w', encoding='utf-8').write(a)
    if '--بناء_بس' in sys.argv:
        print('✓ انبنى بس (بلا نشر) ←', os.path.join(OUT, 'كشوفات_3')); return

    log = {'views': {}, 'reports': {}, 'unbound': [], 'fields': []}
    pm = x('ir.model', 'search', [('model', '=', 'res.partner')])[0]
    for fname, lbl in (('x_period_from', 'فترة الكشف — من تاريخ'), ('x_period_to', 'فترة الكشف — إلى تاريخ')):
        f = x('ir.model.fields', 'search', [('model', '=', 'res.partner'), ('name', '=', fname)])
        if not f:
            f = x('ir.model.fields', 'create', [{'name': fname, 'field_description': lbl, 'model_id': pm, 'ttype': 'date', 'state': 'manual', 'store': True}])
        log['fields'].append(f if isinstance(f, int) else f[0])
    # الداخلية أول (الأصلية والموزّع بينادوا عليها)
    order = ['alawn.stmt_body_customer', 'alawn.stmt_body_vendor', 'alawn.stmt_body_joint', 'alawn.stmt_dispatch',
             'alawn.report_customer_statement', 'alawn.report_vendor_statement', 'alawn.report_joint_statement'] + [m[0] for m in MODES]
    for k in order:
        log['views'][k] = view(k, built[k], k)
    base = x('ir.ui.view', 'search', [('model', '=', 'res.partner'), ('type', '=', 'form'), ('inherit_id', '=', False), ('mode', '=', 'primary')], order='priority, id', limit=1)
    fv = x('ir.ui.view', 'search', [('name', '=', 'alawn.partner_statement_period')])
    if fv: x('ir.ui.view', 'write', fv, {'arch_db': FORM}); fv = fv[0]
    else:
        fv = x('ir.ui.view', 'create', [{'name': 'alawn.partner_statement_period', 'type': 'form', 'model': 'res.partner', 'inherit_id': x('ir.model.data', 'search_read', [('module', '=', 'base'), ('name', '=', 'view_partner_form')], ['res_id'])[0]['res_id'], 'mode': 'extension', 'arch_db': FORM}])
        fv = fv[0] if isinstance(fv, list) else fv
    log['form_view'] = fv
    pf = x('ir.actions.report', 'search_read', [('report_name', '=', 'alawn.report_customer_statement')], ['paperformat_id'])[0]['paperformat_id'][0]
    for key, _, name, prn in MODES:
        vals = {'name': name, 'model': 'res.partner', 'report_type': 'qweb-pdf', 'report_name': key, 'report_file': key,
                'print_report_name': prn, 'binding_model_id': pm, 'binding_type': 'report', 'paperformat_id': pf}
        r = x('ir.actions.report', 'search', [('report_name', '=', key)])
        if r: x('ir.actions.report', 'write', r, vals); r = r[0]
        else:
            r = x('ir.actions.report', 'create', [vals]); r = r[0] if isinstance(r, list) else r
        log['reports'][key] = r
    # الأزرار القديمة بتنشال من قائمة الطباعة (التقارير نفسها بتضل)
    for old in ('alawn.report_customer_statement', 'alawn.report_customer_statement_full', 'alawn.report_vendor_statement', 'alawn.report_joint_statement'):
        r = x('ir.actions.report', 'search_read', [('report_name', '=', old)], ['binding_model_id'])
        if r and r[0]['binding_model_id']:
            x('ir.actions.report', 'write', [r[0]['id']], {'binding_model_id': False}); log['unbound'].append(r[0]['id'])
    json.dump(log, open(os.path.join(HERE, '..', 'odoo_backup', 'ROLLBACK_stmt3_ids_12-10.json'), 'w'), ensure_ascii=False, indent=1)
    from render import audit; audit()
    print('✓ انشر:', json.dumps(log, ensure_ascii=False))


if __name__ == '__main__':
    main()
