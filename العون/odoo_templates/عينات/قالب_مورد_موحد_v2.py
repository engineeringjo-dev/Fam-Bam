# -*- coding: utf-8 -*-
"""تطبيق شكل كشف الحساب المعتمد (07–08/10) على كشف المورّد والكشف الموحّد — صاحب المحل 09/10
البيان موحّد (فاتورة بضاعة/شراء · مرتجع · دفعة نقدية/تحويل كليك/شيك · صرف …) والتفاصيل سطر صغير ·
ألوان أوضح · كود الصنف قبل البند · الرصيد المستحق بصندوق بارز · بدون «ملخّص» (قراره 07/10).
python3 odoo_templates/عينات/قالب_مورد_موحد_v2.py vendor.xml joint.xml out_vendor.xml out_joint.xml"""
import sys
PAY = """<t t-set="jt" t-value="l.move_id.journal_id.type"/>
            <t t-set="nm" t-value="' '.join((l.name or '').split())"/>
            <t t-set="det" t-value="nm.split(': ', 1)[1] if ': ' in nm else nm"/>
            <t t-set="pres" t-value="['دفعة نقدية', 'دفعة تحويل كليك', 'دفعة شيك', 'صرف نقدي', 'صرف تحويل كليك', 'صرف شيك', 'خصم']"/>
            <t t-set="lbl" t-value="([p for p in pres if det.startswith(p)] or [''])[0]"/>
            <t t-set="det" t-value="det[len(lbl):] if lbl else det"/>
            <t t-set="lbl" t-value="lbl or (('صرف نقدي' if l.debit else 'دفعة نقدية') if jt == 'cash' else (('صرف تحويل كليك' if l.debit else 'دفعة تحويل كليك') if jt == 'bank' else 'قيد'))"/>
            <t t-set="det" t-value="det.strip().lstrip('—-·:').strip()"/>
            <div style="font-weight:bold" t-out="lbl"/>
            <div t-if="det" style="font-size:10px;color:#555" t-out="det"/>"""
REFNOTE = """<t t-set="rf" t-value="' '.join((l.move_id.ref or '').split('✓')[0].strip().rstrip('|').split())"/>
            <div t-if="rf" style="font-size:10px;color:#555" t-out="rf"/>"""
ITEM = """<td style="border:1px solid #bbb;padding:3px"><div><span t-if="il.product_id.default_code and il.product_id.default_code != '900900' and not (il.name or '').startswith('[')" style="direction:ltr;unicode-bidi:isolate" t-out="'[%s] ' % il.product_id.default_code"/><span t-out="' '.join((il.name or '').split('\\n')[0].split())"/></div><t t-foreach="(il.name or '').split('\\n')[1:]" t-as="nt"><div t-if="nt.strip()" style="font-size:10px;color:#555" t-out="' '.join(nt.split())"/></t></td>"""
def box(title, expr, cond=None):
    return """<table style="width:45%%;margin-top:14px;margin-right:auto;border-collapse:collapse;page-break-inside:avoid">
    <tr>
      <td style="border:2px solid #305496;background:#305496;color:#fff;padding:9px 12px;font-weight:bold;font-size:15px;text-align:center">%s</td>
      <td style="border:2px solid #305496;background:#eaf0fa;padding:9px 12px;font-weight:bold;font-size:19px;text-align:center;direction:ltr"><span t-out="%s"/> د.أ</td>
    </tr>
  </table>""" % (title, expr)

def rep(a, o, n, cnt=1):
    assert a.count(o) == cnt, (a.count(o), o[:70]); return a.replace(o, n)

v = open(sys.argv[1], encoding='utf-8').read()
v = rep(v, """<tr t-attf-style="{{ 'background:#fdf1e0;' if isref else ('background:#eef7ee;' if l.debit else '') }}">""",
        """<tr t-attf-style="{{ 'background:#fde3c6;' if isref else ('background:#dcf0dc;' if l.debit else '') }}">""")
v = rep(v, """<t t-set="rf" t-value="(l.move_id.ref or '').strip()"/>
            <t t-if="isref"><span t-out="rf or 'مرتجع بضاعة'"/></t>
            <t t-if="isinv"><span t-out="rf or 'فاتورة بضاعة'"/></t>
            <t t-if="not isref and not isinv"><span t-out="' '.join((l.move_id.ref or l.name or 'دفعة').split())"/></t>""",
        """<t t-if="isref"><div style="font-weight:bold">مرتجع بضاعة</div>""" + REFNOTE + """</t>
            <t t-if="isinv"><div style="font-weight:bold">فاتورة بضاعة</div>""" + REFNOTE + """</t>
            <t t-if="not isref and not isinv">""" + PAY + """</t>""")
v = rep(v, """<td style="border:1px solid #bbb;padding:3px" t-out="' '.join((il.name or '').split())"/>""", ITEM, 2)
v = rep(v, """<div style="margin-top:10px;font-size:14px;font-weight:bold;text-align:left">
    الرصيد المستحق علينا: <span t-out="'%.3f' % (sum(lines.mapped('credit')) - sum(lines.mapped('debit')))"/> د.أ
  </div>""", box('الرصيد المستحق علينا', "'%.3f' % (sum(lines.mapped('credit')) - sum(lines.mapped('debit')))"))
open(sys.argv[3], 'w', encoding='utf-8').write(v)

j = open(sys.argv[2], encoding='utf-8').read()
j = rep(j, """<tr t-attf-style="{{ 'background:#eef4fb;' if l.account_id.account_type == 'liability_payable' else ('background:#eef7ee;' if l.credit else '') }}">""",
        """<tr t-attf-style="{{ 'background:#fde3c6;' if mt in ['out_refund','in_refund'] else ('background:#dde8f7;' if mt == 'in_invoice' else ('background:#dcf0dc;' if (l.credit and mt not in ['out_invoice','in_invoice']) else ('background:#fff1c2;' if (l.debit and mt not in ['out_invoice','in_invoice']) else ''))) }}">""")
j = rep(j, """<t t-if="mt == 'out_invoice'"><div><b>فاتورة بيع</b> — بضاعة سحبها</div></t>
            <t t-if="mt == 'out_refund'"><div><b>مرتجع بيع</b></div></t>
            <t t-if="mt == 'in_invoice'"><div><b>فاتورة شراء</b> — بضاعة اشتريناها منه</div></t>
            <t t-if="mt == 'in_refund'"><div><b>مرتجع شراء</b></div></t>
            <t t-if="mt not in ['out_invoice','out_refund','in_invoice','in_refund']">
              <div><b t-out="'دفعنا له' if l.debit else 'قبضنا منه'"/></div>
            </t>
            <t t-set="isdoc" t-value="mt in ['out_invoice','out_refund','in_invoice','in_refund']"/>
            <t t-set="raw" t-value="rf if isdoc else (rf or (l.name or ''))"/>
            <t t-set="sub" t-value="' '.join(raw.split('✓')[0].strip().rstrip('|').rstrip('-').split())"/>
            <div t-if="sub and sub != l.move_id.name" style="font-size:10px;color:#666" t-out="sub"/>""",
        """<t t-if="mt == 'out_invoice'"><div style="font-weight:bold">فاتورة بضاعة</div><div t-if="rf" style="font-size:10px;color:#555" t-out="rf"/></t>
            <t t-if="mt == 'out_refund'"><div style="font-weight:bold">مرتجع بضاعة</div><div t-if="rf" style="font-size:10px;color:#555" t-out="rf"/></t>
            <t t-if="mt == 'in_invoice'"><div style="font-weight:bold">فاتورة شراء</div><div t-if="rf" style="font-size:10px;color:#555" t-out="rf"/></t>
            <t t-if="mt == 'in_refund'"><div style="font-weight:bold">مرتجع شراء</div><div t-if="rf" style="font-size:10px;color:#555" t-out="rf"/></t>
            <t t-if="mt not in ['out_invoice','out_refund','in_invoice','in_refund']">
            """ + PAY + """
            </t>""")
# بدون «ملخّص الطرفين» (قرار 07/10: الرصيد المستحق بارز بس)
s0 = j.index('  <h3 style="font-size:14px;margin:14px 0 6px 0">ملخّص الطرفين</h3>'); s1 = j.index('  <t t-set="adv"')
j = j[:s0] + j[s1:]
j = rep(j, """<div t-attf-style="margin-top:12px;font-size:15px;font-weight:bold;text-align:left;{{ 'color:#b00;' if net &lt; -0.0005 else '' }}">
    <t t-if="net &gt; 0.0005">الرصيد المستحق <u>علــيه</u>: <span t-out="'%.3f' % net"/> د.أ</t>
    <t t-if="net &lt; -0.0005">الرصيد المستحق <u>لــه علينا</u>: <span t-out="'%.3f' % -net"/> د.أ</t>
    <t t-if="abs(net) &lt;= 0.0005">الحساب مُسوّى — الرصيد صفر</t>
  </div>""",
        """<t t-if="net &gt;= -0.0005">""" + box('الرصيد المستحق علــيه', "'%.3f' % abs(net)") + """</t>
  <t t-if="net &lt; -0.0005">""" + box('الرصيد المستحق لــه علينا', "'%.3f' % -net").replace('#305496', '#9c2b2b').replace('#eaf0fa', '#fbeaea') + """</t>""")
open(sys.argv[4], 'w', encoding='utf-8').write(j)
print('✓ القالبين جاهزين')
