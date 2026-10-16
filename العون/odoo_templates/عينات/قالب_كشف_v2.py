import sys
a=open(sys.argv[1]).read()
def rep(o,n):
    global a
    assert a.count(o)==1, (a.count(o), o[:60]); a=a.replace(o,n)
# 1) ألوان أوضح شوي
rep("""<tr t-attf-style="{{ 'background:#fdf1e0;' if isref else ('background:#fdf8e0;' if isout else ('background:#eef7ee;' if l.credit else '')) }}">""",
    """<tr t-attf-style="{{ 'background:#fde3c6;' if isref else ('background:#fff1c2;' if isout else ('background:#dcf0dc;' if l.credit else '')) }}">""")
# 2) المرتجع
rep("""<t t-if="isref"><span t-out="'فاتورة مرتجع بضاعة'"/></t>""","""<t t-if="isref"><div style="font-weight:bold" t-out="'مرتجع بضاعة'"/></t>""")
# 3) الدفعات: تسمية موحّدة + التفاصيل سطر صغير
rep("""<t t-if="not isdon"><span t-out="' '.join((l.name or '').split())"/></t>""",
"""<t t-if="not isdon">
                <t t-set="jt" t-value="l.move_id.journal_id.type"/>
                <t t-set="nm" t-value="' '.join((l.name or '').split())"/>
                <t t-set="det" t-value="nm.split(': ', 1)[1] if ': ' in nm else nm"/>
                <t t-set="lbl" t-value="('صرف نقدي' if l.debit else 'دفعة نقدية') if jt == 'cash' else (('دفعة شيك' if 'شيك' in det else ('صرف' if l.debit else 'دفعة CliQ')) if jt == 'bank' else ('خصم' if 'خصم' in det else (('صرف نقدي' if l.debit else 'دفعة نقدية') if ('كاش' in det or 'نقد' in det or 'سحب' in det) else '')))"/>
                <t t-foreach="['دفعة نقدية', 'دفعة تحويل كليك', 'دفعة تحويل', 'دفعة شيك', 'صرف نقدي', 'دفعة كاش', 'سحب كاش', 'خصم']" t-as="pre"><t t-if="det.startswith(pre)"><t t-set="det" t-value="det[len(pre):]"/></t></t>
                <t t-set="det" t-value="det.strip().lstrip('—-·:').strip()"/>
                <div t-if="lbl" style="font-weight:bold" t-out="lbl"/>
                <div t-if="lbl and det" style="font-size:10px;color:#555" t-out="det"/>
                <span t-if="not lbl" t-out="nm"/>
              </t>""")
# 4) الرصيد المستحق بارز (بدون ملخص الحساب — قرار صاحب المحل 07/10)
rep("""  <div style="margin-top:10px;font-size:14px;font-weight:bold;text-align:left">
    الرصيد المستحق: <span t-out="'%.3f' % (sum(lines.mapped('debit')) - sum(lines.mapped('credit')))"/> د.أ
  </div>""",
"""  <table style="width:45%;margin-top:14px;margin-right:auto;border-collapse:collapse;page-break-inside:avoid">
    <tr>
      <td style="border:2px solid #305496;background:#305496;color:#fff;padding:9px 12px;font-weight:bold;font-size:15px;text-align:center">الرصيد المستحق</td>
      <td style="border:2px solid #305496;background:#eaf0fa;padding:9px 12px;font-weight:bold;font-size:19px;text-align:center;direction:ltr"><span t-out="'%.3f' % (sum(lines.mapped('debit')) - sum(lines.mapped('credit')))"/> د.أ</td>
    </tr>
  </table>""")
open(sys.argv[2],'w').write(a)
