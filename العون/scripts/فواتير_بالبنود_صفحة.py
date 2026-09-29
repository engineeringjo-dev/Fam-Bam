import sys, re, datetime as dt
sys.path.insert(0,'/home/user/Fam-Bam/العون/odoo_templates')
from jrpc import x
from render import render_html
FS=float(sys.argv[1]) if len(sys.argv)>1 else 10
ids=[44,45,46,179,180,252,253,254]
ms=x('account.move','read',ids,['name','invoice_date','amount_total'])
ms.sort(key=lambda m:(m['invoice_date'],m['name']))
WD=['الاثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت','الأحد']
f3=lambda v:'{:,.3f}'.format(v)
def q(v): return ('%g' % v)
blocks=[]; tot=0; nl=0
for m in ms:
    d=dt.date.fromisoformat(m['invoice_date'])
    ls=x('account.move.line','search_read',[('move_id','=',m['id']),('display_type','=','product')],['name','quantity','price_unit','price_subtotal','sequence'],order='sequence,id')
    rows='<tr class="ih"><td colspan="3">فاتورة <span class="ltr">%s</span> &nbsp;·&nbsp; %s <span class="ltr">%s</span></td><td class="n">%s</td></tr>' % (m['name'],WD[d.weekday()],d.strftime('%d/%m/%Y'),f3(m['amount_total']))
    for i,l in enumerate(ls,1):
        nm=re.sub(r'^\[[^\]]*\]\s*','',l['name'] or '')
        rows+='<tr><td class="c">%d</td><td>%s</td><td class="c">%s</td><td class="n">%s</td></tr>' % (i,nm,q(l['quantity']),f3(l['price_subtotal']))
        nl+=1
    tot+=m['amount_total']; blocks.append((len(ls)+1,rows))
half=sum(b[0] for b in blocks)/2; acc=0; cols=['','']
for n_,r_ in blocks:
    cols[0 if acc<half else 1]+=r_; acc+=n_
TH='<thead><tr><th style="width:6%%">#</th><th>الصنف</th><th style="width:10%%">الكمية</th><th style="width:17%%">المجموع</th></tr></thead>'
rows='</tbody></table></td><td style="width:50%%;vertical-align:top;padding-right:4px"><table class="t">'+TH.replace('%%','%')+'<tbody>'
rows=cols[0]+rows+cols[1]
today=dt.date.today()
html='''<!DOCTYPE html><html dir="rtl"><head><meta charset="utf-8"/><style>
body{font-family:'Noto Naskh Arabic','Noto Sans Arabic',sans-serif;font-size:%(fs)spx;color:#1a1a1a;direction:rtl;margin:14px 22px}
.hdr{width:100%%;border-collapse:collapse}.hdr td{border:0;vertical-align:top;font-size:%(fs)spx}
h2{text-align:center;font-size:%(h)spx;margin:3px 0 1px}.sub{text-align:center;font-weight:bold;margin-bottom:1px}.note{text-align:center;color:#555;margin-bottom:4px}
table.t{width:100%%;border-collapse:collapse}.t th{background:#ddd;border:1px solid #888;padding:2px 3px}
.t td{border:1px solid #aaa;padding:1px 4px;line-height:1.25}.c{text-align:center}.n{text-align:left;direction:ltr}.ltr{direction:ltr;display:inline-block}
.ih td{background:#e9eef3;font-weight:bold;border:1px solid #777}
.tot td{background:#f3f3f3;font-weight:bold;border:2px solid #333;font-size:%(h2)spx;padding:3px 4px}
</style></head><body>
<table class="hdr"><tr>
<td style="width:30%%;text-align:right">هاتف : <span class="ltr">0799123618</span> - م.عمر<br/><span class="ltr">0790177761</span> - هاتف المحل</td>
<td style="width:40%%;text-align:center;vertical-align:middle"><div style="font-size:%(h)spx;font-weight:bold">محلات العون لمواد البناء</div></td>
<td style="width:30%%;text-align:left"><b>%(wd)s</b><br/><span class="ltr">%(td)s</span></td>
</tr></table>
<div style="border-bottom:2px solid #333;margin:3px 0 4px"></div>
<h2>فواتير مشروع مسجد علاء الدين — بالبنود</h2>
<div class="note">الفواتير بعد دفعة الـ 800 دينار (17/08/2026)</div>
<table style="width:100%%;border-collapse:collapse"><tr><td style="width:50%%;vertical-align:top;padding-left:4px"><table class="t">%(th)s<tbody>
%(rows)s
</tbody></table></td></tr></table>
<table class="t" style="margin-top:4px"><tr class="tot"><td class="c">مجموع الفواتير — %(n)d فواتير · %(nl)d صنف</td><td class="n" style="width:20%%">%(tot)s</td></tr></table>
<div style="margin-top:6px;border:2px solid #333;padding:4px;text-align:center;font-weight:bold;font-size:%(h2)spx">مجموع الفواتير: %(tot)s دينار أردني</div>
</body></html>''' % dict(fs=FS,h=FS+4,h2=FS+2,wd=WD[today.weekday()],td=today.strftime('%d/%m/%Y'),rows=rows,th=TH.replace('%%','%'),n=len(ms),nl=nl,tot=f3(tot))
out='/home/user/Fam-Bam/العون/فواتير_مسجد_علاء_الدين_بعد_800_بالبنود.pdf'
render_html(html,out)
d=open(out,'rb').read(); print('pages',len(re.findall(rb'/Type\s*/Page[^s]',d)),'lines',nl,'tot',round(tot,3))
