# -*- coding: utf-8 -*-
"""مدير الكشوفات — محلات العون لمواد البناء

البوّابة الوحيدة لإصدار كشوف الحساب. أي كشف بيمرّ من هون أولاً:
المدير بفحص الحساب، وإذا طلع شي بيوقف الإصدار ويقول شو المطلوب.
ما بيطلع كشف إلا بعد ما يعتمده.

  python3 scripts/مدير_الكشوفات.py                    # فحص كل الجهات
  python3 scripts/مدير_الكشوفات.py 19                  # فحص جهة
  python3 scripts/مدير_الكشوفات.py 19 --اصدار          # فحص ثم إصدار PDF
  python3 scripts/مدير_الكشوفات.py "مشروع المدرسة" --اصدار
  python3 scripts/مدير_الكشوفات.py 20 --اصدار          # «كشف حساب منذ آخر تصفير» (الافتراضي)
  python3 scripts/مدير_الكشوفات.py 20 --اصدار --كامل  # «كشف حساب شامل»
  python3 scripts/مدير_الكشوفات.py 48 --اصدار --من 2026-09-01 --الى 2026-09-30   # «كشف حساب عن فترة»

قاعدة صاحب المحل (19/09): ممنوع تسليم كشف حساب فيه بند بسعر صفر
أو بند بلا صنف أو مسودة معلّقة، بدون تنبيه صريح ومعالجة.
وزيادة (23/09): يفحص كمان إذا الفاتورة **اندبلت وانرحّلت مرتين**.

🔴 قاعدة عامة لكل المدراء (صاحب المحل 28/09): **المرتجع بيرجع بنفس سعر بيعه لنفس الورشة** —
بنقارنه بالفواتير المرحّلة لنفس الشريك ونفس الصنف (آخر بيع قبل تاريخ الرجوع). سعر أقل بس إذا الراجع مكسور/مفتوح.
صنف ما انباع لهالورشة = ما بيرجع. التطبيق الآلي: `مسودات_الورشات.py` (سعّر_مرتجع + حاجز بالترحيل).
🔴 وثبّتها 01/10: **سعر بيع الورشة المعنية مش سعر الكتالوج** · **كل مرتجع حركة مستقلة (إشعار دائن) وكل بند بسطر لحاله بالكشف**.
مهمة مدير الكشوفات: قبل أي كشف بيفحص كل مرتجع (`فحص_المرتجعات`): بلا بيع · أغلى من البيع · أقل من البيع (بده موافقة) ·
سطر بالسالب جوّا فاتورة بيع (ممنوع — لازم إشعار دائن). أي واحد منهم بيوقف الكشف.

🔴 قواعد الصياغة بكل الورش (صاحب المحل 07/10) — من مهام مدير الكشوفات (`فحص_الصياغة`):
  1. الفاتورة بتنكتب «فاتورة بضاعة» بس. أي إشي ثاني = ملاحظة صغيرة تحتها (مرجع الفاتورة) **بس إذا صاحب المحل حكاها**.
     وإذا ما حكى ← المرجع فاضي ← الكشف بيطلع «فاتورة بضاعة» لحالها.
  2. الفاتورة اليدوية القديمة = بند اسمه «(فاتورة يدوية سابقة رقم N)» بالقوسين، بالتاريخ اللي بيعطيه صاحب المحل.
  3. المرتجع بينكتب «مرتجع بضاعة» (بقالب الكشف) — ومرجعه فاضي.
  6. الدفعات — صيغة وحدة (صاحب المحل 08/10 «وحّد الدفعات وطريقة الكتابة») · `فحص_الدفعات` · `--وحّد_الدفعات`:
       الواردة:  «دفعة نقدية من (X) الى (Y)» · «دفعة تحويل كليك من (X) الى (Y)» · «دفعة شيك من (X) الى (Y)»  [+ « — ملاحظة»]
       الصادرة:  «صرف نقدي الى (X)» [+ « بامر من (Y)»] [+ « — ملاحظة»] · سحب الحساب الشخصي: «صرف نقدي — التفصيل»
       الموردين (09/10): «صرف نقدي الى (X)» · «صرف تحويل كليك الى (X)» · «صرف شيك الى (X)» [+ « — ملاحظة»]
       النوع لازم يطابق الدفتر (الصندوق = نقدي · البنك = تحويل كليك أو شيك) — ونفس الأنواع بشيت «الدفعات» بنموذج الورشات.
       بالكشف: النوع سطر عريض («دفعة نقدية» · «دفعة تحويل كليك» · «دفعة شيك» · «صرف نقدي» · «خصم» · «مرتجع بضاعة»)
       والتفاصيل («من (X) الى (Y) — …») سطر صغير تحته. الناقص (من/الى) ← سؤال لصاحب المحل، ممنوع أخمّنه.
  4. أي ملاحظة بيحكيها صاحب المحل (لون · نوع · عدة الإيجار · …) ← **جوّا نفس خانة البند بسطر ثاني صغير**
     (اسم السطر = «اسم الصنف⏎الملاحظة») — **ممنوع خانة/سطر جديد بالفاتورة** (لا line_note).
     البراغي بالوزن: الصنف «براغي وزن» والنوع بين قوسين كملاحظة: «براغي وزن⏎(سلف درل)».
  7. 🔴 **كلمة «الدرايف» ممنوعة نهائياً** بأي مستند أو دفعة أو سطر أو نموذج (صاحب المحل 07/10).
  5. **ممنوع أي إضافة من عندي** («حي الله على راسي»): لا مرجع، لا ملاحظة، لا كلمة «الدرايف»، لا وصف من الورقة على اسم البند.
  الفحص بينبّه (⚠️) وما بيوقف الكشف لحد ما تتطبّق القواعد ورشة ورشة بموافقة صاحب المحل.
"""
import sys, os, io
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'odoo_templates'))
from jrpc import x
from collections import defaultdict

SALES = ['out_invoice', 'out_refund']
BILLS = ['in_invoice', 'in_refund']
SERVICES_CATEG = 28          # فئة الخدمات — أصنافها بتتكرر بشكل مشروع

# حالات فحصها صاحب المحل واعتمدها — تُعرض كملاحظة لا كمانع
ALLOW = {
    'BILL/2026/08/0001': 'رصيد مدوّر لكتانة — سطر بلا صنف بشكل مقصود',
    'RINV/2026/00010': 'مرتجع ثروت 21/06 — 17 بند «مرتجع يدوي» بسعر صاحب المحل (بضاعتها جوّا الفواتير اليدوية) · اعتمده 06/10',
}

# التقرير المناسب حسب طبيعة التعامل مع الجهة
REPORTS = {
    'both':     ('alawn.report_joint_statement',    'كشف حساب موحّد'),
    'customer': ('alawn.report_customer_statement', 'كشف حساب العميل'),
    'vendor':   ('alawn.report_vendor_statement',   'كشف حساب المورّد'),
}

BAR = '─' * 78


def _resolve(token):
    """رقم الجهة أو اسمها ← (id, name)."""
    if token.isdigit():
        r = x('res.partner', 'read', [int(token)], ['id', 'name'])
        if not r:
            raise SystemExit('⛔ مدير الكشوفات: ما في جهة بالرقم %s' % token)
        return r[0]['id'], r[0]['name']
    hits = x('res.partner', 'search_read', [('name', 'ilike', token)], ['id', 'name'])
    if not hits:
        raise SystemExit('⛔ مدير الكشوفات: ما لقيت جهة اسمها «%s»' % token)
    if len(hits) > 1:
        raise SystemExit('⛔ مدير الكشوفات: «%s» بتطابق %d جهات: %s'
                         % (token, len(hits), ' · '.join('%d %s' % (h['id'], h['name']) for h in hits)))
    return hits[0]['id'], hits[0]['name']


def فحص(partner_id=None):
    pdom = [('partner_id', '=', partner_id)] if partner_id else []
    out = defaultdict(list)

    def lines(extra, fields):
        return x('account.move.line', 'search_read',
                 [('display_type', '=', 'product'), ('parent_state', '=', 'posted'),
                  ('move_id.move_type', 'in', SALES + BILLS)] + extra + pdom,
                 fields, order='date')

    for l in lines([('price_unit', '=', 0)],
                   ['move_id', 'product_id', 'quantity', 'partner_id', 'date']):
        out[l['partner_id'][1]].append(
            ('سعر صفر', l['date'], l['move_id'][1],
             l['product_id'][1] if l['product_id'] else '(بلا صنف)', l['quantity']))

    for l in lines([('product_id', '=', False)],
                   ['move_id', 'name', 'quantity', 'partner_id', 'date']):
        out[l['partner_id'][1]].append(
            ('بلا صنف', l['date'], l['move_id'][1], (l['name'] or '')[:42], l['quantity']))

    # سطر قيد شريكه جهة مؤرشفة — بيختفي من رصيد المشروع بصمت
    # (لغم 19/09: commercial_partner_id بايت من زلّة parent_id القديمة)
    for l in x('account.move.line', 'search_read',
               [('parent_state', '=', 'posted'), ('partner_id.active', '=', False),
                ('account_id.account_type', 'in', ['asset_receivable', 'liability_payable'])]
               + ([('move_id.partner_id', '=', partner_id)] if partner_id else []),
               ['move_id', 'partner_id', 'debit', 'credit', 'date']):
        mv = x('account.move', 'read', [l['move_id'][0]], ['partner_id'])[0]
        owner = mv['partner_id'][1] if mv['partner_id'] else '(بلا جهة)'
        out[owner].append(
            ('شريك مؤرشف', l['date'], l['move_id'][1],
             'السطر على «%s»' % l['partner_id'][1], '%.3f' % (l['debit'] - l['credit'])))

    # فاتورة مكرّرة — نفس الجهة ونفس اليوم ونفس الصنف بنفس الكمية على مستندين
    # (لغم 23/09: نموذج الورشة بيغطي الورشة من أولها، فبنوده القديمة بتنفوتر مرتين)
    # أصناف الخدمة العامة بتتكرر بشكل مشروع (فاتورة يدوية سابقة · هدية · خصم · تبرّع)
    عام = {p['id'] for p in x('product.product', 'search_read',
           [('categ_id', '=', SERVICES_CATEG)], ['id'])}
    مفاتيح = defaultdict(list)
    for l in x('account.move.line', 'search_read',
               [('display_type', '=', 'product'), ('parent_state', '=', 'posted'),
                ('move_id.move_type', 'in', SALES), ('product_id', '!=', False)] + pdom,
               ['move_id', 'move_type', 'product_id', 'quantity', 'price_unit', 'partner_id', 'date']):
        if l['product_id'][0] in عام:
            continue
        # نوع المستند جزء من المفتاح: مرتجع بنفس يوم البيع (فاتورة + إشعار دائن) مش تكرار
        مفاتيح[(l['move_type'], l['partner_id'][0], l['date'], l['product_id'][0],
                l['quantity'], l['price_unit'])].append(l)
    for (_, pid, d, prod, qty, pu), ls in مفاتيح.items():
        مستندات = {i['move_id'][1] for i in ls}
        if len(مستندات) > 1:
            out[ls[0]['partner_id'][1]].append(
                ('فاتورة مكرّرة', d, ' + '.join(sorted(مستندات)),
                 ls[0]['product_id'][1][:42], qty))

    for m in x('account.move', 'search_read',
               [('state', '=', 'draft'), ('move_type', 'in', SALES + BILLS)] + pdom,
               ['name', 'partner_id', 'invoice_date', 'amount_total'], order='invoice_date'):
        if m['partner_id']:
            out[m['partner_id'][1]].append(
                ('مسودة معلّقة', m['invoice_date'], m['name'] or '(بلا رقم)',
                 'إجمالي %.3f' % m['amount_total'], ''))
    for جهة, نوع, d, doc, what, تفصيل in فحص_المرتجعات(partner_id):
        out[جهة].append((نوع, d, doc, what, تفصيل))
    return out


def فحص_المرتجعات(partner_id=None):
    """🔴 قاعدة ثابتة (صاحب المحل 28/09 و01/10): المرتجع بنفس سعر بيعه **للورشة المعنية** — مش سعر الكتالوج —
    وكل مرتجع **حركة مستقلة (إشعار دائن)** وكل بند **بسطر لحاله** بكشف الحساب.
    بيرجّع [(الجهة, نوع, تاريخ, مستند, الصنف, تفصيل)]."""
    pdom = [('partner_id', '=', partner_id)] if partner_id else []
    عام = {p['id'] for p in x('product.product', 'search_read', [('categ_id', '=', SERVICES_CATEG)], ['id'])}
    flds = ['move_id', 'product_id', 'quantity', 'price_unit', 'discount', 'partner_id', 'date']
    base = [('display_type', '=', 'product'), ('parent_state', '=', 'posted'), ('product_id', '!=', False)]
    def سعر(l): return round(l['price_unit'] * (1 - (l.get('discount') or 0) / 100.0), 3)
    out = []
    ر = [l for l in x('account.move.line', 'search_read', base + [('move_id.move_type', '=', 'out_refund')] + pdom, flds, order='date')
         if l['product_id'][0] not in عام]
    مبيع = defaultdict(list)          # (جهة، صنف) ← أسطر البيع
    if ر:
        for l in x('account.move.line', 'search_read',
                   base + [('move_id.move_type', '=', 'out_invoice'),
                           ('partner_id', 'in', list({l['partner_id'][0] for l in ر})),
                           ('product_id', 'in', list({l['product_id'][0] for l in ر}))], flds, order='date desc, id desc'):
            مبيع[(l['partner_id'][0], l['product_id'][0])].append(l)
    for l in ر:
        جهة, صنف, مستند = l['partner_id'][1], l['product_id'][1][:42], l['move_id'][1]
        ب = [b for b in مبيع[(l['partner_id'][0], l['product_id'][0])] if b['date'] <= l['date']]
        if not ب:
            out.append((جهة, 'مرتجع بلا بيع', l['date'], مستند, صنف, 'ما انباع لهالجهة قبل تاريخ الرجوع')); continue
        أسعار = {سعر(b) for b in ب}; ر_سعر = سعر(l); آخر = ب[0]
        if any(abs(ر_سعر - a) < 0.0005 for a in أسعار): continue
        if ر_سعر > max(أسعار) + 0.0005:
            out.append((جهة, 'مرتجع أغلى من البيع', l['date'], مستند, صنف,
                        'رجع بـ %.3f — انباع بـ %.3f (%s)' % (ر_سعر, سعر(آخر), آخر['move_id'][1].split(' ')[0])))
        else:
            out.append((جهة, 'مرتجع بسعر أقل', l['date'], مستند, صنف,
                        'رجع بـ %.3f — انباع بـ %.3f (%s) · مقبول بس إذا مكسور/مفتوح بموافقته' % (ر_سعر, سعر(آخر), آخر['move_id'][1].split(' ')[0])))
    # مرتجع لازم يكون حركة لحاله — ممنوع سطر بالسالب جوّا فاتورة بيع
    for l in x('account.move.line', 'search_read',
               base + [('move_id.move_type', '=', 'out_invoice'), '|', ('quantity', '<', 0), ('price_unit', '<', 0)] + pdom, flds):
        if l['product_id'][0] in عام: continue
        out.append((l['partner_id'][1], 'مرتجع جوّا فاتورة', l['date'], l['move_id'][1], l['product_id'][1][:42],
                    'سطر بالسالب (%g × %.3f) — لازم ينزل إشعار دائن مستقل' % (l['quantity'], l['price_unit'])))
    return out


# مراجع (ملاحظة صغيرة تحت «فاتورة بضاعة») حكاها صاحب المحل بنفسه
مراجع_معتمدة = {
    'INV/2026/00139': '(عيسى مضاعين شخصي)',   # ثروت — 07/10
    'INV/2025/00012': 'محمد جهاد الصعايدة',   # حمزة شخصي — 08/10 (ملاحظة على الفاتورة كلها، مرة وحدة)
    'INV/2026/00159': 'محيسن',                # المدرسة — 08/10 (ملاحظة على الفاتورة كلها، مرة وحدة)
}


def فحص_الصياغة(partner_id=None):
    """قواعد صاحب المحل 07/10 — بترجّع {الجهة: [(مستند, شو الملاحظة)]}"""
    import re
    dom = [('move_type', 'in', SALES), ('state', '=', 'posted')]
    if partner_id: dom.append(('partner_id', '=', partner_id))
    moves = x('account.move', 'search_read', dom, ['id', 'name', 'ref', 'narration', 'partner_id', 'move_type'])
    out = defaultdict(list)
    if not moves: return out
    byid = {m['id']: m for m in moves}
    for m in moves:
        p = m['partner_id'][1] if m['partner_id'] else '—'
        if m['ref'] and مراجع_معتمدة.get(m['name']) != m['ref']: out[p].append((m['name'], 'مرجع «%s» — المسموح فاضي (أو ملاحظة حكاها صاحب المحل)' % m['ref']))
        if m['narration'] and re.sub('<[^>]+>', '', m['narration']).strip():
            out[p].append((m['name'], 'ملاحظة داخلية على المستند — لازم تنشال'))
    # «الدرايف» بأي سطر لهالجهة — حتى أسطر الدفعات والقيود
    dl = [('name', 'ilike', 'الدرايف')] + ([('partner_id', '=', partner_id)] if partner_id else [])
    for l in x('account.move.line', 'search_read', dl, ['move_id', 'partner_id', 'name']):
        out[l['partner_id'][1] if l['partner_id'] else '—'].append((l['move_id'][1].split(' ')[0], 'كلمة «الدرايف» ممنوعة: %s' % (l['name'] or '')[:50]))
    pdom = [('memo', 'ilike', 'الدرايف')] + ([('partner_id', '=', partner_id)] if partner_id else [])
    for pm in x('account.payment', 'search_read', pdom, ['name', 'partner_id', 'memo']):
        out[pm['partner_id'][1] if pm['partner_id'] else '—'].append((pm['name'], 'كلمة «الدرايف» بملاحظة الدفعة'))
    L = x('account.move.line', 'search_read', [('move_id', 'in', list(byid)), ('display_type', 'in', ['product', 'line_note'])],
          ['move_id', 'name', 'product_id', 'display_type'])
    pids = list({l['product_id'][0] for l in L if l['product_id']})
    pinfo = {p['id']: p for p in x('product.product', 'read', pids, ['name', 'default_code'])} if pids else {}
    for l in L:
        m = byid[l['move_id'][0]]; p = m['partner_id'][1] if m['partner_id'] else '—'; nm = l['name'] or ''
        if l['display_type'] == 'line_note':
            out[p].append((m['name'], 'سطر ملاحظة مستقل — لازم الملاحظة تنزل جوّا خانة البند: %s' % nm[:40])); continue
        if not l['product_id']: continue
        nm = nm.split('\n')[0]   # السطر الثاني وما بعده = ملاحظة صاحب المحل جوّا الخانة
        nm = re.sub(r'^\[[^\]]*\]\s*', '', nm)   # «[الكود] » مش إضافة — الكود لازم يبيّن بالكشف (صاحب المحل 08/10)
        pr = pinfo.get(l['product_id'][0], {})
        if pr.get('default_code') == '900900':
            if not (re.fullmatch(r'\(فاتورة يدوية سابقة رقم \d+\)', nm.strip()) or nm.startswith('فاتورة بضاعة (')
                    or (m.get('move_type') == 'out_refund' and nm.strip() == 'مرتجع بضاعة')):   # مرتجع بالمبلغ بلا بنود (غالب 08/10)
                out[p].append((m['name'], 'الفاتورة اليدوية لازم «(فاتورة يدوية سابقة رقم N)» — هلّق: %s' % nm[:50]))
        elif nm.strip() != (pr.get('name') or '').strip():
            out[p].append((m['name'], 'إضافة على اسم البند (الملاحظة بسطر ثاني جوّا الخانة): %s' % nm[:60]))
    return out


# ───── الدفعات — صيغة وحدة (صاحب المحل 08/10) ─────
import re as _re
_P = r'\([^()]+\)'
صيغ_الدفعات = {   # «من/الى»: الأولوية يكونوا موجودين، وإذا ما في تفاصيل عادي يضلّوا فاضيين (صاحب المحل 08/10)
    'in':   _re.compile(r'^(دفعة نقدية|دفعة تحويل كليك|دفعة شيك)( من %s)?( الى %s)?( — .+)?$' % (_P, _P)),
    'out':  _re.compile(r'^صرف نقدي( الى %s)?( بامر من %s)?( — .+)?$' % (_P, _P)),
    'sup':  _re.compile(r'^(صرف نقدي|صرف تحويل كليك|صرف شيك)( الى %s)?( — .+)?$' % _P),   # دفعات الموردين (09/10)
    'misc': _re.compile(r'^(صرف نقدي( الى %s)?( — .+)?|دفعة نقدية( من %s)?( الى %s)?( — .+)?|خصم( .+)?)$' % (_P, _P, _P)),
}
ناقص_تفاصيل = defaultdict(list)   # دفعات بلا «من/الى» — مقبولة، بتنعرض كمعلومة بس
دفعات_معتمدة = ('تبرّع من فاعل خير', 'باقي من ')   # تبرّع المسجد · رصيد محوّل من ورشة سابقة — صيغ خاصة معتمدة


def صيغة_موحّدة(نص, نوع, دفتر):
    """بترجّع الصيغة الموحّدة لملاحظة دفعة إذا بتنعمل آلياً بلا معلومة جديدة — وإلا None."""
    n = ' '.join((نص or '').split())
    if دفتر == 'bank': n = _re.sub(r'^دفعة تحويل (?!كليك)', 'دفعة تحويل كليك ', n)
    if دفتر == 'bank': n = _re.sub(r'^دفعة كليك(?= |$)', 'دفعة تحويل كليك', n)
    n = _re.sub(r'^دفعة كاش(?= |$)', 'دفعة نقدية', n)
    if دفتر != 'bank': n = _re.sub(r'^دفعة من (?=[^()])', 'دفعة نقدية من ', n)
    n = _re.sub(r'^(دفعة نقدية|دفعة تحويل كليك|دفعة شيك) من ([^()—]+?) \((\d+/\d+)\)$', r'\1 من (\2) — \3', n)   # «من ثروت صباح (26/4)»
    n = _re.sub(r'^(دفعة نقدية|دفعة تحويل كليك|دفعة شيك) من ([^()—]+?)(?= الى | — |$)', r'\1 من (\2)', n)
    n = _re.sub(r'^(.* من \([^()]+\)) الى ([^()—]+?)(?= — |$)', r'\1 الى (\2)', n)
    n = _re.sub(r'^سحب كاش(?= — |$)', 'صرف نقدي', n)
    n = n.replace(' — ملاحظة: ', ' — ')
    n = _re.sub(r'^صرف نقدي الى ([^()—]+?)(?= — |$)', r'صرف نقدي الى (\1)', n)
    if دفتر == 'cash' and نوع == 'out': n = _re.sub(r'^دفعة من المحل الى ([^()—]+)$', r'صرف نقدي الى (\1)', n)
    return n if صيغ_الدفعات[نوع].match(n) else None


def _دفعات(partner_id=None):
    """كل حركة دفع/صرف على حساب الذمم (دفعات + قيود يومية) ← [(مستند, جهة, نوع, دفتر, الملاحظة, payment_id, move_id)]"""
    dom = [('state', '!=', 'cancel')] + ([('partner_id', '=', partner_id)] if partner_id else [])
    out = []
    jtype = {j['id']: j['type'] for j in x('account.journal', 'search_read', [], ['type'])}
    for p in x('account.payment', 'search_read', dom, ['name', 'memo', 'partner_id', 'payment_type', 'journal_id', 'move_id', 'partner_type']):
        jt = jtype[p['journal_id'][0]]
        k = ('sup' if p['payment_type'] == 'outbound' else 'in') if p['partner_type'] == 'supplier' else ('in' if p['payment_type'] == 'inbound' else 'out')
        out.append((p['name'], (p['partner_id'] or [0, '—'])[1], k, jt, p['memo'] or '', p['id'], p['move_id'][0]))
    dom = [('journal_id.type', '=', 'general'), ('account_id.account_type', '=', 'asset_receivable'), ('parent_state', '=', 'posted')] + \
          ([('partner_id', '=', partner_id)] if partner_id else [])
    for l in x('account.move.line', 'search_read', dom, ['move_id', 'partner_id', 'name']):
        out.append((l['move_id'][1].split(' ')[0], (l['partner_id'] or [0, '—'])[1], 'misc', 'general', l['name'] or '', None, l['move_id'][0]))
    return out


def فحص_الدفعات(partner_id=None):
    """{الجهة: [(مستند, الملاحظة الحالية, الصيغة الموحّدة أو None = ناقص معلومة)]} — بس اللي مش موحّد."""
    out = defaultdict(list); ناقص_تفاصيل.clear()
    for doc, partner, نوع, jt, memo, pid, mid in _دفعات(partner_id):
        n = ' '.join(memo.split())
        if n.startswith(دفعات_معتمدة) or صيغ_الدفعات[نوع].match(n):
            if نوع == 'in' and not n.startswith(دفعات_معتمدة) and not (' من (' in n and ' الى (' in n):
                ناقص_تفاصيل[partner].append((doc, memo))   # مقبول — بس الأولوية تنكمّل
            if نوع == 'in' and ((jt == 'cash') != n.startswith('دفعة نقدية')):
                out[partner].append((doc, memo, None))   # النوع مش مطابق للدفتر
            continue
        out[partner].append((doc, memo, صيغة_موحّدة(memo, نوع, jt)))
    return out


def وحّد_الدفعات(partner_id=None):
    """بيطبّق الصيغة الموحّدة على اللي بينعمل آلياً (ملاحظة الدفعة + أسماء أسطر القيد) — مع نسخة تراجع."""
    import json, datetime
    todo = [r for r in _دفعات(partner_id) if not ' '.join(r[4].split()).startswith(دفعات_معتمدة)
            and not صيغ_الدفعات[r[2]].match(' '.join(r[4].split())) and صيغة_موحّدة(r[4], r[2], r[3])]
    if not todo: print('✅ ما في دفعات بدها توحيد آلي'); return 0
    back = []
    f = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'odoo_backup',
                     'ROLLBACK_payments_unify_%s.json' % datetime.datetime.now().strftime('%d-%m_%H%M%S'))
    for doc, partner, نوع, jt, memo, pid, mid in todo:
        new = صيغة_موحّدة(memo, نوع, jt)
        lines = x('account.move.line', 'search_read', [('move_id', '=', mid)], ['name'])
        mv = x('account.move', 'read', [mid], ['ref'])[0]
        back.append({'doc': doc, 'payment_id': pid, 'move_id': mid, 'memo': memo, 'ref': mv['ref'], 'lines': lines})
        if pid: x('account.payment', 'write', [pid], {'memo': new})
        for l in lines:
            nm = ' '.join((l['name'] or '').split()); old = ' '.join(memo.split())
            if nm == old: x('account.move.line', 'write', [l['id']], {'name': new})
            elif nm.endswith(': ' + old): x('account.move.line', 'write', [l['id']], {'name': nm[:-len(old)] + new})   # «طريقة الدفع: الملاحظة»
        if mv['ref'] and ' '.join(mv['ref'].split()) == ' '.join(memo.split()):
            x('account.move', 'write', [mid], {'ref': new})
        json.dump(back, open(f, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)   # أول بأول
        print('  ✓ %-20s %-24s «%s» ← «%s»' % (doc, partner[:24], new, memo), flush=True)
    print('  %d دفعة توحّدت · تراجع: %s' % (len(todo), os.path.relpath(f)))
    return len(todo)


def افرز(findings):
    """يفصل الموانع الحقيقية عن الحالات المعتمدة."""
    blocking, known = {}, []
    for partner, rows in findings.items():
        keep = []
        for r in rows:
            hit = next((k for k in ALLOW if k in (r[2] or '')), None)
            (known.append((partner, hit, ALLOW[hit])) if hit else keep.append(r))
        if keep:
            blocking[partner] = keep
    return blocking, known


def طبيعة_التعامل(partner_id):
    has_sale = x('account.move', 'search_count',
                 [('partner_id', '=', partner_id), ('move_type', 'in', SALES), ('state', '=', 'posted')])
    has_bill = x('account.move', 'search_count',
                 [('partner_id', '=', partner_id), ('move_type', 'in', BILLS), ('state', '=', 'posted')])
    # دفعات/قيود على حساب الدائنين بلا فواتير شراء (مثل يزيد) = تعامل بالاتجاهين برضو (09/10)
    has_pay = x('account.move.line', 'search_count', [('partner_id', '=', partner_id), ('parent_state', '=', 'posted'),
                                                      ('account_id.account_type', '=', 'liability_payable')])
    if has_sale and (has_bill or has_pay):
        return 'both'
    if has_pay and not has_sale:
        return 'vendor'
    return 'vendor' if has_bill else 'customer'


# ═══════════ خيارات الكشف الثلاثة (صاحب المحل 12/10 — «3 ازرار بس ما بدي يطلب زيادة اشتراك») ═══════════
# زر الطباعة بأودو: «كشف حساب شامل» · «كشف حساب عن فترة» · «كشف حساب منذ آخر تصفير» — والقالب بيختار
# عميل/مورّد/موحّد لحاله (alawn.stmt_dispatch) · الباني scripts/بناء_الكشوفات_الثلاثة.py
# الافتراضي هون «منذ آخر تصفير» · --كامل ← شامل · --من YYYY-MM-DD --الى YYYY-MM-DD ← عن فترة (بلا ما نلمس خانات الجهة)
# نقطة التصفير بتنحسب جوّا القالب (آخر يوم الرصيد صار فيه صفر بآخره) · حقل x_statement_from القديم ما عاد مستعمل.
MODES = {'zero': ('alawn.stmt_zero', 'كشف حساب منذ آخر تصفير'),
         'full': ('alawn.stmt_full', 'كشف حساب شامل'),
         'period': ('alawn.stmt_period', 'كشف حساب عن فترة')}


def _arg(flag):
    return sys.argv[sys.argv.index(flag) + 1] if flag in sys.argv and sys.argv.index(flag) + 1 < len(sys.argv) else None


def نقطة_الصفر(partner_id):
    """نفس حساب القالب: اليوم اللي بعد آخر يوم صار فيه الرصيد صفر بآخره (للعرض بس)."""
    import datetime as _dt
    ls = x('account.move.line', 'search_read', [('partner_id', '=', partner_id), ('parent_state', '=', 'posted'),
            ('account_id.account_type', 'in', ['asset_receivable', 'liability_payable'])], ['date', 'balance'], order='date asc, id asc')
    run, آخر = 0.0, None
    for i, l in enumerate(ls):
        run += l['balance']
        if abs(run) < 0.001 and (i == len(ls) - 1 or ls[i + 1]['date'] != l['date']): آخر = l['date']
    return (_dt.date.fromisoformat(آخر) + _dt.timedelta(days=1)).isoformat() if آخر else None


def اصدار(partner_id, name):
    from render import render
    kind = طبيعة_التعامل(partner_id)
    data = None
    if _arg('--من') or _arg('--الى'):
        mode = 'period'; data = {'stmt_from': _arg('--من'), 'stmt_to': _arg('--الى')}
    else:
        mode = 'full' if '--كامل' in sys.argv else 'zero'
    report, label = MODES[mode]
    label = '%s (%s)' % (label, REPORTS[kind][1])
    if mode == 'zero':
        z = نقطة_الصفر(partner_id)
        print('📅 %s' % ('من %s (اليوم اللي بعد آخر تصفير) · رصيد أول المدة 0.000' % z if z else 'ما صار الرصيد صفر من قبل ← من أول حركة'))
    safe = ''.join(c for c in name if c not in '/\\:*?"<>|').strip().replace(' ', '_')
    path = '/home/user/Fam-Bam/العون/%s_%s.pdf' % (label.replace(' ', '_'), safe)
    render(report, partner_id, path, data)
    print('📄 %s → %s' % (label, path))
    return path


def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    do_issue = '--اصدار' in sys.argv or '--issue' in sys.argv

    pid, pname = (None, 'كل الجهات')
    if args:
        pid, pname = _resolve(args[0])

    print(BAR)
    print('  مدير الكشوفات — محلات العون لمواد البناء')
    print('  الطلب: كشف حساب · %s' % pname)
    print(BAR)

    blocking, known = افرز(فحص(pid))

    if known:
        print('\nℹ️  حالات معروفة ومعتمدة سابقاً (مش مانع):')
        for partner, doc, why in known:
            print('    %-26s %-18s %s' % (partner, doc, why))

    صياغة = فحص_الصياغة(pid)
    if صياغة:
        print('\n⚠️  قواعد الصياغة (07/10) — تنبيه، مش مانع لحد ما تتطبّق بموافقة صاحب المحل:')
        for partner in sorted(صياغة):
            print('  ■ %s — %d ملاحظة' % (partner, len(صياغة[partner])))
            for doc, what in صياغة[partner][:40]:
                print('     %-18s %s' % (doc, what))

    if '--وحّد_الدفعات' in sys.argv:
        print('\n🔧 توحيد صيغة الدفعات (الآلي بس):'); وحّد_الدفعات(pid)
    دفعات = فحص_الدفعات(pid)
    if دفعات:
        print('\n⚠️  صيغة الدفعات (08/10) — «دفعة نقدية/تحويل كليك/شيك [من (X)] [الى (Y)]» · «صرف نقدي [الى (X)]»:')
        for partner in sorted(دفعات):
            print('  ■ %s — %d دفعة' % (partner, len(دفعات[partner])))
            for doc, memo, new in دفعات[partner][:40]:
                print('     %-20s «%s» ← %s' % (doc, memo[:60], ('«%s»  (آلي: --وحّد_الدفعات)' % new) if new else '❓ النوع مش واضح أو مش مطابق للدفتر — سؤال لصاحب المحل'))

    if ناقص_تفاصيل:
        print('\nℹ️  دفعات بلا «من/الى» — مقبولة (ما في تفاصيل)، بس الأولوية تنكمّل إذا بتنعرف:')
        for partner in sorted(ناقص_تفاصيل):
            r = ناقص_تفاصيل[partner]
            print('  ■ %s — %s%s' % (partner, ' · '.join('%s «%s»' % (d, m) for d, m in r[:4]), ' … (%d)' % len(r) if len(r) > 4 else ''))

    if pid is None:   # «يعطيني كلشي بناء ع اتفاقاتنا» (صاحب المحل 08/10) — ملخّص جهة جهة
        names = set(صياغة) | set(دفعات) | set(blocking)
        if names:
            print('\n📋 الملخّص حسب الاتفاقات (صياغة · دفعات · موانع):')
            print('    %-30s %6s %6s %6s' % ('الجهة', 'صياغة', 'دفعات', 'موانع'))
            for n in sorted(names):
                print('    %-30s %6d %6d %6d' % (n[:30], len(صياغة.get(n, [])), len(دفعات.get(n, [])), len(blocking.get(n, []))))

    if blocking:
        print('\n⛔ الكشف موقوف — في ملاحظات لازم تتعالج أو يوافق عليها صاحب المحل:\n')
        for partner in sorted(blocking):
            print('  ■ %s' % partner)
            for kind, date, doc, what, qty in blocking[partner]:
                print('     [%-12s] %-11s %-18s %-42s %s' % (kind, date, doc, what, qty))
            print()
        print('  لا يُسلَّم الكشف قبل معالجة اللي فوق.')
        print(BAR)
        return 1

    print('\n✅ الحساب نظيف — ما في بند بسعر صفر ولا بند بلا صنف ولا مسودة معلّقة، وكل المرتجعات بسعر بيعها وبحركة مستقلة.')
    if do_issue and pid:
        print('   الكشف معتمد للتسليم.\n')
        اصدار(pid, pname)
    elif do_issue:
        print('   (الإصدار بدّه جهة محددة — مرّر رقمها أو اسمها)')
    else:
        print('   الكشف معتمد للتسليم. للإصدار: أضف --اصدار')
    print(BAR)
    return 0


if __name__ == '__main__':
    sys.exit(main())
