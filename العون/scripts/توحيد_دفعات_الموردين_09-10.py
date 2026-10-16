# -*- coding: utf-8 -*-
"""توحيد كتابة دفعات الموردين (صاحب المحل 09/10) — نفس صيغة الزبائن:
«صرف نقدي الى (X)» · «صرف تحويل كليك الى (X)» · «صرف شيك الى (X)» + « — ملاحظة» (التفاصيل الموجودة بتضل بالملاحظة)."""
import sys, json, re, datetime
sys.path.insert(0, 'odoo_templates')
from jrpc import x
NEW = {
 'PCSH1/2024/00001': 'صرف نقدي الى (يزيد العمري) — رمضان 2024 (تاريخ تقريبي 01/04)',
 'PBNK1/2024/00001': 'صرف تحويل كليك الى (يزيد العمري) — حوالة بنكية · إيصال رقم 1280-800196758 (مرفق صورة الواتساب)',
 'PBNK1/2024/00002': 'صرف تحويل كليك الى (يزيد العمري)',
 'PBNK1/2024/00003': 'صرف تحويل كليك الى (يزيد العمري)',
 'PBNK1/2026/00020': 'صرف تحويل كليك الى (شركة سليم كتانة واولاد) — سند قبض رقم 9145',
 'PBNK1/2026/00021': 'صرف تحويل كليك الى (شركة سليم كتانة واولاد) — سند قبض رقم 10216',
 'PBNK1/2026/00029': 'صرف تحويل كليك الى (شركة سليم كتانة واولاد) — 24/09/2026 13:58:52',
 'PBNK1/2026/00031': 'صرف تحويل كليك الى (شركة سليم كتانة واولاد) — 30/09/2026',
}
back = []; F = 'odoo_backup/ROLLBACK_vendor_payments_09-10.json'
for name, new in NEW.items():
    p = x('account.payment', 'search_read', [['name', '=', name]], fields=['id', 'memo', 'move_id', 'partner_type', 'amount'])[0]
    assert p['partner_type'] == 'supplier', p
    lines = x('account.move.line', 'search_read', [['move_id', '=', p['move_id'][0]]], fields=['name'])
    mv = x('account.move', 'read', [p['move_id'][0]], fields=['ref'])[0]
    back.append({'name': name, 'memo': p['memo'], 'ref': mv['ref'], 'lines': lines}); json.dump(back, open(F, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    old = ' '.join((p['memo'] or '').split())
    x('account.payment', 'write', [p['id']], {'memo': new})
    for l in lines:
        nm = ' '.join((l['name'] or '').split())
        if nm == old: x('account.move.line', 'write', [l['id']], {'name': new})
        elif nm.endswith(': ' + old): x('account.move.line', 'write', [l['id']], {'name': nm[:-len(old)] + new})
    if mv['ref'] and ' '.join(mv['ref'].split()) == old: x('account.move', 'write', [mv['id']], {'ref': new})
    print('✓', name, p['amount'], '«%s»' % new)
