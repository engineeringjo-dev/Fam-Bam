"""مساعد الاتصال بأودو (alawn.odoo.com) عبر JSON-RPC.

الاستعمال:
    python3 -I script.py /path/to/creds.txt

ملف الاعتماد (creds.txt) سطور بالشكل:
    ODOO_URL=https://alawn.odoo.com
    ODOO_DB=alawn
    ODOO_API_KEY=...   (سرّي — لا يُرفع على المستودع أبداً)

بالسكربتات الثانية:
    exec(open('odoo_helper.py').read().split("# --- end helper ---")[0])
"""
import sys, re, json, ssl, urllib.request

cfg = {}
for line in open(sys.argv[1], encoding='utf-8'):
    m = re.match(r'^(ODOO_[A-Z_]+)=(.*)$', line.strip())
    if m:
        cfg[m.group(1)] = m.group(2).strip()
url, db, key = cfg['ODOO_URL'], cfg['ODOO_DB'], cfg['ODOO_API_KEY']
UID = 2  # alawn depot

# xmlrpc.client ما بيمر من بروكسي البيئة — لازم urllib مع CA bundle
ctx = ssl.create_default_context(cafile='/root/.ccr/ca-bundle.crt')
opener = urllib.request.build_opener(urllib.request.ProxyHandler(), urllib.request.HTTPSHandler(context=ctx))


def call(model, method, args, kw):
    body = json.dumps({'jsonrpc': '2.0', 'method': 'call', 'id': 1, 'params': {
        'service': 'object', 'method': 'execute_kw',
        'args': [db, UID, key, model, method, args, kw]}}).encode()
    r = json.loads(opener.open(urllib.request.Request(
        url + '/jsonrpc', body, {'Content-Type': 'application/json'}), timeout=60).read())
    if 'error' in r:
        raise SystemExit(json.dumps(r['error'].get('data', {}).get('message', r['error']), ensure_ascii=False))
    return r['result']


def q(model, dom, fields, limit=100, order='id desc'):
    return call(model, 'search_read', [dom], {'fields': fields, 'limit': limit, 'order': order})

# --- end helper ---

# أمثلة (معلّقة — ما بتنفّذ إشي):
#
# 1) مسودة فاتورة مورد (بدون ضريبة — الضريبة ضمن التكلفة):
# b = call('account.move', 'create', [{
#     'move_type': 'in_invoice', 'partner_id': PARTNER_ID, 'journal_id': 9,
#     'invoice_date': 'YYYY-MM-DD', 'date': 'YYYY-MM-DD',   # ← لازم date = invoice_date
#     'ref': 'فاتورة <المورد> رقم <الرقم>',
#     'invoice_line_ids': [(0, 0, {'product_id': PID, 'quantity': QTY, 'price_unit': PU,
#                                  'tax_ids': [(6, 0, [])], 'account_id': 134})]}], {})
#
# 2) ترحيل:  call('account.move', 'action_post', [[b]], {})
#
# 3) دفعة:
# ctx = {'active_model': 'account.move', 'active_ids': [b]}
# w = call('account.payment.register', 'create', [{'journal_id': 18, 'payment_date': 'YYYY-MM-DD', 'amount': AMT}], {'context': ctx})
# call('account.payment.register', 'action_create_payments', [[w]], {'context': ctx})
#
# 4) تعديل اسم صنف — لازم باللغتين (المستخدمين على ar_001):
# call('product.template', 'write', [[TID], {'name': NEW}], {})
# call('product.template', 'write', [[TID], {'name': NEW}], {'context': {'lang': 'ar_001'}})
