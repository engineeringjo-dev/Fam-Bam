import json, urllib.request
from pathlib import Path
_e={}
for l in Path('/home/user/Fam-Bam/العون/.env').read_text().splitlines():
    if '=' in l and not l.startswith('#'):
        k,v=l.split('=',1); _e[k.strip()]=v.strip()
URL,DB,USER,KEY=_e['ODOO_URL'],_e['ODOO_DB'],_e['ODOO_USERNAME'],_e['ODOO_API_KEY']
_uid=None
def _call(service,method,args,_tries=5):
    import time as _t
    payload=json.dumps({"jsonrpc":"2.0","method":"call",
        "params":{"service":service,"method":method,"args":args}}).encode()
    last=None
    for i in range(_tries):
        try:
            req=urllib.request.Request(URL+'/jsonrpc',data=payload,
                headers={"Content-Type":"application/json"})
            r=json.loads(urllib.request.urlopen(req,timeout=180).read())
            break
        except Exception as e:                      # انقطاع شبكة — إعادة محاولة تصاعدية
            last=e
            if i==_tries-1: raise
            _t.sleep(2**i)
    if 'error' in r:
        raise RuntimeError(r['error'].get('data',{}).get('message') or json.dumps(r['error'])[:400])
    return r['result']
def uid():
    global _uid
    if _uid is None: _uid=_call('common','authenticate',[DB,USER,KEY,{}])
    return _uid
def x(model,method,*a,**kw):
    return _call('object','execute_kw',[DB,uid(),KEY,model,method,list(a),kw])

# ── مهم: حقل name مترجَم. أي تعديل اسم لازم ينكتب على اللغتين ──
def setname(pid, name):
    for lang in ('en_US', 'ar_001'):
        x('product.template', 'write', [pid], {'name': name}, context={'lang': lang})
    return True
