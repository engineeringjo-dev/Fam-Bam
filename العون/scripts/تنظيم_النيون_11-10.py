# -*- coding: utf-8 -*-
"""تنظيم النيونات (صاحب المحل 11/10 «طبّق ما سبق · الأسعار بعدين · فعّلهم كلهم نقاط البيع»):
لمبات (LED) · 600172 و600195 زي 600166 (LED مبزر) · شصي (مفرد)/(مجوز) جديدين · 702056 ← كهرباء ·
كل النيونات والشصيات available_in_pos · نسخة تراجع."""
import sys, os, json, datetime, importlib.util
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.join(HERE, '..')
sys.path.insert(0, os.path.join(ROOT, 'odoo_templates'))
from jrpc import x
sys.argv = ['x', '--تجاهل-التوأم']
spec = importlib.util.spec_from_file_location('ما', os.path.join(HERE, 'مدير_الأصناف.py')); ما = importlib.util.module_from_spec(spec); spec.loader.exec_module(ما)
ELEC = 10
RENAME = {'600034': 'نيون 10 واط (LED) قصير', '600190': 'نيون 20 واط (LED)', '600033': 'نيون 36 واط (LED)', '600032': 'نيون 40 واط (LED)',
          '600172': 'نيون 54 واط (LED مبزر)', '600166': 'نيون 75 واط (LED مبزر)', '600195': 'نيون 80 واط (LED مبزر)',
          '600107': 'شصي نيون (مجوز) ديكور ثقيل', '702056': 'شصي نيون (4 لمبات) ديكور 120 واط',
          '230047': 'نيون قاتلة ناموس (قطعة غيار) تايلندي'}
KEEP = ['700320']
NEW = [('شصي نيون (مفرد)', 0), ('شصي نيون (مجوز)', 1.5)]
LOG = os.path.join(ROOT, 'odoo_backup', 'ROLLBACK_neon_%s.json' % datetime.datetime.now().strftime('%Y%m%d_%H%M%S'))
log = {'قبل': [], 'جديد': []}; save = lambda: json.dump(log, open(LOG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
F = ['id', 'default_code', 'name', 'categ_id', 'list_price', 'available_in_pos']
for code in list(RENAME) + KEEP:
    t = x('product.template', 'search_read', [['default_code', '=', code]], fields=F)[0]; log['قبل'].append(t); save()
    vals = {'available_in_pos': True}
    if code in RENAME: vals['name'] = RENAME[code]
    if code == '702056': vals['categ_id'] = ELEC
    x('product.template', 'write', [t['id']], vals)
    if code in RENAME: ما.setname(t['id'], RENAME[code])
    print('✏️', code, t['name'], '←', vals.get('name', t['name']))
for n, p in NEW:
    tid = ما.انشاء(n, p, ELEC, اكد=True)
    x('product.template', 'write', [tid], {'available_in_pos': True})
    log['جديد'].append(x('product.template', 'read', [tid], fields=['id', 'default_code', 'name', 'list_price'])[0]); save()
ما.رتّب('نيون')
print('✓ تراجع', os.path.relpath(LOG, ROOT))
