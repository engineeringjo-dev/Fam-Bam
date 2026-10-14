# -*- coding: utf-8 -*-
"""تنظيم أحبال الليد (صاحب المحل 11/10 «طبّق الليد»): نمط «حبل ليد + النوع + (الطول) + اللون» ·
السكك لفئة كهرباء · صنف جديد: حبل ليد بروفايل (لفة 100 متر) صحراوي @70 (حبل طارق) · الأبيض 100 متر ما انعمل (مش مؤكّد) · نسخة تراجع."""
import sys, os, json, datetime, importlib.util
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.join(HERE, '..')
sys.path.insert(0, os.path.join(ROOT, 'odoo_templates'))
from jrpc import x
sys.argv = ['x', '--تجاهل-التوأم']
spec = importlib.util.spec_from_file_location('ما', os.path.join(HERE, 'مدير_الأصناف.py')); ما = importlib.util.module_from_spec(spec); spec.loader.exec_module(ما)
ELEC = 10
RENAME = {'600130': 'حبل ليد (متر) ابيض', '600018': 'حبل ليد (متر) صحراوي',
          '2300147': 'حبل ليد بروفايل (5 متر) حركات مع محول وريموت', '600197': 'وصلة حبل ليد (دبل + تربل)',
          '702220': 'سكة ليد بروفايل (مع غطاء)'}
TO_ELEC = ['702744', '702743', '702220']
LOG = os.path.join(ROOT, 'odoo_backup', 'ROLLBACK_led_%s.json' % datetime.datetime.now().strftime('%Y%m%d_%H%M%S'))
log = {'قبل': [], 'جديد': []}; save = lambda: json.dump(log, open(LOG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
F = ['id', 'default_code', 'name', 'categ_id', 'list_price']
for code in sorted(set(RENAME) | set(TO_ELEC)):
    t = x('product.template', 'search_read', [['default_code', '=', code]], fields=F)[0]; log['قبل'].append(t); save()
    vals = {}
    if code in RENAME: vals['name'] = RENAME[code]
    if code in TO_ELEC: vals['categ_id'] = ELEC
    x('product.template', 'write', [t['id']], vals)
    if 'name' in vals: ما.setname(t['id'], vals['name'])
    print('✏️', code, t['name'], '←', vals.get('name', t['name']), '(كهرباء)' if code in TO_ELEC else '')
tid = ما.انشاء('حبل ليد بروفايل (لفة 100 متر) صحراوي', 70, ELEC, اكد=True)
log['جديد'].append(x('product.template', 'read', [tid], fields=['id', 'default_code', 'name', 'list_price'])[0]); save()
ما.رتّب('حبل ليد')
print('✓ تراجع', os.path.relpath(LOG, ROOT))
