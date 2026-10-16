# -*- coding: utf-8 -*-
"""تنظيم السبوتات (صاحب المحل 11/10): بلا «3D» · يا عادي يا «مسح» · «كبير» = 24 واط ·
تعديل 4 أسماء + 702126 لفئة كهرباء · 12 صنف جديد (السعر 0 لحد ما يعتمد الأسعار) ·
حذف 600066 · 600191 · 600097 (بأمره — ولا سطر ولا حركة مخزون) · بيضوي واب داون بيضلّوا · ترتيب العائلة · نسخة تراجع."""
import sys, os, json, datetime, importlib.util
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.join(HERE, '..')
sys.path.insert(0, os.path.join(ROOT, 'odoo_templates'))
from jrpc import x
sys.argv = ['x', '--تجاهل-التوأم']
spec = importlib.util.spec_from_file_location('ما', os.path.join(HERE, 'مدير_الأصناف.py')); ما = importlib.util.module_from_spec(spec); spec.loader.exec_module(ما)

ELEC = 10
RENAME = [('600096', 'سبوت 7 سم داخلي ابيض'), ('600267', 'سبوت 36 واط داخلي ابيض'),
          ('600185', 'سبوت 36 واط خارجي ابيض'), ('702126', 'سبوت 48 واط خارجي ابيض')]
NEW = ['سبوت 7 سم داخلي صحراوي', 'سبوت 10 سم داخلي ابيض', 'سبوت 10 سم داخلي صحراوي', 'سبوت 36 واط داخلي صحراوي',
       'سبوت 36 واط خارجي صحراوي', 'سبوت 48 واط خارجي صحراوي', 'سبوت 7 سم مسح ابيض', 'سبوت 7 سم مسح صحراوي',
       'سبوت 10 سم مسح ابيض', 'سبوت 10 سم مسح صحراوي', 'سبوت 24 واط مسح ابيض', 'سبوت 24 واط مسح صحراوي']
DELETE = ['600066', '600191', '600097']
LOG = os.path.join(ROOT, 'odoo_backup', 'ROLLBACK_spots_%s.json' % datetime.datetime.now().strftime('%Y%m%d_%H%M%S'))
log = {'تعديل': [], 'جديد': [], 'محذوف': [], 'مؤرشف': []}
save = lambda: json.dump(log, open(LOG, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
F = ['id', 'name', 'default_code', 'barcode', 'categ_id', 'list_price', 'standard_price', 'uom_id', 'taxes_id', 'type', 'sequence']
for code, new in RENAME:
    t = x('product.template', 'search_read', [['default_code', '=', code]], fields=F)[0]
    log['تعديل'].append(t); save()
    vals = {'name': new}
    if code == '702126': vals['categ_id'] = ELEC
    x('product.template', 'write', [t['id']], vals); ما.setname(t['id'], new)
    print('✏️', code, t['name'], '←', new)
for code in DELETE:
    t = x('product.template', 'search_read', [['default_code', '=', code]], fields=F)[0]
    pid = x('product.product', 'search', [['product_tmpl_id', '=', t['id']]])
    assert not x('account.move.line', 'search_count', [['product_id', 'in', pid]]) and not x('stock.move', 'search_count', [['product_id', 'in', pid]]), code
    try:
        x('product.template', 'unlink', [t['id']]); log['محذوف'].append(t); print('🗑️', code, t['name'])
    except Exception as e:
        x('product.template', 'write', [t['id']], {'active': False}); log['مؤرشف'].append(t); print('📦 مؤرشف (ما انحذف):', code, str(e)[:80])
    save()
for n in NEW:
    tid = ما.انشاء(n, 0, ELEC, اكد=True)
    log['جديد'].append(x('product.template', 'read', [tid], fields=['id', 'name', 'default_code'])[0]); save()
ما.رتّب('سبوت')
print('✓ تراجع', os.path.relpath(LOG, ROOT))
