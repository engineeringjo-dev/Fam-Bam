# -*- coding: utf-8 -*-
"""تعديل تاريخ مستند مرحّل على أودو — محلات العون لمواد البناء

الشغلة الوحيدة لهالسكربت: تغيير **تاريخ** مستند واحد مرحّل (فاتورة · إشعار دائن · قيد دفعة)
بأمر صاحب المحل، لمّا يكون انرحّل بتاريخ غلط. ما بيلمس المبالغ ولا الأسطر ولا الشريك.

  python3 scripts/تعديل_تاريخ_مستند.py RINV/2026/00005 2026-09-25          # تجربة (بلا تغيير)
  python3 scripts/تعديل_تاريخ_مستند.py RINV/2026/00005 2026-09-25 --اكد    # التنفيذ

الخطوات: قراءة المستند ← نسخة تراجع بـ odoo_backup/ ← إعادة لمسودة ← كتابة التاريخ ← ترحيل ←
تحقّق إن المبلغ والحالة والرصيد ما تغيّروا. أي خلل بالتحقّق بيطبع ⛔ والتفاصيل.
"""
import sys, os, re, json, datetime
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'odoo_templates'))
from jrpc import x

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
BAR = '─' * 78


def مستند(اسم):
    r = x('account.move', 'search_read', [('name', '=', اسم)],
          ['id', 'name', 'state', 'move_type', 'date', 'invoice_date', 'amount_total', 'partner_id', 'ref'])
    if not r: raise SystemExit('⛔ ما في مستند اسمه %s' % اسم)
    return r[0]


def رصيد(pid):
    return x('res.partner', 'read', [pid], ['credit'])[0]['credit'] if pid else None


def عدّل(اسم, تاريخ, اكد=False):
    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', تاريخ): raise SystemExit('⛔ التاريخ لازم يكون YYYY-MM-DD')
    datetime.date.fromisoformat(تاريخ)
    m = مستند(اسم)
    pid = m['partner_id'][0] if m['partner_id'] else None
    قبل = رصيد(pid)
    print(BAR)
    print('%s — %s' % ('تعديل تاريخ' if اكد else 'تجربة (بلا --اكد)', m['name']))
    print('  النوع %s · الحالة %s · الشريك %s · المبلغ %.3f' % (m['move_type'], m['state'],
          m['partner_id'][1] if m['partner_id'] else '—', m['amount_total']))
    print('  التاريخ: %s ← %s%s' % (m['date'], تاريخ, '' if m['invoice_date'] in (False, m['date']) else
          '  (تاريخ الفاتورة كان %s)' % m['invoice_date']))
    if m['state'] != 'posted': raise SystemExit('⛔ المستند مش مرحّل (الحالة %s) — هالسكربت للمرحّل بس' % m['state'])
    if m['date'] == تاريخ and m['invoice_date'] in (False, تاريخ):
        print('✅ التاريخ أصلاً %s — ما في شي يتعدّل' % تاريخ); print(BAR); return
    if not اكد: print(BAR); return

    os.makedirs(os.path.join(ROOT, 'odoo_backup'), exist_ok=True)
    سجل = os.path.join(ROOT, 'odoo_backup', 'ROLLBACK_date_%s_%s.json' % (
        m['name'].replace('/', '-'), datetime.date.today().isoformat()))
    json.dump({'id': m['id'], 'name': m['name'], 'old_date': m['date'], 'old_invoice_date': m['invoice_date'],
               'new_date': تاريخ, 'amount_total': m['amount_total'], 'balance_before': قبل},
              open(سجل, 'w'), ensure_ascii=False, indent=1)

    x('account.move', 'button_draft', [m['id']])
    vals = {'date': تاريخ}
    if m['move_type'] != 'entry': vals['invoice_date'] = تاريخ
    x('account.move', 'write', [m['id']], vals)
    x('account.move', 'action_post', [m['id']])

    b = مستند(اسم); بعد = رصيد(pid)
    مشاكل = []
    if b['state'] != 'posted': مشاكل.append('الحالة %s' % b['state'])
    if b['date'] != تاريخ: مشاكل.append('التاريخ %s' % b['date'])
    if m['move_type'] != 'entry' and b['invoice_date'] != تاريخ: مشاكل.append('تاريخ الفاتورة %s' % b['invoice_date'])
    if abs(b['amount_total'] - m['amount_total']) > 0.001: مشاكل.append('المبلغ %.3f ≠ %.3f' % (b['amount_total'], m['amount_total']))
    if pid and abs((بعد or 0) - (قبل or 0)) > 0.001: مشاكل.append('الرصيد %.3f ≠ %.3f' % (بعد, قبل))
    if مشاكل:
        print('⛔ التحقّق فشل: ' + ' · '.join(مشاكل)); print('   التراجع: %s' % سجل); print(BAR); raise SystemExit(1)
    print('✅ %s صار بتاريخ %s · المبلغ %.3f · الرصيد %.3f (ما تغيّر)' % (b['name'], b['date'], b['amount_total'], بعد or 0))
    print('   التراجع: %s' % os.path.relpath(سجل, ROOT)); print(BAR)


if __name__ == '__main__':
    a = [s for s in sys.argv[1:] if not s.startswith('--')]
    if len(a) != 2: raise SystemExit(__doc__)
    عدّل(a[0], a[1], '--اكد' in sys.argv)
