# -*- coding: utf-8 -*-
"""أرصدة المشاريع مجمّعة حسب المقاول — محلات العون لمواد البناء

حقل الرصيد `credit` غير مخزّن بأودو، فما بينجمع بقوائم الشاشة.
هذا السكربت بيحسبه ويعطي المجموع لكل مهندس.

    python3 scripts/أرصدة_المقاولين.py
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'odoo_templates'))
from jrpc import x

المقاول = 5        # وسم الأب
BAR = '─' * 62


def main():
    tags = x('res.partner.category', 'search_read', [('parent_id', '=', المقاول)],
             ['id', 'name'], order='name')
    print(BAR)
    print('  أرصدة المشاريع حسب المقاول — محلات العون لمواد البناء')
    print(BAR)
    grand, seen = 0.0, set()
    for t in tags:
        ps = x('res.partner', 'search_read',
               [('category_id', '=', t['id']), ('category_id', '=', 11)],
               ['id', 'name', 'credit'], order='name')
        if not ps:
            continue
        tot = sum(p['credit'] for p in ps)
        print('\n👷 %-24s %d مشروع' % (t['name'], len(ps)))
        for p in ps:
            print('     %-32s %10.3f' % (p['name'][:32], p['credit']))
            seen.add(p['id'])
        print('     %-32s %10.3f' % ('— المجموع', tot))
    print('\n' + BAR)
    all_p = x('res.partner', 'search_read', [('category_id', '=', 11)], ['id', 'name', 'credit'])
    grand = sum(p['credit'] for p in all_p)
    print('  مجموع ذمم كل المشاريع (بلا تكرار): %.3f د.أ' % grand)
    missing = [p['name'] for p in all_p if p['id'] not in seen]
    if missing:
        print('  ⛔ بلا وسم مقاول: ' + ' · '.join(missing))
    else:
        print('  ✅ كل المشاريع موسومة بمقاول')
    print(BAR)
    print('  ملاحظة: المشروع المشترك بيتكرر عند كل مقاول شغّال عليه،')
    print('          فمجموع المقاولين أكبر من مجموع الذمم الحقيقي.')
    print(BAR)


if __name__ == '__main__':
    main()
