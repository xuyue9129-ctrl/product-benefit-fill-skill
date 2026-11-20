#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pbfill: fill and check the 产品利益信息 result / working sheets.

  inspect            count what is filled and what is pending; export the pending list
  fill               apply fill specs to a NEW copy of the workbook (never in place)
  verify             recalculate the new workbook in Excel and compare it with the fill report and the old workbook
  audit              recompute stored indicators from the year-by-year blocks (no Excel needed)
  platform-choose    map a group to the platform's input options and build the request body (no network)
  platform-fetch     send the request to the public illustration endpoint (needs the user's consent)
  platform-convert   check saved replies against the workbook and write a fill spec
"""
import argparse
import csv
import datetime
import glob
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pbfill_lib import __version__, convert, platform, verify          # noqa: E402
from pbfill_lib.fill import Planner, Refuse                             # noqa: E402
from pbfill_lib.layout import Layout, NAMES, pending                    # noqa: E402
from pbfill_lib.xlsx import Book                                        # noqa: E402


def out(o):
    print(json.dumps(o, ensure_ascii=False, indent=1) if not isinstance(o, str) else o)


def dump(path, o):
    d = os.path.dirname(os.path.abspath(path))
    if not os.path.isdir(d):
        os.makedirs(d)
    with open(path, 'w', encoding='utf8') as fh:
        json.dump(o, fh, ensure_ascii=False, indent=1)


def load(path):
    with open(path, encoding='utf-8-sig') as fh:
        return json.load(fh)


def open_layout(path):
    book = Book(path)
    return book, Layout(book)


def cmd_inspect(a):
    book, lay = open_layout(a.workbook)
    st = lay.stats()
    st['workbook_open_in_excel'] = book.locked()
    st['premium'], st['age'], st['accumulation_rate'] = lay.premium, lay.age, lay.rate
    rows = [g.as_dict() for g in lay.groups() if any(pending(v) for v in g.values)]
    if a.json:
        dump(a.json, {'stats': st, 'pending': rows})
    if a.csv:
        with open(a.csv, 'w', encoding='utf-8-sig', newline='') as fh:
            w = csv.writer(fh)
            w.writerow(['结果表行', '序号', '产品', '利益类型', '保险期间', '交费期', '底稿汇总行', '有无逐年块', '待补指标', '口径标记'])
            for r in rows:
                w.writerow([r['row'], r['id'], r['name'], r['type'], r['term'], r['pay_text'], r['summary_row'],
                            '有' if r['block'] else '无', '、'.join(r['pending']), r['label'] or ''])
    out(st)
    if st['uncalculated']:
        out('WARNING: %d indicator cells have no stored value; open the workbook in Excel, let it calculate, save, and inspect again.' % st['uncalculated'])
    return 0


def cmd_fill(a):
    book, lay = open_layout(a.workbook)
    if not a.dry_run and not a.out:
        out('ERROR: --out NEW.xlsx is required (or use --dry-run)')
        return 2
    if a.out and os.path.exists(a.out):
        out('ERROR: output already exists; choose a new file name: %s' % a.out)
        return 2
    specs = []
    files = []
    for pat in a.spec:      # the Windows shells do not expand wildcards
        files += sorted(glob.glob(pat)) if any(ch in pat for ch in '*?[') else [pat]
    if not files:
        out('ERROR: no spec file matches %s' % a.spec)
        return 2
    for f in files:
        s = load(f)
        specs += s if isinstance(s, list) else [s]
    before = lay.stats()
    pl = Planner(lay, a.tol)
    filled, refused, seen = [], [], set()
    for spec in specs:
        row = spec.get('result_row')
        try:
            if row in seen:
                raise Refuse('result row %s appears twice in this run; put everything for one row into one spec' % row)
            rep, ops, rows = pl.plan(spec)
            pl.apply(ops, rows)
        except (Refuse, KeyError, ValueError) as e:
            refused.append({'row': row, 'reason': str(e.args[0]) if e.args else str(e)})
            continue
        seen.add(row)
        lay._groups = None
        lay._ah = None
        filled.append(rep)
    after = lay.stats()
    report = {'tool': 'pbfill %s' % __version__, 'time': datetime.datetime.now().isoformat(timespec='seconds'),
              'source': os.path.abspath(a.workbook), 'output': None, 'dry_run': bool(a.dry_run),
              'before': before, 'after': after, 'filled': filled, 'refused': refused}
    if filled and not a.dry_run:
        report['package_check'] = book.save(a.out)
        report['output'] = os.path.abspath(a.out)
    if a.report:
        dump(a.report, report)
    out({'filled_groups': len(filled), 'filled_indicators': sum(len(r['filled']) for r in filled), 'refused': refused,
         'pending_before': before['pending_guarantee'] + before['pending_dividend'],
         'pending_after': after['pending_guarantee'] + after['pending_dividend'],
         'output': report['output'], 'package_check': report.get('package_check')})
    if not a.dry_run and filled:
        out('NEXT: run "verify" on the new file. The stored values were computed in Python; only an Excel recalculation proves the formulas.')
    return 0 if filled else 1


def cmd_verify(a):
    book, lay = open_layout(a.workbook)
    base_lay = open_layout(a.base)[1] if a.base else None
    filled = load(a.report)['filled'] if a.report else None
    res = {}
    if not a.no_excel:
        rows = [g.row for g in lay.groups()]
        idx = list(book.sheets).index(lay.res_name) + 1
        try:
            rec = verify.excel_recalc(a.workbook, idx, min(rows), max(rows))
            res['excel'] = verify.compare(lay, rec, base_lay, filled)
        except RuntimeError as e:
            res['excel'] = {'ok': False, 'not_run': str(e)}
    checked, diffs, skipped, special = verify.audit(lay)
    res['audit'] = {'indicators_recomputed': checked, 'differences': diffs[:50], 'difference_count': len(diffs),
                    'blocks_not_readable': len(skipped), 'judged_by_other_method': len(special), 'ok': not diffs}
    if a.json:
        dump(a.json, res)
    brief = json.loads(json.dumps(res))
    for k in ('old_changed', 'mismatch'):
        if 'excel' in brief and k in brief['excel']:
            brief['excel'][k] = brief['excel'][k][:20]
    out(brief)
    if a.no_excel:
        out('NOTE: Excel recalculation was skipped; formulas are NOT proven. Say so when reporting.')
        return 0 if res['audit']['ok'] else 1
    return 0 if res['excel'].get('ok') and res['audit']['ok'] else 1


def cmd_audit(a):
    _book, lay = open_layout(a.workbook)
    checked, diffs, skipped, special = verify.audit(lay)
    res = {'indicators_recomputed': checked, 'difference_count': len(diffs), 'differences': diffs, 'blocks_not_readable': skipped,
           'judged_by_other_method': special}
    if a.json:
        dump(a.json, res)
    out({'indicators_recomputed': checked, 'difference_count': len(diffs), 'differences': diffs[:30], 'blocks_not_readable': len(skipped),
         'judged_by_other_method': len(special)})
    return 0 if not diffs else 1


def _sets(items):
    o = {}
    for it in items or []:
        k, _, v = it.partition('=')
        try:
            o[k] = json.loads(v)
        except ValueError:
            o[k] = v
    return o


def cmd_choose(a):
    init = load(a.init)
    sets, notes = platform.choose(init, a.pay, a.term, a.premium, a.age, a.sex, _sets(a.set))
    if sets is None:
        out({'offered': False, 'notes': notes})
        return 1
    pb = platform.build_pbdata(init, sets, a.age, a.sex)
    dump(a.out, pb)
    out({'offered': True, 'chosen': sets, 'notes': notes, 'pbdata': os.path.abspath(a.out)})
    return 0


def cmd_fetch(a):
    if not (a.consent or '').strip():
        out('ERROR: --consent is required. Ask the user first; pass who agreed and when, e.g. --consent "张三 2026-01-05 同意本次取数".')
        return 2
    takes = [t.strip() for t in a.takes.split(',') if t.strip()]
    try:
        st = platform.fetch(load(a.pbdata), takes, a.annuity, a.out, a.state, a.base_url,
                            tag='%s | consent: %s' % (a.tag or '', a.consent.strip()), gap=a.gap)
    except platform.Stop as e:
        out({'stopped': True, 'reason': str(e),
             'instruction': 'Stop here and tell the user. Do not retry, switch entry point, parameters, account or network.'})
        return 3
    out(st)
    return 0


def cmd_convert(a):
    _book, lay = open_layout(a.workbook)
    g = lay.group(a.row)
    cols = {}
    reasons = {}
    for take in ('bonus_acc', 'bonus_add'):
        f = os.path.join(a.dir, 'detail_%s.json' % take)
        if not os.path.exists(f):
            cols[take] = None
            continue
        try:
            cols[take] = platform.columns(f)
        except ValueError as e:
            cols[take] = None
            reasons[take] = str(e)
    try:
        spec, info = convert.build(lay, g, cols['bonus_acc'], cols['bonus_add'], a.scale, a.demo_rate, a.pricing_rate, a.share)
    except (Refuse, ValueError) as e:
        out({'usable': False, 'reason': str(e), 'platform_messages': reasons})
        return 1
    date = a.date or datetime.date.today().isoformat()
    pay = '趸交' if g.pay == 1 else '%d年交' % g.pay
    spec = dict({'result_row': g.row, 'label': convert.label(info, date, info.pop('scale_note')),
                 'source': '%s补填来源：%s 利益演示接口 https://%s%s ；参数：%d岁，%s保费%s元，%s，红利中档；直接调用公开预览演示接口取数，未登录、未生成计划书；原始返回归档 %s。'
                           % (date, a.entry or '计划书平台公开预览产品条目', platform.HOST, platform.DETAIL_PATH, lay.age,
                              '趸交' if g.pay == 1 else '年交', ('%f' % lay.premium).rstrip('0').rstrip('.'), pay, a.archive or a.dir)},
                **spec)
    usable = bool(spec.get('accumulate') or spec.get('paidup'))
    res = {'usable': usable, 'row': g.row, 'product': g.name, 'pay': g.pay_text, 'term': g.term,
           'pending': [n for n, v in zip(NAMES, g.values) if pending(v)], 'info': info, 'platform_messages': reasons}
    if usable:
        dump(a.out, spec)
        res['spec'] = os.path.abspath(a.out)
    out(res)
    if usable:
        out('NEXT: read the label draft in the spec, correct it if needed, then run "fill".')
    return 0 if usable else 1


def main(argv=None):
    for s in (sys.stdout, sys.stderr):
        if hasattr(s, 'reconfigure'):
            s.reconfigure(encoding='utf-8')
    p = argparse.ArgumentParser(prog='pbfill', description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument('--version', action='version', version=__version__)
    sub = p.add_subparsers(dest='cmd', required=True)

    s = sub.add_parser('inspect')
    s.add_argument('workbook')
    s.add_argument('--json')
    s.add_argument('--csv')
    s.set_defaults(fn=cmd_inspect)

    s = sub.add_parser('fill')
    s.add_argument('workbook')
    s.add_argument('--spec', nargs='+', required=True, help='JSON file(s); each holds one spec or a list of specs')
    s.add_argument('--out', help='new workbook to write; must not exist')
    s.add_argument('--dry-run', action='store_true')
    s.add_argument('--tol', type=float, default=1.0, help='allowed difference (yuan) between new data and figures already entered')
    s.add_argument('--report', help='JSON report; needed later by verify')
    s.set_defaults(fn=cmd_fill)

    s = sub.add_parser('verify')
    s.add_argument('workbook', help='the NEW workbook')
    s.add_argument('--base', help='the workbook before filling')
    s.add_argument('--report', help='report written by fill')
    s.add_argument('--no-excel', action='store_true')
    s.add_argument('--json')
    s.set_defaults(fn=cmd_verify)

    s = sub.add_parser('audit')
    s.add_argument('workbook')
    s.add_argument('--json')
    s.set_defaults(fn=cmd_audit)

    s = sub.add_parser('platform-choose')
    s.add_argument('--init', required=True, help="the product's input configuration, supplied by the user")
    s.add_argument('--pay', type=int, required=True, help='premium term in years; 1 = single premium')
    s.add_argument('--term', required=True, help='insurance period text of the group, e.g. 终身 / 保至105周岁；60岁起年领')
    s.add_argument('--premium', type=float, default=100000)
    s.add_argument('--age', type=int, default=40)
    s.add_argument('--sex', type=int, default=1, help='1 = male, 2 = female')
    s.add_argument('--set', action='append', help='override an input element, key=value')
    s.add_argument('--out', required=True, help='pbdata.json to write')
    s.set_defaults(fn=cmd_choose)

    s = sub.add_parser('platform-fetch')
    s.add_argument('--pbdata', required=True)
    s.add_argument('--out', required=True, help='directory for the raw replies of this group')
    s.add_argument('--state', required=True, help='directory shared by all fetches: request log and stop marker')
    s.add_argument('--takes', default='bonus_acc,bonus_add')
    s.add_argument('--annuity', action='store_true', help='the product is an annuity (annuity taken in cash)')
    s.add_argument('--consent', help='who agreed to this fetch and when (required)')
    s.add_argument('--tag', help='free text written to the request log')
    s.add_argument('--gap', type=float, default=platform.GAP_SECONDS, help='seconds between requests (never below %s)' % platform.GAP_SECONDS)
    s.add_argument('--base-url', help=argparse.SUPPRESS)
    s.set_defaults(fn=cmd_fetch)

    s = sub.add_parser('platform-convert')
    s.add_argument('workbook')
    s.add_argument('--row', type=int, required=True, help='row number of the group in the result sheet')
    s.add_argument('--dir', required=True, help='directory written by platform-fetch')
    s.add_argument('--out', required=True, help='spec JSON to write')
    s.add_argument('--scale', choices=['all', 'base'])
    s.add_argument('--demo-rate', type=float, default=0.035)
    s.add_argument('--pricing-rate', type=float, default=0.0175)
    s.add_argument('--share', type=float, default=0.7)
    s.add_argument('--date')
    s.add_argument('--entry', help='how the source names the platform entry (id, name, the platform’s own label)')
    s.add_argument('--archive', help='archive path to quote in the source note (default: --dir)')
    s.set_defaults(fn=cmd_convert)

    a = p.parse_args(argv)
    return a.fn(a)


if __name__ == '__main__':
    sys.exit(main())
