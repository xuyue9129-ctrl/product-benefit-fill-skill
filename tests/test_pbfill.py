# -*- coding: utf-8 -*-
"""Tests on a synthetic workbook and a local fake of the illustration endpoint. No real product data, no internet."""
import contextlib
import http.server
import io
import json
import os
import shutil
import sys
import tempfile
import threading
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.join(HERE, '..', 'skills', 'product-benefit-fill')
sys.path.insert(0, os.path.join(SKILL, 'examples'))
sys.path.insert(0, os.path.join(SKILL, 'scripts'))

import make_demo_workbook as mkbook                        # noqa: E402
import pbfill                                              # noqa: E402
from pbfill_lib import platform, verify                    # noqa: E402
from pbfill_lib.calc import irr                            # noqa: E402
from pbfill_lib.layout import Layout                       # noqa: E402
from pbfill_lib.xlsx import Book, shift_formula            # noqa: E402

N = 45
E = mkbook.life_cv(N)


def run(*argv):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        code = pbfill.main(list(argv))
    return code, buf.getvalue()


def jload(path):
    with open(path, encoding='utf8') as fh:
        return json.load(fh)


def read(path, mode='r', encoding='utf8'):
    with open(path, mode, **({} if 'b' in mode else {'encoding': encoding})) as fh:
        return fh.read()


def values(path, row):
    lay = Layout(Book(path))
    return lay.group(row).values, lay


class Base(unittest.TestCase):
    def setUp(self):
        self.d = tempfile.mkdtemp(prefix='pbfill_test_')
        self.wb = mkbook.build(os.path.join(self.d, 'book.xlsx'))

    def tearDown(self):
        shutil.rmtree(self.d, ignore_errors=True)

    def p(self, name):
        return os.path.join(self.d, name)

    def spec(self, name, obj):
        with open(self.p(name), 'w', encoding='utf8') as fh:
            json.dump(obj, fh, ensure_ascii=False)
        return self.p(name)


class Inspect(Base):
    def test_counts_and_pending_list(self):
        code, _ = run('inspect', self.wb, '--json', self.p('i.json'), '--csv', self.p('p.csv'))
        self.assertEqual(code, 0)
        d = jload(self.p('i.json'))
        st = d['stats']
        self.assertEqual((st['groups'], st['products'], st['pending_guarantee'], st['pending_dividend']), (3, 3, 3, 12))
        self.assertEqual(st['uncalculated'], 0)
        self.assertEqual([r['row'] for r in d['pending']], [9, 10, 11])
        self.assertTrue(read(self.p('p.csv'), encoding='utf-8-sig').startswith('结果表行'))

    def test_guaranteed_values_of_the_synthetic_book(self):
        v, _ = values(self.wb, 9)
        self.assertAlmostEqual(v[0], 0.02, places=7)       # 100000 growing at 2% a year
        self.assertEqual(v[1:3], [1, 36])                  # 1.02^35 < 2 < 1.02^36


class Fill(Base):
    def test_dividends_on_an_existing_block(self):
        s = self.spec('s.json', {'result_row': 9, 'label': '示例标注', 'source': '示例来源',
                                 'accumulate': {'dividends': [1000.0] * N},
                                 'paidup': {'total_cv': [round(1.1 * e, 2) for e in E[1:]]}})
        code, _ = run('fill', self.wb, '--spec', s, '--out', self.p('new.xlsx'), '--report', self.p('r.json'))
        self.assertEqual(code, 0)
        v, lay = values(self.p('new.xlsx'), 9)
        bal = lambda y: 1000.0 * (1.0175 ** y - 1) / 0.0175          # closed form, independent of the tool's recursion
        self.assertAlmostEqual(v[3], ((E[40] + bal(40)) / 100000.0) ** (1 / 40.0) - 1, places=7)
        self.assertEqual(v[4], 1)
        self.assertEqual(v[5], next(y for y in range(1, N + 1) if E[y] + bal(y) > 200000))
        self.assertAlmostEqual(v[6], (1.1 * 1.02 ** 40) ** (1 / 40.0) - 1, places=7)
        self.assertEqual(v[7:9], [1, 31])                             # 1.1 * 1.02^31 is the first above 2
        self.assertAlmostEqual(v[0], 0.02, places=7)                  # guaranteed side untouched
        g = lay.group(9)
        self.assertEqual(g.label, '示例标注〔此前标注：此前标注示例〕')
        self.assertTrue(lay.work.value('AC%d' % g.summary_row).endswith('示例来源'))
        self.assertEqual(lay.work.formula('O%d' % g.summary_row), 'IRR(R10:R50)')
        self.assertEqual(lay.stats()['pending_dividend'], 6)
        checked, diffs, _skipped, _special = verify.audit(lay)
        self.assertEqual(diffs, [])
        self.assertGreaterEqual(checked, 9)

    def test_new_block_for_a_group_without_one(self):
        n = 60
        s = self.spec('s.json', {'result_row': 10, 'label': '示例标注', 'source': '示例来源',
                                 'guarantee': {'cash_value': [100000.0] * n, 'survival': [5000.0] * n}})
        code, _ = run('fill', self.wb, '--spec', s, '--out', self.p('new.xlsx'))
        self.assertEqual(code, 0)
        v, lay = values(self.p('new.xlsx'), 10)
        self.assertAlmostEqual(v[0], 0.05, places=7)                  # 5000 a year on 100000, capital back at the end
        self.assertEqual(v[1:3], [1, 21])                             # 100000 + 5000y > 200000 first at y = 21
        self.assertEqual(v[3:], ['-'] * 6)
        g = lay.group(10)
        self.assertEqual(g.block[2] - g.block[0], n)
        self.assertEqual(lay.kind_of(g), 'ann')
        self.assertEqual(verify.audit(lay)[1], [])

    def test_official_points_are_kept_and_must_agree(self):
        f10 = (1.0175 ** 10 - 1) / 0.0175
        d1 = 20000.0 / f10
        d2 = (50000.0 - 20000.0 * 1.0175 ** 10) / f10
        good = [d1] * 10 + [d2] * 10 + [1500.0] * (N - 20)
        code, _ = run('fill', self.wb, '--spec', self.spec('g.json', {'result_row': 11, 'label': 'L', 'source': 'S', 'accumulate': {'dividends': good}}),
                      '--out', self.p('new.xlsx'), '--report', self.p('r.json'))
        self.assertEqual(code, 0)
        rep = jload(self.p('r.json'))
        self.assertEqual(rep['filled'][0]['info']['acc']['official_points_kept'], 2)
        lay = Layout(Book(self.p('new.xlsx')))
        start = lay.group(11).block[0]
        self.assertEqual(lay.work.value('P%d' % (start + 10)), 20000.0)
        bad = [1000.0] * N
        code, o = run('fill', self.wb, '--spec', self.spec('b.json', {'result_row': 11, 'label': 'L', 'source': 'S', 'accumulate': {'dividends': bad}}),
                      '--out', self.p('new2.xlsx'))
        self.assertEqual(code, 1)
        self.assertIn('official figure', o)
        self.assertFalse(os.path.exists(self.p('new2.xlsx')))

    def test_years_without_data_are_not_guessed(self):
        bal = [None] * N
        bal[9], bal[19], bal[39] = 20000.0, 50000.0, 120000.0          # only three disclosed years
        s = self.spec('s.json', {'result_row': 11, 'label': 'L', 'source': 'S', 'accumulate': {'balance': bal}})
        code, _ = run('fill', self.wb, '--spec', s, '--out', self.p('new.xlsx'), '--report', self.p('r.json'))
        self.assertEqual(code, 0)
        rep = jload(self.p('r.json'))['filled'][0]
        self.assertEqual(rep['filled'], ['累生IRR'])
        self.assertEqual(rep['undetermined'], ['累生返本', '累生翻倍'])
        v, lay = values(self.p('new.xlsx'), 11)
        self.assertIsInstance(v[3], float)
        self.assertEqual(v[4:6], ['待补演示', '待补演示'])
        self.assertEqual(verify.audit(lay)[1], [])

    def test_refusals(self):
        specs = [
            {'result_row': 9, 'source': 'S', 'accumulate': {'dividends': [1.0] * N}},                                  # no label
            {'result_row': 9, 'label': 'L', 'accumulate': {'dividends': [1.0] * N}},                                   # no source
            {'result_row': 10, 'label': 'L', 'source': 'S', 'accumulate': {'dividends': [1.0]}},                       # no block yet
            {'result_row': 9, 'label': 'L', 'source': 'S', 'paidup': {'total_cv': [0.9 * e for e in E[1:]]}},          # below guaranteed
            {'result_row': 9, 'label': 'L', 'source': 'S', 'guarantee': {'cash_value': [1.0] * N}},                    # guaranteed already there
            {'result_row': 9, 'label': 'L', 'source': 'S', 'accumulate': {'dividends': [1000.0] * (N + 1)}},           # too many years
            {'result_row': 9, 'label': 'L', 'source': 'S', 'accumulate': {'dividends': [1000.0] * 30}},                # stops before the IRR year
            {'result_row': 4, 'label': 'L', 'source': 'S', 'accumulate': {'dividends': [1.0]}},                        # not a group row
        ]
        code, o = run('fill', self.wb, '--spec', self.spec('s.json', specs), '--out', self.p('new.xlsx'))
        self.assertEqual(code, 1)
        self.assertEqual(len(json.loads(o)['refused']), len(specs))
        self.assertFalse(os.path.exists(self.p('new.xlsx')))

    def test_filled_values_are_never_overwritten(self):
        s = self.spec('s.json', {'result_row': 9, 'label': 'L', 'source': 'S', 'accumulate': {'dividends': [1000.0] * N}})
        self.assertEqual(run('fill', self.wb, '--spec', s, '--out', self.p('new.xlsx'))[0], 0)
        code, o = run('fill', self.p('new.xlsx'), '--spec', s, '--out', self.p('new2.xlsx'))
        self.assertEqual(code, 1)
        self.assertIn('not pending', o)

    def test_source_is_untouched_and_output_is_never_replaced(self):
        before = read(self.wb, 'rb')
        s = self.spec('s.json', {'result_row': 9, 'label': 'L', 'source': 'S', 'accumulate': {'dividends': [1000.0] * N}})
        self.assertEqual(run('fill', self.wb, '--spec', s, '--dry-run')[0], 0)
        self.assertEqual(os.listdir(self.d).count('book.xlsx'), 1)
        self.assertEqual(run('fill', self.wb, '--spec', s, '--out', self.p('new.xlsx'))[0], 0)
        self.assertEqual(read(self.wb, 'rb'), before)
        self.assertEqual(run('fill', self.wb, '--spec', s, '--out', self.p('new.xlsx'))[0], 2)
        self.assertEqual(run('fill', self.wb, '--spec', s, '--out', self.wb)[0], 2)
        with zipfile.ZipFile(self.wb) as a, zipfile.ZipFile(self.p('new.xlsx')) as b:
            self.assertEqual(a.namelist(), b.namelist())
            for n in a.namelist():
                if 'sheet' not in n:
                    self.assertEqual(a.read(n), b.read(n))

    def test_verify_without_excel(self):
        s = self.spec('s.json', {'result_row': 9, 'label': 'L', 'source': 'S', 'accumulate': {'dividends': [1000.0] * N}})
        run('fill', self.wb, '--spec', s, '--out', self.p('new.xlsx'), '--report', self.p('r.json'))
        code, o = run('verify', self.p('new.xlsx'), '--base', self.wb, '--report', self.p('r.json'), '--no-excel')
        self.assertEqual(code, 0)
        self.assertIn('NOT proven', o)


class Examples(Base):
    def test_shipped_examples_run(self):
        ex = lambda n: os.path.join(SKILL, 'examples', n)
        code, o = run('fill', self.wb, '--spec', ex('spec_dividend.json'), ex('spec_guarantee.json'), ex('spec_official_points.json'),
                      '--out', self.p('new.xlsx'), '--report', self.p('r.json'))
        self.assertEqual(code, 0, o)
        rep = jload(self.p('r.json'))
        self.assertEqual(rep['refused'], [])
        self.assertEqual(sum(len(r['filled']) for r in rep['filled']), 6 + 3 + 1)
        st = Layout(Book(self.p('new.xlsx'))).stats()
        self.assertEqual((st['pending_guarantee'], st['pending_dividend']), (0, 5))
        code, o = run('platform-choose', '--init', ex('product_config_example.json'), '--pay', '1', '--term', '终身', '--out', self.p('pb.json'))
        self.assertEqual(code, 0, o)


class Arithmetic(unittest.TestCase):
    def test_irr(self):
        self.assertAlmostEqual(irr([-100.0, 0.0, 121.0]), 0.1, places=9)
        self.assertIsNone(irr([-100.0, -1.0]))

    def test_shift_formula(self):
        self.assertEqual(shift_formula('IF(C10<$J$5,$B$2,0)', 3, 0), 'IF(C13<$J$5,$B$2,0)')
        self.assertEqual(shift_formula('SUM(V$10:V12)+"A1"', 1, 1), 'SUM(W$10:W13)+"A1"')


INIT = {'data': {'pbId': 999001, 'initData': {'inputData': {'productGroupList': [{'key': 'allMainInsData', 'prodOrderList': [{
    'key': 5001, 'name': '示例终身寿险（分红型）', 'companyName': '示例公司', 'mulInsOrderList': [{
        'key': 7001, 'required': True, 'name': '示例终身寿险（分红型）', 'eleOrderList': [
            {'key': 'pPeriod', 'value': 'b5', 'opts': [{'val': 'a1000', 'desc': '趸交'}, {'val': 'b5', 'desc': '5年交'}]},
            {'key': 'iDuration', 'value': 'a999', 'opts': [{'val': 'a999', 'desc': '终身'}]},
            {'key': 'baof', 'value': 10000}]}]}]}]}}}}


def reply(take, factor=1.0):
    cv = [e * factor for e in E[1:]]
    cols = [{'key': 'baof', 'name': '保费', 'values': [100000.0] + [0.0] * (N - 1)},
            {'key': 'cash_value', 'name': '现金价值', 'values': cv}]
    if take == 'bonus_acc':
        bonus = [round(0.01225 * E[y - 1] * factor, 2) for y in range(1, N + 1)]
        s, acc = 0.0, []
        for b in bonus:
            s = s * 1.0175 + b
            acc.append(round(s, 2))
        cols += [{'key': 'bonus', 'name': '当年红利', 'values': bonus}, {'key': 'bonus_sum', 'name': '累积红利', 'values': acc}]
    else:
        pua = [round(0.005 * y * E[y] * factor, 2) for y in range(1, N + 1)]
        cols += [{'key': 'pua_cash_value', 'name': '交清增额现金价值', 'values': pua},
                 {'key': 'total_cash_value', 'name': '总现金价值', 'values': [round(c + p, 2) for c, p in zip(cv, pua)]}]
    return {'success': True, 'code': '0', 'data': {'productDataList': cols, 'userInfo': {'mobile': 'x'}}}


class Fake(http.server.BaseHTTPRequestHandler):
    mode, hits, factor = 'ok', 0, 1.0

    def do_POST(self):
        Fake.hits += 1
        body = json.loads(self.rfile.read(int(self.headers['Content-Length'])).decode('utf8'))
        assert self.path == platform.DETAIL_PATH and isinstance(json.loads(body['pbData']), dict)
        assert 'Cookie' not in self.headers and 'Authorization' not in self.headers
        if Fake.mode == 'ok':
            out = reply(body['bonusTake']['value'], Fake.factor)
        elif Fake.mode == 'captcha':
            out = {'success': False, 'code': '002021', 'msg': '请完成滑动验证'}
        else:
            out = {'success': False, 'code': '015001', 'msg': '案例数据处理'}
        raw = json.dumps(out, ensure_ascii=False).encode('utf8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def log_message(self, *a):
        pass


class Platform(Base):
    def setUp(self):
        Base.setUp(self)
        Fake.mode, Fake.hits, Fake.factor = 'ok', 0, 1.0
        self.srv = http.server.HTTPServer(('127.0.0.1', 0), Fake)
        threading.Thread(target=self.srv.serve_forever, daemon=True).start()
        self.url = 'http://127.0.0.1:%d' % self.srv.server_address[1]
        self.spec('init.json', INIT)
        code, o = run('platform-choose', '--init', self.p('init.json'), '--pay', '1', '--term', '终身', '--out', self.p('pb.json'))
        self.assertEqual(code, 0, o)

    def tearDown(self):
        self.srv.shutdown()
        self.srv.server_close()
        Base.tearDown(self)

    def fetch(self, out='raw', consent='测试人 同意'):
        argv = ['platform-fetch', '--pbdata', self.p('pb.json'), '--out', self.p(out), '--state', self.p('state'), '--gap', '0', '--base-url', self.url]
        return run(*(argv + (['--consent', consent] if consent else [])))

    def test_choose_maps_the_group_to_offered_options(self):
        pb = jload(self.p('pb.json'))
        ins = pb['productGroupList'][0]['prodOrderList'][0]['mulInsOrderList'][0]
        self.assertEqual((ins['pPeriod'], ins['iDuration'], ins['baof']), ('a1000', 'a999', 100000.0))
        self.assertEqual(pb['personData']['personOrderList'][0]['age'], 40)
        self.assertNotIn('companyName', pb['productGroupList'][0]['prodOrderList'][0])
        code, o = run('platform-choose', '--init', self.p('init.json'), '--pay', '10', '--term', '终身', '--out', self.p('x.json'))
        self.assertEqual(code, 1)
        self.assertFalse(os.path.exists(self.p('x.json')))

    def test_no_request_without_consent(self):
        code, _ = self.fetch(consent=None)
        self.assertEqual((code, Fake.hits), (2, 0))

    def test_fetch_convert_fill(self):
        code, o = self.fetch()
        self.assertEqual((code, Fake.hits), (0, 2), o)
        saved = read(self.p(os.path.join('raw', 'detail_bonus_acc.json')))
        self.assertNotIn('mobile', saved)                                # personal fields are dropped before saving
        self.assertEqual(self.fetch()[0], 0)
        self.assertEqual(Fake.hits, 2)                                   # saved replies are not requested again
        log = read(self.p(os.path.join('state', platform.LOG_FILE)))
        self.assertIn('测试人 同意', log)
        code, o = run('platform-convert', self.wb, '--row', '9', '--dir', self.p('raw'), '--out', self.p('spec.json'), '--date', '2026-01-05')
        self.assertEqual(code, 0, o)
        spec = jload(self.p('spec.json'))
        self.assertIn('1.225%', spec['label'])
        self.assertIn('按比值认定', spec['label'])
        self.assertIn('未登录、未生成计划书', spec['source'])
        code, o = run('fill', self.wb, '--spec', self.p('spec.json'), '--out', self.p('new.xlsx'))
        self.assertEqual(code, 0, o)
        v, lay = values(self.p('new.xlsx'), 9)
        self.assertEqual(lay.stats()['pending_dividend'], 6)
        self.assertTrue(0.02 < v[3] < 0.035 and 0.02 < v[6] < 0.035)
        self.assertEqual(verify.audit(lay)[1], [])

    def test_basis_is_measured_not_assumed(self):
        self.fetch()
        code, o = run('platform-convert', self.wb, '--row', '9', '--dir', self.p('raw'), '--out', self.p('spec.json'), '--pricing-rate', '0.02')
        self.assertEqual(code, 0)
        self.assertIn('未能认定', jload(self.p('spec.json'))['label'])

    def test_other_sum_insured_needs_a_decision(self):
        Fake.factor = 0.5
        self.fetch()
        code, o = run('platform-convert', self.wb, '--row', '9', '--dir', self.p('raw'), '--out', self.p('spec.json'))
        self.assertEqual(code, 1)
        self.assertIn('--scale', o)
        code, o = run('platform-convert', self.wb, '--row', '9', '--dir', self.p('raw'), '--out', self.p('spec.json'), '--scale', 'all')
        self.assertEqual(code, 0, o)
        self.assertIn('折算', jload(self.p('spec.json'))['label'])

    def test_refusal_stops_everything(self):
        Fake.mode = 'captcha'
        code, o = self.fetch()
        self.assertEqual((code, Fake.hits), (3, 1))                      # stopped after the first reply, no second request
        self.assertTrue(os.path.exists(self.p(os.path.join('state', platform.STOP_FILE))))
        Fake.mode = 'ok'
        code, o = self.fetch(out='raw2')
        self.assertEqual((code, Fake.hits), (3, 1))                      # still stopped, nothing sent

    def test_platform_without_data_is_reported_not_guessed(self):
        Fake.mode = 'nodata'
        self.assertEqual(self.fetch()[0], 0)
        code, o = run('platform-convert', self.wb, '--row', '9', '--dir', self.p('raw'), '--out', self.p('spec.json'))
        self.assertEqual(code, 1)
        self.assertFalse(os.path.exists(self.p('spec.json')))

    def test_daily_cap(self):
        pb = jload(self.p('pb.json'))
        cl = platform.Client(self.p('state'), self.url, 0, daily_cap=1)
        cl.post_detail(platform.detail_body(pb, 'bonus_acc', False))
        with self.assertRaises(platform.Stop):
            cl.post_detail(platform.detail_body(pb, 'bonus_add', False))
        self.assertEqual(Fake.hits, 1)


if __name__ == '__main__':
    unittest.main()
