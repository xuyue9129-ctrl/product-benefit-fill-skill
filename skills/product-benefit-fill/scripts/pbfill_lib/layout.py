# -*- coding: utf-8 -*-
"""Locate the two benefit sheets and describe every product group in them.

Result sheet  (name contains 产品利益信息 + 结果): one row per group; H:P are formulas that point at
               L:T of the group's summary row in the working sheet; Q holds the per-row method label.
Working sheet (name contains 产品利益信息 + 底稿): parameter row, summary rows, then year-by-year blocks.
"""
import re

from . import xlsx
from .calc import Block

RES_COLS = 'HIJKLMNOP'          # result sheet: 9 indicators
SUM_COLS = 'LMNOPQRST'          # working sheet summary row: the same 9 indicators
NAMES = ['保证IRR', '保证返本', '保证翻倍', '累生IRR', '累生返本', '累生翻倍', '增额IRR', '增额返本', '增额翻倍']
NOT_OFFERED = ('无60/65岁起领', '40岁不可投保')


def pending(v):
    return isinstance(v, str) and v.startswith('待补')


def num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


class Group(object):
    def __init__(self, **kw):
        self.__dict__.update(kw)

    def as_dict(self):
        d = {k: v for k, v in self.__dict__.items() if not k.startswith('_')}
        d['pending'] = [n for n, v in zip(NAMES, self.values) if pending(v)]
        return d


class Layout(object):
    def __init__(self, book):
        self.book = book
        self.res_name = book.find_sheet('产品利益信息', '结果')
        self.work_name = book.find_sheet('产品利益信息', '底稿')
        self.res = book.sheet(self.res_name)
        self.work = book.sheet(self.work_name)
        self.premium, self.age, self.rate = 100000.0, 40, 0.0175
        self.param_row = self.summary_header = self.block_header = None
        for rn in sorted(self.work.rows):
            a = self.work.value('A%d' % rn)
            if a == '保费（元）' and self.param_row is None:
                self.param_row = rn
                self.premium = float(self.work.value('B%d' % rn))
                if num(self.work.value('D%d' % rn)):
                    self.age = int(self.work.value('D%d' % rn))
                if num(self.work.value('H%d' % rn)):
                    self.rate = float(self.work.value('H%d' % rn))
            elif a == '原序号' and self.summary_header is None:
                self.summary_header = rn
            elif a == '组合号':
                self.block_header = rn
                break
        if not (self.param_row and self.summary_header and self.block_header):
            raise ValueError('working sheet layout not recognised (parameter row %s, summary header %s, block header %s)'
                             % (self.param_row, self.summary_header, self.block_header))
        self._groups = None
        self._ah = None

    # ---------- groups ----------
    def _resolve(self, rrow, i):
        """Value of an indicator: the result cell's cached value, else the summary cell it points at."""
        v = self.res.value('%s%d' % (RES_COLS[i], rrow))
        if v is None:
            f = self.res.formula('%s%d' % (RES_COLS[i], rrow)) or ''
            m = re.search(r'!\$?([A-Z]+)\$?(\d+)$', f)
            if m:
                v = self.work.value('%s%s' % (m.group(1), m.group(2)))
        return v

    def _ah_index(self):
        """summary row -> (first, last) block row, from the usage note in column AH of each block row."""
        if self._ah is None:
            idx = {}
            for rn in sorted(self.work.rows):
                if rn <= self.block_header:
                    continue
                t = self.work.value('AH%d' % rn)
                m = re.search(r'底稿汇总行第(\d+)行', t) if isinstance(t, str) else None
                if m:
                    sr = int(m.group(1))
                    a, b = idx.get(sr, (rn, rn))
                    idx[sr] = (min(a, rn), max(b, rn))
            self._ah = idx
        return self._ah

    def _block_range(self, sr):
        fl = self.work.formula('L%d' % sr) or ''
        m = re.match(r'^IRR\(L(\d+):L(\d+)\)$', fl)
        start = terminal = end = None
        if m:
            start, terminal = int(m.group(1)), int(m.group(2))
            for c in 'MN':
                rr = re.findall(r'[A-Z]+(\d+):[A-Z]+(\d+)', self.work.formula('%s%d' % (c, sr)) or '')
                if rr:
                    end = max(int(b) for a, b in rr)
            if end is None:
                end = terminal
                while self.work.value('A%d' % (end + 1)) == self.work.value('A%d' % start) and self.work.value('C%d' % (end + 1)) is not None:
                    end += 1
        elif sr in self._ah_index():
            start, end = self._ah_index()[sr]
            o = self.work.value('O%d' % start)
            terminal = start + int(o) if num(o) else None
        if start is None or terminal is None:
            return None
        return start, terminal, end

    def groups(self):
        if self._groups is not None:
            return self._groups
        out = []
        pid = name = typ = None
        for rn in sorted(self.res.rows):
            f = self.res.formula('H%d' % rn)
            m = re.search(r'!\$?L\$?(\d+)$', f or '')
            if not m:
                continue
            sr = int(m.group(1))
            b = self.res.value('B%d' % rn)
            if b is not None:
                pid = int(b) if num(b) else b
                name = self.res.value('C%d' % rn)
                typ = self.res.value('D%d' % rn)
            j = self.work.value('J%d' % sr)
            g = Group(row=rn, id=pid, name=name, type=typ, term=self.res.value('E%d' % rn), pay_text=self.res.value('F%d' % rn),
                      pay=int(j) if num(j) else None, summary_row=sr, values=[self._resolve(rn, i) for i in range(9)],
                      label=self.res.value('Q%d' % rn), block=self._block_range(sr))
            out.append(g)
        self._groups = out
        return out

    def group(self, row):
        for g in self.groups():
            if g.row == row:
                return g
        raise KeyError('result row %s is not a product group row' % row)

    # ---------- blocks ----------
    def kind_of(self, g):
        """'ann' = annuity (cash value + maturity + cumulative payments); 'H'/'E' = life (cash value only)."""
        if g.block:
            kf = self.work.formula('K%d' % (g.block[0] + 1)) or ''
            if re.search(r'E\d+\+G\d+\+I\d+', kf):
                return 'ann'
            if re.fullmatch(r'E\d+', kf):
                return 'E'
            if re.fullmatch(r'IF\(ISNUMBER\(H\d+\),H\d+,"待补"\)', kf):
                return 'H'
            raise ValueError('row %d: unrecognised total-benefit formula in column K: %s' % (g.row, kf))
        return 'ann' if '年金' in (g.name or '') else 'H'

    def read_block(self, g):
        """Return (Block, raw) for a group that has a year-by-year block; raw holds the cached cell values."""
        if not g.block:
            raise ValueError('result row %d has no year-by-year block' % g.row)
        start, terminal, end = g.block
        n = end - start
        w = self.work
        col = lambda c: [w.value('%s%d' % (c, start + i)) for i in range(n + 1)]
        C = col('C')
        if not all(num(C[i]) and int(C[i]) == i for i in range(n + 1)):
            raise ValueError('row %d: policy years in column C are not 0..%d' % (g.row, n))
        E_ = col('E')
        if not all(num(v) for v in E_):
            raise ValueError('row %d: guaranteed cash value (column E) is incomplete' % g.row)
        F = [v if num(v) else 0.0 for v in col('F')]
        G = [v if num(v) else 0.0 for v in col('G')]
        H = [v if num(v) else 0.0 for v in col('H')]
        blk = Block(self.premium, g.pay, E_, F, G, H, terminal - start, self.kind_of(g))
        raw = {c: col(c) for c in ('K', 'L', 'J', 'P', 'U', 'V', 'W', 'AB')}
        for i in range(n + 1):
            for c, mine in (('K', blk.K), ('L', blk.L), ('J', blk.J)):
                if num(raw[c][i]) and abs(raw[c][i] - mine[i]) > 0.01:
                    raise ValueError('row %d: column %s year %d is %r in the sheet but %r by the standard formula'
                                     % (g.row, c, i, raw[c][i], mine[i]))
        return blk, raw

    def miss_text(self, g, n):
        """Text used when payback/doubling is never reached inside the block."""
        if g.block:
            for c in 'MNPQST':
                m = re.search(r'=0,"([^"]+)"', self.work.formula('%s%d' % (c, g.summary_row)) or '')
                if m:
                    return m.group(1)
        m = re.match(r'保至(\d+)', str(g.term or ''))
        if m:
            return '%s岁内未达' % m.group(1)
        if '终身' in str(g.term or ''):
            return '%d岁内未达' % (self.age + n)
        return '期内未达'

    # ---------- statistics ----------
    def stats(self):
        gs = self.groups()
        s = {'groups': len(gs), 'products': len({g.id for g in gs}), 'pending_guarantee': 0, 'pending_dividend': 0,
             'filled_dividend': 0, 'complete_groups': 0, 'pending_products': set(), 'pending_groups': 0, 'uncalculated': 0}
        for g in gs:
            pg = sum(1 for v in g.values[:3] if pending(v))
            pd_ = sum(1 for v in g.values[3:] if pending(v))
            s['pending_guarantee'] += pg
            s['pending_dividend'] += pd_
            s['uncalculated'] += sum(1 for v in g.values if v is None)
            if g.type == '分红':
                s['filled_dividend'] += sum(1 for v in g.values[3:] if v is not None and not pending(v) and v != '-')
            if pg + pd_:
                s['pending_products'].add(g.id)
                s['pending_groups'] += 1
            elif not any(v in NOT_OFFERED for v in g.values[:3]) and all(v is not None for v in g.values):
                s['complete_groups'] += 1
        s['pending_products'] = len(s['pending_products'])
        return s
