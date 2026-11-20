# -*- coding: utf-8 -*-
"""Turn a fill spec into cell writes for one group at a time and apply them to the in-memory workbook.

A group is written completely or not at all. Indicators that already hold a value are never overwritten:
if the new data would change one of them, the group is refused and the reason is reported.
"""
from .calc import Block, ND, balance
from .layout import RES_COLS, SUM_COLS, NAMES, pending, num

F = {
    'D': 'IF(C{r}<$J${sr},$B${p},0)',
    'J': 'MIN(C{r},$J${sr})*$B${p}',
    'K_ann': 'IF(ISNUMBER(E{r}),E{r}+G{r}+I{r},"待补")',
    'K_life': 'IF(ISNUMBER(H{r}),H{r},"待补")',
    'L': 'IF(AND(C{r}=O{r},NOT(ISNUMBER(E{r}))),"待补",-D{r}+F{r}+G{r}+IF(C{r}=O{r},E{r},0))',
    'M': 'IF(C{r}=0,0,IF(ISNUMBER(K{r}),IF(ROUND(K{r},6)>ROUND(J{r},6),1,0),0))',
    'N': 'IF(C{r}=0,0,IF(ISNUMBER(K{r}),IF(ROUND(K{r},6)>ROUND(2*J{r},6),1,0),0))',
    'Q': 'IF(ISNUMBER(P{r}),K{r}+P{r},"未披露")',
    'R': 'IF(C{r}=O{r},IF(ISNUMBER(P{r}),L{r}+P{r},"未披露"),L{r})',
    'S': 'IF(C{r}=0,0,IF(ISNUMBER(Q{r}),IF(ROUND(Q{r},6)>ROUND(1*J{r},6),1,0),"未披露"))',
    'T': 'IF(C{r}=0,0,IF(ISNUMBER(Q{r}),IF(ROUND(Q{r},6)>ROUND(2*J{r},6),1,0),"未披露"))',
    'X_life': 'IF(ISNUMBER(U{r}),U{r}+IF(K{r}>E{r},K{r}-E{r}+W{r},0),"未披露")',
    'X_ann': 'IF(AND(ISNUMBER(U{r}),COUNT(V{s}:V{r})={n},ISNUMBER(W{r})),U{r}+I{r}+SUM(V{s}:V{r})+G{r}+W{r},"未披露")',
    'Y': 'IF(AND(ISNUMBER(V{r}),ISNUMBER(W{r})),IF(C{r}=O{r},IF(ISNUMBER(U{r}),-D{r}+F{r}+V{r}+G{r}+W{r}+U{r},"未披露"),-D{r}+F{r}+V{r}+G{r}+W{r}),"未披露")',
    'Z': 'IF(C{r}=0,0,IF(ISNUMBER(X{r}),IF(ROUND(X{r},6)>ROUND(1*J{r},6),1,0),"未披露"))',
    'AA': 'IF(C{r}=0,0,IF(ISNUMBER(X{r}),IF(ROUND(X{r},6)>ROUND(2*J{r},6),1,0),"未披露"))',
}
# a new cell in an existing block row takes the number format of a neighbour that holds the same kind of value
STYLE_FROM = {'P': 'E', 'U': 'E', 'V': 'E', 'W': 'E', 'AB': 'E', 'Q': 'K', 'R': 'K', 'X': 'K', 'Y': 'K',
              'S': 'M', 'T': 'M', 'Z': 'M', 'AA': 'M'}


class Refuse(Exception):
    pass


def need(cond, msg):
    if not cond:
        raise Refuse(msg)


def _series(vals, n, what, allow_none=True):
    """Spec arrays start at policy year 1; return a list indexed by policy year with year 0 = 0."""
    vals = list(vals or [])
    need(len(vals) <= n, '%s has %d years but the block has only %d' % (what, len(vals), n))
    for v in vals:
        need(v is None and allow_none or num(v), '%s contains a non-numeric value: %r' % (what, v))
    return [0.0] + [None if v is None else float(v) for v in vals]


def _same(old, new):
    if num(old) and num(new):
        return abs(old - new) < 5e-5 if (abs(old) < 1 or abs(new) < 1) else old == new
    return old == new


class Planner(object):
    def __init__(self, lay, tol=1.0):
        self.lay, self.tol = lay, tol
        self.undetermined = []
        self.w, self.r = lay.work, lay.res

    # ---------- helpers ----------
    def _reach(self, flag_col, a, z, year, miss):
        if year is not None:
            return 'INDEX(C{a}:C{z},MATCH(1,{c}{a}:{c}{z},0))'.format(a=a, z=z, c=flag_col), year
        return 'IF(COUNTIF({c}{a}:{c}{z},1)=0,"{m}","核查错误")'.format(a=a, z=z, c=flag_col, m=miss), miss

    def _summary(self, g, ops, cols, res, flow_col, flag_cols, start, terminal, end, miss, values, base, constants):
        """Write the three summary cells of one benefit basis; existing values must not change."""
        fi = ('IRR({c}{s}:{c}{t})'.format(c=flow_col, s=start, t=terminal), res['irr'])
        f1 = self._reach(flag_cols[0], start + 1, end, res['payback'], miss)
        f2 = self._reach(flag_cols[1], start + 1, end, res['double'], miss)
        filled = []
        for k, (formula, val) in enumerate((fi, f1, f2)):
            i = base + k
            old = values[i]
            if not res['known'][k]:
                # years without data come before the answer: leave the indicator as it is and say so
                if pending(old):
                    self.undetermined.append(NAMES[i])
                continue
            if pending(old):
                addr = '%s%d' % (cols[k], g.summary_row)
                if constants:
                    ops.append(('w', addr, 's' if isinstance(val, str) else 'v', val, None, None))
                else:
                    ops.append(('w', addr, 'f', formula, val, None))
                values[i] = val
                filled.append(NAMES[i])
            else:
                need(constants or _same(old, val), '%s is already %r; the new data would make it %r' % (NAMES[i], old, val))
        return filled

    def _column(self, ops, col, start, vals, first=0):
        """Plain values into a block column, leaving numbers that are already there."""
        for i in range(first, len(vals)):
            addr = '%s%d' % (col, start + i)
            if num(self.w.value(addr)):
                continue
            v = ND if vals[i] is None else vals[i]
            ops.append(('w', addr, 's' if isinstance(v, str) else 'v', v, None, '%s%d' % (STYLE_FROM[col], start + i)))

    def _formulas(self, ops, col, start, texts, caches):
        for i, (t, c) in enumerate(zip(texts, caches)):
            addr = '%s%d' % (col, start + i)
            c = ND if c is None else c
            cell = self.w.cell(addr)
            if cell is None:
                ops.append(('w', addr, 'f', t, c, '%s%d' % (STYLE_FROM[col], start + i)))
            else:
                need(self.w.formula(addr) is not None, '%s holds a constant where a formula is expected' % addr)
                ops.append(('w', addr, 'c', c, None, None))

    def _merge_points(self, col, start, mine, what):
        """Keep numbers already entered from an official source; they must agree with the new series."""
        out = list(mine)
        kept = 0
        for i in range(1, len(mine)):
            old = self.w.value('%s%d' % (col, start + i))
            if num(old):
                need(mine[i] is not None and abs(old - mine[i]) <= self.tol,
                     '%s: year %d is %r in the sheet (official figure) but %r in the new data; difference exceeds %.2f'
                     % (what, i, old, mine[i], self.tol))
                out[i] = old
                kept += 1
        return out, kept

    # ---------- one group ----------
    def plan(self, spec):
        lay, w = self.lay, self.w
        g = lay.group(int(spec['result_row']))
        need(spec.get('label'), 'label is required: every filled row must say where the numbers come from and on what basis')
        need(spec.get('source'), 'source is required: document name, page or URL, parameters')
        need(g.pay, 'summary row %d has no premium term in column J' % g.summary_row)
        need(all(v is not None for v in g.values), 'this workbook has never been calculated (no cached values); open and save it in Excel first')
        ops, rows, info = [], [], {}
        self.undetermined = []
        values = list(g.values)
        filled = []
        sr, p = g.summary_row, lay.param_row
        acc, add, gua = spec.get('accumulate'), spec.get('paidup'), spec.get('guarantee')
        need(acc or add or gua, 'nothing to fill: give guarantee, accumulate and/or paidup')

        if gua:
            need(g.block is None, 'the group already has a year-by-year block; guaranteed data cannot be replaced')
            need(all(pending(v) for v in values[:3]), 'guaranteed indicators are not pending')
            n = len(gua.get('cash_value') or [])
            need(n >= 1, 'guarantee.cash_value is empty')
            E_ = _series(gua['cash_value'], n, 'guarantee.cash_value', allow_none=False)
            Fs = [v or 0.0 for v in _series(gua.get('survival'), n, 'guarantee.survival')]
            Gs = [v or 0.0 for v in _series(gua.get('maturity'), n, 'guarantee.maturity')]
            Fs += [0.0] * (n + 1 - len(Fs))
            Gs += [0.0] * (n + 1 - len(Gs))
            blk = Block(lay.premium, g.pay, E_, Fs, Gs, None, gua.get('terminal'), lay.kind_of(g))
            start = w.last_row() + 1
            new_block = True
        else:
            blk, _raw = lay.read_block(g)
            start = g.block[0]
            new_block = False
        n, terminal, end = blk.N, start + blk.O, start + blk.N
        miss = lay.miss_text(g, n)
        ann = blk.kind == 'ann'
        newcols = {}   # column -> [(kind, value, cache)] per year, only for an appended block

        if gua:
            res = blk.guarantee()
            need(res['irr'] is not None, 'guaranteed cash flows have no IRR')
            fl = [('IRR(L%d:L%d)' % (start, terminal), res['irr']), self._reach('M', start + 1, end, res['payback'], miss),
                  self._reach('N', start + 1, end, res['double'], miss)]
            for k, (formula, val) in enumerate(fl):
                ops.append(('w', '%s%d' % (SUM_COLS[k], sr), 'f', formula, val, None))
                values[k] = val
                filled.append(NAMES[k])
            rr = lambda i: start + i
            newcols.update({
                'C': [('v', i, None) for i in range(n + 1)],
                'D': [('f', F['D'].format(r=rr(i), sr=sr, p=p), blk.D[i]) for i in range(n + 1)],
                'E': [('v', blk.E[i], None) for i in range(n + 1)],
                'F': [('v', blk.F[i], None) for i in range(n + 1)],
                'G': [('v', blk.G[i], None) for i in range(n + 1)],
                'H': [('v', blk.H[i], None) for i in range(n + 1)],
                'I': [('f', 'F%d' % rr(i) if i == 0 else 'I%d+F%d' % (rr(i) - 1, rr(i)), blk.I[i]) for i in range(n + 1)],
                'J': [('f', F['J'].format(r=rr(i), sr=sr, p=p), blk.J[i]) for i in range(n + 1)],
                'K': [('f', F['K_ann' if ann else 'K_life'].format(r=rr(i)), blk.K[i]) for i in range(n + 1)],
                'L': [('f', F['L'].format(r=rr(i)), blk.L[i]) for i in range(n + 1)],
                'M': [('f', F['M'].format(r=rr(i)), res['f1'][i]) for i in range(n + 1)],
                'N': [('f', F['N'].format(r=rr(i)), res['f2'][i]) for i in range(n + 1)],
                'O': [('v', blk.O, None) for i in range(n + 1)],
            })

        if acc:
            need(any(pending(v) for v in values[3:6]), 'accumulation indicators are not pending')
            constants = acc.get('mode') == 'constants'
            need(acc.get('mode') in (None, 'series', 'constants'), 'accumulate.mode must be series or constants')
            div = acc.get('dividends')
            bal = acc.get('balance')
            need(div is not None or bal is not None, 'accumulate needs dividends and/or balance')
            if div is not None:
                d = _series(div, n, 'accumulate.dividends', allow_none=False)[1:]
                P = balance(d, lay.rate)
                if bal is not None:
                    b = _series(bal, n, 'accumulate.balance')
                    dev = max([abs(P[i] - b[i]) for i in range(1, min(len(P), len(b))) if b[i] is not None] or [0.0])
                    need(dev < 1.0, 'balance re-accumulated at %.4g differs from the supplied balance by up to %.2f' % (lay.rate, dev))
                    info['balance_check'] = dev
                P = [round(v, 2) for v in P]
                AB = [0.0] + [round(v, 2) for v in d]
            else:
                P = _series(bal, n, 'accumulate.balance')
                AB = None
            P += [None] * (n + 1 - len(P))
            kept = 0
            if not new_block and not constants:
                P, kept = self._merge_points('P', start, P, 'accumulated dividend')
            need(P[blk.O] is not None, 'no accumulated dividend at the IRR end year %d' % blk.O)
            res = blk.accumulate(P)
            need(res['irr'] is not None, 'accumulation cash flows have no IRR')
            info['acc'] = {'years': sum(1 for v in P[1:] if v is not None), 'official_points_kept': kept}
            cols = {'Q': [F['Q'].format(r=start + i) for i in range(n + 1)], 'R': [F['R'].format(r=start + i) for i in range(n + 1)],
                    'S': [F['S'].format(r=start + i) for i in range(n + 1)], 'T': [F['T'].format(r=start + i) for i in range(n + 1)]}
            caches = {'Q': res['total'], 'R': res['flow'], 'S': res['f1'], 'T': res['f2']}
            if new_block:
                newcols['P'] = [('s' if v is None else 'v', ND if v is None else v, None) for v in P]
                if AB:
                    newcols['AB'] = [None] + [('v', v, None) for v in AB[1:]] + [None] * (n + 1 - len(AB))
                for c in 'QRST':
                    newcols[c] = [('f', cols[c][i], ND if caches[c][i] is None else caches[c][i]) for i in range(n + 1)]
            elif not constants:
                self._column(ops, 'P', start, P)
                if AB:
                    self._column(ops, 'AB', start, AB, first=1)
                for c in 'QRST':
                    self._formulas(ops, c, start, cols[c], caches[c])
            filled += self._summary(g, ops, 'OPQ', res, 'R', 'ST', start, terminal, end, miss, values, 3, constants)

        if add:
            need(any(pending(v) for v in values[6:9]), 'paid-up addition indicators are not pending')
            constants = add.get('mode') == 'constants'
            need(add.get('mode') in (None, 'series', 'constants'), 'paidup.mode must be series or constants')
            need(add.get('total_cv'), 'paidup.total_cv is required')
            U = _series(add['total_cv'], n, 'paidup.total_cv')
            V = [v or 0.0 for v in _series(add.get('extra_annuity'), n, 'paidup.extra_annuity')]
            W = [v or 0.0 for v in _series(add.get('extra_maturity'), n, 'paidup.extra_maturity')]
            U += [None] * (n + 1 - len(U))
            V += [0.0] * (n + 1 - len(V))
            W += [0.0] * (n + 1 - len(W))
            need(ann or not any(V), 'extra_annuity is only meaningful for an annuity')
            kept = 0
            if not new_block and not constants:
                U, kept = self._merge_points('U', start, U, 'total cash value with paid-up additions')
            need(U[blk.O] is not None, 'no total cash value at the IRR end year %d' % blk.O)
            low = [i for i in range(1, n + 1) if U[i] is not None and blk.G[i] == 0 and U[i] < blk.E[i] - 1]
            need(not low, 'total cash value is below the guaranteed cash value in year %s' % low[:5])
            res = blk.paidup(U, V, W)
            need(res['irr'] is not None, 'paid-up addition cash flows have no IRR')
            info['add'] = {'years': sum(1 for v in U[1:] if v is not None), 'official_points_kept': kept, 'extra_annuity_total': sum(V)}
            xf = [(F['X_ann'].format(r=start + i, s=start, n=i + 1) if ann else F['X_life'].format(r=start + i)) for i in range(n + 1)]
            cols = {'X': xf, 'Y': [F['Y'].format(r=start + i) for i in range(n + 1)],
                    'Z': [F['Z'].format(r=start + i) for i in range(n + 1)], 'AA': [F['AA'].format(r=start + i) for i in range(n + 1)]}
            caches = {'X': res['total'], 'Y': res['flow'], 'Z': res['f1'], 'AA': res['f2']}
            if new_block:
                newcols['U'] = [('s' if v is None else 'v', ND if v is None else v, None) for v in U]
                newcols['V'] = [('v', v, None) for v in V]
                newcols['W'] = [('v', v, None) for v in W]
                for c in ('X', 'Y', 'Z', 'AA'):
                    newcols[c] = [('f', cols[c][i], ND if caches[c][i] is None else caches[c][i]) for i in range(n + 1)]
            elif not constants:
                self._column(ops, 'U', start, U)
                self._column(ops, 'V', start, V)
                self._column(ops, 'W', start, W)
                for c in ('X', 'Y', 'Z', 'AA'):
                    self._formulas(ops, c, start, cols[c], caches[c])
            filled += self._summary(g, ops, 'RST', res, 'Y', ('Z', 'AA'), start, terminal, end, miss, values, 6, constants)

        need(filled, 'no pending indicator would be filled' + ('; years without data come before the answer for %s'
                                                               % '、'.join(self.undetermined) if self.undetermined else ''))
        if new_block:
            tpl = w.last_row()
            need(num(w.value('C%d' % tpl)) and tpl > lay.block_header, 'last row of the working sheet is not a block row; cannot take cell styles from it')
            st = lambda c: w.style_of('%s%d' % (c, tpl)) or w.style_of('%s%d' % (STYLE_FROM.get(c, c), tpl))
            a = w.value('A%d' % tpl)
            blk_id = int(a) + 1 if num(a) else len(lay.groups()) + 1
            usage = '结果表第%d行引用（底稿汇总行第%d行）' % (g.row, sr)
            for i in range(n + 1):
                cells = [('A', 'v', blk_id, None, st('A')), ('B', 'v', g.id if num(g.id) else 0, None, st('B')),
                         ('AE', 's', g.name or '', None, st('AE')), ('AF', 's', g.term or '', None, st('AF')),
                         ('AG', 's', g.pay_text or '', None, st('AG')), ('AH', 's', usage, None, st('AH'))]
                for c, col in newcols.items():
                    if col[i] is not None:
                        cells.append((c, col[i][0], col[i][1], col[i][2], st(c)))
                rows.append(cells)
        # result sheet: calculated values of the nine formulas, and the method label
        for i in range(9):
            if not _same(values[i], g.values[i]):
                ops.append(('r', '%s%d' % (RES_COLS[i], g.row), 'c', values[i], None, None))
        mark = spec['label'] + ('〔此前标注：%s〕' % g.label if g.label else '')
        ops.append(('r', 'Q%d' % g.row, 's', mark, None, 'P%d' % g.row))
        for col, extra in (('U', spec.get('note') or spec['label']), ('AC', spec['source'])):
            old = w.value('%s%d' % (col, sr))
            sep = (' ' if col == 'U' else '\n\n') if isinstance(old, str) and old else ''
            ops.append(('w', '%s%d' % (col, sr), 's', (old if isinstance(old, str) else '') + sep + extra, None, None))
        return {'row': g.row, 'id': g.id, 'name': g.name, 'pay': g.pay_text, 'term': g.term, 'summary_row': sr,
                'filled': filled, 'undetermined': list(self.undetermined), 'values': values, 'label': mark, 'info': info,
                'block': [start, terminal, end], 'new_block': new_block}, ops, rows

    def apply(self, ops, rows):
        for cells in rows:
            self.w.append_row(cells)
        for sheet, addr, kind, value, cache, style_from in ops:
            s = self.w if sheet == 'w' else self.r
            style = s.style_of(style_from) if style_from else None
            s.put(addr, kind, value, cache, style)
