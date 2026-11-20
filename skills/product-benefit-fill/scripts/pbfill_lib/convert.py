# -*- coding: utf-8 -*-
"""Check saved platform replies against one workbook group and turn them into a fill spec.

Nothing is taken on trust: the platform's premium and guaranteed cash value must reproduce the figures already
in the working sheet (taken from the official cash-value table) before any dividend column is used, and the
illustration basis is measured from the numbers rather than read from the platform's label.
"""
import statistics

from .calc import dividend_ratio
from .fill import Refuse, need
from .layout import num

CV_KEYS = ('cash_value_tmp', 'base_cash_value', 'cash_value1', 'cash_value', 'total_cash_value')
BONUS_KEYS = ('bonus', 'bonus1', 'bonus_zs')
PUA_CV_KEYS = ('pua_cash_value', 'bonus_baoe_cv', 'bonus_cash_value')
PUA_ANNUITY_KEYS = ('pua_annuity', 'bonus_live_end', 'pua_live_end')
NOT_SCALED = ('baof', 'baof_sum', 'baof_start', 'baof_start_sum')


def _v(cols, k):
    return list(cols[k].get('values') or [])


def _lim(blk):
    # in the maturity year the sheet shows 0 cash value and the maturity benefit separately; platforms differ there
    return blk.N - 1 if blk.G[blk.N] else blk.N


def premium_ok(cols, blk):
    if 'baof' not in cols:
        return True
    return all(num(x) and abs(x - (blk.premium if i < blk.pay else 0.0)) < 1 for i, x in enumerate(_v(cols, 'baof')[:blk.N]))


def cv_match(cols, blk):
    """Which platform column equals the guaranteed cash value of the block. Returns (key, years, basis) where
    basis is 'E' (after the year's survival benefit, as the sheet stores it) or 'H' (as published, before it)."""
    lim = _lim(blk)
    cands = [k for k in CV_KEYS if k in cols]
    for ref, basis in ((blk.E, 'E'), (blk.H, 'H')):
        if basis == 'H' and not any(abs(blk.H[i] - blk.E[i]) > 1 for i in range(1, blk.N + 1)):
            break
        for k in cands:
            cv = _v(cols, k)
            n = len(cv)
            if n < lim - 1:
                continue
            if all(num(cv[i - 1]) and abs(cv[i - 1] - ref[i]) <= max(1.0, 1e-5 * ref[i]) for i in range(1, min(lim, n) + 1)):
                return k, n, basis
    return None, 0, None


def find_scale(cols, blk):
    """Constant factor between the sheet's and the platform's guaranteed cash value, if there is one."""
    lim = _lim(blk)
    for k in CV_KEYS[:4]:
        if k not in cols:
            continue
        cv = _v(cols, k)
        m = min(lim, len(cv))
        rr = [blk.E[i] / cv[i - 1] for i in range(1, m + 1) if num(cv[i - 1]) and cv[i - 1] > 0 and blk.E[i] > 0]
        if not rr or len(rr) < m - 1:
            continue
        f = statistics.median(rr)
        dev = max(abs(f * cv[i - 1] - blk.E[i]) for i in range(1, m + 1) if num(cv[i - 1]))
        if dev < max(0.5, 5e-4 * max(blk.E[1:])):
            return f, dev, k
    return None


def rescale(cols, sc, mode, blk):
    """mode 'all': every benefit column is listed per another sum insured -> multiply all of them.
    mode 'base': only the guaranteed cash value is; dividend columns are already on the premium scale."""
    f, _dev, key = sc
    lim = min(_lim(blk), blk.N)
    if mode == 'all':
        return {k: (c if k in NOT_SCALED else dict(c, values=[v * f if num(v) else v for v in _v(cols, k)])) for k, c in cols.items()}
    base = _v(cols, key)
    out = {}
    for k, c in cols.items():
        vals = _v(cols, k)
        if k == key:
            out[k] = dict(c, values=[blk.E[i + 1] if i + 1 <= lim else (v * f if num(v) else v) for i, v in enumerate(vals)])
        elif k in ('total_cash_value', 'cash_value', 'cash_value1'):
            out[k] = dict(c, values=[(v - base[i] + (blk.E[i + 1] if i + 1 <= lim else base[i] * f))
                                     if num(v) and i < len(base) and num(base[i]) else v for i, v in enumerate(vals)])
        else:
            out[k] = c
    return out


def _years(cols, blk):
    need(premium_ok(cols, blk), '平台年度保费与底稿保费、交费期不符')
    key, n, basis = cv_match(cols, blk)
    if key is None:
        sc = find_scale(cols, blk)
        need(sc is None, '平台保证现金价值与底稿各年度成固定比例 %.6f（平台按另一保额列示）；需人工判断红利列是否同一口径后用 --scale all 或 --scale base 重新转换'
             % (sc[0] if sc else 0))
        need(False, '平台保证现金价值与底稿不一致，也不成固定比例；不是同一产品/同一组合，不能使用')
    n = min(n, blk.N)
    need(n >= blk.N - 1, '平台只返回%d个保单年度，底稿逐年块有%d个' % (n, blk.N))
    return key, n, basis


def accumulate_part(cols, blk, expect):
    """Dividends left to accumulate. Returns (spec part, info)."""
    _key, n, basis = _years(cols, blk)
    bkey = next((k for k in BONUS_KEYS if k in cols), None)
    need(bkey, '平台该条目只返回保证利益，没有当年红利数据')
    bonus = _v(cols, bkey)
    need(len(bonus) >= n and all(num(bonus[i]) for i in range(n)), '平台当年红利列有缺失年度')
    part = {'dividends': bonus[:n]}
    if 'bonus_sum' in cols:
        bs = _v(cols, 'bonus_sum')
        part['balance'] = [bs[i] if i < len(bs) and num(bs[i]) else None for i in range(n)]
    post = dividend_ratio(bonus[:n], blk.E, blk.pay)
    need(post, '无法计算红利比值（保证现金价值为0）')
    lo, hi, med = min(post), max(post), statistics.median(post)
    info = {'years': n, 'short': n < blk.N, 'ratio_min': lo, 'ratio_max': hi, 'ratio_median': med, 'ratio_expected': expect,
            'verified': abs(med - expect) <= 5e-5 and lo >= expect - 2e-4 and hi <= expect + 2e-4, 'cv_basis': basis}
    return part, info


def paidup_part(cols, blk):
    """Dividends used to buy paid-up additions. Returns (spec part, info)."""
    key, n, basis = _years(cols, blk)
    E, F, G, H, N = blk.E, blk.F, blk.G, blk.H, blk.N
    ann = blk.kind == 'ann'
    m = min(_lim(blk), n)
    need('total_cash_value' in cols, '平台没有总现金价值列')
    tv = _v(cols, 'total_cash_value')
    need(len(tv) >= n and all(num(tv[i]) for i in range(n)), '平台总现金价值列有缺失年度')
    base = _v(cols, key)
    bk = next((k for k in PUA_CV_KEYS if k in cols), None)
    need(bk is not None or any(abs(tv[i] - base[i]) > 1.5 for i in range(m)),
         '平台该条目只返回保证利益（总现价与保证现价相同），没有红利保额/交清增额的现金价值')
    bc = _v(cols, bk) if bk else [tv[i] - base[i] for i in range(n)]
    terminal_bonus = 'bonus_end' in cols and any(abs(tv[i] - base[i] - bc[i]) > 1.5 for i in range(m))
    if not terminal_bonus:
        bad = [i + 1 for i in range(m) if abs(tv[i] - base[i] - bc[i]) > 1.5]
        need(not bad, '平台总现价≠保证现价+红利现价（%d个年度，如第%s年）' % (len(bad), bad[:3]))
    U = [0.0] + [round(tv[i] - ((H[i + 1] - E[i + 1]) if basis == 'H' else 0.0), 2) for i in range(n)]
    V = [0.0] * (n + 1)
    W = [0.0] * (n + 1)
    vnote = None
    if ann:
        ay = _v(cols, 'annuity_year') if 'annuity_year' in cols else None
        v1 = [round((ay[i] or 0) - F[i + 1] - G[i + 1], 2) for i in range(n)] if ay and len(ay) >= n else None
        bl = next((k for k in PUA_ANNUITY_KEYS if k in cols), None)
        if bl:
            V = [0.0] + [round(_v(cols, bl)[i] or 0, 2) for i in range(n)]
            vnote = '红利年金取平台单列的“%s”' % (cols[bl].get('name') or bl)
        elif v1 and all(v >= -0.01 for v in v1[:m]):
            V = [0.0] + [max(0.0, v) for v in v1]
            vnote = '红利年金=平台年金列−底稿保证给付' if sum(V[:m + 1]) else '平台年金列与底稿保证给付相同，未列示红利年金'
        else:
            need(False, '平台年金列与底稿保证给付不一致，且没有单列的红利年金')
    wnote = None
    if G[N] and n >= N:
        if E[N] == 0:
            tvn = tv[N - 1]
            if 'pua_expire_insurance' in cols:
                W[N] = round(_v(cols, 'pua_expire_insurance')[N - 1] or 0, 2)
                wnote = '平台交清增额满期金列'
            elif terminal_bonus and 'bonus_expire_insurance' in cols and abs(tvn - _v(cols, 'bonus_end')[N - 1]) < 1.5:
                W[N] = round(_v(cols, 'bonus_expire_insurance')[N - 1] + _v(cols, 'bonus_end')[N - 1], 2)
                wnote = '平台红利满期金+终了红利'
            elif 'bonus_expire_insurance' in cols and not terminal_bonus:
                W[N] = round(_v(cols, 'bonus_expire_insurance')[N - 1] or 0, 2)
                wnote = '平台红利满期金列'
            elif abs(tvn) < 1.5:
                W[N] = 0.0
            elif tvn >= G[N] - 1.5:
                W[N] = round(tvn - G[N], 2)
                wnote = '平台满期年度总现价−底稿保证满期金'
            else:
                need(bk and abs(tvn - bc[N - 1]) < 1.5, '满期年度平台总现价无法拆分出红利满期金')
                W[N] = round(bc[N - 1], 2)
                wnote = '平台满期年度红利现价'
            U[N] = 0.0
            if ann and not vnote.startswith('红利年金取'):
                V[N] = 0.0
    need(all(U[i] >= E[i] - 1 for i in range(1, m + 1)), '平台总现价低于底稿保证现价')
    rr = []
    for i in range(max(blk.pay + 2, 3), min(n, 45)):
        if E[i - 1] > 0 and E[i] > 0:
            rr.append((bc[i - 1] - bc[i - 2] * E[i] / E[i - 1]) / (E[i - 1] + bc[i - 2]))
    info = {'years': n, 'short': n < N, 'terminal_bonus': terminal_bonus, 'annuity_note': vnote, 'maturity_note': wnote,
            'implied_ratio_median': statistics.median(rr) if rr else None, 'cv_basis': basis, 'extra_annuity_total': sum(V)}
    return {'total_cv': U[1:], 'extra_annuity': V[1:], 'extra_maturity': W[1:]}, info


def pct(x):
    return '%.3f%%' % (x * 100)


def label(info, date, scale_note=''):
    """Draft of the per-row basis label. The user of the skill may reword it but must keep every stated fact."""
    out = []
    a = info.get('accumulate')
    if a:
        rng = pct(a['ratio_median']) if a['ratio_max'] - a['ratio_min'] < 2e-4 else \
            '%s—%s（中位%s）' % (pct(a['ratio_min']), pct(a['ratio_max']), pct(a['ratio_median']))
        if a['verified']:
            s = '累积生息—平台公开预览演示逐年数据；当年红利÷上年末保证现价交费期满后为%s，与该演示利率对应的%s一致，按比值认定' % (rng, pct(a['ratio_expected']))
        else:
            s = '累积生息—取自平台公开预览演示逐年数据，按平台原样填列；当年红利÷上年末保证现价交费期满后为%s，不是目标演示利率对应的%s，未能认定为目标口径' \
                % (rng, pct(a['ratio_expected']))
        if a['short']:
            s += '；平台数据比底稿逐年块少1个年度'
        out.append(s)
    b = info.get('paidup')
    if b:
        s = '交清增额—取自平台公开预览演示逐年数据，按平台原样填列'
        if b['implied_ratio_median'] is not None:
            s += '；按红利现价逐年反推的红利比值中位约%s' % pct(b['implied_ratio_median'])
        s += '；同一条目累积生息已按比值认定' if a and a['verified'] else '，未能认定为目标口径'
        if b['terminal_bonus']:
            s += '；含终了红利（平台总现价=保证现价+红利现价+终了红利）'
        if b['annuity_note']:
            s += '；' + b['annuity_note']
        if b['cv_basis'] == 'H':
            s += '；平台现价为当年生存金给付前口径，已扣除当年保证生存金'
        if b['maturity_note']:
            s += '；满期年度增加的满期金取' + b['maturity_note']
        if b['short']:
            s += '；平台数据比底稿逐年块少1个年度'
        out.append(s)
    for k, name in (('accumulate', '累积生息'), ('paidup', '交清增额')):
        if info.get(k + '_refused'):
            out.append('%s仍无数据：%s' % (name, info[k + '_refused']))
    return '%s填列：%s%s' % (date, '；'.join(out), scale_note)


def build(lay, g, acc_cols, add_cols, scale=None, demo=0.035, pricing=0.0175, share=0.7):
    """Returns (spec fragment, info). A part that cannot be used is left out and the reason is in info."""
    blk, _raw = lay.read_block(g)
    expect = share * (demo - pricing)
    spec, info = {}, {'ratio_expected': expect}
    scale_note = ''
    if scale:
        need(scale in ('all', 'base'), '--scale must be all or base')
        fs = [find_scale(c, blk) for c in (acc_cols, add_cols) if c is not None]
        need(fs and all(fs) and all(abs(f[0] - fs[0][0]) < 1e-7 * fs[0][0] for f in fs), '平台保证现金价值与底稿不成固定比例，不能折算')
        acc_cols = rescale(acc_cols, fs[0], scale, blk) if acc_cols is not None else None
        add_cols = rescale(add_cols, fs[0], scale, blk) if add_cols is not None else None
        info['scale'] = {'factor': fs[0][0], 'max_deviation': fs[0][1], 'mode': scale}
        scale_note = ('；平台按另一基本保额列示全部利益，已乘以恒定系数%.6f（底稿保证现价÷平台保证现价，各年度恒定）折算，属折算值' if scale == 'all'
                      else '；平台按另一基本保额列示保证现价（底稿保证现价÷平台保证现价各年度恒为%.6f），红利数值按平台原值取用') % fs[0][0]
    for name, cols, fn in (('accumulate', acc_cols, lambda c: accumulate_part(c, blk, expect)), ('paidup', add_cols, lambda c: paidup_part(c, blk))):
        if cols is None:
            continue
        try:
            spec[name], info[name] = fn(cols)
        except Refuse as e:
            info[name + '_refused'] = str(e)
    info['scale_note'] = scale_note
    return spec, info
