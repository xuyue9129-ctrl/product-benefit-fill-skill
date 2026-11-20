# -*- coding: utf-8 -*-
"""Indicator arithmetic, written to give the same numbers as the workbook formulas."""

ND = '未披露'


def irr(cf):
    """Annual IRR of cash flows at t = 0, 1, 2, ... (same root Excel's IRR finds); None when no sign change."""
    def npv(r):
        return sum(c / (1.0 + r) ** t for t, c in enumerate(cf))
    lo, hi = -0.9, 1.0
    if npv(lo) * npv(hi) > 0:
        return None
    for _ in range(200):
        mid = (lo + hi) / 2.0
        if npv(lo) * npv(mid) <= 0:
            hi = mid
        else:
            lo = mid
    return (lo + hi) / 2.0


def first_year(flags):
    """(year, known): the first policy year whose flag is 1. A year without data before the first hit makes the
    answer unknown, because the threshold may have been passed in that year; nothing is interpolated."""
    for y in range(1, len(flags)):
        if flags[y] == 1:
            return y, True
        if flags[y] != 0:
            return None, False
    return None, True


def exceeds(total, paid, k):
    """Strictly greater after rounding both sides to 6 decimals, as the sheet does."""
    return 1 if round(total, 6) > round(k * paid, 6) else 0


def balance(dividends, rate):
    """Accumulated dividend balance: dividends[i] is credited at the end of policy year i+1."""
    out = [0.0]
    for d in dividends:
        out.append(out[-1] * (1.0 + rate) + d)
    return out


class Block(object):
    """Year-by-year data of one group; index = policy year, 0 = issue date."""

    def __init__(self, premium, pay, E, F, G, H=None, terminal=None, kind='ann'):
        n = len(E) - 1
        self.premium, self.pay, self.N, self.kind = float(premium), int(pay), n, kind
        self.E = [float(v) for v in E]
        self.F = [float(v or 0) for v in F]
        self.G = [float(v or 0) for v in G]
        self.H = [float(v) for v in (H if H is not None else E)]
        self.O = int(terminal) if terminal is not None else (40 if n >= 40 else n)
        if not (len(self.F) == len(self.G) == len(self.H) == n + 1):
            raise ValueError('E, F, G, H must have the same length')
        if not 1 <= self.O <= n:
            raise ValueError('IRR end year %s is outside the block (1..%d)' % (self.O, n))
        self.D = [self.premium if y < self.pay else 0.0 for y in range(n + 1)]
        self.J = [min(y, self.pay) * self.premium for y in range(n + 1)]
        self.I = []
        s = 0.0
        for y in range(n + 1):
            s += self.F[y]
            self.I.append(s)
        if kind == 'ann':      # annuity: cash value after payment + maturity + everything received so far
            self.K = [self.E[y] + self.G[y] + self.I[y] for y in range(n + 1)]
        elif kind == 'E':
            self.K = list(self.E)
        else:                  # life: the cash value as published in the source
            self.K = list(self.H)
        self.L = [-self.D[y] + self.F[y] + self.G[y] + (self.E[y] if y == self.O else 0.0) for y in range(n + 1)]

    def _flags(self, total, k):
        return [0] + [ND if total[y] is None else exceeds(total[y], self.J[y], k) for y in range(1, self.N + 1)]

    def _triple(self, flow, total):
        f1, f2 = self._flags(total, 1), self._flags(total, 2)
        (y1, k1), (y2, k2) = first_year(f1), first_year(f2)
        return {'irr': irr(flow[:self.O + 1]), 'payback': y1, 'double': y2, 'known': (True, k1, k2), 'f1': f1, 'f2': f2}

    def guarantee(self):
        r = self._triple(self.L, self.K)
        r.update(total=self.K, flow=self.L)
        return r

    def accumulate(self, P):
        """P[y]: accumulated dividend balance at the end of year y (None = not disclosed)."""
        P = list(P) + [None] * (self.N + 1 - len(P))
        if P[self.O] is None:
            raise ValueError('no accumulated dividend at the IRR end year %d' % self.O)
        Q = [None if P[y] is None else self.K[y] + P[y] for y in range(self.N + 1)]
        R = [self.L[y] + (P[y] if y == self.O else 0.0) for y in range(self.N + 1)]
        r = self._triple(R, Q)
        r.update(total=Q, flow=R, P=P)
        return r

    def paidup(self, U, V=None, W=None):
        """U[y]: total cash value with paid-up additions; V[y]: extra annuity paid in cash; W[y]: extra maturity benefit."""
        n = self.N
        U = list(U) + [None] * (n + 1 - len(U))
        V = [float(v or 0) for v in (V or [])] + [0.0] * (n + 1 - len(V or []))
        W = [float(v or 0) for v in (W or [])] + [0.0] * (n + 1 - len(W or []))
        if U[self.O] is None:
            raise ValueError('no total cash value at the IRR end year %d' % self.O)
        X = []
        sv = 0.0
        for y in range(n + 1):
            sv += V[y]
            if U[y] is None:
                X.append(None)
            elif self.kind == 'ann':
                X.append(U[y] + self.I[y] + sv + self.G[y] + W[y])
            else:
                X.append(U[y] + ((self.K[y] - self.E[y] + W[y]) if self.K[y] > self.E[y] else 0.0))
        Y = [-self.D[y] + self.F[y] + V[y] + self.G[y] + W[y] + (U[y] if y == self.O else 0.0) for y in range(n + 1)]
        r = self._triple(Y, X)
        r.update(total=X, flow=Y, U=U, V=V, W=W)
        return r


def dividend_ratio(dividends, cash_value, pay):
    """Dividend of year y divided by the guaranteed cash value at the end of year y-1, for years after the
    premium term. With a 3.5% illustration, 1.75% accumulation rate and a 70% policyholder share the ratio
    is 70% x (3.5% - 1.75%) = 1.225%."""
    out = []
    for y in range(2, len(dividends) + 1):
        if y - 1 < len(cash_value) and cash_value[y - 1] and cash_value[y - 1] > 0 and dividends[y - 1] is not None:
            out.append((y, dividends[y - 1] / cash_value[y - 1]))
    post = [r for y, r in out if y > pay + 1] or [r for y, r in out]
    return post
