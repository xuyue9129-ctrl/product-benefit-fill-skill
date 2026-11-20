# -*- coding: utf-8 -*-
"""Independent checks of a filled workbook.

1. Excel (Windows, via COM): open the file read-only, recalculate every formula, read the result sheet.
   This proves that the formulas, not just the stored numbers, give the expected indicators.
2. Python audit: recompute every group's indicators from its year-by-year block and compare them with what the
   result sheet shows. Works anywhere, but only checks stored values.
"""
import json
import os
import re
import subprocess
import sys
import tempfile

from .layout import SUM_COLS, NAMES, pending, num

PS = r'''
param([string]$Path,[int]$Sheet,[int]$First,[int]$Last,[string]$Out)
$ErrorActionPreference='Stop'
$xl=New-Object -ComObject Excel.Application
if($xl.Workbooks.Count -ne 0){ Write-Output 'ABORT: this Excel instance already has open workbooks'; exit 3 }
$xl.Visible=$false; $xl.DisplayAlerts=$false; $xl.AskToUpdateLinks=$false
try{
  $wb=$xl.Workbooks.Open($Path,0,$true)
  $xl.CalculateFull()
  $ws=$wb.Worksheets.Item($Sheet)
  $v=$ws.Range(('H'+$First+':Q'+$Last)).Value2
  $n=$Last-$First+1
  $rows=New-Object System.Collections.ArrayList
  for($i=1;$i -le $n;$i++){ $r=@(); for($j=1;$j -le 10;$j++){ $x=$v[$i,$j]; if($x -is [int]){ $x=[string]('#ERR'+$x) }; $r+=,$x }; [void]$rows.Add($r) }
  [System.IO.File]::WriteAllText($Out,(ConvertTo-Json @{rows=$rows} -Depth 4 -Compress),(New-Object System.Text.UTF8Encoding($false)))
  $wb.Close($false)
  Write-Output 'OK'
} finally { $xl.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }
'''


STANDARD = re.compile(r'^(IRR\([A-Z]+\d+:[A-Z]+\d+\)|INDEX\(C\d+:C\d+,MATCH\(1,[A-Z]+\d+:[A-Z]+\d+,0\)\)'
                      r'|IF\(COUNTIF\([A-Z]+\d+:[A-Z]+\d+,1\)=0,"[^"]+","核查错误"\))$')


def same(a, b):
    if num(a) and num(b):
        return abs(a - b) < 5e-7
    return a == b


def excel_recalc(path, sheet_index, first, last):
    """{row: [H..Q]} after a full recalculation in Excel, or raise RuntimeError when Excel cannot be used."""
    if sys.platform != 'win32':
        raise RuntimeError('Excel COM is only available on Windows')
    d = tempfile.mkdtemp(prefix='pbfill_')
    ps, out = os.path.join(d, 'recalc.ps1'), os.path.join(d, 'vals.json')
    with open(ps, 'w', encoding='ascii') as fh:
        fh.write(PS)
    try:
        r = subprocess.run(['powershell', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps, os.path.abspath(path),
                            str(sheet_index), str(first), str(last), out], capture_output=True, timeout=1800)
        if r.returncode != 0 or not os.path.exists(out):
            msg = (r.stdout + r.stderr).decode('utf8', 'replace').strip()[-400:]
            raise RuntimeError('Excel recalculation failed: %s' % msg)
        with open(out, encoding='utf-8-sig') as fh:
            rows = json.load(fh)['rows']
    finally:
        for f in (ps, out):
            if os.path.exists(f):
                os.remove(f)
        os.rmdir(d)
    return {first + i: row for i, row in enumerate(rows)}


def compare(new_lay, recalc, base_lay=None, filled=None):
    """recalc: {row: [H..Q]} from Excel. base_lay: the workbook before filling. filled: the fill report's list."""
    res = {'groups': 0, 'error_cells': 0, 'pending': 0, 'old_changed': [], 'mismatch': [], 'newly_filled': 0, 'label_mismatch': []}
    base = {g.row: g for g in base_lay.groups()} if base_lay else {}
    exp = {r['row']: r for r in (filled or [])}
    for g in new_lay.groups():
        row = recalc.get(g.row)
        if row is None:
            res['mismatch'].append({'row': g.row, 'why': 'row missing from the recalculated range'})
            continue
        res['groups'] += 1
        for i in range(9):
            v = row[i]
            if isinstance(v, str) and v.startswith('#ERR'):
                res['error_cells'] += 1
            if pending(v):
                res['pending'] += 1
            b = base.get(g.row)
            if b is not None:
                if pending(b.values[i]):
                    res['newly_filled'] += 0 if pending(v) else 1
                elif not same(b.values[i], v):
                    res['old_changed'].append({'row': g.row, 'indicator': NAMES[i], 'before': b.values[i], 'after': v})
            e = exp.get(g.row)
            if e is not None and not same(e['values'][i], v):
                res['mismatch'].append({'row': g.row, 'indicator': NAMES[i], 'expected': e['values'][i], 'excel': v})
        e = exp.get(g.row)
        if e is not None and row[9] != e['label']:
            res['label_mismatch'].append(g.row)
    res['ok'] = not (res['error_cells'] or res['old_changed'] or res['mismatch'] or res['label_mismatch'])
    return res


def audit(lay):
    """Recompute stored indicators from the year-by-year blocks. Returns (checked, differences, skipped, special)."""
    checked, diffs, skipped, special = 0, [], [], []
    for g in lay.groups():
        if not g.block:
            continue
        try:
            blk, raw = lay.read_block(g)
        except ValueError as e:
            skipped.append({'row': g.row, 'why': str(e)})
            continue
        miss = lay.miss_text(g, blk.N)
        parts = [(0, blk.guarantee())]
        P = [v if num(v) else None for v in raw['P']]
        U = [v if num(v) else None for v in raw['U']]
        if P[blk.O] is not None and lay.work.formula('O%d' % g.summary_row):
            parts.append((3, blk.accumulate(P)))
        if U[blk.O] is not None and lay.work.formula('R%d' % g.summary_row):
            parts.append((6, blk.paidup(U, raw['V'], raw['W'])))
        for base, r in parts:
            mine = [r['irr'], r['payback'] if r['payback'] is not None else miss, r['double'] if r['double'] is not None else miss]
            for k in range(3):
                shown = g.values[base + k]
                if pending(shown) or shown is None:
                    continue
                # constants and bespoke formulas were judged by a method recorded in the row's label; they are
                # listed, not recomputed, because the standard arithmetic does not apply to them
                if not r['known'][k] or not STANDARD.match(lay.work.formula('%s%d' % (SUM_COLS[base + k], g.summary_row)) or ''):
                    special.append({'row': g.row, 'indicator': NAMES[base + k], 'sheet': shown})
                    continue
                checked += 1
                ok = abs(shown - mine[k]) < 5e-7 if num(shown) and num(mine[k]) else shown == mine[k]
                if not ok:
                    diffs.append({'row': g.row, 'indicator': NAMES[base + k], 'sheet': shown, 'recomputed': mine[k]})
    return checked, diffs, skipped, special
