#!/usr/bin/env python3
"""Offline, fresh-process smoke test. Works with only the installed skill directory."""
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile


def main():
    skill = Path(__file__).resolve().parent.parent
    cli = skill / 'scripts' / 'pbfill.py'
    env = dict(os.environ, PYTHONIOENCODING='utf-8')
    env.pop('PYTHONPATH', None)
    with tempfile.TemporaryDirectory(prefix='pbfill_smoke_') as tmp:
        work = Path(tmp) / 'portable space'
        work.mkdir()

        def run(*args):
            result = subprocess.run([sys.executable, *map(str, args)], cwd=str(work), env=env,
                                    stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=120)
            if result.returncode:
                raise RuntimeError((result.stdout + result.stderr).decode('utf-8', 'replace'))

        run(skill / 'examples' / 'make_demo_workbook.py', 'demo.xlsx')
        run(cli, 'inspect', 'demo.xlsx', '--json', 'before.json')
        run(cli, 'fill', 'demo.xlsx', '--spec', skill / 'examples' / 'spec_*.json',
            '--out', 'filled.xlsx', '--report', 'report.json')
        run(cli, 'verify', 'filled.xlsx', '--base', 'demo.xlsx', '--report', 'report.json',
            '--no-excel', '--json', 'verify.json')
        run(cli, 'inspect', 'filled.xlsx', '--json', 'after.json')
        with (work / 'report.json').open(encoding='utf-8') as fh:
            report = json.load(fh)
        with (work / 'after.json').open(encoding='utf-8') as fh:
            after = json.load(fh)['stats']
        with (work / 'verify.json').open(encoding='utf-8') as fh:
            audit = json.load(fh)['audit']
        if (report['refused'] or sum(len(r['filled']) for r in report['filled']) != 10
                or after['pending_guarantee'] != 0 or after['pending_dividend'] != 5
                or not audit['ok'] or audit['blocks_not_readable']):
            raise RuntimeError('Synthetic example did not meet expected fill/audit results')
    print('PASS: standalone skill, fresh Python processes, paths with spaces, 10 indicators filled; offline audit passed.')
    print('Excel recalculation and live platform access were NOT tested.')


if __name__ == '__main__':
    main()
