#!/usr/bin/env python3
"""Copy the self-contained skill to an explicitly chosen directory. Never overwrite."""
import argparse
from pathlib import Path
import shutil


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--dest', required=True, help='Full destination directory, including product-benefit-fill')
    args = parser.parse_args()
    source = Path(__file__).resolve().parent / 'skills' / 'product-benefit-fill'
    dest = Path(args.dest).expanduser().resolve()
    if dest == source or source in dest.parents:
        parser.error('destination must be outside the source skill directory')
    if dest.exists():
        parser.error('destination already exists; choose a new directory')
    shutil.copytree(source, dest, ignore=shutil.ignore_patterns('__pycache__', '*.pyc'))
    print('Installed: %s' % dest)
    print('Next: install requirements.txt with the Python interpreter that will run scripts/selftest.py')


if __name__ == '__main__':
    main()
