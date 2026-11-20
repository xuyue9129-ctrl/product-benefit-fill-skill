"""Validate the installed package from an unrelated cwd, with no repository imports."""
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


class Portability(unittest.TestCase):
    def test_standalone_install(self):
        root = Path(__file__).resolve().parent.parent
        env = dict(os.environ, PYTHONIOENCODING='utf-8')
        env.pop('PYTHONPATH', None)
        with tempfile.TemporaryDirectory(prefix='pbfill_install_') as tmp:
            target = Path(tmp) / '技能 with spaces' / 'product-benefit-fill'

            def run(*args):
                return subprocess.run([sys.executable, *map(str, args)], cwd=tmp, env=env,
                                      stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=180)

            installed = run(root / 'install.py', '--dest', target)
            self.assertEqual(installed.returncode, 0, installed.stderr.decode('utf-8', 'replace'))
            self.assertTrue((target / 'requirements.txt').is_file())
            smoke = run(target / 'scripts' / 'selftest.py')
            self.assertEqual(smoke.returncode, 0, (smoke.stdout + smoke.stderr).decode('utf-8', 'replace'))
            marker = target / 'keep.txt'
            marker.write_text('user file', encoding='utf-8')
            refused = run(root / 'install.py', '--dest', target)
            self.assertNotEqual(refused.returncode, 0)
            self.assertEqual(marker.read_text(encoding='utf-8'), 'user file')


if __name__ == '__main__':
    unittest.main()
