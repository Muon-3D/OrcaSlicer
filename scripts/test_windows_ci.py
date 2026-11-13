"""Check generated CPython linker settings and independent Windows test gates."""
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

import yaml

ROOT = Path(__file__).resolve().parents[1]


class WindowsCI(unittest.TestCase):
    def python_environment(self, arch, debug=False):
        cmake = shutil.which('cmake')
        self.assertIsNotNone(cmake, 'CMake is required for dependency configuration probes')
        with tempfile.TemporaryDirectory(prefix='orca-python-config-') as directory:
            root = Path(directory)
            output = root / 'environment.txt'
            script = root / 'probe.cmake'
            script.write_text(
                'set(WIN32 TRUE)\nset(MSVC_VERSION 1951)\n'
                f'set(CMAKE_SYSTEM_PROCESSOR {arch})\nset(CMAKE_SIZEOF_VOID_P 8)\n'
                f'set(DEP_DEBUG {"ON" if debug else "OFF"})\n'
                'macro(ExternalProject_Add)\nendmacro()\n'
                f'include("{(ROOT / "deps/python3/python3.cmake").as_posix()}")\n'
                f'file(WRITE "{output.as_posix()}" "${{_python_env_args}}")\n')
            environment = dict(os.environ, _LINK_='/DEBUG')
            subprocess.run([cmake, '-P', str(script)], env=environment, check=True,
                           capture_output=True, text=True, timeout=30)
            # Execute the actual generated CMake environment, rather than just reading source text.
            generated = output.read_text().split(';')
            result = subprocess.run([cmake, '-E', 'env', *generated, sys.executable,
                                     '-c', 'import os,json;print(json.dumps(os.environ.get("_LINK_")))'],
                                    env=environment, check=True, capture_output=True,
                                    text=True, timeout=30)
            return json.loads(result.stdout)

    def test_arm_release_limits_ltcg_threads_and_preserves_existing_flags(self):
        flags = self.python_environment('ARM64')
        self.assertIn('/DEBUG', flags)
        self.assertIn('/CGTHREADS:1', flags)

    def test_other_python_builds_preserve_linker_environment(self):
        for arch, debug in [('AMD64', False), ('ARM64', True)]:
            with self.subTest(arch=arch, debug=debug):
                self.assertEqual(self.python_environment(arch, debug), '/DEBUG')

    def test_each_arch_test_depends_only_on_its_build(self):
        jobs = yaml.safe_load((ROOT / '.github/workflows/check_windows_build.yml').read_text())['jobs']
        for arch in ('x64', 'arm64'):
            with self.subTest(arch=arch):
                builder, tests = f'windows_build_{arch}', f'windows_tests_{arch}'
                self.assertIn(builder, jobs)
                self.assertIn(tests, jobs)
                self.assertEqual(jobs[tests]['needs'], builder)
                self.assertEqual(jobs[builder]['with']['arch'], arch)
                self.assertEqual(jobs[tests]['with']['artifact'], '${{ github.sha }}-tests-windows-' + arch)
                self.assertEqual(jobs[tests]['with']['test-dir'], 'build-arm64/tests' if arch == 'arm64' else 'build/tests')
        producer = yaml.safe_load((ROOT / '.github/workflows/build_orca.yml').read_text())['jobs']['build_orca']['steps']
        upload = next(step for step in producer if step.get('name') == 'Upload Test Artifact Win')
        self.assertEqual(upload['with']['name'], '${{ github.sha }}-tests-windows-${{ inputs.arch }}')


if __name__ == '__main__':
    unittest.main()
