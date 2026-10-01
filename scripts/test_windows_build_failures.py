"""Exercise the actual Windows batch build with small, compiler-free projects."""
import os
import pathlib
import shutil
import subprocess
import tempfile
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]


@unittest.skipUnless(os.name == 'nt', 'Windows batch execution required')
class WindowsBuildFailures(unittest.TestCase):
    def run_build(self, project, target='deps'):
        self.assertIsNotNone(shutil.which('cmake'))
        self.assertIsNotNone(shutil.which('ninja'))
        with tempfile.TemporaryDirectory(prefix='orca-build-failure-') as directory:
            root = pathlib.Path(directory)
            shutil.copyfile(ROOT / 'build_release_vs.bat', root / 'build_release_vs.bat')
            (root / 'deps').mkdir()
            (root / 'scripts').mkdir()
            (root / 'scripts/run_gettext.bat').write_text('@exit /b 0\n')
            source = root / 'deps' if target == 'deps' else root
            (source / 'CMakeLists.txt').write_text(
                'cmake_minimum_required(VERSION 3.20)\nproject(FailureProbe NONE)\n' + project)
            environment = dict(os.environ)
            environment.pop('ORCA_UPDATER_SIG_KEY', None)
            return subprocess.run(['cmd.exe', '/d', '/c', 'build_release_vs.bat', target, '-x'],
                                  cwd=root, env=environment, capture_output=True, text=True, timeout=60)

    def test_failed_dependency_configure_returns_failure(self):
        result = self.run_build('message(FATAL_ERROR "intentional configure failure")\n')
        self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_failed_dependency_build_returns_failure(self):
        result = self.run_build('add_custom_target(deps COMMAND ${CMAKE_COMMAND} -E false)\n')
        self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_failed_slicer_install_returns_failure(self):
        result = self.run_build('add_custom_target(probe ALL)\n'
                                'install(CODE "message(FATAL_ERROR \\\"intentional install failure\\\")")\n',
                                target='slicer')
        self.assertNotEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_successful_ninja_slicer_build_and_install(self):
        result = self.run_build('add_custom_target(probe ALL COMMAND ${CMAKE_COMMAND} -E echo "slicer probe built")\n'
                                'install(CODE "message(STATUS \\\"slicer probe installed\\\")")\n', target='slicer')
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn('slicer probe built', result.stdout)
        self.assertIn('slicer probe installed', result.stdout)
        self.assertNotIn('unknown target', result.stderr)

    def test_successful_dependency_build_returns_success(self):
        result = self.run_build('add_custom_target(deps COMMAND ${CMAKE_COMMAND} -E echo "probe complete")\n')
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn('probe complete', result.stdout)


if __name__ == '__main__':
    unittest.main()
