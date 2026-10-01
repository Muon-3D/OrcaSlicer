import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]


class ComparisonTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        spec = importlib.util.spec_from_file_location('comparison', ROOT / 'scripts/run_linux_cli_comparison.py')
        cls.module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(cls.module)

    def repo(self, path):
        path.mkdir()
        subprocess.run(['git', 'init', '-q', str(path)], check=True)
        (path / 'input.txt').write_text('fixed input\n')
        subprocess.run(['git', '-C', str(path), 'add', 'input.txt'], check=True)
        subprocess.run(['git', '-C', str(path), '-c', 'user.name=Test', '-c', 'user.email=test@example.invalid',
                        'commit', '-qm', 'fixture'], check=True)
        return subprocess.check_output(['git', '-C', str(path), 'rev-parse', 'HEAD'], text=True).strip()

    def test_real_failed_child_retains_cases_and_exact_identities(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, tests = root / 'source', root / 'tests'
            source_sha, tests_sha = self.repo(source), self.repo(tests)
            binary = root / 'binary'
            binary.write_bytes(b'exact executable bytes')
            output = root / 'result'
            xml = '<testsuites><testsuite><testcase classname="cli" name="pass"/><testcase classname="cli" name="crash"><failure message="signal 11"/></testcase><testcase classname="cli" name="known"><skipped type="pytest.xfail"/></testcase></testsuite></testsuites>'
            command = [sys.executable, '-c', 'import pathlib,sys;pathlib.Path(sys.argv[1]).write_text(sys.argv[2]);print("real failure");sys.exit(1)',
                       str(output / 'junit.xml'), xml]
            code = self.module.run_suite(source, source_sha, tests, tests_sha, binary, output, command=command)
            self.assertEqual(code, 1)
            report = json.loads((output / 'report.json').read_text())
            self.assertEqual(report['source_sha'], source_sha)
            self.assertEqual(report['test_sha'], tests_sha)
            self.assertEqual(report['executable_sha256'], hashlib.sha256(binary.read_bytes()).hexdigest())
            self.assertEqual([c['status'] for c in report['cases']], ['passed', 'failed', 'xfailed'])
            self.assertIn('real failure', (output / 'pytest.log').read_text())

    def test_wrong_revision_and_dirty_source_are_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / 'source'
            sha = self.repo(root)
            with self.assertRaises(ValueError):
                self.module.verify_checkout(root, '0' * 40)
            (root / 'input.txt').write_text('changed')
            with self.assertRaises(ValueError):
                self.module.verify_checkout(root, sha)

    def test_missing_junit_cannot_be_a_success(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, tests = root / 'source', root / 'tests'
            source_sha, tests_sha = self.repo(source), self.repo(tests)
            binary = root / 'binary'
            binary.write_bytes(b'binary')
            code = self.module.run_suite(source, source_sha, tests, tests_sha, binary, root / 'result',
                                         command=[sys.executable, '-c', 'print("no result")'])
            self.assertNotEqual(code, 0)
            report = json.loads((root / 'result/report.json').read_text())
            self.assertEqual(report['status'], 'incomplete')

    def test_failed_exit_with_only_passing_cases_is_incomplete(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, tests = root / 'source', root / 'tests'
            source_sha, tests_sha = self.repo(source), self.repo(tests)
            binary = root / 'binary'
            binary.write_bytes(b'binary')
            output = root / 'result'
            command = [sys.executable, '-c', 'import pathlib,sys;pathlib.Path(sys.argv[1]).write_text(sys.argv[2]);sys.exit(1)',
                       str(output / 'junit.xml'), '<testsuite><testcase name="pass"/></testsuite>']
            code = self.module.run_suite(source, source_sha, tests, tests_sha, binary, output, command=command)
            self.assertEqual(code, 2)
            self.assertEqual(json.loads((output / 'report.json').read_text())['status'], 'incomplete')

    def report(self, cases, **changes):
        return dict(source_sha='a' * 40, test_sha='b' * 40, test_tree='c' * 40,
                    executable_sha256='d' * 64, status='complete', cases=cases,
                    environment={'cache_key': 'same', 'packages': 'same'}, **changes)

    def test_matching_baseline_failures_are_not_promoted_to_pass(self):
        cases = [{'id': 'cli:crash', 'status': 'failed'}, {'id': 'cli:ok', 'status': 'passed'}]
        result = self.module.compare_reports(self.report(cases), self.report(cases))
        self.assertEqual(result['status'], 'baseline-failures')
        self.assertEqual(result['new_failures'], [])
        self.assertEqual(result['matching_failures'], ['cli:crash'])
        self.assertFalse(result['verification_passed'])

    def test_new_failure_is_reported(self):
        base = self.report([{'id': 'cli:ok', 'status': 'passed'}])
        candidate = self.report([{'id': 'cli:ok', 'status': 'failed'}])
        result = self.module.compare_reports(base, candidate)
        self.assertEqual(result['status'], 'regression')
        self.assertEqual(result['new_failures'], ['cli:ok'])

    def test_mismatched_harness_environment_or_cases_is_rejected(self):
        base = self.report([{'id': 'one', 'status': 'passed'}])
        for change in ({'test_sha': 'e' * 40}, {'environment': {'cache_key': 'different', 'packages': 'same'}},
                       {'cases': []}, {'status': 'incomplete'}):
            with self.subTest(change=change), self.assertRaises(ValueError):
                self.module.compare_reports(base, {**base, **change})

    def test_partial_inventory_is_rejected_even_when_both_reports_match(self):
        base = self.report([{'id': 'one', 'status': 'passed'}])
        base['expected_case_count'] = 93
        with self.assertRaises(ValueError):
            self.module.compare_reports(base, dict(base))


if __name__ == '__main__':
    unittest.main()
