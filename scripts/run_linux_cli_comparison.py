"""Run and compare an immutable pair of Linux CLI regression builds."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import xml.etree.ElementTree as ET


def git(path, *args):
    return subprocess.check_output(['git', '-C', str(path), *args], text=True, timeout=30).strip()


def verify_checkout(path, expected):
    if not re.fullmatch('[0-9a-f]{40}', expected) or git(path, 'rev-parse', 'HEAD') != expected:
        raise ValueError('Checkout does not match the requested full revision')
    if subprocess.run(['git', '-C', str(path), 'diff', '--quiet', 'HEAD'], timeout=30).returncode:
        raise ValueError('Tracked source differs from the requested revision')
    return git(path, 'rev-parse', 'HEAD^{tree}')


def sha256(path):
    digest = hashlib.sha256()
    with Path(path).open('rb') as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b''):
            digest.update(block)
    return digest.hexdigest()


def read_cases(path):
    cases = []
    seen = set()
    for case in ET.parse(path).iter('testcase'):
        identity = case.get('classname', '') + ':' + case.get('name', '')
        if identity in seen:
            raise ValueError('Duplicate JUnit case: ' + identity)
        seen.add(identity)
        failure, error, skipped = case.find('failure'), case.find('error'), case.find('skipped')
        status = 'failed' if failure is not None else 'error' if error is not None else 'passed'
        if skipped is not None:
            status = 'xfailed' if skipped.get('type') == 'pytest.xfail' else 'skipped'
        detail = failure if failure is not None else error if error is not None else skipped
        cases.append({'id': identity, 'status': status, 'seconds': case.get('time'),
                      'detail': '' if detail is None else (detail.get('message', '') + '\n' + (detail.text or ''))[:4000]})
    if not cases:
        raise ValueError('JUnit contains no test cases')
    return cases


def run_suite(source, source_sha, tests, test_sha, binary, output, command=None, timeout=1200):
    source, tests, binary, output = map(lambda p: Path(p).resolve(), (source, tests, binary, output))
    output.mkdir(parents=True)  # Refuse a reused output directory and its stale receipts.
    report = {'source_sha': source_sha, 'test_sha': test_sha, 'cases': [], 'status': 'incomplete',
              'started_at': datetime.datetime.now(datetime.timezone.utc).isoformat(),
              'run_id': os.environ.get('GITHUB_RUN_ID'), 'harness_sha': os.environ.get('HARNESS_SHA'),
              'expected_case_count': int(os.environ['EXPECTED_CASE_COUNT']) if 'EXPECTED_CASE_COUNT' in os.environ else None}
    code = 2
    try:
        report['source_tree'] = verify_checkout(source, source_sha)
        report['test_tree'] = verify_checkout(tests, test_sha)
        report['executable_sha256'] = sha256(binary)
        report['source_file_sha256'] = {str(p): sha256(source / p) for p in (
            Path('src/OrcaSlicer.cpp'), Path('src/libslic3r/Config.cpp'), Path('src/libslic3r/PrintConfig.cpp'))
            if (source / p).is_file()}
        packages = subprocess.check_output([sys.executable, '-m', 'pip', 'freeze'], text=True, timeout=30)
        report['environment'] = {'cache_key': os.environ.get('DEPS_CACHE_KEY'),
                                 'packages': sorted(packages.splitlines()),
                                 'python': sys.version, 'compiler': os.environ.get('COMPILER_ID'),
                                 'cmake': os.environ.get('CMAKE_ID'), 'build_command': './build_linux.sh -istrlL'}
        if command is None:
            command = [sys.executable, '-m', 'pytest', '.', '-c', 'pytest.ini', '--orca-bin', str(binary),
                       '--orca-source', str(source), '-v', '-o', 'xfail_strict=false', '-rxX',
                       '-n', '2', '--dist', 'loadfile', '--junitxml', str(output / 'junit.xml')]
        with (output / 'pytest.log').open('w', encoding='utf-8') as stream:
            result = subprocess.run(command, cwd=tests, stdout=stream, stderr=subprocess.STDOUT,
                                    timeout=timeout, text=True)
        report['test_exit_code'] = result.returncode
        report['cases'] = read_cases(output / 'junit.xml')
        if report['expected_case_count'] is not None and len(report['cases']) != report['expected_case_count']:
            raise ValueError('The pinned test inventory is incomplete')
        failed = any(c['status'] in ('failed', 'error') for c in report['cases'])
        if result.returncode != (1 if failed else 0):
            raise ValueError('Test exit status is inconsistent or the suite did not complete')
        report['status'] = 'complete'
        code = result.returncode
    except (OSError, ValueError, subprocess.SubprocessError, ET.ParseError) as error:
        report['failure'] = repr(error)
    finally:
        report['finished_at'] = datetime.datetime.now(datetime.timezone.utc).isoformat()
        (output / 'report.json').write_text(json.dumps(report, indent=2), encoding='utf-8')
    print(json.dumps({'status': report['status'], 'source_sha': source_sha, 'test_sha': test_sha,
                      'exit_code': code, 'case_count': len(report['cases'])}), flush=True)
    return code


def compare_reports(base, candidate):
    for report in (base, candidate):
        if report['status'] != 'complete' or not report['cases']:
            raise ValueError('Both complete case reports are required')
        if report.get('expected_case_count') is not None and len(report['cases']) != report['expected_case_count']:
            raise ValueError('The pinned test inventory is incomplete')
    if base.get('expected_case_count') != candidate.get('expected_case_count'):
        raise ValueError('The expected test inventories differ')
    for key in ('test_sha', 'test_tree', 'environment'):
        if base[key] != candidate[key]:
            raise ValueError('Comparison inputs differ: ' + key)
    left, right = ({c['id']: c['status'] for c in report['cases']} for report in (base, candidate))
    if left.keys() != right.keys():
        raise ValueError('The test case inventories differ')
    failures = lambda cases: {key for key, status in cases.items() if status in ('failed', 'error')}
    old, new = failures(left), failures(right)
    introduced, resolved, matching = new - old, old - new, old & new
    status = 'regression' if introduced else 'baseline-failures' if new else 'passed'
    return {'status': status, 'verification_passed': not new,
            'base_source': base['source_sha'], 'candidate_source': candidate['source_sha'],
            'base_executable_sha256': base['executable_sha256'], 'candidate_executable_sha256': candidate['executable_sha256'],
            'test_sha': base['test_sha'], 'test_tree': base['test_tree'],
            'matching_failures': sorted(matching), 'new_failures': sorted(introduced),
            'resolved_failures': sorted(resolved),
            'cases': [{'id': key, 'base': left[key], 'candidate': right[key]} for key in sorted(left)]}


def main():
    parser = argparse.ArgumentParser()
    commands = parser.add_subparsers(dest='mode', required=True)
    run = commands.add_parser('run')
    for name in ('source', 'tests', 'binary', 'output'):
        run.add_argument('--' + name, type=Path, required=True)
    for name in ('source-sha', 'test-sha'):
        run.add_argument('--' + name, required=True)
    compare = commands.add_parser('compare')
    for name in ('base', 'candidate', 'output'):
        compare.add_argument('--' + name, type=Path, required=True)
    args = parser.parse_args()
    if args.mode == 'run':
        return run_suite(args.source, args.source_sha, args.tests, args.test_sha, args.binary, args.output)
    result = compare_reports(json.loads(args.base.read_text()), json.loads(args.candidate.read_text()))
    args.output.write_text(json.dumps(result, indent=2), encoding='utf-8')
    print(json.dumps(result), flush=True)
    return 0 if result['verification_passed'] else 1


if __name__ == '__main__':
    sys.exit(main())
