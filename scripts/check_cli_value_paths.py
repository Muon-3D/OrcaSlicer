"""Exercise the native CLI's boolean values and refusal of blocked directories."""
import argparse
import json
import os
from pathlib import Path
import subprocess
import tempfile


def main():
    parser = argparse.ArgumentParser()
    for name in ('binary', 'tests', 'output'):
        parser.add_argument('--' + name, type=Path, required=True)
    args = parser.parse_args()
    results = []
    environment = {k: v for k, v in os.environ.items() if k not in ('DISPLAY', 'WAYLAND_DISPLAY')}
    model = args.tests / 'test_projects/synthetic/cube20.stl'
    with tempfile.TemporaryDirectory(prefix='cli-value-paths-') as scratch:
        root = Path(scratch)
        cases = [(word, '1') for word in ('1', 'true', 'YES', 'on', 'enabled')]
        cases += [(word, '0') for word in ('0', 'false', 'NO', 'off', 'disabled')]
        for word, expected in cases:
            directory = root / word
            directory.mkdir()
            output = directory / 'settings.json'
            command = [str(args.binary), '--datadir', str(directory / 'data'), '--outputdir', str(directory),
                       '--enable-support=' + word, '--export-settings', str(output), '--info', str(model)]
            result = subprocess.run(command, env=environment, capture_output=True, text=True, timeout=60)
            actual = json.loads(output.read_text()).get('enable_support') if output.is_file() else None
            results.append({'case': 'boolean-' + word, 'rc': result.returncode, 'expected': expected,
                            'actual': actual, 'passed': result.returncode == 0 and actual == expected,
                            'stderr': result.stderr})
        blocker = root / 'file-instead-of-directory'
        blocker.write_text('preserve these bytes')
        for kind in ('datadir', 'outputdir'):
            directory = root / ('blocked-' + kind)
            directory.mkdir()
            data = blocker / 'nested' if kind == 'datadir' else directory / 'data'
            output = blocker / 'nested' if kind == 'outputdir' else directory
            command = [str(args.binary), '--datadir', str(data), '--outputdir', str(output),
                       '--info' if kind == 'datadir' else '--export-stl', str(model)]
            result = subprocess.run(command, env=environment, capture_output=True, text=True, timeout=60)
            results.append({'case': 'blocked-' + kind, 'rc': result.returncode,
                            'passed': result.returncode > 0 and blocker.read_text() == 'preserve these bytes',
                            'stderr': result.stderr})
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(results, indent=2))
    print(json.dumps({'passed': sum(r['passed'] for r in results), 'total': len(results)}))
    return 0 if all(r['passed'] for r in results) else 1


if __name__ == '__main__':
    raise SystemExit(main())
