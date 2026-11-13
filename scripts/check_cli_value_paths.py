"""Exercise the native CLI: booleans, blocked directories, a sinking-object slice and per-feature filament use."""
import argparse
import json
import os
from pathlib import Path
import subprocess
import tempfile
import zipfile

CUBE_CORNERS = [(0, 0, 0), (20, 0, 0), (20, 20, 0), (0, 20, 0), (0, 0, 20), (20, 0, 20), (20, 20, 20), (0, 20, 20)]
CUBE_TRIANGLES = [(0, 2, 1), (0, 3, 2), (4, 5, 6), (4, 6, 7), (0, 1, 5), (0, 5, 4),
                  (1, 2, 6), (1, 6, 5), (2, 3, 7), (2, 7, 6), (3, 0, 4), (3, 4, 7)]


def write_cube_stl(path):
    lines = ['solid cube20']
    for triangle in CUBE_TRIANGLES:
        lines.append('facet normal 0 0 0')
        lines.append('outer loop')
        lines.extend('vertex %d %d %d' % CUBE_CORNERS[i] for i in triangle)
        lines.append('endloop')
        lines.append('endfacet')
    lines.append('endsolid cube20')
    path.write_text('\n'.join(lines) + '\n')


def write_cube_3mf(path, z):
    """A 20 mm cube whose single build item sits at height z (negative: sunk into the bed)."""
    vertices = ''.join('<vertex x="%d" y="%d" z="%d"/>' % c for c in CUBE_CORNERS)
    triangles = ''.join('<triangle v1="%d" v2="%d" v3="%d"/>' % t for t in CUBE_TRIANGLES)
    model = ('<?xml version="1.0" encoding="UTF-8"?>'
             '<model unit="millimeter" xml:lang="en-US" xmlns="http://schemas.microsoft.com/3dmanufacturing/core/2015/02">'
             '<resources><object id="1" type="model"><mesh><vertices>' + vertices + '</vertices><triangles>' + triangles +
             '</triangles></mesh></object></resources>'
             '<build><item objectid="1" transform="1 0 0 0 1 0 0 0 1 90 90 ' + str(z) + '"/></build></model>')
    types = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
             '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
             '<Default Extension="model" ContentType="application/vnd.ms-package.3dmanufacturing-3dmodel+xml"/></Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Target="/3D/3dmodel.model" Id="rel0" '
            'Type="http://schemas.microsoft.com/3dmanufacturing/2013/01/3dmodel"/></Relationships>')
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as archive:
        archive.writestr('[Content_Types].xml', types)
        archive.writestr('_rels/.rels', rels)
        archive.writestr('3D/3dmodel.model', model)


def run_slice(binary, environment, directory, model, extra):
    command = [str(binary), '--datadir', str(directory / 'data'), '--outputdir', str(directory),
               '--slice', '0'] + extra + [str(model)]
    return subprocess.run(command, env=environment, capture_output=True, text=True, timeout=1500)


def check_sinking_object(binary, environment, root):
    """A headless slice of an object sunk 5 mm into the bed must finish and write G-code (PartPlate::check_outside)."""
    directory = root / 'sinking'
    directory.mkdir()
    model = directory / 'sunk-cube.3mf'
    write_cube_3mf(model, -5)
    result = run_slice(binary, environment, directory, model, [])
    gcode = sorted(directory.glob('*.gcode'))
    return {'case': 'sinking-object-slice', 'rc': result.returncode, 'gcode': [g.name for g in gcode],
            'passed': result.returncode == 0 and bool(gcode) and gcode[0].stat().st_size > 0,
            'stderr': result.stderr[-2000:]}


def check_wall_filament_used(binary, environment, root, model):
    """Wall filament 2 with support off: filament 2 must count as used and appear in the G-code (PartPlate::get_extruders_under_cli)."""
    directory = root / 'wall-filament'
    directory.mkdir()
    filaments = []
    for name in ('cli-filament-a', 'cli-filament-b'):
        path = directory / (name + '.json')
        path.write_text(json.dumps({'type': 'filament', 'name': name, 'from': 'User'}))
        filaments.append(str(path))
    extra = ['--load-filaments', ';'.join(filaments), '--outer-wall-filament-id=2', '--inner-wall-filament-id=2',
             '--enable-support=0']
    result = run_slice(binary, environment, directory, model, extra)
    gcode = sorted(directory.glob('*.gcode'))
    used = None
    if gcode:
        for line in gcode[0].read_text(errors='replace').splitlines():
            if line.startswith('; filament used [mm] ='):
                used = [float(v) for v in line.split('=', 1)[1].split(',')]
                break
    return {'case': 'wall-filament-2-counts-as-used', 'rc': result.returncode, 'used_mm': used,
            'passed': result.returncode == 0 and used is not None and len(used) >= 2 and used[1] > 0,
            'stderr': result.stderr[-2000:]}


def main():
    parser = argparse.ArgumentParser()
    for name in ('binary', 'output'):
        parser.add_argument('--' + name, type=Path, required=True)
    parser.add_argument('--tests', type=Path, help='unused, kept for older callers')
    args = parser.parse_args()
    results = []
    environment = {k: v for k, v in os.environ.items() if k not in ('DISPLAY', 'WAYLAND_DISPLAY')}
    with tempfile.TemporaryDirectory(prefix='cli-value-paths-') as scratch:
        root = Path(scratch)
        model = root / 'cube20.stl'
        write_cube_stl(model)
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
        results.append(check_sinking_object(args.binary, environment, root))
        results.append(check_wall_filament_used(args.binary, environment, root, model))
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(results, indent=2))
    print(json.dumps({'passed': sum(r['passed'] for r in results), 'total': len(results)}))
    for r in results:
        if not r['passed']:
            print('FAILED', r['case'], 'rc=%s' % r.get('rc'), r.get('stderr', '')[-500:])
    return 0 if all(r['passed'] for r in results) else 1


if __name__ == '__main__':
    raise SystemExit(main())
