"""Check the Muon3D M1 start gcode leaves the purge to the printer (KAN-990).

The M1's start gcode used to draw its own purge, at fixed coordinates along
the bed's front edge at Z 0.15. That edge is outside every adaptive bed mesh,
and the M1's bed cannot be levelled, so the nozzle met the plate and the
printer's contact guard aborted the print. The printer now plans the purge:
PRINT_START works out a lane next to the print, probes it with the mesh, and
PURGE_LINE draws along it (Muon3D_Klipper, core/M1/macros/print.cfg).

So for every Muon3D printer profile this checks that the effective
machine_start_gcode:

* calls PRINT_START with the adaptive mesh bounds, which is what the printer
  plans the lane from, and PURGE_LANE=1, which asks it to;
* calls PURGE_LINE once, after PRINT_START;
* has no move of its own that extrudes or goes below purge height.

    python3 scripts/check_muon_m1_start_gcode.py
"""

import json
import re
import sys
from pathlib import Path

VENDOR = Path(__file__).resolve().parent.parent / "resources" / "profiles" / "Muon3D"
# At or above this the nozzle is travelling, not purging.
PURGE_HEIGHT = 1.0
LANE_PARAMS = (
    "MESH_MIN={adaptive_bed_mesh_min[0]},{adaptive_bed_mesh_min[1]}",
    "MESH_MAX={adaptive_bed_mesh_max[0]},{adaptive_bed_mesh_max[1]}",
    "PURGE_LANE=1",
)
MOVE = re.compile(r"^G[0-3](?:\s|$)", re.I)
WORD = re.compile(r"([A-Z])\s*(-?[0-9.]+)", re.I)


def problems(gcode):
    """What is wrong with one start gcode, as a list of messages."""
    found = []
    lines = [line.split(";", 1)[0].strip() for line in gcode.splitlines()]
    commands = [line.split()[0].upper() if line else "" for line in lines]

    starts = [i for i, c in enumerate(commands) if c == "PRINT_START"]
    purges = [i for i, c in enumerate(commands) if c == "PURGE_LINE"]
    if len(starts) != 1:
        found.append("calls PRINT_START %d times, not once" % len(starts))
    else:
        for param in LANE_PARAMS:
            if param not in lines[starts[0]].split():
                found.append("PRINT_START does not pass " + param)
    if len(purges) != 1:
        found.append("calls PURGE_LINE %d times, not once" % len(purges))
    elif starts and purges[0] < starts[0]:
        found.append("calls PURGE_LINE before PRINT_START")

    for number, line in enumerate(lines, 1):
        if not MOVE.match(line):
            continue
        words = {letter.upper(): float(value)
                 for letter, value in WORD.findall(line[2:])}
        if words.get("E", 0.0) > 0.0:
            found.append("line %d extrudes: %s" % (number, line))
        if "Z" in words and words["Z"] < PURGE_HEIGHT:
            found.append("line %d goes to purge height: %s" % (number, line))
    return found


def profiles():
    """Every Muon3D printer profile, by name, with inherits resolved."""
    by_name = {}
    for path in sorted((VENDOR / "machine").glob("*.json")):
        data = json.loads(path.read_text(encoding="utf-8"))
        by_name[data["name"]] = data

    def resolve(name, seen=()):
        data = by_name[name]
        parent = data.get("inherits")
        if not parent:
            return dict(data)
        if parent in seen:
            raise ValueError("inherits loop at " + name)
        merged = resolve(parent, seen + (name,))
        merged.update(data)
        return merged

    return {name: resolve(name) for name, data in by_name.items()
            if data.get("instantiation") == "true"}


# The purge this replaced, as it opened. The check has to refuse it.
OLD_PURGE = """PRINT_START MESH_MIN={adaptive_bed_mesh_min[0]},{adaptive_bed_mesh_min[1]} MESH_MAX={adaptive_bed_mesh_max[0]},{adaptive_bed_mesh_max[1]}
G0 X130.0 Y-1.5 F8000
G0 Z0.15
G1 X134 E5 F300
"""


def main():
    if not problems(OLD_PURGE):
        print("check_muon_m1_start_gcode: the check accepts the old purge")
        return 1
    printers = profiles()
    if not printers:
        print("check_muon_m1_start_gcode: no Muon3D printer profiles found")
        return 1
    failed = False
    for name, data in sorted(printers.items()):
        gcode = data.get("machine_start_gcode")
        found = ["has no machine_start_gcode"] if gcode is None else problems(gcode)
        for message in found:
            print("check_muon_m1_start_gcode: %s: %s" % (name, message))
        failed = failed or bool(found)
    if failed:
        return 1
    print("check_muon_m1_start_gcode: %d printer profiles leave the purge to"
          " PURGE_LINE" % len(printers))
    return 0


if __name__ == "__main__":
    sys.exit(main())
