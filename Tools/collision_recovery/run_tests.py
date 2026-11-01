#!/usr/bin/env python3
"""Run impact unit tests and integration tests against the built SITL EKF.

Requires system GoogleTest (libgtest-dev), g++, and a completed px4_sitl_default
build. Uses a temporary output directory and never starts PX4 or connects to hardware.
"""
import json
from pathlib import Path
import shlex
import subprocess
import tempfile


def main():
    root = Path(__file__).resolve().parents[2]
    build = root / 'build/px4_sitl_default'
    database = json.loads((build / 'compile_commands.json').read_text())
    entry = next(item for item in database if item['file'].endswith('/EKF/covariance.cpp')
                 and 'CMakeFiles/ecl_EKF.dir/' in item['command'])
    arguments = shlex.split(entry['command'])
    flags = []
    for index, argument in enumerate(arguments):
        if argument.startswith(('-I', '-D')):
            flags.append(argument)
        if argument == '-include':
            flags.extend([argument, arguments[index + 1]])
    libraries = [
        'src/modules/ekf2/EKF/libecl_EKF.a',
        'src/modules/ekf2/EKF/bias_estimator/libbias_estimator.a',
        'src/modules/ekf2/EKF/output_predictor/liboutput_predictor.a',
        'src/modules/ekf2/EKF/yaw_estimator/libyaw_estimator.a',
        'src/lib/geo/libgeo.a',
        'src/lib/lat_lon_alt/liblat_lon_alt.a',
        'src/lib/world_magnetic_model/libworld_magnetic_model.a',
    ]
    with tempfile.TemporaryDirectory(prefix='px4-impact-tests-') as output:
        unit = str(Path(output) / 'unit')
        integration = str(Path(output) / 'ekf')
        subprocess.run([
            'g++', '-std=gnu++14', '-O1', '-g', '-Wall', '-Wextra', '-Werror',
            '-fsanitize=address,undefined', '-fno-omit-frame-pointer',
            str(root / 'src/lib/collision_recovery/collision_recovery_test.cpp'),
            '-lgtest_main', '-lgtest', '-pthread', '-o', unit,
        ], check=True)
        subprocess.run([unit], check=True)
        # These tests use the real production EKF library and its build configuration.
        # Only the standalone unit binary above is sanitizer-instrumented.
        subprocess.run([
            'g++', '-std=gnu++14', '-O1', '-g', *flags,
            '-I' + str(root / 'src/modules/ekf2'),
            str(root / 'src/modules/ekf2/test/test_EKF_impact.cpp'),
            '-Wl,--start-group', *[str(build / lib) for lib in libraries],
            '-Wl,--end-group', '-lgtest_main', '-lgtest', '-pthread', '-o', integration,
        ], check=True)
        subprocess.run([integration], check=True)


if __name__ == '__main__':
    main()
