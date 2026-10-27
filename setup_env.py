"""Install this project's isolated Python environment and Chromium."""
import argparse
import os
from pathlib import Path
import subprocess
import sys
import venv

ROOT = Path(__file__).resolve().parent
ENV = ROOT / '.venv'
PYTHON = ENV / ('Scripts/python.exe' if os.name == 'nt' else 'bin/python')
READY = ENV / '.xray-ready'


def install():
    if sys.version_info < (3, 11):
        raise RuntimeError('Install Python 3.11 or newer, then run setup again.')
    READY.unlink(missing_ok=True)
    venv.EnvBuilder(with_pip=True).create(ENV)
    subprocess.run([str(PYTHON), '-m', 'pip', 'install', '-r', str(ROOT / 'requirements.txt')], check=True)
    command = [str(PYTHON), '-m', 'playwright', 'install']
    if sys.platform.startswith('linux'):
        command.append('--with-deps')
    subprocess.run(command + ['chromium'], check=True)
    READY.write_text('ready\n', encoding='utf-8')
    print('Setup finished. Put Xray XLSX files in "xray helium10", then run run.bat or run.sh.')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run', action='store_true', help='Install only when needed, then run agent.py')
    args, extra = parser.parse_known_args()
    if extra and not args.run:
        parser.error('Unexpected arguments: ' + ' '.join(extra))
    try:
        if not PYTHON.is_file() or not READY.is_file():
            install()
        if args.run:
            return subprocess.run([str(PYTHON), str(ROOT / 'agent.py'), *extra], cwd=ROOT).returncode
        return 0
    except (OSError, RuntimeError, subprocess.CalledProcessError) as exc:
        print(f'Setup failed: {exc}', file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
