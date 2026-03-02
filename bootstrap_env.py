#!/usr/bin/env python3
import sys
import os
import subprocess
import shutil
import time

def resource_paths(*relative_paths):
    """
    Return existing absolute paths for all relative folder names
    Works for PyInstaller (_MEIPASS) or normal execution
    """
    base_path = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    paths = []
    for rel in relative_paths:
        full_path = os.path.join(base_path, rel)
        if os.path.isdir(full_path):
            paths.append(full_path)
    return paths

def run(cmd):
    print(f"▶ {' '.join(cmd)}")
    subprocess.check_call(cmd)

def main():
    start_time = time.time()
    root = os.getcwd()
    env_dir = os.path.join(root, '.venv')

    # Detect module folders
    module_dirs = resource_paths('modules', 'modules13')
    if not module_dirs:
        print("❌ No module folders found (modules or modules13).")
        sys.exit(1)
    print("Modules folders detected:", module_dirs)
    for md in module_dirs:
        print(f"{md} contains: {os.listdir(md)}")

    # Locate host Python
    if os.name == 'nt':
        host_py = shutil.which('py') or shutil.which('python')
    else:
        host_py = shutil.which('python')

    if not host_py:
        print("❌ No Python interpreter found.")
        sys.exit(1)

    print(f"Using host Python: {host_py}")

    # Check if pip exists
    try:
        subprocess.check_call([host_py, '-m', 'pip', '--version'],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        have_pip = True
    except Exception:
        have_pip = False

    # Bootstrap pip if missing
    if not have_pip:
        get_pip = None
        for md in module_dirs:
            candidate = os.path.join(md, 'get-pip.py')
            if os.path.isfile(candidate):
                get_pip = candidate
                break
        if get_pip:
            run([host_py, get_pip, '--no-index', '--find-links', md])
        else:
            print("❌ get-pip.py not found.")
            sys.exit(1)

    # Create virtual environment if missing
    if not os.path.isdir(env_dir):
        run([host_py, '-m', 'venv', env_dir])

    # Determine venv python & pip
    venv_py = os.path.join(env_dir, 'Scripts', 'python.exe') if os.name=='nt' else os.path.join(env_dir, 'bin', 'python')
    pip_cmd = [venv_py, '-m', 'pip']

    # Print Python info
    py_version = subprocess.check_output([venv_py, '--version']).decode().strip()
    print(f"Using venv Python: {venv_py}, version: {py_version}")

    # Install wheels
    skip_prefixes = [
        'pyinstaller', 'altgraph', 'pefile',
        'packaging', 'pyinstaller_hooks_contrib', 'pywin32_ctypes'
    ]
    ordered_prefixes = [
        'wheel', 'setuptools', 'tzdata', 'six',
        'python_dateutil', 'pytz', 'et_xmlfile',
        'openpyxl', 'xlsxwriter', 'numpy', 'pandas',
        'pywin32', 'xlwings'
    ]

    installed = set()
    # Gather all wheel files from all module folders
    all_wheels = []
    for md in module_dirs:
        for fname in sorted(os.listdir(md)):
            if fname.lower().endswith('.whl'):
                all_wheels.append((md, fname))

    # Pass 1: ordered installation
    for prefix in ordered_prefixes:
        for md, fname in all_wheels:
            if fname in installed:
                continue
            if not fname.lower().startswith(prefix):
                continue
            path = os.path.join(md, fname)
            print(f"🆕 Installing {fname} ...")
            try:
                run(pip_cmd + ['install', '--no-index', path])
                installed.add(fname)
            except subprocess.CalledProcessError:
                print(f"⚠️ Failed to install {fname}, skipping.")

    # Pass 2: remaining wheels
    for md, fname in all_wheels:
        if fname in installed:
            continue
        if any(fname.lower().startswith(pref) for pref in skip_prefixes):
            continue
        path = os.path.join(md, fname)
        print(f"🆕 Installing {fname} ...")
        try:
            run(pip_cmd + ['install', '--no-index', path])
            installed.add(fname)
        except subprocess.CalledProcessError:
            print(f"⚠️ Failed to install {fname}, skipping.")

    # Runtime info
    end_time = time.time()
    duration = end_time - start_time
    print(f"\nRUNTIME: {round(duration, 2)} seconds")
    print("\n✅ Environment ready! Activate with:")
    print(f"   source .venv/Scripts/activate   (Git Bash)")

if __name__ == '__main__':
    main()
