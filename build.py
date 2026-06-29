#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build script for LebaoLabelPrinter EXE
Handles Python version detection, dependency install, and PyInstaller build
with explicit VC runtime DLL bundling for Python 3.13+
"""

import subprocess
import sys
import os
import shutil
import glob


def find_compatible_python():
    """Find a Python 3.10-3.13 for building. Returns (executable_path, version_string)."""
    # Try py launcher versions in priority order
    for ver in ['3.13', '3.12', '3.11', '3.10']:
        try:
            result = subprocess.run(
                ['py', f'-{ver}', '-c', 'import sys; print(sys.executable)'],
                capture_output=True, text=True, timeout=10
            )
            if result.returncode == 0 and result.stdout.strip():
                exe = result.stdout.strip()
                # Verify it actually works
                test = subprocess.run(
                    [exe, '-c', 'print("ok")'],
                    capture_output=True, text=True, timeout=10
                )
                if test.returncode == 0:
                    ver_result = subprocess.run(
                        [exe, '--version'],
                        capture_output=True, text=True, timeout=10
                    )
                    return exe, ver_result.stdout.strip()
        except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
            continue

    # Try current interpreter
    v = sys.version_info
    if v.major == 3 and v.minor <= 13:
        return sys.executable, f"Python {v.major}.{v.minor}.{v.micro}"

    # Try generic python
    try:
        result = subprocess.run(
            ['python', '-c', 'import sys; v=sys.version_info; print(sys.executable if v.minor<=13 else "")'],
            capture_output=True, text=True, timeout=10
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip(), "Python (generic)"
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        pass

    return None, None


def find_vc_runtime_dlls(py_base):
    """Find vcruntime DLLs in Python installation directory."""
    dlls = []
    # Check main directory
    for name in ['vcruntime140.dll', 'vcruntime140_1.dll', 'vcruntime140_2.dll']:
        path = os.path.join(py_base, name)
        if os.path.exists(path):
            dlls.append(path)
    # Check DLLs subdirectory
    dlls_dir = os.path.join(py_base, 'DLLs')
    if os.path.exists(dlls_dir):
        for name in ['vcruntime140.dll', 'vcruntime140_1.dll', 'vcruntime140_2.dll']:
            path = os.path.join(dlls_dir, name)
            if os.path.exists(path) and path not in dlls:
                dlls.append(path)
    # Also check .libs directory (newer Python versions)
    libs_dir = os.path.join(py_base, '.libs')
    if os.path.exists(libs_dir):
        for f in os.listdir(libs_dir):
            if f.startswith('vcruntime') and f.endswith('.dll'):
                path = os.path.join(libs_dir, f)
                if path not in dlls:
                    dlls.append(path)
    return dlls


def main():
    print("=" * 44)
    print("  Lebao QA - Label Printer Tool v2.0")
    print("  Build EXE Script")
    print("=" * 44)
    print()

    # 1. Find compatible Python
    py_exe, py_ver = find_compatible_python()
    if not py_exe:
        print("[ERROR] No compatible Python found!")
        print("  Need Python 3.10, 3.11, 3.12, or 3.13")
        print("  Python 3.14+ is NOT supported by PyInstaller")
        print("  Install from: https://www.python.org/downloads/")
        input("Press Enter to exit...")
        sys.exit(1)

    # Check version
    v_info = subprocess.run(
        [py_exe, '-c', 'import sys; v=sys.version_info; print(f"{v.major}.{v.minor}")'],
        capture_output=True, text=True
    )
    major_minor = v_info.stdout.strip()
    minor = int(major_minor.split('.')[1])

    if minor >= 14:
        print(f"[WARNING] Python {major_minor} is NOT supported by PyInstaller!")
        print("  The EXE will fail on other machines (python3xx.dll error)")
        print("  Please install Python 3.12 or 3.13 from python.org")
        print()
        confirm = input("Continue anyway? (y/N): ").strip().lower()
        if confirm != 'y':
            sys.exit(1)

    print(f"[OK] Using: {py_ver}")
    print(f"      Path: {py_exe}")
    print()

    # 2. Get Python base prefix
    base_result = subprocess.run(
        [py_exe, '-c', 'import sys; print(sys.base_exec_prefix)'],
        capture_output=True, text=True
    )
    py_base = base_result.stdout.strip()
    print(f"      Base: {py_base}")
    print()

    # 3. Find VC runtime DLLs
    vc_dlls = find_vc_runtime_dlls(py_base)
    if vc_dlls:
        print(f"[OK] Found VC runtime DLLs:")
        for dll in vc_dlls:
            print(f"      {dll}")
    else:
        print("[WARNING] No vcruntime DLLs found - EXE may fail on other machines!")
        print("  Install Microsoft Visual C++ Redistributable on the build machine:")
        print("  https://aka.ms/vs/17/release/vc_redist.x64.exe")
    print()

    # 4. Install dependencies
    print("[1/4] Installing dependencies...")
    subprocess.run(
        [py_exe, '-m', 'pip', 'install', 'pandas', 'openpyxl',
         '--quiet', '--no-warn-script-location'],
        check=True
    )

    print("[2/4] Installing PyInstaller...")
    subprocess.run(
        [py_exe, '-m', 'pip', 'install', 'pyinstaller',
         '--quiet', '--no-warn-script-location'],
        check=True
    )

    print("[OK] Dependencies installed")
    print()

    # 5. Clean old builds
    print("[3/4] Cleaning old builds...")
    for d in ['build', 'dist']:
        if os.path.exists(d):
            shutil.rmtree(d)
            print(f"      Removed {d}/")
    spec = 'LebaoLabelPrinter.spec'
    if os.path.exists(spec):
        os.remove(spec)
        print(f"      Removed {spec}")
    print("[OK] Cleaned")
    print()

    # 6. Ensure output dir exists
    if not os.path.exists('output'):
        os.makedirs('output')

    # 7. Build with PyInstaller
    print("[4/4] Building EXE (please wait 2-5 minutes)...")
    print("      Using --onedir mode (2-3x faster startup)")
    print()

    cmd = [
        py_exe, '-m', 'PyInstaller',
        '--onedir',
        '--windowed',
        '--name=LebaoLabelPrinter',
        '--add-data=config.ini;.',
        '--add-data=README.txt;.',
        '--add-data=output;output',
        '--hidden-import=pandas',
        '--hidden-import=openpyxl',
        '--hidden-import=pandas._libs.tslibs',
        '--clean',
    ]

    # Add VC runtime DLLs explicitly (critical for Python 3.13+)
    for dll in vc_dlls:
        cmd.append(f'--add-binary={dll};.')

    cmd.append('label_printer.py')

    result = subprocess.run(cmd)
    print()

    # 8. Verify build result
    exe_path = os.path.join('dist', 'LebaoLabelPrinter', 'LebaoLabelPrinter.exe')
    if not os.path.exists(exe_path):
        print("[ERROR] Build failed - EXE not found!")
        print("  Common causes:")
        print("  1. Not enough disk space (need 500MB+)")
        print("  2. Antivirus blocking (disable temporarily)")
        print("  3. Python version incompatible (use 3.10-3.13)")
        input("Press Enter to exit...")
        sys.exit(1)

    # 9. Also copy any remaining needed DLLs from Python dir to _internal
    internal_dir = os.path.join('dist', 'LebaoLabelPrinter', '_internal')
    if os.path.exists(internal_dir):
        # Copy vcruntime DLLs directly to _internal as well (belt and suspenders)
        for dll in vc_dlls:
            dst = os.path.join(internal_dir, os.path.basename(dll))
            if not os.path.exists(dst):
                shutil.copy2(dll, dst)
                print(f"      Copied {os.path.basename(dll)} to _internal/")

        # Copy python3.dll (sometimes needed as a shim)
        python3_dll = os.path.join(py_base, 'python3.dll')
        if os.path.exists(python3_dll):
            dst = os.path.join(internal_dir, 'python3.dll')
            if not os.path.exists(dst):
                shutil.copy2(python3_dll, dst)
                print(f"      Copied python3.dll to _internal/")

        # For Python 3.13+, also check for libcrypto and libssl
        dlls_dir = os.path.join(py_base, 'DLLs')
        if os.path.exists(dlls_dir):
            for pattern in ['libcrypto*.dll', 'libssl*.dll']:
                for f in glob.glob(os.path.join(dlls_dir, pattern)):
                    dst = os.path.join(internal_dir, os.path.basename(f))
                    if not os.path.exists(dst):
                        shutil.copy2(f, dst)
                        print(f"      Copied {os.path.basename(f)} to _internal/")

    print()
    print("=" * 44)
    print("  Build SUCCESS!")
    print("=" * 44)
    print()
    print(f"  Output folder: dist\\LebaoLabelPrinter\\")
    print(f"  Main EXE:      dist\\LebaoLabelPrinter\\LebaoLabelPrinter.exe")
    print()
    print("  IMPORTANT: Edit dist\\LebaoLabelPrinter\\config.ini before use!")
    print()
    print("  Copy the ENTIRE dist\\LebaoLabelPrinter\\ folder to any PC.")
    print("  For distribution, zip the folder.")
    print()
    print("  Note: --onedir mode starts much faster than --onefile.")
    print("=" * 44)
    print()

    # Auto-open dist folder on Windows
    if sys.platform == 'win32':
        dist_dir = os.path.join(os.getcwd(), 'dist', 'LebaoLabelPrinter')
        os.startfile(dist_dir)

    input("Press Enter to exit...")


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print(f"\n[ERROR] Build script crashed: {e}")
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")
        sys.exit(1)
