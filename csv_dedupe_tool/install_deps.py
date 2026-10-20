#!/usr/bin/env python3
"""
Install dependencies for the CSV/Excel duplicate column remover.
Safe to run multiple times. Works on Windows, macOS, Linux.
"""

import subprocess
import sys
import os
from pathlib import Path


MIN_PYTHON = (3, 7)
REQUIRED = ["openpyxl>=3.0.0"]  # only needed for Excel (.xlsx) files


def check_python():
    if sys.version_info < MIN_PYTHON:
        print(f"✗ Python {MIN_PYTHON[0]}.{MIN_PYTHON[1]}+ required. You have {sys.version.split()[0]}")
        print(f"  Download from https://www.python.org/downloads/")
        return False
    print(f"✓ Python {sys.version.split()[0]}")
    return True


def ensure_pip():
    try:
        subprocess.run([sys.executable, "-m", "pip", "--version"],
                       check=True, capture_output=True)
        print("✓ pip available")
        return True
    except (subprocess.CalledProcessError, FileNotFoundError):
        print("! pip not found, trying ensurepip...")
        try:
            subprocess.run([sys.executable, "-m", "ensurepip", "--upgrade"],
                           check=True, capture_output=True)
            print("✓ pip installed via ensurepip")
            return True
        except Exception as e:
            print(f"✗ Could not install pip: {e}")
            print("  Install Python with 'pip' from https://www.python.org/downloads/")
            return False


def install_packages():
    print(f"Installing: {', '.join(REQUIRED)}")
    result = subprocess.run(
        [sys.executable, "-m", "pip", "install", "--upgrade", *REQUIRED],
        capture_output=False
    )
    if result.returncode == 0:
        print("✓ Packages installed")
        return True
    print("✗ Installation failed")
    return False


def verify():
    try:
        import openpyxl
        print(f"✓ openpyxl {openpyxl.__version__} verified")
        return True
    except ImportError:
        print("✗ openpyxl not importable after install")
        return False


def main():
    print("=" * 52)
    print("  Dependency Installer")
    print("=" * 52)
    print()

    if not check_python():
        sys.exit(1)
    if not ensure_pip():
        sys.exit(1)
    print()
    if not install_packages():
        sys.exit(1)
    print()
    if not verify():
        sys.exit(1)

    print()
    print("=" * 52)
    print("  All dependencies ready.")
    print("  Run:  python dedupe_regulator_cols.py")
    print("=" * 52)


if __name__ == "__main__":
    main()
