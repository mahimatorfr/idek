#!/usr/bin/env python3
"""
Environment checker and launcher for dedupe_regulator_cols.py
Verifies Python version, checks script exists, and runs it.
Zero external dependencies - uses only stdlib.
"""

import sys
import subprocess
import os
from pathlib import Path


MIN_PYTHON = (3, 7)  # pathlib, f-strings, argparse all available


def check_python():
    """Verify Python version meets minimum."""
    if sys.version_info < MIN_PYTHON:
        print(f"❌ Python {MIN_PYTHON[0]}.{MIN_PYTHON[1]}+ required. You have {sys.version.split()[0]}")
        return False
    print(f"✓ Python {sys.version.split()[0]} OK")
    return True


def check_script(script_name):
    """Verify the main script exists."""
    script_path = Path(__file__).parent / script_name
    if not script_path.exists():
        print(f"❌ Main script not found: {script_path}")
        return False
    print(f"✓ Found {script_name}")
    return script_path


def run_main(script_path, args):
    """Execute the main script with given arguments."""
    try:
        # Pass through all arguments to the main script
        result = subprocess.run([sys.executable, str(script_path)] + args, check=False)
        return result.returncode == 0
    except KeyboardInterrupt:
        print("\n\nCancelled.")
        return False
    except Exception as e:
        print(f"❌ Error running script: {e}")
        return False


def main():
    print("=" * 50)
    print("  CSV Duplicate Column Remover - Launcher")
    print("=" * 50)
    print()

    # Check Python
    if not check_python():
        sys.exit(1)

    # Find main script
    main_script = check_script("dedupe_regulator_cols.py")
    if not main_script:
        sys.exit(1)

    print()
    print("All checks passed. Launching...")
    print()

    # Forward all command-line args to the main script
    success = run_main(main_script, sys.argv[1:])

    if success:
        print("\n✓ Completed successfully")
    else:
        print("\n❌ Completed with errors")
        sys.exit(1)


if __name__ == "__main__":
    main()