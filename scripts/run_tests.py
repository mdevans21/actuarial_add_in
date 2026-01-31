#!/usr/bin/env python3
"""
Unified test runner for Actuarial Add-In.

Runs all tests from WSL:
  - C# unit tests (via Windows dotnet)
  - Python validation tests (pytest)
  - Source comparison report

Usage:
    python scripts/run_tests.py          # Run all tests
    python scripts/run_tests.py --csharp # C# tests only
    python scripts/run_tests.py --python # Python tests only
    python scripts/run_tests.py --report # Generate validation report
"""

import argparse
import subprocess
import sys
import os
from pathlib import Path

# Paths
REPO_ROOT = Path(__file__).parent.parent
CSHARP_TEST_DIR = REPO_ROOT / "src" / "ActuarialAddIn.Tests"
PYTHON_TEST_DIR = REPO_ROOT / "tests"
RESULTS_FILE = REPO_ROOT / "test_results.md"


def run_csharp_tests() -> bool:
    """Run C# tests using Windows dotnet from WSL."""
    print("\n" + "=" * 60)
    print("Running C# Tests (generates test_results.md)")
    print("=" * 60 + "\n")

    # Use the Windows path via /mnt/c
    win_test_dir = Path("/mnt/c/Users/matth/Code/actuarial_add_in/src/ActuarialAddIn.Tests")
    win_results = Path("/mnt/c/Users/matth/Code/actuarial_add_in/test_results.md")

    # Run dotnet via cmd.exe from the Windows directory
    result = subprocess.run(
        ["cmd.exe", "/c", "dotnet run"],
        cwd=str(win_test_dir)
    )

    if result.returncode == 0:
        # Copy the generated test_results.md to WSL repo (if different paths)
        if win_results.exists() and win_results.resolve() != RESULTS_FILE.resolve():
            import shutil
            shutil.copy(win_results, RESULTS_FILE)
            print(f"\nResults copied to: {RESULTS_FILE}")
        else:
            print(f"\nResults written to: {RESULTS_FILE}")
        return True
    else:
        print(f"\nC# tests failed with exit code {result.returncode}")
        return False


def run_python_tests() -> bool:
    """Run Python pytest validation tests."""
    print("\n" + "=" * 60)
    print("Running Python Validation Tests (pytest)")
    print("=" * 60 + "\n")

    result = subprocess.run(
        [sys.executable, "-m", "pytest", str(PYTHON_TEST_DIR), "-v"],
        cwd=str(REPO_ROOT)
    )

    return result.returncode == 0


def generate_validation_report() -> bool:
    """Generate the validation comparison report."""
    print("\n" + "=" * 60)
    print("Generating Validation Report")
    print("=" * 60 + "\n")

    compare_script = PYTHON_TEST_DIR / "compare_sources.py"
    if not compare_script.exists():
        print(f"Error: {compare_script} not found")
        return False

    result = subprocess.run(
        [sys.executable, str(compare_script), "--generate-report"],
        cwd=str(REPO_ROOT)
    )

    if result.returncode == 0:
        report_path = PYTHON_TEST_DIR / "reports" / "VALIDATION_REPORT.md"
        print(f"\nReport generated: {report_path}")
        return True
    return False


def main():
    parser = argparse.ArgumentParser(
        description="Run Actuarial Add-In tests",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Test Types:
  C# Tests     Runs the .NET test suite which exercises all Excel UDF
               functions and generates test_results.md with detailed output.

  Python Tests Runs pytest against JSON fixtures to validate statistical
               functions against scipy and reference values from literature.

  Report       Generates a detailed validation report comparing C# output
               against Python/scipy and published reference values.
"""
    )
    parser.add_argument("--csharp", action="store_true", help="Run C# tests only")
    parser.add_argument("--python", action="store_true", help="Run Python tests only")
    parser.add_argument("--report", action="store_true", help="Generate validation report only")
    parser.add_argument("--all", action="store_true", help="Run all tests (default)")

    args = parser.parse_args()

    # Default to all if no specific test selected
    run_all = args.all or not (args.csharp or args.python or args.report)

    results = []

    if run_all or args.csharp:
        results.append(("C# Tests", run_csharp_tests()))

    if run_all or args.python:
        results.append(("Python Tests", run_python_tests()))

    if run_all or args.report:
        results.append(("Validation Report", generate_validation_report()))

    # Summary
    print("\n" + "=" * 60)
    print("Test Summary")
    print("=" * 60)

    all_passed = True
    for name, passed in results:
        status = "PASSED" if passed else "FAILED"
        print(f"  {name}: {status}")
        if not passed:
            all_passed = False

    print()
    return 0 if all_passed else 1


if __name__ == "__main__":
    sys.exit(main())
