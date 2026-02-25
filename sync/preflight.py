"""
Local pre-deployment checks for the Azure Function (no Azure CLI / Core Tools required).

Usage:
  python preflight.py
  python preflight.py --check-env
"""

from __future__ import annotations

import argparse
import importlib
import os
import py_compile
import sys
from pathlib import Path
from typing import Iterable, List

ROOT = Path(__file__).resolve().parent
REQUIRED_FILES = ["host.json", "function_app.py", "requirements.txt"]
REQUIRED_ENV_VARS = [
    "SHAREPOINT_SITE_URL",
    "AZURE_STORAGE_ACCOUNT_NAME",
]


def _check_python_version() -> List[str]:
    errors: List[str] = []
    if sys.version_info < (3, 9):
        errors.append(
            f"Python {sys.version_info.major}.{sys.version_info.minor} not supported. "
            "Use Python 3.9+ (3.11/3.12 recommended for Azure Functions)."
        )
    return errors


def _check_required_files() -> List[str]:
    errors: List[str] = []
    for filename in REQUIRED_FILES:
        if not (ROOT / filename).exists():
            errors.append(f"Missing required file: {filename}")
    return errors


def _iter_python_files() -> Iterable[Path]:
    for path in ROOT.glob("*.py"):
        if path.name == "preflight.py":
            continue
        yield path


def _compile_all() -> List[str]:
    errors: List[str] = []
    for path in _iter_python_files():
        try:
            py_compile.compile(str(path), doraise=True)
        except Exception as exc:
            errors.append(f"Compile failed for {path.name}: {exc}")
    return errors


def _check_function_discovery() -> List[str]:
    errors: List[str] = []
    cwd = Path.cwd()
    try:
        os.chdir(ROOT)
        module = importlib.import_module("function_app")
        app = getattr(module, "app", None)
        if app is None:
            return ["function_app.py loaded but variable 'app' is missing"]

        get_functions = getattr(app, "get_functions", None)
        if callable(get_functions):
            functions = get_functions()
            if len(functions) == 0:
                errors.append("No Azure Function discovered from function_app.app")
            else:
                names = [getattr(f, "get_function_name", lambda: "<unknown>")() for f in functions]
                print(f"[OK] Discovered functions: {', '.join(names)}")
        else:
            print("[WARN] Could not introspect function list (get_functions not available), import succeeded.")
    except ModuleNotFoundError as exc:
        errors.append(
            "Dependency missing for function discovery "
            f"({exc}). Run: pip install -r requirements.txt"
        )
    except Exception as exc:
        errors.append(f"Import/discovery failed for function_app.py: {exc}")
    finally:
        os.chdir(cwd)
    return errors


def _check_env() -> List[str]:
    missing = [name for name in REQUIRED_ENV_VARS if not os.environ.get(name)]
    if missing:
        return [f"Missing required env vars for runtime config check: {', '.join(missing)}"]
    return []


def main() -> int:
    parser = argparse.ArgumentParser(description="Run local Azure Function preflight checks")
    parser.add_argument(
        "--check-env",
        action="store_true",
        help="also validate a minimal set of required environment variables",
    )
    args = parser.parse_args()

    all_errors: List[str] = []

    print("[1/4] Checking Python version...")
    all_errors.extend(_check_python_version())

    print("[2/4] Checking required files...")
    all_errors.extend(_check_required_files())

    print("[3/4] Compiling Python files...")
    all_errors.extend(_compile_all())

    print("[4/4] Importing and discovering Azure Functions...")
    all_errors.extend(_check_function_discovery())

    if args.check_env:
        print("[extra] Checking required environment variables...")
        all_errors.extend(_check_env())

    if all_errors:
        print("\nPreflight FAILED:")
        for err in all_errors:
            print(f"- {err}")
        return 1

    print("\nPreflight PASSED: function package is ready for deployment.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
