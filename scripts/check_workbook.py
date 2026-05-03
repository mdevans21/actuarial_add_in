r"""check_workbook.py — Recalc the workbook in Excel and dump every ACT_ cell.

ZERO-CONFIG WORKFLOW
====================
Drop this file anywhere in (or alongside) the repo, install xlwings once,
and run it. No arguments.

    pip install xlwings
    python check_workbook.py

The script looks for the workbook and the .xll in the obvious places:

  * its own folder (flat layout — workbook + xll + script all together)
  * `..\excel\actuarial_add_in.xlsx`
        and `..\src\ActuarialAddIn\bin\Release\net6.0-windows\publish\ActuarialAddIn-AddIn64-packed.xll`
        (when the script is in the repo's `scripts\` directory)
  * `.\excel\` and `.\src\...\publish\`
        (when run from the repo root)

It produces `actuarial_add_in_dump.json` next to itself — one record per
`ACT_*` formula in the workbook, with sheet, cell, formula, and the
value Excel actually computed under the loaded add-in. Send the dump
back for review.

If you've never opened the .xll before, Windows blocks it as untrusted.
The script will tell you when this has happened. The fix is one line:

    Unblock-File <path_to_packed_xll>
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

try:
    import xlwings as xw
except ImportError:
    print("ERROR: xlwings is not installed. Run:  pip install xlwings")
    sys.exit(1)


HERE = Path(__file__).resolve().parent
ACT_RE = re.compile(r"^=\s*([A-Z][A-Z0-9_]*)\s*\(")

# Hunt order: same folder; repo-from-scripts (../excel, ../src/.../publish);
# repo-from-root (./excel, ./src/.../publish).
WORKBOOK_CANDIDATES = [
    HERE / "actuarial_add_in.xlsx",
    HERE.parent / "excel" / "actuarial_add_in.xlsx",
    HERE / "excel" / "actuarial_add_in.xlsx",
]
XLL_CANDIDATES = [
    HERE / "ActuarialAddIn-AddIn64-packed.xll",
    HERE.parent / "src" / "ActuarialAddIn" / "bin" / "Release" / "net6.0-windows"
        / "publish" / "ActuarialAddIn-AddIn64-packed.xll",
    HERE.parent / "src" / "ActuarialAddIn" / "bin" / "Debug" / "net6.0-windows"
        / "publish" / "ActuarialAddIn-AddIn64-packed.xll",
    HERE / "src" / "ActuarialAddIn" / "bin" / "Release" / "net6.0-windows"
        / "publish" / "ActuarialAddIn-AddIn64-packed.xll",
]
DUMP = HERE / "actuarial_add_in_dump.json"


def find_first(paths: list[Path], description: str) -> Path | None:
    for p in paths:
        if p.exists():
            return p
    print(f"ERROR: could not find {description}. Looked in:")
    for p in paths:
        print(f"  {p}")
    return None


def to_2d(v, n_rows: int, n_cols: int):
    """Normalise xlwings range.value/.formula into a 2-D list."""
    if n_rows == 1 and n_cols == 1:
        return [[v]]
    if n_rows == 1:
        return [list(v)]
    if n_cols == 1:
        return [[x] for x in v]
    return [list(r) for r in v]


def main() -> int:
    workbook = find_first(WORKBOOK_CANDIDATES, "workbook (actuarial_add_in.xlsx)")
    if workbook is None:
        return 1
    xll = find_first(XLL_CANDIDATES, "packed XLL (ActuarialAddIn-AddIn64-packed.xll)")
    if xll is None:
        print()
        print("If you haven't built the add-in yet, run this first from the repo root:")
        print("  dotnet build src\\ActuarialAddIn\\ActuarialAddIn.csproj -c Release")
        return 1

    print(f"Workbook: {workbook}")
    print(f"XLL:      {xll}")
    print()
    print("Opening Excel (hidden)...")
    calls: list[dict] = []
    n_errors = 0

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False

        print(f"Registering XLL...")
        ok = app.api.RegisterXLL(str(xll))
        if not ok:
            print()
            print("ERROR: Excel refused to load the .xll. Two usual causes:")
            print(f"  1. Mark of the Web — run in PowerShell:")
            print(f"        Unblock-File '{xll}'")
            print(f"  2. .NET 6 Desktop Runtime not installed.")
            print(f"     Check: dotnet --list-runtimes")
            print(f"     Install: winget install Microsoft.DotNet.DesktopRuntime.6")
            return 2

        wb = app.books.open(str(workbook))
        print("Recalculating (this may take a few seconds under the live add-in)...")
        app.api.CalculateFullRebuild()

        for sht in wb.sheets:
            used = sht.used_range
            n_rows = used.last_cell.row - used.row + 1
            n_cols = used.last_cell.column - used.column + 1
            if n_rows == 0 or n_cols == 0:
                continue
            try:
                formulas = to_2d(used.formula, n_rows, n_cols)
                values = to_2d(used.value, n_rows, n_cols)
            except Exception as e:  # noqa: BLE001
                print(f"  ! {sht.name}: {e}; skipping")
                continue

            r0, c0 = used.row, used.column
            n_act = 0
            for i, row in enumerate(formulas):
                for j, formula in enumerate(row):
                    if not isinstance(formula, str):
                        continue
                    m = ACT_RE.match(formula)
                    if not m:
                        continue
                    v = values[i][j]
                    if hasattr(v, "isoformat"):
                        v = v.isoformat()
                    if isinstance(v, str) and v.startswith("#"):
                        n_errors += 1
                    cell_addr = xw.utils.col_name(c0 + j) + str(r0 + i)
                    calls.append({
                        "sheet": sht.name,
                        "cell": cell_addr,
                        "function": m.group(1),
                        "formula": formula,
                        "value": v,
                    })
                    n_act += 1
            if n_act:
                print(f"  {sht.name}: {n_act} ACT_ formulas")

        wb.close()

    DUMP.write_text(json.dumps(calls, indent=2, default=str))
    print()
    print(f"Dumped {len(calls)} ACT_ formulas to:")
    print(f"  {DUMP}")

    # Distinguish a few failure modes from the value distribution.
    n_none = sum(1 for c in calls if c["value"] is None)
    none_pct = (n_none / len(calls) * 100) if calls else 0.0
    if n_errors:
        print()
        print(f"!! {n_errors} cell(s) returned an Excel error (#NAME? / #VALUE! / etc).")
        print("   Usually means the add-in didn't fully load.")
        return 3
    if none_pct >= 90:
        print()
        print(f"!! {n_none} of {len(calls)} cells ({none_pct:.0f}%) came back as None.")
        print("   Excel computed nothing — the add-in did not load even though")
        print("   RegisterXLL succeeded. Most common cause: the .NET 6 Windows")
        print("   Desktop Runtime is missing.")
        print()
        print("   Install:")
        print("     winget install Microsoft.DotNet.DesktopRuntime.6")
        print("   Verify:")
        print("     dotnet --list-runtimes  (look for Microsoft.WindowsDesktop.App 6.x)")
        print("   Then close all Excel processes and re-run:")
        print("     Get-Process excel | Stop-Process -Force")
        return 4

    print()
    print("All cells calculated cleanly.  Send the dump JSON back for review.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
