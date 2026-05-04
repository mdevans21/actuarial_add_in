#!/usr/bin/env bash
# Drive the Windows-side dump test from WSL.
#
# Downloads the .xll + .xlsx for a given release tag (default: latest), copies
# them and dump_workbook.ps1 to a Windows-visible scratch dir, invokes Excel
# via PowerShell COM, and copies the resulting JSON back to
# tests/actuarial_add_in_dump.json.
#
# Requirements (one-time):
#   * WSL2 with /mnt/c interop enabled (default).
#   * Excel installed on the Windows host.
#   * .NET Desktop Runtime matching the XLL's TargetFramework (net6.0-windows
#     today -> Microsoft.WindowsDesktop.App 6.x).
#   * gh CLI on the WSL side, authenticated for the repo.
#
# Usage:
#   scripts/run_dump.sh                 # latest release
#   scripts/run_dump.sh v0.7.0          # specific tag
#   KEEP=1 scripts/run_dump.sh v0.7.0   # leave scratch dir for inspection
#   SAVE_WORKBOOK=1 scripts/run_dump.sh # write the post-recalc workbook back to
#                                       # excel/actuarial_add_in.xlsx so cached
#                                       # values land in the repo copy.

set -euo pipefail

TAG="${1:-}"
REPO="mdevans21/actuarial_add_in"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
SCRATCH_WIN="C:\\Users\\${USER}\\actuarial_dump_test"
SCRATCH_WSL="/mnt/c/Users/${USER}/actuarial_dump_test"

if [[ -z "$TAG" ]]; then
    TAG=$(gh release view --repo "$REPO" --json tagName -q .tagName)
    echo "Using latest release: $TAG"
fi

if ! command -v powershell.exe >/dev/null; then
    echo "ERROR: powershell.exe not on PATH. Are you running outside WSL2?" >&2
    exit 2
fi

rm -rf "$SCRATCH_WSL"
mkdir -p "$SCRATCH_WSL"

echo ">> Downloading $TAG assets via gh..."
gh release download "$TAG" --repo "$REPO" --pattern '*' --dir "$SCRATCH_WSL"

# LOCAL_WORKBOOK=1 uses the in-repo excel/actuarial_add_in.xlsx instead of the
# release asset. Useful when iterating on workbook content (formula tweaks,
# new sections, chart layout) without shipping a fresh tag for each change —
# the released XLL still provides the C# implementation.
if [[ -n "${LOCAL_WORKBOOK:-}" ]]; then
    echo ">> LOCAL_WORKBOOK=1 — using in-repo workbook instead of released asset"
    cp "$ROOT/excel/actuarial_add_in.xlsx" "$SCRATCH_WSL/actuarial_add_in.xlsx"
fi

cp "$ROOT/scripts/dump_workbook.ps1" "$SCRATCH_WSL/dump_workbook.ps1"

echo ">> Invoking Excel via PowerShell COM..."
SAVE_FLAG=()
if [[ -n "${SAVE_WORKBOOK:-}" ]]; then SAVE_FLAG=(-SaveWorkbook); fi
powershell.exe -NoProfile -ExecutionPolicy Bypass -File \
    "${SCRATCH_WIN}\\dump_workbook.ps1" \
    -Xll      "${SCRATCH_WIN}\\ActuarialAddIn-AddIn64-packed.xll" \
    -Workbook "${SCRATCH_WIN}\\actuarial_add_in.xlsx" \
    -Output   "${SCRATCH_WIN}\\dump.json" \
    "${SAVE_FLAG[@]}"

DEST="$ROOT/tests/actuarial_add_in_dump.json"
cp "$SCRATCH_WSL/dump.json" "$DEST"
cp "$SCRATCH_WSL/dump.cells.json" "$ROOT/tests/actuarial_add_in_cells.json"
echo ">> Wrote $DEST"
echo ">> Wrote $ROOT/tests/actuarial_add_in_cells.json"

if [[ -n "${SAVE_WORKBOOK:-}" ]]; then
    cp "$SCRATCH_WSL/actuarial_add_in.xlsx" "$ROOT/excel/actuarial_add_in.xlsx"
    echo ">> Wrote $ROOT/excel/actuarial_add_in.xlsx (with Excel-cached calc values)"
fi

if [[ -z "${KEEP:-}" ]]; then
    # Excel may still hold a Windows file handle on the .xll for a few seconds
    # after Quit(); rm can race that. The scratch dir is recreated next run, so
    # ignore failures rather than fail the whole pipeline.
    rm -rf "$SCRATCH_WSL" 2>/dev/null || true
fi
