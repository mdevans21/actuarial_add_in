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
#   * For the net48 build (RUNTIME=net48, default): nothing — .NET Framework 4.8
#     is preinstalled on Windows 10 1903+ and Windows 11.
#     For the net8 perf variant (RUNTIME=net8): .NET 8 Desktop Runtime
#     (Microsoft.WindowsDesktop.App 8.x).
#   * gh CLI on the WSL side, authenticated for the repo.
#
# Usage:
#   scripts/run_dump.sh                          # latest release, net48 default
#   scripts/run_dump.sh v0.8.0                   # specific tag
#   RUNTIME=net8 scripts/run_dump.sh v0.8.0      # exercise the net8 variant
#   KEEP=1 scripts/run_dump.sh v0.8.0            # leave scratch dir for inspection
#   SAVE_WORKBOOK=1 scripts/run_dump.sh          # write the post-recalc workbook back to
#                                                # excel/actuarial_add_in.xlsx so cached
#                                                # values land in the repo copy.
#   LOCAL_XLL=1 scripts/run_dump.sh              # skip the gh download; use the locally-
#                                                # built XLL under src/ActuarialAddIn/bin/
#                                                # Release/<tfm>/publish/ for the selected
#                                                # RUNTIME. Pairs naturally with
#                                                # LOCAL_WORKBOOK=1 to test pre-release.

set -euo pipefail

TAG="${1:-}"
RUNTIME="${RUNTIME:-net48}"
REPO="mdevans21/actuarial_add_in"
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
SCRATCH_WIN="C:\\Users\\${USER}\\actuarial_dump_test"
SCRATCH_WSL="/mnt/c/Users/${USER}/actuarial_dump_test"

case "$RUNTIME" in
    net48) XLL_NAME="ActuarialAddIn-AddIn64-packed.xll" ;;
    net8)  XLL_NAME="ActuarialAddIn-AddIn64-packed-net8.xll" ;;
    *) echo "ERROR: RUNTIME must be 'net48' or 'net8' (got '$RUNTIME')." >&2; exit 2 ;;
esac

if [[ -z "${LOCAL_XLL:-}" && -z "$TAG" ]]; then
    TAG=$(gh release view --repo "$REPO" --json tagName -q .tagName)
    echo "Using latest release: $TAG"
fi
echo ">> Runtime: $RUNTIME ($XLL_NAME)"

if ! command -v powershell.exe >/dev/null; then
    echo "ERROR: powershell.exe not on PATH. Are you running outside WSL2?" >&2
    exit 2
fi

rm -rf "$SCRATCH_WSL"
mkdir -p "$SCRATCH_WSL"

case "$RUNTIME" in
    net48) LOCAL_TFM_DIR="net48" ;;
    net8)  LOCAL_TFM_DIR="net8.0-windows" ;;
esac

if [[ -n "${LOCAL_XLL:-}" ]]; then
    LOCAL_XLL_PATH="$ROOT/src/ActuarialAddIn/bin/Release/$LOCAL_TFM_DIR/publish/ActuarialAddIn-AddIn64-packed.xll"
    if [[ ! -f "$LOCAL_XLL_PATH" ]]; then
        echo "ERROR: locally-built XLL not found at $LOCAL_XLL_PATH" >&2
        echo "       Run 'dotnet build src/ActuarialAddIn/ActuarialAddIn.csproj -c Release' first." >&2
        exit 2
    fi
    echo ">> LOCAL_XLL=1 — copying locally-built $RUNTIME XLL into scratch as $XLL_NAME"
    cp "$LOCAL_XLL_PATH" "$SCRATCH_WSL/$XLL_NAME"
else
    echo ">> Downloading $TAG assets via gh..."
    gh release download "$TAG" --repo "$REPO" --pattern '*' --dir "$SCRATCH_WSL"
fi

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
    -Xll      "${SCRATCH_WIN}\\${XLL_NAME}" \
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
