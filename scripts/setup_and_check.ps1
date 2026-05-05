<#
.SYNOPSIS
    Drop-and-run Windows runner: download the released add-in, recalc the
    workbook in Excel, dump every ACT_ cell to JSON.

.DESCRIPTION
    Place this .ps1 next to check_workbook.py in any Windows folder, then
    run it. The script:

      1. Sanity-checks prerequisites (gh, python; .NET Framework 4.8 is
         preinstalled on Windows 10/11 — only the optional net8 perf
         variant needs the .NET 8 Desktop Runtime).
      2. Resolves the requested release tag (default: latest).
      3. Downloads the packed .xll and the worked-example workbook from
         the GitHub Release into the script's folder.
      4. Unblocks the .xll so Excel will load it.
      5. Installs xlwings via pip if missing.
      6. Runs check_workbook.py — opens hidden Excel, registers the
         add-in, recalculates the workbook, dumps every ACT_ formula
         and value to actuarial_add_in_dump.json next to itself.

    No build step, no git clone, no temp directory — everything lands
    next to this script.

.PARAMETER Tag
    Release tag to fetch (default: 'latest', resolves to the most recent
    GitHub release). Pass an explicit tag like 'v0.5.0' to pin.

.PARAMETER Repo
    GitHub repo in 'owner/name' form. Default: mdevans21/actuarial_add_in.

.EXAMPLE
    .\setup_and_check.ps1
    # Latest release.

.EXAMPLE
    .\setup_and_check.ps1 -Tag v0.5.0
    # Pinned tag.

.NOTES
    Authentication: this script uses the GitHub CLI (gh) so that private
    repos work without exposing tokens. Run `gh auth login` once if
    you've never done so.
#>

param(
    [string]$Tag = "latest",
    [string]$Repo = "mdevans21/actuarial_add_in"
)

$ErrorActionPreference = "Stop"

$Here = $PSScriptRoot
if (-not $Here) { $Here = (Get-Location).Path }
Set-Location $Here

Write-Host "Working folder: $Here"
Write-Host "Repo:           $Repo"
Write-Host "Tag:            $Tag"
Write-Host ""

# --- 1. Prerequisite checks ---------------------------------------------
Write-Host "Checking prerequisites..."
foreach ($cmd in @("gh", "python", "pip")) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
        Write-Error "$cmd is not on PATH. Install it and re-run."
        exit 1
    }
}

# The default .xll targets net48 (.NET Framework 4.8 — preinstalled on Win 10/11).
# Only warn if the user is opting into the net8 perf variant by setting
# $env:RUNTIME = "net8" before invoking this script.
$runtime = if ($env:RUNTIME) { $env:RUNTIME } else { "net48" }
if ($runtime -eq "net8") {
    $runtimes = & dotnet --list-runtimes 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "dotnet not on PATH; cannot verify .NET 8 Desktop Runtime."
    } elseif ($runtimes -notmatch "Microsoft\.WindowsDesktop\.App 8\.") {
        Write-Warning ".NET 8 Windows Desktop Runtime not detected (required for the net8 build)."
        Write-Warning "  Install via:  winget install Microsoft.DotNet.DesktopRuntime.8"
        Write-Warning "Continuing anyway in case you have it via another route..."
    }
}

# Verify check_workbook.py is alongside this script
$pyPath = Join-Path $Here "check_workbook.py"
if (-not (Test-Path $pyPath)) {
    Write-Error "check_workbook.py not found at: $pyPath"
    Write-Error "Copy it from the repo's scripts/ folder alongside this .ps1 and re-run."
    exit 1
}

# --- 2. Resolve release tag ---------------------------------------------
if ($Tag -eq "latest") {
    $resolved = & gh api "repos/$Repo/releases/latest" --jq '.tag_name' 2>$null
    if ($LASTEXITCODE -ne 0 -or -not $resolved) {
        Write-Error "Could not resolve latest release. Try 'gh auth login' or pass an explicit -Tag."
        exit 1
    }
    Write-Host "Latest release: $resolved"
    $Tag = $resolved
}

# --- 3. Download .xll and workbook --------------------------------------
# Default to the net48 build (zero install). Set $env:RUNTIME = "net8" before
# invoking to exercise the perf variant instead.
if ($runtime -eq "net8") {
    $xllName = "ActuarialAddIn-AddIn64-packed-net8.xll"
} else {
    $xllName = "ActuarialAddIn-AddIn64-packed.xll"
}
$xlsxName = "actuarial_add_in.xlsx"
$xllPath = Join-Path $Here $xllName
$xlsxPath = Join-Path $Here $xlsxName

Write-Host "Downloading $xllName from release $Tag..."
& gh release download $Tag --repo $Repo --pattern $xllName --dir $Here --clobber
if ($LASTEXITCODE -ne 0) {
    Write-Error "gh release download failed for $xllName. Has the release workflow finished?"
    Write-Error "Check: https://github.com/$Repo/releases/tag/$Tag"
    exit 1
}

Write-Host "Downloading $xlsxName from release $Tag..."
& gh release download $Tag --repo $Repo --pattern $xlsxName --dir $Here --clobber
if ($LASTEXITCODE -ne 0) {
    Write-Error "gh release download failed for $xlsxName. The release may not have the workbook attached."
    Write-Error "Check: https://github.com/$Repo/releases/tag/$Tag"
    exit 1
}

# --- 4. Unblock the .xll (MOTW from gh download) ------------------------
try { Unblock-File $xllPath -ErrorAction Stop } catch { }

# --- 5. xlwings install if missing --------------------------------------
Write-Host "Ensuring xlwings is installed..."
& python -m pip install --quiet --disable-pip-version-check xlwings
if ($LASTEXITCODE -ne 0) {
    Write-Warning "pip install xlwings reported non-zero. Continuing anyway."
}

# --- 6. Run the check ---------------------------------------------------
Write-Host ""
Write-Host "Running check_workbook.py..."
Write-Host "----------------------------------------------------------"
& python $pyPath
$checkExit = $LASTEXITCODE
Write-Host "----------------------------------------------------------"

Write-Host ""
Write-Host "Done. Folder contents:"
Get-ChildItem $Here | Format-Table Name, Length, LastWriteTime -AutoSize

if ($checkExit -ne 0) {
    Write-Warning "check_workbook.py exited with code $checkExit. See output above."
}
exit $checkExit
