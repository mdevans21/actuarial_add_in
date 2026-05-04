<#
.SYNOPSIS
  Open the Actuarial Add-In workbook under a freshly-loaded XLL, recalculate,
  and dump every ACT_*-prefixed cell (sheet, cell, function, formula, value)
  to a JSON file. PowerShell-only — no Python or xlwings dependency.

.DESCRIPTION
  Mirrors the original scripts/check_workbook.py pipeline so the same dump
  format is produced. Invokes Excel via COM, calls RegisterXLL, opens the
  workbook, calls CalculateFullRebuild, then walks the used range on every
  sheet collecting cells whose formula matches `^=\s*ACT_[A-Z0-9_]+\(`.

.PARAMETER Xll
  Full Windows path to the packed AddIn64 .xll. Will be Unblock-File'd
  (clears Mark-of-the-Web) before Excel sees it.

.PARAMETER Workbook
  Full Windows path to the .xlsx workbook to dump. Also unblocked.

.PARAMETER Output
  Full Windows path of the JSON output file.

.EXAMPLE
  pwsh -File dump_workbook.ps1 `
       -Xll      C:\dump\ActuarialAddIn-AddIn64-packed.xll `
       -Workbook C:\dump\actuarial_add_in.xlsx `
       -Output   C:\dump\actuarial_add_in_dump.json
#>
param(
    [Parameter(Mandatory)] [string] $Xll,
    [Parameter(Mandatory)] [string] $Workbook,
    [Parameter(Mandatory)] [string] $Output
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

foreach ($p in @($Xll, $Workbook)) {
    if (-not (Test-Path -LiteralPath $p)) { throw "Path not found: $p" }
    Unblock-File -LiteralPath $p
}

# Match the first ACT_*(...) call anywhere in the formula. This catches both
# direct calls (=ACT_VERSION()) and wrapped calls (=INDEX(ACT_CL_FACTORS(...), 1));
# without this, every Chain Ladder / Copulas cell would be missed because they
# are all INDEX-wrapped to extract a single value from the array result.
$rxAct = [regex] '\b(ACT_[A-Z0-9_]+)\s*\('

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.ScreenUpdating = $false
$xl.AutomationSecurity = 1   # msoAutomationSecurityLow
$xl.AskToUpdateLinks = $false

try {
    Write-Host "RegisterXLL: $Xll"
    if (-not $xl.RegisterXLL($Xll)) {
        throw "RegisterXLL returned False. Confirm the .NET Desktop runtime " +
              "matches the XLL's target framework (project currently builds " +
              "net6.0-windows; .NET 6 WindowsDesktop must be installed)."
    }

    Write-Host "Opening: $Workbook"
    $book = $xl.Workbooks.Open($Workbook, 0, $false)
    try {
        Write-Host "CalculateFullRebuild..."
        $xl.CalculateFullRebuild()

        $records = New-Object System.Collections.Generic.List[object]
        # Sidecar: every non-empty cell's resolved Value2 (sheet, cell -> value).
        # Lets a Linux-side comparator resolve B7:K16 references in ACT_* formulas
        # without needing to re-open the xlsx through openpyxl.
        $cellValues = New-Object System.Collections.Generic.List[object]

        foreach ($sht in $book.Worksheets) {
            $used = $sht.UsedRange
            $rows = $used.Rows.Count
            $cols = $used.Columns.Count
            if ($rows -eq 0 -or $cols -eq 0) { continue }

            # Pull the entire used range as a single COM call — orders of
            # magnitude faster than per-cell access on large sheets.
            $formulaGrid = $used.Formula
            $valueGrid   = $used.Value2
            $r0 = $used.Row
            $c0 = $used.Column

            $localCount = 0
            for ($ri = 1; $ri -le $rows; $ri++) {
                for ($ci = 1; $ci -le $cols; $ci++) {
                    if ($rows -eq 1 -and $cols -eq 1) {
                        $f = $formulaGrid; $v = $valueGrid
                    } elseif ($rows -eq 1) {
                        $f = $formulaGrid[$ci]; $v = $valueGrid[$ci]
                    } elseif ($cols -eq 1) {
                        $f = $formulaGrid[$ri]; $v = $valueGrid[$ri]
                    } else {
                        $f = $formulaGrid[$ri, $ci]; $v = $valueGrid[$ri, $ci]
                    }
                    # Always log a sidecar entry for every non-empty cell.
                    if ($null -ne $v) {
                        $cellAddrAll = '{0}{1}' -f (
                            & {
                                param($n)
                                $s = ''
                                while ($n -gt 0) {
                                    $n--
                                    $s = [char](65 + ($n % 26)) + $s
                                    $n = [int]($n / 26)
                                }
                                $s
                            } ($c0 + $ci - 1)
                        ), ($r0 + $ri - 1)
                        $vAll = if ($v -is [double]) {
                            if ([double]::IsNaN($v))                    { 'NaN' }
                            elseif ([double]::IsPositiveInfinity($v))   { 'Infinity' }
                            elseif ([double]::IsNegativeInfinity($v))   { '-Infinity' }
                            else { $v }
                        } elseif ($v -is [datetime]) { $v.ToString('o') }
                        else { $v }
                        $cellValues.Add([pscustomobject]@{
                            sheet = $sht.Name; cell = $cellAddrAll; value = $vAll
                        })
                    }

                    if (-not ($f -is [string])) { continue }
                    if (-not $f.StartsWith('=')) { continue }
                    $m = $rxAct.Match($f)
                    if (-not $m.Success) { continue }
                    $fnName = $m.Groups[1].Value

                    # Excel sometimes returns DBNull, doubles, dates, errors —
                    # normalise so JSON round-trips cleanly. NaN/Infinity become
                    # the same string sentinels the C# harness uses.
                    if ($null -eq $v) { $vOut = $null }
                    elseif ($v -is [double]) {
                        if ([double]::IsNaN($v))      { $vOut = 'NaN' }
                        elseif ([double]::IsPositiveInfinity($v)) { $vOut = 'Infinity' }
                        elseif ([double]::IsNegativeInfinity($v)) { $vOut = '-Infinity' }
                        else { $vOut = $v }
                    }
                    elseif ($v -is [datetime]) { $vOut = $v.ToString('o') }
                    else { $vOut = "$v" }

                    $cellAddr = '{0}{1}' -f (
                        # Convert 1-based column index to letter(s).
                        & {
                            param($n)
                            $s = ''
                            while ($n -gt 0) {
                                $n--
                                $s = [char](65 + ($n % 26)) + $s
                                $n = [int]($n / 26)
                            }
                            $s
                        } ($c0 + $ci - 1)
                    ), ($r0 + $ri - 1)

                    $records.Add([pscustomobject]@{
                        sheet    = $sht.Name
                        cell     = $cellAddr
                        function = $fnName
                        formula  = $f
                        value    = $vOut
                    })
                    $localCount++
                }
            }
            if ($localCount -gt 0) { Write-Host "  $($sht.Name): $localCount ACT_ formulas" }
        }

        # ConvertTo-Json's default depth is 2 — bump it; structure is shallow.
        $json = $records | ConvertTo-Json -Depth 5 -Compress:$false
        # PowerShell wraps single-element arrays awkwardly; force [...] form.
        if ($records.Count -eq 1) { $json = "[$json]" }
        # Set-Content -Encoding UTF8 emits a BOM, which `json.load` rejects.
        # WriteAllText with UTF8 (no BOM by default in .NET 6+) is what we want.
        $utf8NoBom = New-Object System.Text.UTF8Encoding $false
        [System.IO.File]::WriteAllText($Output, $json, $utf8NoBom)
        Write-Host "Wrote $($records.Count) ACT_* records to $Output"

        # Sidecar: cells.json next to dump.json with every non-empty cell's value.
        $cellsPath = [System.IO.Path]::ChangeExtension($Output, '.cells.json')
        $cellsJson = $cellValues | ConvertTo-Json -Depth 3 -Compress:$false
        if ($cellValues.Count -eq 1) { $cellsJson = "[$cellsJson]" }
        [System.IO.File]::WriteAllText($cellsPath, $cellsJson, $utf8NoBom)
        Write-Host "Wrote $($cellValues.Count) cell values to $cellsPath"

        # Surface failure modes the way the Python original did.
        $errors = @($records | Where-Object { $_.value -is [string] -and $_.value -match '^#' })
        $nones  = @($records | Where-Object { $null -eq $_.value })
        if ($errors.Count -gt 0) {
            Write-Host "!! $($errors.Count) cell(s) returned an Excel error (#NAME?, #VALUE!, ...)" -ForegroundColor Yellow
        }
        $nonePct = if ($records.Count -gt 0) { 100.0 * $nones.Count / $records.Count } else { 0 }
        if ($nonePct -ge 90) {
            throw ("$($nones.Count) of $($records.Count) cells ($([math]::Round($nonePct,1))%) " +
                   "came back as null. Excel computed nothing; the add-in did not load. " +
                   "Most common cause: the .NET Desktop Runtime matching the XLL's target " +
                   "framework is missing. Run ``dotnet --list-runtimes`` to verify.")
        }
    }
    finally {
        $book.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($book) | Out-Null
    }
}
finally {
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
