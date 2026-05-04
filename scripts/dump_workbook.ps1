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
    [Parameter(Mandatory)] [string] $Output,
    # If set, save the workbook after CalculateFullRebuild so the .xlsx file
    # carries Excel's cached calc values. Lets the workbook open cold (no XLL
    # loaded) and still show real numbers instead of #NAME? in every ACT_* cell.
    [switch] $SaveWorkbook
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

# Excel COM Variant CVErr codes -> human-readable error names. When a cell
# evaluates to one of the seven Excel errors, Range.Value2 returns the integer
# below; we want "#VALUE!" in the JSON, not "-2146826273".
$xlErrCodes = @{
    -2146826281 = '#DIV/0!'
    -2146826246 = '#N/A'
    -2146826259 = '#NAME?'
    -2146826288 = '#NULL!'
    -2146826252 = '#NUM!'
    -2146826265 = '#REF!'
    -2146826273 = '#VALUE!'
}

function Format-CellValue($v) {
    if ($null -eq $v) { return $null }
    # Excel COM Variant CVErr codes for the seven Excel errors arrive typed as
    # [double] (not [int]) when read via Range.Value2 in PowerShell. The hash
    # table uses int32 keys (PowerShell @{} default), so we must cast to int32
    # for ContainsKey to match — int64 silently fails the lookup.
    if ($v -is [double]) {
        if ([double]::IsNaN($v))                  { return 'NaN' }
        if ([double]::IsPositiveInfinity($v))     { return 'Infinity' }
        if ([double]::IsNegativeInfinity($v))     { return '-Infinity' }
        if ($v -eq [math]::Floor($v) -and $v -ge -2147483648 -and $v -le 2147483647) {
            $asInt = [int32]$v
            if ($xlErrCodes.ContainsKey($asInt))  { return $xlErrCodes[$asInt] }
        }
        return $v
    }
    if ($v -is [int] -or $v -is [long]) {
        $asInt = [int32]$v
        if ($xlErrCodes.ContainsKey($asInt))      { return $xlErrCodes[$asInt] }
        return $v
    }
    if ($v -is [datetime]) { return $v.ToString('o') }
    return "$v"
}

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
                        $cellValues.Add([pscustomobject]@{
                            sheet = $sht.Name; cell = $cellAddrAll
                            value = Format-CellValue $v
                        })
                    }

                    if (-not ($f -is [string])) { continue }
                    if (-not $f.StartsWith('=')) { continue }
                    $m = $rxAct.Match($f)
                    if (-not $m.Success) { continue }
                    $fnName = $m.Groups[1].Value

                    $vOut = Format-CellValue $v

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

        # Surface failure modes prominently. Earlier versions only matched on
        # string values starting with `#`, but Excel COM Variant CVErr codes
        # come through as integers (-2146826273 = #VALUE!, etc); Format-CellValue
        # now decodes those to the human name, so the regex is reliable.
        $errors = @($records | Where-Object { $_.value -is [string] -and $_.value -match '^#' })
        $nones  = @($records | Where-Object { $null -eq $_.value })
        if ($errors.Count -gt 0) {
            Write-Host ""
            Write-Host "!! $($errors.Count) cell(s) returned an Excel error:" -ForegroundColor Yellow
            $errors | Group-Object value | Sort-Object Count -Descending | ForEach-Object {
                Write-Host ("   {0,-10} {1,4} cells, e.g. {2}!{3} {4}" -f
                    $_.Name, $_.Count,
                    $_.Group[0].sheet, $_.Group[0].cell, $_.Group[0].function) -ForegroundColor Yellow
            }
            Write-Host ""
        }
        $nonePct = if ($records.Count -gt 0) { 100.0 * $nones.Count / $records.Count } else { 0 }
        if ($nonePct -ge 90) {
            throw ("$($nones.Count) of $($records.Count) cells ($([math]::Round($nonePct,1))%) " +
                   "came back as null. Excel computed nothing; the add-in did not load. " +
                   "Most common cause: the .NET Desktop Runtime matching the XLL's target " +
                   "framework is missing. Run ``dotnet --list-runtimes`` to verify.")
        }

        if ($SaveWorkbook) {
            Write-Host "Saving workbook with cached calc values: $Workbook"
            $book.Save()
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
