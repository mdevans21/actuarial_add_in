# Developing the Actuarial Add-In

This document covers the build, test, release, and spreadsheet-regeneration
workflow for contributors. For day-to-day user documentation see the
[top-level README](../README.md); for repository conventions and the
canonical test parameters see [`../CONTRIBUTING.md`](../CONTRIBUTING.md).

---

## Build

The add-in targets `net6.0-windows`. Excel-DNA bundles a Windows Desktop
runtime dependency, so the `.xll` only runs on Windows Excel even though
`dotnet build` itself runs cross-platform.

### On Windows

```bash
dotnet restore ActuarialAddIn.sln
dotnet build ActuarialAddIn.sln --configuration Release
```

Output locations:

- 32-bit: `src/ActuarialAddIn/bin/Release/net6.0-windows/publish/ActuarialAddIn-AddIn-packed.xll`
- 64-bit: `src/ActuarialAddIn/bin/Release/net6.0-windows/publish/ActuarialAddIn-AddIn64-packed.xll`

### On WSL / Linux

The .NET 6 SDK must be on `PATH`, and you need the
`EnableWindowsTargeting` property:

```bash
dotnet build ActuarialAddIn.sln \
  --configuration Release \
  -p:EnableWindowsTargeting=true
```

This compiles and packs `.xll` artefacts even on Linux, but they will
only load inside Windows Excel.

### Invoking `dotnet` from WSL via PowerShell (legacy workflow)

Earlier setups invoked the Windows `dotnet` from WSL via
`powershell.exe`. This still works and sometimes avoids WSL/Windows
filesystem cache issues (see below):

```bash
powershell.exe -Command "cd 'C:\Users\matth\Code\actuarial_add_in'; dotnet build src/ActuarialAddIn/ActuarialAddIn.csproj --configuration Release"
```

**Always rebuild before committing any C# change.** The compiled
`ActuarialAddIn.dll` is required for the Excel examples in
`excel/actuarial_add_in.xlsx` to evaluate correctly on the next Excel
open.

---

## Tests

### C# test harness (authoritative)

```bash
dotnet run --project src/ActuarialAddIn.Tests -- test_results.md
```

Produces two files:

- `test_results.md` — human-readable markdown tables of every function's
  outputs on the canonical test parameters (Poisson λ=5, Taylor-Ashe
  triangle, etc.).
- `tests/fixtures/addin_outputs.json` — structured JSON snapshot
  consumed by the reconciliation notebook (see below).

The harness lives in `src/ActuarialAddIn.Tests/`:

- `Program.cs` — markdown generator; one big imperative script.
- `AddinOutputsEmitter.cs` — JSON emitter; systematic
  (function, args, result) records grouped by family.

### Reconciliation notebook (CI gate)

```bash
cd tests
pip install -r requirements.txt

# Regenerate from the Python source (edit this, not the .ipynb directly)
python build_reconciliation_notebook.py

# Execute — raises AssertionError on any Basic failure
papermill reconciliation.ipynb reconciliation_executed.ipynb --log-output
```

The notebook is regenerated from
`tests/build_reconciliation_notebook.py` so it stays reviewable as
plain Python. Re-running the generator after adding a reconciliation is
idempotent.

The final cell asserts every harness function name (less six trivial
metadata exemptions) has at least one `record(SECTION, Result(...))`
call. Adding a new C# function without wiring it through the notebook
fails the build.

### End-to-end dump test (live Excel + XLL)

The C# harness exercises the same code on Linux through MathNet/dotnet
directly — useful but not the same as running through Excel-DNA's
hosted CLR inside Excel. The dump test runs the workbook *under a
freshly-loaded XLL in a real Excel COM session* and dumps every cell.

From WSL2 (the typical setup — bash side handles `gh` auth, Windows
side runs Excel):

```bash
# Latest release
scripts/run_dump.sh

# Specific tag
scripts/run_dump.sh v0.7.0

# Keep the scratch dir for inspection
KEEP=1 scripts/run_dump.sh v0.7.0
```

Mechanics: `run_dump.sh` `gh release download`s the .xll + .xlsx,
copies them and `scripts/dump_workbook.ps1` to a Windows-visible
scratch dir, invokes `powershell.exe` over WSL interop, and copies
the resulting JSON back to:

- `tests/actuarial_add_in_dump.json` — `{sheet, cell, function,
  formula, value}` for every cell whose formula contains an `ACT_*(`
  call (matches direct calls *and* `INDEX(ACT_X(...), n)` wrappers).
- `tests/actuarial_add_in_cells.json` — sidecar map of every
  non-empty cell's resolved value, used by the audit comparator to
  resolve cell refs in formulas like `=ACT_CL_FACTORS(B7:K16)` to the
  corresponding harness records.

Requirements (one-time, all on the Windows host):
- Excel installed (any modern version).
- `Microsoft.WindowsDesktop.App 6.x` (matches the XLL's target
  framework). Verify with `dotnet --list-runtimes`.
- `gh` CLI on the WSL side, authenticated.

A typical run takes ~10 s end-to-end and produces 1372 records on the
v0.7.0 workbook. The dump should reconcile bit-exact against the C#
harness for every cell whose args appear in both surfaces.

### Test-data reference

Shared parameters between the C# harness, the reconciliation notebook,
and the Excel examples:

| Family | Parameters |
|---|---|
| Poisson | λ = 5 |
| Negative Binomial | r = 5, p = 0.3 |
| Lognormal | μ = 0, σ = 1 |
| Gamma | α = 2, β = 1 |
| Pareto I | α = 2, xm = 1 |
| MBBEFD | b = 2, g = 3 |
| Swiss Re | curves 1–5 |
| Lloyd's | Y1–Y4 |
| XOL layer | attach = 1M, limit = 5M |
| Student-t copula | df = 5, 7×7 corr matrix ρ = 0.6 decay, seed = 42 |
| Panjer Poisson | λ = 2, max_s = 40 |
| Panjer NegBin | r = 2, p = 0.5, max_s = 40 |
| Panjer Binomial | n = 5, p = 0.3, max_s = 40 |
| Aggregate stats | h = 0.5 |
| Chain ladder triangle | Taylor-Ashe 10×10 cumulative paid |
| Expected LDFs | 3.491, 1.747, 1.457, 1.174, 1.104, 1.086, 1.054, 1.077, 1.018 |
| Expected total IBNR | 18,680,856 |

---

## Release

1. **Update `src/ActuarialAddIn/VersionInfo.cs`:**
   - `CurrentVersion` — new tag (e.g. `"0.4.0"`).
   - `BuildDate` — today (ISO date).
   - `GetCommitHistory()` — paste the output of
     `git log --format='new CommitInfo("%h", "%cs", "%s"),' -20`.
2. **Update [`CHANGELOG.md`](../CHANGELOG.md)** — move items from
   `[Unreleased]` to the new version section, date-stamp the header.
3. **Regenerate the Versions tab in the spreadsheet** (runs automatically
   from `scripts/populate_examples.py`).
4. **Commit** the version bump alone:
   ```bash
   git commit -am "Bump version to v0.4.0"
   ```
5. **Ask the user whether the next push should trigger a release.** If
   yes, tag and push:
   ```bash
   git tag v0.4.0
   git push origin main --tags
   ```
   The `release` workflow (`.github/workflows/release.yml`) builds
   Release-mode artefacts and publishes `.xll` binaries.

If the user just wants to push without releasing:

```bash
git push origin main
```

---

## Spreadsheet regeneration

Generating `excel/actuarial_add_in.xlsx` is a **two-step process**
because `openpyxl` cannot emit Excel 365 dynamic-array metadata — every
formula ends up with an `@` (implicit intersection) prefix until
`xlwings` rewrites them via `Formula2`.

```bash
# Step 1 — WSL/Linux Python: build the workbook with openpyxl
python3 scripts/populate_examples.py

# Step 2 — Windows Python with Excel installed: upgrade formulas to Formula2
python3.exe scripts/fix_array_formulas.py
```

### Why two steps?

- `openpyxl` writes cells/formulas/charts cleanly but does not know how
  to write the dynamic-array XML metadata Excel 365 uses.
- `xlwings` drives Excel itself via COM, which knows exactly how to set
  `Formula2` — the property that turns `=ACT_CL_FACTORS(…)` into a
  spilling array formula instead of `=@ACT_CL_FACTORS(…)`.

### Step 2 requirements

- `pip install xlwings` on Windows Python (not WSL Python).
- Excel must be **closed** before running — xlwings opens it in the
  background.
- Idempotent: strips existing wrappers before re-adding, so safe to run
  multiple times.

### Special handling in step 2

- Fit functions (`ACT_DIST_*_FIT`) return multi-element arrays: wrapped
  in `TRANSPOSE()` with parameter labels.
- Other array returns (copulas, return-period tables): wrapped in
  `INDEX(..., 1, 1)` where a scalar is expected.
- All other tabs: plain `Formula2` upgrade.

**Do not** use ZIP-level postprocessing to modify the `.xlsx` after
`openpyxl` saves it — rewriting the archive with Python's `zipfile`
module corrupts dynamic-array metadata and triggers Excel recovery.

### Chart styling

Every openpyxl chart must have `chart.roundedCorners = False` set
immediately after creation — the default is rounded, which we don't
want.

---

## Troubleshooting

### WSL / Windows filesystem cache divergence

WSL and Windows occasionally disagree on whether a file has been
updated, especially when sources live on the Windows side of the
`\\wsl.localhost\` bridge and the build runs from WSL. Symptoms:

- Build succeeds but new functions throw `#NAME` in Excel.
- `ActuarialAddIn.dll` `LastWriteTime` is stale despite a successful
  build.
- `strings` / PowerShell search doesn't find expected function names in
  the DLL.

Diagnose:

```powershell
# From PowerShell
Get-Item 'C:\Users\matth\Code\actuarial_add_in\src\ActuarialAddIn\bin\Release\net6.0-windows\ActuarialAddIn.dll' | Select-Object LastWriteTime
```

```powershell
# Find every ACT_* symbol in the compiled DLL
$bytes = [IO.File]::ReadAllBytes('…\ActuarialAddIn.dll')
$text  = [Text.Encoding]::ASCII.GetString($bytes)
[regex]::Matches($text, 'ACT_[A-Z_]+') | ForEach-Object { $_.Value } | Sort-Object -Unique
```

Fix:

```bash
# Copy source across the WSL/Windows boundary
powershell.exe -Command "Copy-Item -Path '\\wsl.localhost\Ubuntu\home\matth\Code\actuarial_add_in\src\ActuarialAddIn\Functions\*.cs' -Destination 'C:\Users\matth\Code\actuarial_add_in\src\ActuarialAddIn\Functions\' -Force"

# Nuke caches and rebuild
rm -rf src/ActuarialAddIn/bin src/ActuarialAddIn/obj
powershell.exe -Command "cd 'C:\Users\matth\Code\actuarial_add_in'; dotnet build src/ActuarialAddIn/ActuarialAddIn.csproj --configuration Release"
```

If Excel has the `.xll` locked (build fails with "access denied"),
close Excel before rebuilding.

### Re-loading the add-in after a build

1. Close Excel if open.
2. Reopen `excel/actuarial_add_in.xlsx`.
3. Excel will reload the registered add-in automatically. If it's not
   registered, load it once via **File → Options → Add-ins → Browse**.

---

## Modifying the `.xlsx` outside openpyxl

Occasionally you may need to hand-edit `xl/worksheets/*.xml`. The file
is a ZIP archive of XML:

```bash
unzip -o excel/actuarial_add_in.xlsx -d /tmp/xlsx_extract
# edit /tmp/xlsx_extract/xl/worksheets/sheetN.xml
cd /tmp/xlsx_extract && zip -r ../updated.xlsx . -x "*.DS_Store"
cp /tmp/updated.xlsx excel/actuarial_add_in.xlsx
```

The zip must be created from **inside** the extracted directory so
paths are relative. Verify the output loads in Excel before committing.

---

## Dynamic arrays — C# patterns

Functions that return a column or row should accept an optional
`vertical = true` argument and return a 2D `object[,]` so Excel 365
spills correctly:

```csharp
public static object ACT_CL_FACTORS(double[,] triangle, bool vertical = true)
{
    var factors = GetFactorsArray(triangle);   // private helper
    int n = factors.Length;

    if (vertical)
    {
        var result = new object[n, 1];
        for (int i = 0; i < n; i++) result[i, 0] = factors[i];
        return result;
    }
    else
    {
        var result = new object[1, n];
        for (int i = 0; i < n; i++) result[0, i] = factors[i];
        return result;
    }
}
```

### Internal helper pattern

When public Excel functions need to call each other internally, route
through private helpers returning native types (`double[]`, not
`object[,]`):

```csharp
public static object ACT_CL_ULTIMATE(double[,] triangle, bool vertical = true)
{
    var factors = GetFactorsArray(triangle);   // NOT ACT_CL_FACTORS
    // ...
}

private static double[] GetFactorsArray(double[,] triangle)
{
    // pure computation
}
```

This avoids boxing, type-conversion bugs, and circular dependencies
when a function's output shape changes.

---

## Where things live

```
actuarial_add_in/
├── src/
│   ├── ActuarialAddIn/
│   │   ├── Functions/          # [ExcelFunction] entry points (10 files, ~183 functions)
│   │   ├── Helpers/            # ArrayHelpers.cs
│   │   ├── Ribbon.xml          # custom ribbon tab
│   │   ├── VersionInfo.cs      # updated on every release
│   │   └── ActuarialAddIn-AddIn.dna  # Excel-DNA manifest
│   └── ActuarialAddIn.Tests/
│       ├── Program.cs              # markdown test report (authoritative on function values)
│       └── AddinOutputsEmitter.cs  # JSON emitter for the notebook
├── tests/
│   ├── reconciliation.ipynb              # CI-gated reconciliation
│   ├── build_reconciliation_notebook.py  # generator for the above
│   ├── requirements.txt
│   └── fixtures/
│       └── addin_outputs.json  # produced by the C# emitter on each build
├── excel/
│   └── actuarial_add_in.xlsx        # worked examples
├── scripts/
│   ├── populate_examples.py    # step 1 of spreadsheet regen
│   ├── fix_array_formulas.py   # step 2 of spreadsheet regen
│   ├── dump_workbook.ps1       # PowerShell: load XLL via Excel COM, dump cells
│   ├── run_dump.sh             # WSL: gh-download + drive dump_workbook.ps1
│   ├── check_workbook.py       # legacy: same job, xlwings-based
│   └── setup_and_check.ps1     # legacy: pure-Windows wrapper for check_workbook.py
├── .github/workflows/
│   ├── build.yml      # build + reconcile job
│   └── release.yml    # tag → GitHub release
├── README.md
├── CHANGELOG.md
├── CONTRIBUTING.md      # repo conventions, canonical test parameters
└── docs/DEVELOPING.md   # you are here
```
