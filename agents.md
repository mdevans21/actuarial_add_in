# Agent Instructions

This document contains instructions for AI agents working on this project.

## Excel Examples Consistency

**Important:** The examples in `excel/actuarial_add_in.xlsm` must be kept consistent with the test suite in `src/ActuarialAddIn.Tests/Program.cs`.

When updating the test suite:
1. Update the corresponding Excel sheet(s) with the same test values
2. Ensure parameter values match exactly (e.g., lambda=5 for Poisson, r=5/p=0.3 for Negative Binomial)
3. Run `scripts/populate_examples.py` to regenerate examples, or update manually

When adding new functions:
1. Add tests to `src/ActuarialAddIn.Tests/Program.cs`
2. Add a new sheet or section to the Excel workbook with matching examples
3. Update `scripts/populate_examples.py` to include the new function

## Excel Chart Styling

**Use Excel defaults for charts.** Do not override chart styling properties like `roundedCorners`. Let Excel use its default appearance. When creating charts in `populate_examples.py`:
- Set title, axes, dimensions as needed
- Do NOT set `chart.roundedCorners` or similar styling overrides
- Keep charts simple and functional

## Test Data Reference

The following test parameters are used consistently across the test suite and Excel examples:

| Distribution | Parameters | Functions |
|-------------|------------|-----------|
| Poisson | lambda=5 | ACT_DIST_POISSON_* |
| Negative Binomial | r=5, p=0.3 | ACT_DIST_NEGBIN_* |
| Lognormal | mu=0, sigma=1 | ACT_DIST_LOGNORM_* |
| Gamma | alpha=2, beta=1 | ACT_DIST_GAMMA_* |
| Pareto | alpha=2, xm=1 | ACT_DIST_PARETO_* |

| Exposure Curves | Parameters | Functions |
|----------------|------------|-----------|
| MBBEFD | b=2, g=3 | ACT_EXPOSURE_MBBEFD |
| Swiss Re | curves 1-5 | ACT_EXPOSURE_SWISSRE |
| Lloyd's | Y1-Y4 | ACT_EXPOSURE_LLOYDS |

| Reinsurance | Parameters | Functions |
|-------------|------------|-----------|
| XOL Layer | attachment=1M, limit=5M | ACT_XOL_LAYER_LOSS |

| Chain Ladder | Data |
|--------------|------|
| Triangle | Taylor-Ashe 10x10 cumulative paid losses |
| Source | England & Verrall (2002), Taylor & Ashe (1983) |
| Expected Factors | 3.491, 1.747, 1.457, 1.174, 1.104, 1.086, 1.054, 1.077, 1.018 |
| Expected Total IBNR | 18,680,856 |

## ODP Bootstrap Reference Values

**Target:** Non-Constant Scale Parameter ODP Bootstrap (England & Verrall 2002)

Reference: England, P.D. and Verrall, R.J. (2002) "Stochastic Claims Reserving in General Insurance", British Actuarial Journal, 8(3), 443-518.

### E&V ODP Bootstrap Results (Taylor-Ashe Triangle)

| AY | IBNR | Constant SE | Const % | Non-Const SE | Non-Const % |
|----|------|-------------|---------|--------------|-------------|
| 1 | 0 | 0 | 0.0% | 0 | 0.0% |
| 2 | 94,634 | 112,552 | 119.0% | 43,882 | 46.4% |
| 3 | 469,511 | 217,547 | 46.3% | 109,449 | 23.3% |
| 4 | 709,638 | 262,934 | 37.1% | 141,509 | 19.9% |
| 5 | 984,889 | 306,595 | 31.1% | 256,031 | 26.0% |
| 6 | 1,419,459 | 375,745 | 26.5% | 398,377 | 28.1% |
| 7 | 2,177,641 | 500,332 | 23.0% | 529,898 | 24.3% |
| 8 | 3,920,301 | 791,481 | 20.2% | 735,245 | 18.8% |
| 9 | 4,278,972 | 1,060,473 | 24.8% | 809,457 | 18.9% |
| 10 | 4,625,811 | 2,025,898 | 43.8% | 1,285,560 | 27.8% |
| **Total** | **18,680,856** | **2,992,296** | **16.0%** | **2,228,677** | **11.9%** |

### Implementation Notes

- **Target methodology:** Non-Constant Scale Parameter ODP Bootstrap
- **Key difference:** Non-constant uses period-specific scale parameters (phi_j) instead of a single global phi
- **Our implementation:** Stratifies residuals by development period, uses period-specific phi for process variance
- **Scale parameter phi:** Constant = 75,019 (single value); Non-constant = varies by development period (8k to 132k)

### Benchmark Results (10,000 iterations, seed=123)

| Implementation | Total SE | vs E&V Non-Const |
|----------------|----------|------------------|
| **Our C#** | **2,161,659** | **97%** |
| E&V Non-Constant | 2,228,677 | 100% |
| chainladder hat_adj=False | 2,480,570 | 111% |
| chainladder hat_adj=True | 2,971,846 | 133% |
| E&V Constant | 2,992,296 | 134% |

### Why chainladder differs from E&V non-constant

The Python `chainladder` package does NOT implement E&V's non-constant scale:
- chainladder pools all residuals together (not stratified by period)
- chainladder uses a single global phi for process variance
- chainladder's `hat_adj=True` matches E&V **constant** scale
- chainladder's `hat_adj=False` is between the two E&V variants

Our C# implementation follows the true E&V non-constant approach:
1. Period-specific φⱼ values calculated separately for each development period
2. Residuals standardized within each period and sampled from within period (stratified)
3. Process variance uses period-specific φⱼ in Gamma sampling

| Copula | Parameters |
|--------|------------|
| Student-t | df=5, 7x7 correlation matrix (rho=0.6 decay), seed=42 |

## Test Suite Consistency

Three test sources must remain consistent:
1. **C# Test Suite** (`src/ActuarialAddIn.Tests/Program.cs`) - Primary test definitions
2. **Python Benchmarks** (`tests/run_benchmarks.py`) - Comparison against scipy/chainladder
3. **Excel Examples** (`excel/actuarial_add_in.xlsm`) - Visual examples (regenerate with `scripts/populate_examples.py`)

## Git Push and Releases

**Before pushing to the remote repository, ask the user if this push should trigger a release.**

- Pushing to `main` updates the repository but does not create a release
- Pushing a version tag (e.g., `v0.1.0`) triggers the GitHub Actions workflow to build and publish a release

If the user wants a release:
```bash
git tag v<version>
git push origin main --tags
```

If the user just wants to push changes without a release:
```bash
git push origin main
```

## Building the Add-In from WSL

The project runs in WSL but requires Windows .NET SDK for building. Use PowerShell to invoke dotnet:

```bash
powershell.exe -Command "cd 'C:\Users\matth\Code\actuarial_add_in'; dotnet build src/ActuarialAddIn/ActuarialAddIn.csproj --configuration Release"
```

**Output locations:**
- 32-bit: `src/ActuarialAddIn/bin/Release/net6.0-windows/publish/ActuarialAddIn-AddIn-packed.xll`
- 64-bit: `src/ActuarialAddIn/bin/Release/net6.0-windows/publish/ActuarialAddIn-AddIn64-packed.xll`

**IMPORTANT: Always rebuild before committing.** When any C# code in `src/ActuarialAddIn/` has been modified, run the build command above before creating a git commit. This ensures the compiled add-in stays in sync with the source code.

### WSL/Windows Filesystem Sync Issues

WSL and Windows may have different views of the same files due to filesystem caching. This can cause builds to use stale code even when source files appear updated in WSL.

**Symptoms of sync issues:**
- Build succeeds but new functions give #NAME errors in Excel
- DLL timestamp is old despite successful build
- `strings` or PowerShell search doesn't find expected function names in DLL

**Prevention - verify after every build:**

1. Check the DLL was actually updated (timestamp should be recent):
```bash
powershell.exe -Command "Get-Item 'C:\Users\matth\Code\actuarial_add_in\src\ActuarialAddIn\bin\Release\net6.0-windows\ActuarialAddIn.dll' | Select-Object LastWriteTime"
```

2. Verify new functions exist in the compiled DLL:
```bash
powershell.exe -Command "\$bytes = [System.IO.File]::ReadAllBytes('C:\Users\matth\Code\actuarial_add_in\src\ActuarialAddIn\bin\Release\net6.0-windows\ActuarialAddIn.dll'); \$text = [System.Text.Encoding]::ASCII.GetString(\$bytes); [regex]::Matches(\$text, 'ACT_[A-Z_]+') | ForEach-Object { \$_.Value } | Sort-Object -Unique"
```

**If sync issues occur:**

1. Force copy files from WSL to Windows:
```bash
powershell.exe -Command "Copy-Item -Path '\\wsl.localhost\Ubuntu\home\matth\Code\actuarial_add_in\src\ActuarialAddIn\Functions\*.cs' -Destination 'C:\Users\matth\Code\actuarial_add_in\src\ActuarialAddIn\Functions\' -Force"
```

2. Delete build cache and rebuild:
```bash
rm -rf src/ActuarialAddIn/bin src/ActuarialAddIn/obj
powershell.exe -Command "cd 'C:\Users\matth\Code\actuarial_add_in'; dotnet build src/ActuarialAddIn/ActuarialAddIn.csproj --configuration Release"
```

3. If Excel has the add-in locked (build fails with "access denied"), close Excel completely before rebuilding.

**After building:**
1. Close Excel if open
2. Reopen the spreadsheet
3. Excel will load the updated add-in automatically (if configured) or load manually via File > Options > Add-ins

## Modifying Excel .xlsm Files

Excel .xlsm files are ZIP archives containing XML files. To modify formulas or structure:

### 1. Extract the workbook

```bash
mkdir -p /tmp/xlsm_extract
unzip -o excel/actuarial_add_in.xlsm -d /tmp/xlsm_extract
```

### 2. Understand the structure

- `xl/workbook.xml` - Sheet names and order
- `xl/worksheets/sheet1.xml`, `sheet2.xml`, etc. - Individual sheet content
- `xl/sharedStrings.xml` - String table (cell text values reference indices here)
- `xl/calcChain.xml` - Formula calculation chain

### 3. Modify sheet XML

Formulas are stored in `<f>` elements within `<c>` (cell) elements:

```xml
<c r="B21"><f>ACT_CL_FACTORS(B7:K16,,,,FALSE)</f></c>
```

To find specific formulas:
```bash
grep -o '<c r="[^"]*"[^>]*><f>[^<]*</f>' /tmp/xlsm_extract/xl/worksheets/sheet7.xml
```

### 4. Repackage the workbook

```bash
cd /tmp/xlsm_extract
zip -r ../updated_workbook.xlsm . -x "*.DS_Store"
cp /tmp/updated_workbook.xlsm /home/matth/Code/actuarial_add_in/excel/actuarial_add_in.xlsm
```

**Important:** The zip must be created from within the extracted directory to preserve the correct structure.

### 5. Verify changes

```python
import zipfile
with zipfile.ZipFile('excel/actuarial_add_in.xlsm', 'r') as z:
    content = z.read('xl/worksheets/sheet7.xml').decode('utf-8')
    # Check for expected formulas
    print('ACT_CL_FACTORS' in content)
```

## Dynamic Arrays and Excel 365

When designing Excel functions for Excel 365:

1. **Return 2D arrays** (`object[,]`) from C# for spill behavior
2. **Add a `vertical` parameter** (default TRUE) for column vs row output
3. **Remove INDEX wrappers** - let formulas spill naturally
4. **The @ prefix** is NOT stored in formulas - Excel 365 displays it dynamically for implicit intersection

Example C# pattern:
```csharp
public static object MyFunction(double[,] data, bool vertical = true)
{
    var values = ComputeValues(data);
    int n = values.Length;

    if (vertical)
    {
        var result = new object[n, 1];
        for (int i = 0; i < n; i++)
            result[i, 0] = values[i];
        return result;
    }
    else
    {
        var result = new object[1, n];
        for (int i = 0; i < n; i++)
            result[0, i] = values[i];
        return result;
    }
}
```

## Internal Helper Functions

When public Excel functions call each other internally, create private helper functions that return native types (`double[]`) instead of `object` or `object[,]`:

```csharp
// Public Excel function
public static object ACT_CL_FACTORS(double[,] triangle, bool vertical = true)
{
    var factors = GetFactorsArray(triangle);  // Use helper
    // ... format as 2D array for Excel
}

// Private helper for internal use
private static double[] GetFactorsArray(double[,] triangle)
{
    // ... compute and return double[]
}

// Other functions use the helper
public static object ACT_CL_ULTIMATE(double[,] triangle, bool vertical = true)
{
    var factors = GetFactorsArray(triangle);  // Not ACT_CL_FACTORS
    // ...
}
```

This avoids type conversion issues and circular dependencies between public functions.

## Validation Testing Framework

### Overview

The validation framework compares function outputs across three sources:
1. **Excel formulas** (C# Add-In)
2. **Python implementations** (scipy, chainladder)
3. **Online reference values** (E&V 2002, academic papers)

### Fixture Files

Reference values are stored in JSON files under `tests/fixtures/`:

| File | Contents |
|------|----------|
| `distributions.json` | scipy expected values for all 10 distributions |
| `chain_ladder.json` | E&V (2002) reference values, Taylor-Ashe triangle |
| `copulas.json` | Copula CDF and tail dependence values |
| `exposure_curves.json` | Bernegger (1997), Swiss Re, Lloyd's curves |
| `credibility.json` | CAS exam and textbook reference values |
| `sources.json` | URLs and citations for all validation sources |

### Running Validation

```bash
# Generate full validation report
python tests/compare_sources.py --generate-report

# Run pytest tests (if pytest installed)
python -m pytest tests/test_validation.py -v

# Output: tests/reports/VALIDATION_REPORT.md
```

### Validation Sources by Category

| Category | Primary Source | Functions |
|----------|---------------|-----------|
| Distributions | scipy.stats | ACT_DIST_* (all 10) |
| Chain Ladder | E&V (2002) | ACT_CL_*, ACT_MACK_* |
| Copulas | McNeil et al (2015) | ACT_COPULA_* |
| Exposure Curves | Bernegger (1997) | ACT_EXPOSURE_* |
| Credibility | CAS Exam 5 | ACT_CREDIBILITY_* |

### Adding New Validation Tests

1. **Add reference value to fixture** (`tests/fixtures/*.json`)
2. **Add validator method** (`tests/compare_sources.py`)
3. **Add pytest test** (`tests/test_validation.py`)

See `docs/TESTING.md` for detailed documentation.

### Online Reference Sources

**Chain Ladder:**
- E&V (2002): https://doi.org/10.1017/S1357321700003809
- chainladder-python: https://chainladder-python.readthedocs.io/
- R ChainLadder: https://cran.r-project.org/web/packages/ChainLadder/

**Distributions:**
- scipy.stats: https://docs.scipy.org/doc/scipy/reference/stats.html
- R actuar: https://cran.r-project.org/web/packages/actuar/

**Exposure Curves:**
- Bernegger (1997): https://doi.org/10.2143/AST.27.1.563208
- Swiss Re Sigma: https://www.swissre.com/institute/research/sigma-research.html

**Copulas:**
- McNeil et al (2015): Quantitative Risk Management, Princeton University Press
- R copula: https://cran.r-project.org/web/packages/copula/

**Credibility:**
- CAS Exam 5: https://www.casact.org/examinations/
- NCCI Experience Rating: https://www.ncci.com/
