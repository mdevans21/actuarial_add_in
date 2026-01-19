# Actuarial Excel Add-In

A comprehensive Excel add-in for general insurance actuarial calculations, built with C# and Excel-DNA.

This code is in beta and offered without warranty. Please raise bug reports and feature requests through GitHub. 

To install the addin, follow the instructions below, or download the example excel. The example has the add in installed already and shows example usage. 

## Features

### Statistical Distributions
PDF, CDF, and Inverse CDF for:
- **Poisson** - `ACT_DIST_POISSON_PDF`, `ACT_DIST_POISSON_CDF`, `ACT_DIST_POISSON_INV`
- **Negative Binomial** - `ACT_DIST_NEGBIN_PDF`, `ACT_DIST_NEGBIN_CDF`, `ACT_DIST_NEGBIN_INV`
- **Lognormal** - `ACT_DIST_LOGNORM_PDF`, `ACT_DIST_LOGNORM_CDF`, `ACT_DIST_LOGNORM_INV`
- **Gamma** - `ACT_DIST_GAMMA_PDF`, `ACT_DIST_GAMMA_CDF`, `ACT_DIST_GAMMA_INV`
- **Pareto** - `ACT_DIST_PARETO_PDF`, `ACT_DIST_PARETO_CDF`, `ACT_DIST_PARETO_INV`
- **GPD** - `ACT_DIST_GPD_PDF`, `ACT_DIST_GPD_CDF`, `ACT_DIST_GPD_INV`
- **Weibull** - `ACT_DIST_WEIBULL_PDF`, `ACT_DIST_WEIBULL_CDF`, `ACT_DIST_WEIBULL_INV`
- **Beta** - `ACT_DIST_BETA_PDF`, `ACT_DIST_BETA_CDF`, `ACT_DIST_BETA_INV`
- **Exponential** - `ACT_DIST_EXP_PDF`, `ACT_DIST_EXP_CDF`, `ACT_DIST_EXP_INV`
- **Burr Type XII** - `ACT_DIST_BURR_PDF`, `ACT_DIST_BURR_CDF`, `ACT_DIST_BURR_INV`

### Exposure Curves
- **MBBEFD/Swiss Re** - `ACT_EXPOSURE_MBBEFD`, `ACT_EXPOSURE_SWISSRE`
- **Lloyd's** - `ACT_EXPOSURE_LLOYDS`
- **Power** - `ACT_EXPOSURE_POWER`, `ACT_EXPOSURE_INVERSE_POWER`
- **Pareto** - `ACT_EXPOSURE_PARETO`
- **Riebesell** - `ACT_EXPOSURE_RIEBESELL`, `ACT_EXPOSURE_RIEBESELL_INV`
- **Layer Rating** - `ACT_EXPOSURE_LAYER_RATE`

### Copulas
- **Gaussian Copula** - `ACT_COPULA_GAUSSIAN`, `ACT_COPULA_GAUSSIAN_SINGLE`
- **Student-t Copula** - `ACT_COPULA_STUDENT_T`, `ACT_COPULA_STUDENT_T_SINGLE`
- **Clayton Copula** - `ACT_COPULA_CLAYTON`, `ACT_COPULA_CLAYTON_CDF`
- **Frank Copula** - `ACT_COPULA_FRANK`, `ACT_COPULA_FRANK_CDF`
- **Gumbel Copula** - `ACT_COPULA_GUMBEL`, `ACT_COPULA_GUMBEL_CDF`
- **Utilities** - `ACT_COPULA_TAU_TO_THETA`, `ACT_COPULA_TAIL_LOWER`, `ACT_COPULA_TAIL_UPPER`

### Chain Ladder Reserving
- **Development Factors** - `ACT_CL_FACTORS`
- **Ultimates** - `ACT_CL_ULTIMATE`
- **IBNR** - `ACT_CL_IBNR`
- **Latest Diagonal** - `ACT_CL_LATEST`
- **Mack Standard Errors** - `ACT_MACK_FACTOR_SE`, `ACT_MACK_RESERVE_SE`
- **Bootstrap** - `ACT_CL_BOOTSTRAP`, `ACT_CL_BOOTSTRAP_ORIGIN`
- **Bornhuetter-Ferguson** - `ACT_BF_ULTIMATE`
- **Cape Cod** - `ACT_CAPECOD_ULTIMATE`, `ACT_CAPECOD_ELR`
- **Berquist-Sherman** - `ACT_BERQUIST_SHERMAN`
- **Triangle Utilities** - `ACT_TRIANGLE_TO_INCREMENTAL`, `ACT_INCREMENTAL_TO_CUMULATIVE`, `ACT_TRIANGLE_DIAGONAL`, `ACT_TRIANGLE_LINK_RATIOS`

### Reinsurance
- **Excess of Loss** - `ACT_XOL_LAYER_LOSS`, `ACT_XOL_EXPECTED_LOSS`
- **ILF** - `ACT_ILF_PARETO`

### Return Periods
- **Loss Interpolation** - `ACT_RETURN_PERIOD_LOSS`, `ACT_RETURN_PERIOD_TABLE`
- **AAL Calculation** - `ACT_AAL_FROM_OEP`

### Interpolation
- **Linear** - `ACT_INTERP` (with FLAT/GRADIENT/ERROR extrapolation)
- **2D Bilinear** - `ACT_INTERP2D`
- **Log-Linear** - `ACT_INTERP_LOG`

## Installation

### From Release
1. Download the latest `.xll` file from [Releases](../../releases)
2. Open Excel
3. Go to File → Options → Add-ins
4. Select "Excel Add-ins" at the bottom → Go
5. Click Browse and select the downloaded `.xll` file
6. For 64-bit Excel, use `ActuarialAddIn-AddIn64-packed.xll`
7. For 32-bit Excel, use `ActuarialAddIn-AddIn-packed.xll`

### Building from Source
```bash
# Restore dependencies
dotnet restore ActuarialAddIn.sln

# Build
dotnet build ActuarialAddIn.sln --configuration Release

# Output files will be in:
# src/ActuarialAddIn/bin/Release/net6.0-windows/publish/
```

## Examples

See `excel/actuarial_add_in_v0.2.xlsm` for working examples of all functions.

Each function category has its own worksheet:
- **Distributions** - PDF, CDF, Inverse CDF examples for all distributions
- **Exposure Curves** - MBBEFD, Swiss Re, Lloyd's curves
- **Reinsurance** - XOL layer calculations
- **Interpolation** - Linear interpolation with extrapolation options
- **Chain Ladder** - Development triangles and reserve calculations
- **Copulas** - Correlated random number generation
- **Return Periods** - EP curve interpolation and AAL

## Testing

Run the test suite to generate a markdown report:

```bash
dotnet run --project src/ActuarialAddIn.Tests/ActuarialAddIn.Tests.csproj -- test_results.md
```

## Requirements

- .NET 6.0 SDK (for building)
- Microsoft Excel (for using the add-in)
- Windows (Excel-DNA add-ins are Windows-only)

## Dependencies

- [Excel-DNA](https://excel-dna.net/) - Excel add-in framework
- [MathNet.Numerics](https://numerics.mathdotnet.com/) - Numerical computing library

## License

MIT
