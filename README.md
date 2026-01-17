# Actuarial Excel Add-In

A comprehensive Excel add-in for actuarial calculations, built with C# and Excel-DNA.

## Features

### Statistical Distributions
PDF, CDF, and Inverse CDF for:
- **Poisson** - `ACT_POISSON_PDF`, `ACT_POISSON_CDF`, `ACT_POISSON_INV`
- **Negative Binomial** - `ACT_NEGBIN_PDF`, `ACT_NEGBIN_CDF`, `ACT_NEGBIN_INV`
- **Lognormal** - `ACT_LOGNORM_PDF`, `ACT_LOGNORM_CDF`, `ACT_LOGNORM_INV`
- **Gamma** - `ACT_GAMMA_PDF`, `ACT_GAMMA_CDF`, `ACT_GAMMA_INV`
- **Pareto** - `ACT_PARETO_PDF`, `ACT_PARETO_CDF`, `ACT_PARETO_INV`

### Exposure Curves
- **MBBEFD/Swiss Re** - `ACT_MBBEFD`, `ACT_SWISSRE_CURVE`
- **Lloyd's** - `ACT_LLOYDS_CURVE`
- **Power** - `ACT_POWER_CURVE`, `ACT_INVERSE_POWER_CURVE`
- **Pareto** - `ACT_PARETO_EXPOSURE`
- **Layer Rating** - `ACT_LAYER_RATE_ON_LINE`

### Copulas
- **Student-t Copula** - `ACT_STUDENT_T_COPULA`, `ACT_STUDENT_T_COPULA_SINGLE`

### Chain Ladder Reserving
- **Development Factors** - `ACT_CL_FACTORS`
- **Ultimates** - `ACT_CL_ULTIMATE`
- **IBNR** - `ACT_CL_IBNR`
- **Mack Bootstrap** - `ACT_MACK_FACTOR_SE`, `ACT_MACK_RESERVE_SE`
- **Bootstrap** - `ACT_BOOTSTRAP_CL`
- **Berquist-Sherman** - `ACT_BERQUIST_SHERMAN`

### Reinsurance
- **Excess of Loss** - `ACT_XOL_LAYER_LOSS`, `ACT_XOL_EXPECTED_LOSS`
- **Quota Share** - `ACT_QS_CEDED`
- **Aggregate** - `ACT_AGGREGATE_LAYER`
- **ILF** - `ACT_ILF_PARETO`

### Return Periods
- **Loss Interpolation** - `ACT_RETURN_PERIOD_LOSS`
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

See `excel/actuarial_add_in_v0.0.xlsm` for working examples of all functions.

Each function category has its own worksheet:
- **Distributions** - PDF, CDF, Inverse CDF examples
- **Exposure Curves** - MBBEFD, Swiss Re, Lloyd's curves
- **Reinsurance** - XOL, quota share, aggregate layers
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
