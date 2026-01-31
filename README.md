# Actuarial Excel Add-In

A comprehensive Excel add-in for general insurance actuarial calculations, built with C# and Excel-DNA.

The aim of this add-in is to bring models normally restricted to the statistical worlds of R and python to 
the more user-accessible Microsoft Excel.

This code is in beta and offered without warranty. Please raise bug reports and feature requests through GitHub.

## Quick Start

1. Download the latest `.xll` file from [Releases](../../releases)
2. Open Excel → File → Options → Add-ins
3. Select "Excel Add-ins" → Go → Browse
4. Select `ActuarialAddIn-AddIn64-packed.xll` (64-bit) or `ActuarialAddIn-AddIn-packed.xll` (32-bit)

See `excel/actuarial_add_in_v0.2.xlsm` for working examples of all functions.

---

## Features Overview

| Category      | Functions | Description                              |
|---------------|-----------|------------------------------------------|
| Distributions | 30+       | PDF, CDF, Inverse for 10 distributions   |
| Exposure Curves | 10+     | MBBEFD, Swiss Re, Lloyd's, Power, Pareto |
| Chain Ladder  | 15+       | Factors, Ultimates, Mack SE, ODP Bootstrap |
| Reinsurance   | 5+        | XOL, ILF, Return Periods                 |
| Copulas       | 10+       | Gaussian, Student-t, Clayton, Frank, Gumbel |
| Interpolation | 5+        | Linear, Log-linear, 2D Bilinear          |

---

## 1. Statistical Distributions

PDF, CDF, and Inverse CDF functions for actuarial modeling.

### Test Parameters (consistent across test suite and Excel examples)

| Distribution      | Parameters  | Mean  | Variance |
|-------------------|-------------|-------|----------|
| Poisson           | λ=5         | 5     | 5        |
| Negative Binomial | r=5, p=0.3  | 11.67 | 38.89    |
| Lognormal         | μ=0, σ=1    | 1.649 | 4.671    |
| Gamma             | α=2, β=1    | 2     | 2        |
| Pareto            | α=2, xm=1   | 2     | ∞        |

### Functions

| Distribution      | PDF                                      | CDF                                      | Inverse                                    |
|-------------------|------------------------------------------|------------------------------------------|--------------------------------------------|
| Poisson           | `ACT_DIST_POISSON_PDF(k, lambda)`        | `ACT_DIST_POISSON_CDF(k, lambda)`        | `ACT_DIST_POISSON_INV(p, lambda)`          |
| Negative Binomial | `ACT_DIST_NEGBIN_PDF(k, r, p)`           | `ACT_DIST_NEGBIN_CDF(k, r, p)`           | `ACT_DIST_NEGBIN_INV(prob, r, p)`          |
| Lognormal         | `ACT_DIST_LOGNORM_PDF(x, mu, sigma)`     | `ACT_DIST_LOGNORM_CDF(x, mu, sigma)`     | `ACT_DIST_LOGNORM_INV(p, mu, sigma)`       |
| Gamma             | `ACT_DIST_GAMMA_PDF(x, alpha, beta)`     | `ACT_DIST_GAMMA_CDF(x, alpha, beta)`     | `ACT_DIST_GAMMA_INV(p, alpha, beta)`       |
| Pareto            | `ACT_DIST_PARETO_PDF(x, alpha, xm)`      | `ACT_DIST_PARETO_CDF(x, alpha, xm)`      | `ACT_DIST_PARETO_INV(p, alpha, xm)`        |
| GPD               | `ACT_DIST_GPD_PDF(x, xi, sigma)`         | `ACT_DIST_GPD_CDF(x, xi, sigma)`         | `ACT_DIST_GPD_INV(p, xi, sigma)`           |
| Weibull           | `ACT_DIST_WEIBULL_PDF(x, k, lambda)`     | `ACT_DIST_WEIBULL_CDF(x, k, lambda)`     | `ACT_DIST_WEIBULL_INV(p, k, lambda)`       |
| Beta              | `ACT_DIST_BETA_PDF(x, alpha, beta)`      | `ACT_DIST_BETA_CDF(x, alpha, beta)`      | `ACT_DIST_BETA_INV(p, alpha, beta)`        |
| Exponential       | `ACT_DIST_EXP_PDF(x, lambda)`            | `ACT_DIST_EXP_CDF(x, lambda)`            | `ACT_DIST_EXP_INV(p, lambda)`              |
| Burr Type XII     | `ACT_DIST_BURR_PDF(x, c, k, lambda)`     | `ACT_DIST_BURR_CDF(x, c, k, lambda)`     | `ACT_DIST_BURR_INV(p, c, k, lambda)`       |

---

## 2. Exposure Curves

First loss scales and exposure rating curves for reinsurance pricing.

### Test Parameters

| Curve Type       | Parameters | G(0.5) |
|------------------|------------|--------|
| MBBEFD           | b=2, g=3   | 0.5957 |
| Swiss Re Curve 1 | -          | 0.5251 |
| Swiss Re Curve 3 | -          | 0.7139 |
| Lloyd's Y1       | -          | 0.6464 |
| Lloyd's Y4       | -          | 0.9375 |

### Functions

| Function                                              | Description                    | Parameters                              |
|-------------------------------------------------------|--------------------------------|-----------------------------------------|
| `ACT_EXPOSURE_MBBEFD(d, b, g)`                        | MBBEFD exposure curve          | d: damage ratio, b,g: shape             |
| `ACT_EXPOSURE_SWISSRE(d, curve)`                      | Swiss Re standard curves       | curve: 1-5                              |
| `ACT_EXPOSURE_LLOYDS(d, curve)`                       | Lloyd's standard curves        | curve: Y1-Y4                            |
| `ACT_EXPOSURE_POWER(d, n)`                            | Power curve G(d) = d^n         | n: exponent                             |
| `ACT_EXPOSURE_INVERSE_POWER(G, n)`                    | Inverse power curve            | G: exposure, n: exponent                |
| `ACT_EXPOSURE_PARETO(d, alpha)`                       | Pareto exposure curve          | alpha: tail parameter                   |
| `ACT_EXPOSURE_RIEBESELL(d, c)`                        | Riebesell curve                | c: parameter                            |
| `ACT_EXPOSURE_LAYER_RATE(xs, xl, M, curve_type, params)` | Layer rate from exposure curve | xs: attachment, xl: limit, M: EPI    |

---

## 3. Chain Ladder Reserving

Comprehensive reserving functions including deterministic and stochastic methods.

### Reference Data: Taylor-Ashe Triangle

All Chain Ladder examples use the Taylor-Ashe 10x10 triangle:
- **Source**: Taylor & Ashe (1983), England & Verrall (2002)
- **Total IBNR**: 18,680,856
- **Development Factors**: 3.491, 1.747, 1.457, 1.174, 1.104, 1.086, 1.054, 1.077, 1.018

### Deterministic Functions

| Function                                                            | Description                       | Output          |
|---------------------------------------------------------------------|-----------------------------------|-----------------|
| `ACT_CL_FACTORS(triangle, [n_periods], [exclude_recent], [exclude_high_low])` | Development factors     | Array of factors |
| `ACT_CL_ULTIMATE(triangle)`                                         | Projected ultimates               | Array by origin |
| `ACT_CL_IBNR(triangle)`                                             | IBNR reserves                     | Array by origin |
| `ACT_CL_LATEST(triangle)`                                           | Latest diagonal values            | Array by origin |
| `ACT_TRIANGLE_TO_INCREMENTAL(triangle)`                             | Convert cumulative to incremental | Matrix          |
| `ACT_INCREMENTAL_TO_CUMULATIVE(triangle)`                           | Convert incremental to cumulative | Matrix          |
| `ACT_TRIANGLE_LINK_RATIOS(triangle)`                                | Individual link ratios            | Matrix          |

### Mack Chain Ladder

Distribution-free method for measuring reserve uncertainty.

| Function                    | Description                        | Reference                        |
|-----------------------------|------------------------------------|----------------------------------|
| `ACT_MACK_FACTOR_SE(triangle)` | Standard error of development factors | Mack (1993)                   |
| `ACT_MACK_RESERVE_SE(triangle)` | Reserve standard error by origin  | E&V Table: Total SE = 2,447,095 |

### ODP Bootstrap (England & Verrall 2002)

**Implementation**: Non-constant scale parameter ODP Bootstrap

Key features:
- Period-specific scale parameters (φⱼ) instead of single global φ
- Stratified residual sampling by development period
- Gamma process variance with period-specific φⱼ

| Function                                           | Description                   | Parameters                     |
|----------------------------------------------------|-------------------------------|--------------------------------|
| `ACT_CL_BOOTSTRAP(triangle, iterations, [seed])`   | Total reserve distribution    | Returns percentiles + mean/std |
| `ACT_CL_BOOTSTRAP_ORIGIN(triangle, iterations, [seed])` | Reserve distribution by origin year | Returns stats by origin  |

### Bootstrap Benchmark Results (Taylor-Ashe, 10,000 iterations)

| AY        | IBNR           | Our SE    | E&V Non-Const | Ratio |
|-----------|----------------|-----------|---------------|-------|
| 1         | 0              | 0         | 0             | -     |
| 2         | 94,634         | 35,021    | 43,882        | 80%   |
| 3         | 469,511        | 98,084    | 109,449       | 90%   |
| 4         | 709,638        | 117,269   | 141,509       | 83%   |
| 5         | 984,889        | 247,779   | 256,031       | 97%   |
| 6         | 1,419,459      | 397,985   | 398,377       | 100%  |
| 7         | 2,177,641      | 541,311   | 529,898       | 102%  |
| 8         | 3,920,301      | 828,128   | 735,245       | 113%  |
| 9         | 4,278,972      | 933,906   | 809,457       | 115%  |
| 10        | 4,625,811      | 1,273,483 | 1,285,560     | 99%   |
| **Total** | **18,680,856** | **2,161,659** | **2,228,677** | **97%** |

**Reference**: England, P.D. and Verrall, R.J. (2002) "Stochastic Claims Reserving in General Insurance", British Actuarial Journal, 8(3), 443-518.

### Other Reserving Methods

| Function                                       | Description                        |
|------------------------------------------------|------------------------------------|
| `ACT_BF_ULTIMATE(triangle, factors, a_priori)` | Bornhuetter-Ferguson ultimates     |
| `ACT_CAPECOD_ULTIMATE(triangle, premiums)`     | Cape Cod method ultimates          |
| `ACT_CAPECOD_ELR(triangle, premiums)`          | Cape Cod expected loss ratio       |
| `ACT_BERQUIST_SHERMAN(paid, case_reserves, trend)` | Case reserve adequacy adjustment |

---

## 4. Reinsurance Functions

### Excess of Loss

| Function                                                  | Description               |
|-----------------------------------------------------------|---------------------------|
| `ACT_XOL_LAYER_LOSS(ground_up_loss, attachment, limit)`   | Calculate layer loss      |
| `ACT_XOL_EXPECTED_LOSS(mean, cv, attachment, limit, [distribution])` | Expected layer loss |
| `ACT_ILF_PARETO(limit, alpha)`                            | Increased limit factor (Pareto) |

### Return Periods and EP Curves

| Function                                                 | Description              |
|----------------------------------------------------------|--------------------------|
| `ACT_RETURN_PERIOD_LOSS(return_periods, losses, target_rp, [method])` | Interpolate EP curve |
| `ACT_RETURN_PERIOD_TABLE(return_periods, losses, targets, [method])` | Generate RP table   |
| `ACT_AAL_FROM_OEP(return_periods, oep_losses)`           | Calculate AAL from OEP curve |

---

## 5. Copulas

Generate correlated random samples for simulation models.

### Test Parameters
- Correlation matrix: 7x7 with ρ=0.6 decay
- Degrees of freedom: 5 (Student-t)
- Seed: 42 for reproducibility

### Functions

| Function                                       | Description                    |
|------------------------------------------------|--------------------------------|
| `ACT_COPULA_GAUSSIAN(corr_matrix, n_samples, [seed])` | Gaussian copula samples   |
| `ACT_COPULA_STUDENT_T(corr_matrix, df, n_samples, [seed])` | Student-t copula samples |
| `ACT_COPULA_CLAYTON(u1, u2, theta)`            | Clayton copula CDF             |
| `ACT_COPULA_FRANK(u1, u2, theta)`              | Frank copula CDF               |
| `ACT_COPULA_GUMBEL(u1, u2, theta)`             | Gumbel copula CDF              |
| `ACT_COPULA_TAU_TO_THETA(tau, copula_type)`    | Convert Kendall's tau to theta |
| `ACT_COPULA_TAIL_LOWER(theta, copula_type)`    | Lower tail dependence          |
| `ACT_COPULA_TAIL_UPPER(theta, copula_type)`    | Upper tail dependence          |

---

## 6. Interpolation

### Functions

| Function                                     | Description               | Extrapolation Options   |
|----------------------------------------------|---------------------------|-------------------------|
| `ACT_INTERP(x_vals, y_vals, x, [extrap])`    | Linear interpolation      | FLAT, GRADIENT, ERROR   |
| `ACT_INTERP_LOG(x_vals, y_vals, x, [extrap])` | Log-linear interpolation | FLAT, GRADIENT, ERROR   |
| `ACT_INTERP2D(x_vals, y_vals, z_matrix, x, y)` | Bilinear 2D interpolation | -                     |

---

## Testing

### Python Benchmark Tests

Compare against scipy and chainladder:

```bash
cd tests
pip install -r requirements.txt
python run_benchmarks.py
```

Output includes:
- Distribution function comparisons vs scipy
- Chain ladder comparisons vs chainladder package
- ODP Bootstrap comparison vs E&V (2002) published results

### C# Test Suite

Generate markdown test report:

```bash
dotnet run --project src/ActuarialAddIn.Tests/ActuarialAddIn.Tests.csproj -- test_results.md
```

---

## Building from Source

```bash
# Restore and build
dotnet restore ActuarialAddIn.sln
dotnet build ActuarialAddIn.sln --configuration Release

# Output files:
# src/ActuarialAddIn/bin/Release/net6.0-windows/publish/ActuarialAddIn-AddIn64-packed.xll
```

---

## Requirements

- .NET 6.0 SDK (for building)
- Microsoft Excel 2016+ (for using the add-in)
- Windows (Excel-DNA add-ins are Windows-only)

## Dependencies

- [Excel-DNA](https://excel-dna.net/) - Excel add-in framework
- [MathNet.Numerics](https://numerics.mathdotnet.com/) - Numerical computing library

## References

- England, P.D. and Verrall, R.J. (2002) "Stochastic Claims Reserving in General Insurance", British Actuarial Journal, 8(3), 443-518.
- Mack, T. (1993) "Distribution-free calculation of the standard error of chain ladder reserve estimates", ASTIN Bulletin, 23(2), 213-225.
- Taylor, G. and Ashe, F.R. (1983) "Second moments of estimates of outstanding claims", Journal of Econometrics, 23, 37-61.

## License

MIT
