# Plan: Implement LEV Functions for All Distributions

## Overview

Add Limited Expected Value (LEV) functions for all 10 distributions in the add-in. LEV is defined as E[min(X, d)] - the expected value of X capped at limit d.

## Key Sources

1. **R actuar package** - Primary reference for formulas (Klugman, Panjer & Willmot "Loss Models")
2. **GEMAct Python package** - For validation benchmarks
3. **scipy.stats** - Numerical integration fallback via `expect()`

## Distributions and Formulas

### Closed-Form (Direct Implementation)

| Distribution | Formula | Notes |
|--------------|---------|-------|
| **Exponential** | `LEV(d) = (1/λ)[1 - e^(-λd)]` | Simple closed form |
| **Pareto** | `LEV(d) = αθ/(α-1) × [1 - (θ/(d+θ))^(α-1)]` for α≠1 | Already in Reinsurance.cs |
| **Weibull** | `LEV(d) = λΓ(1+1/k) × γ(1+1/k, (d/λ)^k) / Γ(1+1/k)` | Uses incomplete gamma |
| **Gamma** | `LEV(d) = αβ × P(α+1, d/β) + d × [1 - P(α, d/β)]` | P is regularized incomplete gamma |
| **Lognormal** | `LEV(d) = e^(μ+σ²/2) × Φ((ln(d)-μ-σ²)/σ) + d × [1 - Φ((ln(d)-μ)/σ)]` | Uses normal CDF |
| **GPD** | `LEV(d) = σ/(1+ξ) × [1 - (1 + ξd/σ)^(-1/ξ-1)]` for ξ≠0 | Special case ξ=0 is exponential |
| **Beta** | `LEV(d) = d - ∫₀^d F(x)dx = d×(1-I_d(α,β)) + α/(α+β) × I_d(α+1,β)` | I is regularized incomplete beta |
| **Burr XII** | `LEV(d) = c×λ^(1/c) × B(k-1/c, 1+1/c) × [1 - I(k-1/c, 1+1/c; 1/(1+(d/λ)^c))]` | Uses incomplete beta |

### Discrete Distributions (Summation)

| Distribution | Method |
|--------------|--------|
| **Poisson** | `LEV(d) = Σ_{k=0}^{floor(d)-1} P(X≥k)` or `= Σ_{k=0}^{floor(d)} k×P(k) + d×P(X>d)` |
| **Negative Binomial** | Same summation approach |

## Implementation Plan

### File: `src/ActuarialAddIn/Functions/Distributions.cs`

Add LEV functions alongside existing PDF/CDF/INV functions:

```
ACT_DIST_EXP_LEV(limit, lambda)
ACT_DIST_PARETO_LEV(limit, alpha, xm)
ACT_DIST_LOGNORM_LEV(limit, mu, sigma)
ACT_DIST_GAMMA_LEV(limit, alpha, beta)
ACT_DIST_WEIBULL_LEV(limit, k, lambda)
ACT_DIST_GPD_LEV(limit, xi, sigma)
ACT_DIST_BETA_LEV(limit, alpha, beta)
ACT_DIST_BURR_LEV(limit, c, k, lambda)
ACT_DIST_POISSON_LEV(limit, lambda)
ACT_DIST_NEGBIN_LEV(limit, r, p)
```

### Implementation Order (by complexity)

1. **Exponential** - Simplest closed form
2. **Pareto** - Already have formula in Reinsurance.cs (refactor to shared helper)
3. **GPD** - Similar to Pareto structure
4. **Gamma** - Uses MathNet.Numerics incomplete gamma
5. **Lognormal** - Uses normal CDF (already available)
6. **Weibull** - Uses incomplete gamma
7. **Beta** - Uses incomplete beta
8. **Burr XII** - Most complex, uses incomplete beta
9. **Poisson** - Direct summation
10. **Negative Binomial** - Direct summation

### MathNet.Numerics Functions Needed

- `SpecialFunctions.GammaLowerRegularized(a, x)` - for Gamma, Weibull
- `SpecialFunctions.BetaRegularized(a, b, x)` - for Beta, Burr
- Already using `Normal.CDF()` - for Lognormal

## Validation

### Python Benchmarks

Add to `tests/fixtures/distributions.json`:
```json
{
  "lev": {
    "exponential": {"lambda": 2, "limit": 1, "expected": 0.4323},
    "pareto": {"alpha": 2, "xm": 1, "limit": 2, "expected": 1.333},
    ...
  }
}
```

Generate expected values using:
```python
from scipy.stats import expon, pareto, lognorm, gamma, weibull_min, genpareto, beta
from scipy.integrate import quad

def lev(dist, limit):
    """LEV via numerical integration: integral of survival function"""
    return quad(lambda x: 1 - dist.cdf(x), 0, limit)[0]
```

### Test Cases

Add to `src/ActuarialAddIn.Tests/Program.cs`:
- LEV at various limits for each distribution
- Edge cases: limit=0, limit=infinity, limit < 0

### Excel Examples

Add LEV section to `excel/actuarial_add_in.xlsm`:
- Table showing LEV(limit) for each distribution
- Chart showing LEV curve vs limit

## Files to Modify

1. `src/ActuarialAddIn/Functions/Distributions.cs` - Add 10 LEV functions
2. `src/ActuarialAddIn.Tests/Program.cs` - Add LEV tests
3. `tests/fixtures/distributions.json` - Add LEV expected values
4. `tests/compare_sources.py` - Add LEV validators
5. `scripts/populate_examples.py` - Add LEV examples to Excel
6. `README.md` - Document new functions
7. `agents.md` - Add LEV test parameters

## Verification

1. Run C# tests: `dotnet run --project src/ActuarialAddIn.Tests`
2. Run Python validation: `python tests/compare_sources.py --generate-report`
3. Rebuild add-in and verify in Excel
4. Check all LEV functions return sensible values (positive, monotonic in limit)
