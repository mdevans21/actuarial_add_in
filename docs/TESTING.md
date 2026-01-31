# Actuarial Add-In Testing Methodology

## Overview

This document describes the testing methodology for validating the Actuarial Add-In functions against multiple authoritative sources.

## Testing Architecture

### Three-Way Validation

Each function is validated against up to three sources:

```
┌─────────────────┐     ┌─────────────────┐     ┌─────────────────┐
│  Excel Formula  │     │ Python (scipy/  │     │ Online Reference│
│  (C# Add-In)    │────▶│  chainladder)   │◀────│ (E&V, Papers)  │
└─────────────────┘     └─────────────────┘     └─────────────────┘
         │                      │                        │
         └──────────────────────┼────────────────────────┘
                               ▼
                     ┌─────────────────┐
                     │ Validation Report│
                     │ (Markdown)      │
                     └─────────────────┘
```

### Fixture Files

Reference values are stored in JSON fixture files:

```
tests/fixtures/
├── distributions.json    # scipy expected values for all distributions
├── chain_ladder.json     # E&V (2002) reference values
├── copulas.json          # Copula CDF and tail dependence values
├── exposure_curves.json  # Bernegger (1997) and Swiss Re curves
├── credibility.json      # CAS exam and textbook reference values
└── sources.json          # URLs and citations for all sources
```

## Running Tests

### Unified Test Runner (Recommended)

Run all tests from WSL using the unified script:

```bash
# Run all tests (C# + Python + Report)
python scripts/run_tests.py

# Run only C# tests (generates test_results.md)
python scripts/run_tests.py --csharp

# Run only Python validation tests
python scripts/run_tests.py --python

# Generate validation report only
python scripts/run_tests.py --report
```

### Individual Commands

If you prefer to run tests individually:

```bash
# C# Tests (from repo root, requires Windows dotnet)
cd src/ActuarialAddIn.Tests && dotnet run

# Python Tests
python -m pytest tests/test_validation.py -v

# Validation Report
python tests/compare_sources.py --generate-report
```

Output files:
- `test_results.md` - C# test output with reconciliation tables
- `tests/reports/VALIDATION_REPORT.md` - Detailed validation report

## Validation Sources by Category

### Distributions

| Function | Primary Source | Secondary Source |
|----------|---------------|------------------|
| ACT_DIST_POISSON_* | scipy.stats.poisson | - |
| ACT_DIST_NEGBIN_* | scipy.stats.nbinom | R actuar |
| ACT_DIST_LOGNORM_* | scipy.stats.lognorm | - |
| ACT_DIST_GAMMA_* | scipy.stats.gamma | - |
| ACT_DIST_PARETO_* | scipy.stats.pareto | - |
| ACT_DIST_WEIBULL_* | scipy.stats.weibull_min | - |
| ACT_DIST_BETA_* | scipy.stats.beta | - |
| ACT_DIST_EXP_* | scipy.stats.expon | - |
| ACT_DIST_GPD_* | scipy.stats.genpareto | - |
| ACT_DIST_BURR_* | scipy.stats.burr12 | R actuar |

### Chain Ladder

| Function | Primary Source | Notes |
|----------|---------------|-------|
| ACT_CL_FACTORS | E&V (2002) / Taylor-Ashe | Standard benchmark |
| ACT_CL_ULTIMATE | E&V (2002) | Cumulative development |
| ACT_CL_IBNR | E&V (2002) | Total: 18,680,856 |
| ACT_MACK_RESERVE_SE | Mack (1993) | Distribution-free |
| ACT_CL_BOOTSTRAP | E&V (2002) Table 3 | Non-constant scale |

**Important**: chainladder-python uses constant scale and does NOT match E&V non-constant methodology. Our implementation targets E&V non-constant (97% match).

### Copulas

| Function | Primary Source | Secondary Source |
|----------|---------------|------------------|
| ACT_COPULA_GAUSSIAN | scipy multivariate_normal | McNeil et al (2015) |
| ACT_COPULA_STUDENT_T | scipy multivariate_t | McNeil et al (2015) |
| ACT_COPULA_CLAYTON | Analytical formula | R copula package |
| ACT_COPULA_FRANK | Analytical formula | R copula package |
| ACT_COPULA_GUMBEL | Analytical formula | R copula package |
| ACT_COPULA_TAIL_* | McNeil et al (2015) | Tail dependence formulas |

### Exposure Curves

| Function | Primary Source | Notes |
|----------|---------------|-------|
| ACT_EXPOSURE_MBBEFD | Bernegger (1997) | ASTIN Bulletin |
| ACT_EXPOSURE_SWISSRE | Swiss Re Sigma | Industry standard |
| ACT_EXPOSURE_LLOYDS | Lloyd's Market Bulletins | Y1-Y4 curves |
| ACT_EXPOSURE_POWER | Analytical | G(d) = d^n |
| ACT_EXPOSURE_PARETO | Analytical | Based on Pareto severity |

### Credibility

| Function | Primary Source | Notes |
|----------|---------------|-------|
| ACT_CREDIBILITY_BUHLMANN | CAS Exam 5 | Z = n/(n+k) |
| ACT_CREDIBILITY_K | Bühlmann & Gisler (2005) | K-parameter estimation |
| ACT_FULL_CREDIBILITY_STANDARD | CAS Exam 5 | k = (z/r)² |
| ACT_EXPERIENCE_MOD | NCCI manuals | Workers comp |
| ACT_SCHEDULE_RATING | ISO manual | Commercial lines |

## Adding New Validation Tests

### 1. Add Reference Value to Fixture

```json
// In tests/fixtures/distributions.json
"new_distribution": {
  "source_url": "https://...",
  "parameters": {"param1": 1.0},
  "test_values": {
    "pdf_x1": {"x": 1.0, "expected": 0.5, "tolerance": 1e-10}
  }
}
```

### 2. Add Validator Method

```python
# In tests/compare_sources.py
def validate_new_distribution(self) -> List[ValidationResult]:
    data = self.fixtures.get("new_distribution", {})
    # ... validation logic
```

### 3. Add Pytest Test

```python
# In tests/test_validation.py
def test_new_distribution(self, distributions_fixture):
    data = distributions_fixture["new_distribution"]
    # ... test assertions
```

## Tolerance Guidelines

| Value Type | Typical Tolerance | Notes |
|------------|-------------------|-------|
| Probability (PDF/CDF) | 1e-10 | High precision required |
| Quantile (Inverse) | 1e-8 | Numerical inversion |
| Integer (counts) | 0 | Exact match |
| Reserve estimates | 1000-10000 | Depends on triangle size |
| Bootstrap SE | 5-10% | Stochastic variation |

## Continuous Integration

### Recommended CI Pipeline

```yaml
# .github/workflows/validate.yml
name: Validation

on: [push, pull_request]

jobs:
  validate:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-python@v4
        with:
          python-version: '3.10'
      - run: pip install scipy numpy
      - run: python tests/compare_sources.py --generate-report
      - run: python -m pytest tests/test_validation.py -v
      - uses: actions/upload-artifact@v3
        with:
          name: validation-report
          path: tests/reports/VALIDATION_REPORT.md
```

## References

1. England, P.D. and Verrall, R.J. (2002). "Stochastic claims reserving in general insurance." British Actuarial Journal 8(3): 443-518.
2. Mack, T. (1993). "Distribution-free calculation of the standard error of chain ladder reserve estimates." ASTIN Bulletin 23(2): 213-225.
3. Bernegger, S. (1997). "The Swiss Re Exposure Curves and the MBBEFD Distribution Class." ASTIN Bulletin 27(1): 99-111.
4. McNeil, A.J., Frey, R., and Embrechts, P. (2015). "Quantitative Risk Management." Princeton University Press.
5. Bühlmann, H. and Gisler, A. (2005). "A Course in Credibility Theory and its Applications." Springer.
6. scipy.stats documentation: https://docs.scipy.org/doc/scipy/reference/stats.html
7. chainladder-python documentation: https://chainladder-python.readthedocs.io/
