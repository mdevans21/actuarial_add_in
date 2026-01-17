# Agent Instructions

This document contains instructions for AI agents working on this project.

## Excel Examples Consistency

**Important:** The examples in `excel/actuarial_add_in_v0.0.xlsm` must be kept consistent with the test suite in `src/ActuarialAddIn.Tests/Program.cs`.

When updating the test suite:
1. Update the corresponding Excel sheet(s) with the same test values
2. Ensure parameter values match exactly (e.g., lambda=5 for Poisson, r=5/p=0.3 for Negative Binomial)
3. Run `scripts/populate_examples.py` to regenerate examples, or update manually

When adding new functions:
1. Add tests to `src/ActuarialAddIn.Tests/Program.cs`
2. Add a new sheet or section to the Excel workbook with matching examples
3. Update `scripts/populate_examples.py` to include the new function

## Test Data Reference

The following test parameters are used consistently across the test suite and Excel examples:

| Distribution | Parameters |
|-------------|------------|
| Poisson | lambda=5 |
| Negative Binomial | r=5, p=0.3 |
| Lognormal | mu=0, sigma=1 |
| Gamma | alpha=2, beta=1 |
| Pareto | alpha=2, xm=1 |

| Exposure Curves | Parameters |
|----------------|------------|
| MBBEFD | b=2, g=3 |
| Swiss Re | curves 1-5 |
| Lloyd's | Y1-Y4 |

| Reinsurance | Parameters |
|-------------|------------|
| XOL Layer | attachment=1M, limit=5M |
| Aggregate | deductible=2M, limit=10M |

| Chain Ladder | Data |
|--------------|------|
| Triangle | 5x5 cumulative paid losses |

| Copula | Parameters |
|--------|------------|
| Student-t | df=5, 3x3 correlation matrix, seed=42 |
