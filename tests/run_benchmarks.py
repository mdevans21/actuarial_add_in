#!/usr/bin/env python3
"""
Automated benchmark tests for the Actuarial Add-In.
Compares add-in function outputs against scipy and chainladder.

Usage:
    python run_benchmarks.py
"""

import numpy as np
import pandas as pd
from scipy import stats
import chainladder as cl
import sys
from dataclasses import dataclass
from typing import List, Tuple

@dataclass
class TestResult:
    name: str
    passed: bool
    max_diff: float
    tolerance: float
    details: str = ""

def test_poisson_distribution() -> TestResult:
    """Test Poisson distribution against scipy."""
    lambda_param = 5
    k_values = list(range(11))

    # Expected values from add-in
    addin_pmf = [0.006738, 0.033690, 0.084224, 0.140374, 0.175467,
                 0.175467, 0.146223, 0.104445, 0.065278, 0.036266, 0.018133]

    scipy_pmf = [stats.poisson.pmf(k, lambda_param) for k in k_values]
    max_diff = max(abs(a - s) for a, s in zip(addin_pmf, scipy_pmf))

    tolerance = 1e-5
    return TestResult(
        name="Poisson Distribution (λ=5)",
        passed=max_diff < tolerance,
        max_diff=max_diff,
        tolerance=tolerance
    )

def test_negbin_distribution() -> TestResult:
    """Test Negative Binomial distribution against scipy."""
    r, p = 5, 0.3
    k_values = list(range(11))

    # Expected values from add-in
    addin_pmf = [0.002430, 0.008505, 0.017860, 0.029172, 0.040841,
                 0.051460, 0.060036, 0.066040, 0.069342, 0.070112, 0.068710]

    scipy_pmf = [stats.nbinom.pmf(k, r, p) for k in k_values]
    max_diff = max(abs(a - s) for a, s in zip(addin_pmf, scipy_pmf))

    tolerance = 1e-5
    return TestResult(
        name="Negative Binomial (r=5, p=0.3)",
        passed=max_diff < tolerance,
        max_diff=max_diff,
        tolerance=tolerance
    )

def test_lognormal_distribution() -> TestResult:
    """Test Lognormal distribution against scipy."""
    mu, sigma = 0, 1
    x_values = [0.5, 1, 1.5, 2, 3, 5]

    # Expected values from add-in
    addin_pdf = [0.627496, 0.398942, 0.244974, 0.156874, 0.072728, 0.021851]

    scipy_pdf = [stats.lognorm.pdf(x, sigma, scale=np.exp(mu)) for x in x_values]
    max_diff = max(abs(a - s) for a, s in zip(addin_pdf, scipy_pdf))

    tolerance = 1e-5
    return TestResult(
        name="Lognormal (μ=0, σ=1)",
        passed=max_diff < tolerance,
        max_diff=max_diff,
        tolerance=tolerance
    )

def test_gamma_distribution() -> TestResult:
    """Test Gamma distribution against scipy."""
    alpha, beta = 2, 1
    x_values = [0.5, 1, 2, 3, 5]

    # Expected values from add-in
    addin_pdf = [0.303265, 0.367879, 0.270671, 0.149361, 0.033690]

    scipy_pdf = [stats.gamma.pdf(x, alpha, scale=1/beta) for x in x_values]
    max_diff = max(abs(a - s) for a, s in zip(addin_pdf, scipy_pdf))

    tolerance = 1e-5
    return TestResult(
        name="Gamma (α=2, β=1)",
        passed=max_diff < tolerance,
        max_diff=max_diff,
        tolerance=tolerance
    )

def test_pareto_distribution() -> TestResult:
    """Test Pareto distribution against scipy."""
    alpha, xm = 2, 1
    x_values = [1, 1.5, 2, 3, 5, 10]

    # Expected values from add-in
    addin_pdf = [2.000000, 0.592593, 0.250000, 0.074074, 0.016000, 0.002000]

    scipy_pdf = [stats.pareto.pdf(x/xm, alpha) / xm for x in x_values]
    max_diff = max(abs(a - s) for a, s in zip(addin_pdf, scipy_pdf))

    tolerance = 1e-5
    return TestResult(
        name="Pareto (α=2, xm=1)",
        passed=max_diff < tolerance,
        max_diff=max_diff,
        tolerance=tolerance
    )

def test_chain_ladder_factors() -> TestResult:
    """Test chain ladder development factors against chainladder."""
    triangle_data = np.array([
        [100, 150, 170, 180, 185],
        [110, 165, 190, 200, np.nan],
        [120, 180, 210, np.nan, np.nan],
        [130, 195, np.nan, np.nan, np.nan],
        [140, np.nan, np.nan, np.nan, np.nan]
    ])

    triangle = cl.Triangle(
        pd.DataFrame(triangle_data,
                     index=[2019, 2020, 2021, 2022, 2023],
                     columns=[1, 2, 3, 4, 5]),
        origin='index',
        development='columns',
        cumulative=True
    )

    dev = cl.Development().fit(triangle)
    cl_factors = dev.ldf_.values.flatten()[:4]

    # Expected from add-in
    addin_factors = np.array([1.5000, 1.1515, 1.0556, 1.0278])

    max_diff = np.max(np.abs(cl_factors - addin_factors))
    tolerance = 0.001  # 0.1% tolerance for factors

    return TestResult(
        name="Chain Ladder Development Factors",
        passed=max_diff < tolerance,
        max_diff=max_diff,
        tolerance=tolerance,
        details=f"CL: {cl_factors.round(4)}, Add-in: {addin_factors}"
    )

def test_chain_ladder_ultimate() -> TestResult:
    """Test chain ladder ultimates against chainladder."""
    triangle_data = np.array([
        [100, 150, 170, 180, 185],
        [110, 165, 190, 200, np.nan],
        [120, 180, 210, np.nan, np.nan],
        [130, 195, np.nan, np.nan, np.nan],
        [140, np.nan, np.nan, np.nan, np.nan]
    ])

    triangle = cl.Triangle(
        pd.DataFrame(triangle_data,
                     index=[2019, 2020, 2021, 2022, 2023],
                     columns=[1, 2, 3, 4, 5]),
        origin='index',
        development='columns',
        cumulative=True
    )

    dev = cl.Development().fit(triangle)
    cl_model = cl.Chainladder().fit(dev.transform(triangle))
    cl_ultimates = cl_model.ultimate_.values.flatten()

    # Expected from add-in
    addin_ultimates = np.array([185.00, 205.56, 227.82, 243.60, 262.34])

    max_diff = np.max(np.abs(cl_ultimates - addin_ultimates))
    tolerance = 0.5  # Allow 0.5 difference due to rounding

    return TestResult(
        name="Chain Ladder Ultimates",
        passed=max_diff < tolerance,
        max_diff=max_diff,
        tolerance=tolerance,
        details=f"CL: {cl_ultimates.round(2)}, Add-in: {addin_ultimates}"
    )

def test_mbbefd_curve() -> TestResult:
    """Test MBBEFD exposure curve."""
    def mbbefd_curve(d, b, g):
        if d <= 0: return 0
        if d >= 1: return 1
        if b == 1 and g == 1: return d
        if b == 1: return np.log(1 + (g - 1) * d) / np.log(g)
        if g == 1: return (1 - b**d) / (1 - b)
        return np.log((g - 1 + b**d) / g) / np.log(b)

    b, g = 2, 3
    d_values = np.arange(0, 1.1, 0.1)

    python_values = [mbbefd_curve(d, b, g) for d in d_values]

    # Expected from add-in
    addin_values = [0.000000, 0.149001, 0.278578, 0.394114, 0.499046,
                    0.595704, 0.685740, 0.770368, 0.850505, 0.926862, 1.000000]

    max_diff = max(abs(p - a) for p, a in zip(python_values, addin_values))
    tolerance = 1e-5

    return TestResult(
        name="MBBEFD Curve (b=2, g=3)",
        passed=max_diff < tolerance,
        max_diff=max_diff,
        tolerance=tolerance
    )

def test_taylor_ashe() -> TestResult:
    """Test against Taylor-Ashe sample data."""
    ta = cl.load_sample('taylor_ashe')
    ta_dev = cl.Development().fit(ta)
    ta_cl = cl.Chainladder().fit(ta_dev.transform(ta))

    total_ibnr = float(ta_cl.ibnr_.sum())

    # Expected total IBNR from literature: approximately 18,680,856
    expected_ibnr = 18680856
    diff_pct = abs(total_ibnr - expected_ibnr) / expected_ibnr * 100

    tolerance = 1.0  # 1% tolerance

    return TestResult(
        name="Taylor-Ashe Total IBNR",
        passed=diff_pct < tolerance,
        max_diff=diff_pct,
        tolerance=tolerance,
        details=f"Calculated: {total_ibnr:,.0f}, Expected: {expected_ibnr:,}"
    )

def run_all_tests() -> List[TestResult]:
    """Run all benchmark tests."""
    tests = [
        test_poisson_distribution,
        test_negbin_distribution,
        test_lognormal_distribution,
        test_gamma_distribution,
        test_pareto_distribution,
        test_chain_ladder_factors,
        test_chain_ladder_ultimate,
        test_mbbefd_curve,
        test_taylor_ashe,
    ]

    results = []
    for test_func in tests:
        try:
            result = test_func()
            results.append(result)
        except Exception as e:
            results.append(TestResult(
                name=test_func.__name__,
                passed=False,
                max_diff=float('inf'),
                tolerance=0,
                details=f"Error: {str(e)}"
            ))

    return results

def main():
    print("=" * 70)
    print("ACTUARIAL ADD-IN BENCHMARK TESTS")
    print("=" * 70)
    print()
    print(f"chainladder version: {cl.__version__}")
    print()

    results = run_all_tests()

    passed = sum(1 for r in results if r.passed)
    total = len(results)

    print("-" * 70)
    print(f"{'Test Name':<40} {'Status':<10} {'Max Diff':<15}")
    print("-" * 70)

    for result in results:
        status = "PASS" if result.passed else "FAIL"
        status_color = "" if result.passed else ""
        print(f"{result.name:<40} {status:<10} {result.max_diff:.2e}")
        if result.details and not result.passed:
            print(f"  Details: {result.details}")

    print("-" * 70)
    print(f"Results: {passed}/{total} tests passed")
    print("=" * 70)

    return 0 if passed == total else 1

if __name__ == "__main__":
    sys.exit(main())
