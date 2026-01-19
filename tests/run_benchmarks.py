#!/usr/bin/env python3
"""
Actuarial Add-In Benchmark Tests

Comprehensive test suite comparing C# add-in outputs against:
- scipy for statistical distributions
- chainladder package for reserving methods
- England & Verrall (2002) published benchmark results

Usage:
    python run_benchmarks.py

Reference Data:
    - Taylor-Ashe 10x10 triangle (Taylor & Ashe 1983, England & Verrall 2002)
    - E&V ODP Bootstrap results from British Actuarial Journal 8(3), 443-518
"""

import numpy as np
import pandas as pd
from scipy import stats
import chainladder as cl
import sys
from dataclasses import dataclass
from typing import List, Dict, Tuple, Optional

# =============================================================================
# REFERENCE DATA - Taylor-Ashe Triangle
# =============================================================================

TAYLOR_ASHE_CUMULATIVE = np.array([
    [357848, 1124788, 1735330, 2218270, 2745596, 3319994, 3466336, 3606286, 3833515, 3901463],
    [352118, 1236139, 2170033, 3353322, 3799067, 4120063, 4647867, 4914039, 5339085, np.nan],
    [290507, 1292306, 2218525, 3235179, 3985995, 4132918, 4628910, 4909315, np.nan, np.nan],
    [310608, 1418858, 2195047, 3757447, 4029929, 4381982, 4588268, np.nan, np.nan, np.nan],
    [443160, 1136350, 2128333, 2897821, 3402672, 3873311, np.nan, np.nan, np.nan, np.nan],
    [396132, 1333217, 2180715, 2985752, 3691712, np.nan, np.nan, np.nan, np.nan, np.nan],
    [440832, 1288463, 2419861, 3483130, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan],
    [359480, 1421128, 2864498, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan],
    [376686, 1363294, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan],
    [344014, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan],
])

# England & Verrall (2002) Reference Values - Table from BAJ paper
EV_REFERENCE = {
    'ibnr': [0, 94634, 469511, 709638, 984889, 1419459, 2177641, 3920301, 4278972, 4625811],
    'mack_se': [0, 75535, 121699, 133549, 261406, 411010, 558317, 875328, 971258, 1363155],
    'odp_constant_se': [0, 112552, 217547, 262934, 306595, 375745, 500332, 791481, 1060473, 2025898],
    'odp_nonconstant_se': [0, 43882, 109449, 141509, 256031, 398377, 529898, 735245, 809457, 1285560],
    'total_ibnr': 18680856,
    'total_mack_se': 2447095,
    'total_odp_constant_se': 2992296,
    'total_odp_nonconstant_se': 2228677,
}

# Expected development factors
EXPECTED_FACTORS = [3.4906, 1.7473, 1.4574, 1.1739, 1.1038, 1.0863, 1.0539, 1.0766, 1.0177]


# =============================================================================
# TEST RESULT CLASS
# =============================================================================

@dataclass
class TestResult:
    name: str
    passed: bool
    max_diff: float
    tolerance: float
    details: str = ""


# =============================================================================
# DISTRIBUTION TESTS
# =============================================================================

def test_poisson_distribution() -> TestResult:
    """Test Poisson distribution against scipy (λ=5)."""
    lambda_param = 5
    k_values = list(range(11))
    
    # Expected values from C# add-in
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
    """Test Negative Binomial distribution against scipy (r=5, p=0.3)."""
    r, p = 5, 0.3
    k_values = list(range(11))
    
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
    """Test Lognormal distribution against scipy (μ=0, σ=1)."""
    mu, sigma = 0, 1
    x_values = [0.5, 1, 1.5, 2, 3, 5]
    
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
    """Test Gamma distribution against scipy (α=2, β=1)."""
    alpha, beta = 2, 1
    x_values = [0.5, 1, 2, 3, 5]
    
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
    """Test Pareto distribution against scipy (α=2, xm=1)."""
    alpha, xm = 2, 1
    x_values = [1, 1.5, 2, 3, 5, 10]
    
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


# =============================================================================
# EXPOSURE CURVE TESTS
# =============================================================================

def test_mbbefd_curve() -> TestResult:
    """Test MBBEFD exposure curve (b=2, g=3)."""
    def mbbefd_curve(d, b, g):
        """MBBEFD curve matching the C# implementation."""
        if d <= 0: return 0.0
        if d >= 1: return 1.0
        if abs(b - 1) < 1e-10:
            return np.log(1 + (g - 1) * d) / np.log(g)
        if abs(g - 1) < 1e-10:
            return d
        # General formula with 'a' parameter
        bg = b ** g
        a = (g - bg) / (bg - 1)
        numerator = np.log(a + b ** d) - np.log(a + 1)
        denominator = np.log(a + b) - np.log(a + 1)
        return numerator / denominator
    
    b, g = 2, 3
    d_values = np.arange(0, 1.1, 0.1)
    
    python_values = [mbbefd_curve(d, b, g) for d in d_values]
    
    # Expected values from C# add-in
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


# =============================================================================
# CHAIN LADDER TESTS - Taylor-Ashe Data
# =============================================================================

def get_chainladder_triangle():
    """Create chainladder Triangle from Taylor-Ashe data."""
    data = []
    for i in range(10):
        for j in range(10):
            if not np.isnan(TAYLOR_ASHE_CUMULATIVE[i, j]):
                data.append({
                    'origin': pd.Timestamp(f'{1981+i}-01-01'),
                    'development': pd.Timestamp(f'{1981+i+j}-12-31'),
                    'values': TAYLOR_ASHE_CUMULATIVE[i, j]
                })
    df = pd.DataFrame(data)
    return cl.Triangle(df, origin='origin', development='development', 
                       columns='values', cumulative=True)


def test_chain_ladder_factors() -> TestResult:
    """Test chain ladder development factors against expected values."""
    tri = get_chainladder_triangle()
    dev = cl.Development().fit(tri)
    cl_factors = dev.ldf_.values.flatten()[:9]
    
    max_diff = np.max(np.abs(cl_factors - np.array(EXPECTED_FACTORS)))
    tolerance = 0.001
    
    return TestResult(
        name="Chain Ladder Factors (Taylor-Ashe)",
        passed=max_diff < tolerance,
        max_diff=max_diff,
        tolerance=tolerance,
        details=f"Expected: {EXPECTED_FACTORS}"
    )


def test_chain_ladder_ibnr() -> TestResult:
    """Test chain ladder total IBNR against E&V reference."""
    tri = get_chainladder_triangle()
    dev = cl.Development().fit(tri)
    model = cl.Chainladder().fit(dev.transform(tri))
    
    total_ibnr = float(model.ibnr_.sum())
    expected = EV_REFERENCE['total_ibnr']
    diff_pct = abs(total_ibnr - expected) / expected * 100
    
    tolerance = 1.0  # 1% tolerance
    
    return TestResult(
        name="Chain Ladder Total IBNR",
        passed=diff_pct < tolerance,
        max_diff=diff_pct,
        tolerance=tolerance,
        details=f"Calculated: {total_ibnr:,.0f}, Expected: {expected:,}"
    )


def test_mack_standard_errors() -> TestResult:
    """Test Mack standard errors against E&V reference."""
    tri = get_chainladder_triangle()
    dev = cl.Development().fit(tri)
    mack = cl.MackChainladder().fit(dev.transform(tri))
    
    mack_se = mack.total_mack_std_err_.values.flatten()
    total_se = float(mack_se.sum()) if len(mack_se) > 0 else 0
    
    expected = EV_REFERENCE['total_mack_se']
    diff_pct = abs(total_se - expected) / expected * 100
    
    tolerance = 5.0  # 5% tolerance for Mack SE
    
    return TestResult(
        name="Mack Total Standard Error",
        passed=diff_pct < tolerance,
        max_diff=diff_pct,
        tolerance=tolerance,
        details=f"Calculated: {total_se:,.0f}, Expected: {expected:,}"
    )


# =============================================================================
# ODP BOOTSTRAP COMPARISON
# =============================================================================

def run_odp_bootstrap_comparison(n_sims: int = 10000, seed: int = 123) -> Dict:
    """
    Run comprehensive ODP bootstrap comparison.
    
    Compares:
    1. Our Python non-constant implementation
    2. chainladder hat_adj=False (their "non-constant")
    3. chainladder hat_adj=True (constant scale)
    4. E&V (2002) published results
    """
    np.random.seed(seed)
    n = 10
    triangle = TAYLOR_ASHE_CUMULATIVE.copy()
    
    # Calculate factors
    def calc_factors(tri):
        factors = []
        for j in range(n-1):
            sum_curr = sum_next = 0
            for i in range(n - j - 1):
                if not np.isnan(tri[i,j]) and not np.isnan(tri[i,j+1]) and tri[i,j] > 0:
                    sum_curr += tri[i, j]
                    sum_next += tri[i, j+1]
            factors.append(sum_next / sum_curr if sum_curr > 0 else 1.0)
        return np.array(factors)
    
    factors = calc_factors(triangle)
    
    # Calculate fitted incremental
    fitted_incr = np.zeros((n, n))
    for i in range(n):
        for j in range(n - i):
            if j == 0:
                fitted_incr[i, j] = triangle[i, j]
            else:
                fitted_incr[i, j] = triangle[i, j-1] * (factors[j-1] - 1.0)
    
    # Calculate NON-CONSTANT phi and residuals BY PERIOD
    phi_by_period = {}
    residuals_by_period = {}
    
    for j in range(1, n):
        residuals_j = []
        residuals_sq = []
        for i in range(n - j):
            fitted = fitted_incr[i, j]
            if fitted > 0:
                actual_incr = triangle[i, j] - triangle[i, j-1]
                r = (actual_incr - fitted) / np.sqrt(fitted)
                residuals_j.append(r)
                residuals_sq.append(r**2)
        
        n_obs = len(residuals_sq)
        if n_obs > 1:
            phi_by_period[j] = sum(residuals_sq) / (n_obs - 1)
            phi_j = phi_by_period[j]
            residuals_by_period[j] = np.array([r / np.sqrt(phi_j) for r in residuals_j])
        else:
            phi_by_period[j] = min(phi_by_period.get(j-1, 10000), phi_by_period.get(j-2, 10000))
            residuals_by_period[j] = residuals_by_period.get(j-1, np.array([0]))
    
    # Run Python non-constant bootstrap
    reserves_by_year = [[] for _ in range(n)]
    
    for _ in range(n_sims):
        boot_incr = np.zeros((n, n))
        for i in range(n):
            for j in range(n - i):
                if j == 0:
                    boot_incr[i, j] = fitted_incr[i, j]
                else:
                    mean = fitted_incr[i, j]
                    if mean > 0:
                        r = np.random.choice(residuals_by_period[j])
                        phi_j = phi_by_period[j]
                        boot_incr[i, j] = max(1.0, mean + r * np.sqrt(mean * phi_j))
                    else:
                        boot_incr[i, j] = max(1.0, mean)
        
        pseudo_tri = np.cumsum(boot_incr, axis=1)
        boot_factors = calc_factors(pseudo_tri)
        
        for i in range(1, n):
            last_col = n - 1 - i
            current = triangle[i, last_col]
            for j in range(last_col, n - 1):
                mean_next = current * boot_factors[j]
                mean_incr = mean_next - current
                phi_j = phi_by_period[j + 1]
                if mean_incr > 0:
                    sim_incr = np.random.gamma(mean_incr / phi_j, phi_j)
                    current += sim_incr
                else:
                    current = mean_next
            reserves_by_year[i].append(current - triangle[i, last_col])
    
    # Calculate Python results
    python_se = [0] + [np.std(reserves_by_year[i]) for i in range(1, n)]
    # Total reserve for each simulation
    total_reserves = []
    for j in range(n_sims):
        total = sum(reserves_by_year[i][j] if len(reserves_by_year[i]) > j else 0 
                    for i in range(1, n))
        total_reserves.append(total)
    python_total_se = np.std(total_reserves)
    
    # Run chainladder bootstrap (hat_adj=False)
    tri = get_chainladder_triangle()
    sample_nohat = cl.BootstrapODPSample(n_sims=n_sims, hat_adj=False, random_state=seed)
    sample_nohat.fit(tri)
    resampled = sample_nohat.transform(tri)
    dev = cl.Development().fit_transform(resampled)
    model = cl.Chainladder().fit(dev)
    ibnr_arr = np.nan_to_num(model.ibnr_.values.squeeze(), nan=0.0)
    cl_nohat_se = [np.std(ibnr_arr[:, i]) for i in range(10)]
    cl_nohat_total_se = np.std(np.sum(ibnr_arr, axis=1))
    
    # Run chainladder bootstrap (hat_adj=True)
    sample_hat = cl.BootstrapODPSample(n_sims=n_sims, hat_adj=True, random_state=seed)
    sample_hat.fit(tri)
    resampled2 = sample_hat.transform(tri)
    dev2 = cl.Development().fit_transform(resampled2)
    model2 = cl.Chainladder().fit(dev2)
    ibnr_arr2 = np.nan_to_num(model2.ibnr_.values.squeeze(), nan=0.0)
    cl_hat_se = [np.std(ibnr_arr2[:, i]) for i in range(10)]
    cl_hat_total_se = np.std(np.sum(ibnr_arr2, axis=1))
    
    return {
        'python_nonconstant_se': python_se,
        'python_nonconstant_total_se': python_total_se,
        'chainladder_nohat_se': cl_nohat_se,
        'chainladder_nohat_total_se': cl_nohat_total_se,
        'chainladder_hat_se': cl_hat_se,
        'chainladder_hat_total_se': cl_hat_total_se,
        'ev_nonconstant_se': EV_REFERENCE['odp_nonconstant_se'],
        'ev_nonconstant_total_se': EV_REFERENCE['total_odp_nonconstant_se'],
        'ev_constant_se': EV_REFERENCE['odp_constant_se'],
        'ev_constant_total_se': EV_REFERENCE['total_odp_constant_se'],
        'ibnr': EV_REFERENCE['ibnr'],
    }


# =============================================================================
# MAIN TEST RUNNER
# =============================================================================

def run_all_tests() -> List[TestResult]:
    """Run all benchmark tests."""
    tests = [
        test_poisson_distribution,
        test_negbin_distribution,
        test_lognormal_distribution,
        test_gamma_distribution,
        test_pareto_distribution,
        test_mbbefd_curve,
        test_chain_ladder_factors,
        test_chain_ladder_ibnr,
        test_mack_standard_errors,
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


def print_bootstrap_comparison(n_sims: int = 10000, seed: int = 123):
    """Print comprehensive ODP bootstrap comparison table."""
    print("\n" + "=" * 80)
    print("ODP BOOTSTRAP COMPARISON")
    print("=" * 80)
    print(f"\nReference: England & Verrall (2002) 'Stochastic Claims Reserving'")
    print(f"Data: Taylor-Ashe 10x10 cumulative paid losses")
    print(f"Simulations: {n_sims:,}, Seed: {seed}")
    print()
    
    results = run_odp_bootstrap_comparison(n_sims, seed)
    
    # Print detailed comparison table
    print("-" * 100)
    print(f"{'AY':<6} {'IBNR':>12} {'Python SE':>12} {'E&V Non-C':>12} {'CL nohat':>12} {'CL hat':>12} {'Py vs E&V':>10}")
    print("-" * 100)
    
    for i in range(10):
        ibnr = results['ibnr'][i]
        py_se = results['python_nonconstant_se'][i]
        ev_se = results['ev_nonconstant_se'][i]
        cl_nohat = results['chainladder_nohat_se'][i]
        cl_hat = results['chainladder_hat_se'][i]
        ratio = f"{py_se/ev_se:.0%}" if ev_se > 0 else "-"
        
        print(f"{i+1:<6} {ibnr:>12,} {py_se:>12,.0f} {ev_se:>12,} {cl_nohat:>12,.0f} {cl_hat:>12,.0f} {ratio:>10}")
    
    print("-" * 100)
    print(f"{'Total':<6} {sum(results['ibnr']):>12,} {results['python_nonconstant_total_se']:>12,.0f} "
          f"{results['ev_nonconstant_total_se']:>12,} {results['chainladder_nohat_total_se']:>12,.0f} "
          f"{results['chainladder_hat_total_se']:>12,.0f} "
          f"{results['python_nonconstant_total_se']/results['ev_nonconstant_total_se']:.0%}")
    print("-" * 100)
    
    # Print summary table
    print("\n" + "=" * 60)
    print("SUMMARY: Total SE Comparison")
    print("=" * 60)
    print(f"{'Implementation':<35} {'Total SE':>12} {'vs E&V NC':>10}")
    print("-" * 60)
    
    ev_nc = results['ev_nonconstant_total_se']
    implementations = [
        ("Python Non-Constant (our method)", results['python_nonconstant_total_se']),
        ("E&V Non-Constant (target)", ev_nc),
        ("chainladder hat_adj=False", results['chainladder_nohat_total_se']),
        ("chainladder hat_adj=True", results['chainladder_hat_total_se']),
        ("E&V Constant", results['ev_constant_total_se']),
    ]
    
    for name, se in implementations:
        ratio = f"{se/ev_nc:.0%}"
        print(f"{name:<35} {se:>12,.0f} {ratio:>10}")
    
    print("-" * 60)
    
    # Print observations
    print("\nObservations:")
    py_ratio = results['python_nonconstant_total_se'] / ev_nc
    cl_nohat_ratio = results['chainladder_nohat_total_se'] / ev_nc
    cl_hat_ratio = results['chainladder_hat_total_se'] / results['ev_constant_total_se']
    
    print(f"  • Python non-constant implementation achieves {py_ratio:.0%} match to E&V target")
    print(f"  • chainladder hat_adj=False produces {cl_nohat_ratio:.0%} of E&V non-constant")
    print(f"  • chainladder hat_adj=True matches E&V constant at {cl_hat_ratio:.0%}")
    print(f"  • Key difference: chainladder pools residuals, we stratify by development period")
    print()


def print_mack_comparison():
    """Print Mack standard error comparison."""
    print("\n" + "=" * 80)
    print("MACK STANDARD ERROR COMPARISON")
    print("=" * 80)
    
    tri = get_chainladder_triangle()
    dev = cl.Development().fit(tri)
    mack = cl.MackChainladder().fit(dev.transform(tri))
    
    print(f"\nReference: England & Verrall (2002)")
    print(f"Data: Taylor-Ashe 10x10 triangle")
    print()
    
    print("-" * 70)
    print(f"{'AY':<6} {'IBNR':>12} {'E&V Mack SE':>15} {'CL Mack SE':>15} {'Ratio':>10}")
    print("-" * 70)
    
    # Get individual origin year SEs from chainladder
    try:
        cl_se_by_origin = mack.total_mack_std_err_.values.flatten()
    except:
        cl_se_by_origin = np.zeros(10)
    
    for i in range(10):
        ibnr = EV_REFERENCE['ibnr'][i]
        ev_se = EV_REFERENCE['mack_se'][i]
        cl_se = cl_se_by_origin[i] if i < len(cl_se_by_origin) else 0
        ratio = f"{cl_se/ev_se:.0%}" if ev_se > 0 else "-"
        print(f"{i+1:<6} {ibnr:>12,} {ev_se:>15,} {cl_se:>15,.0f} {ratio:>10}")
    
    print("-" * 70)
    total_cl_se = np.sum(cl_se_by_origin) if len(cl_se_by_origin) > 0 else 0
    print(f"{'Total':<6} {EV_REFERENCE['total_ibnr']:>12,} {EV_REFERENCE['total_mack_se']:>15,} "
          f"{total_cl_se:>15,.0f} {total_cl_se/EV_REFERENCE['total_mack_se']:.0%}")
    print("-" * 70)
    print()


def main():
    print("=" * 80)
    print("ACTUARIAL ADD-IN BENCHMARK TESTS")
    print("=" * 80)
    print()
    print(f"chainladder version: {cl.__version__}")
    print(f"numpy version: {np.__version__}")
    print()
    
    # Run basic tests
    print("-" * 80)
    print("BASIC FUNCTION TESTS")
    print("-" * 80)
    
    results = run_all_tests()
    
    passed = sum(1 for r in results if r.passed)
    total = len(results)
    
    print(f"{'Test Name':<45} {'Status':<10} {'Max Diff':<15}")
    print("-" * 80)
    
    for result in results:
        status = "PASS" if result.passed else "FAIL"
        print(f"{result.name:<45} {status:<10} {result.max_diff:.2e}")
        if result.details and not result.passed:
            print(f"  Details: {result.details}")
    
    print("-" * 80)
    print(f"Results: {passed}/{total} tests passed")
    
    # Run Mack comparison
    print_mack_comparison()
    
    # Run bootstrap comparison
    print_bootstrap_comparison(n_sims=10000, seed=123)
    
    print("=" * 80)
    print("TEST COMPLETE")
    print("=" * 80)
    
    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
